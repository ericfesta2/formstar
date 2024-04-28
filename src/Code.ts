type ActionRequest = {
    postData: {
        contents: string,
        type: 'recaptcha_v2' | 'recaptcha_v3'
    }
};

type CaptchaInfo = {
    type: 'cloudflare_turnstile' | 'recaptcha_v2' | 'recaptcha_v3' | 'none',
    data: {
        secretKey: string
    }
};

let sheetName = '';
let emailSubject = 'New submission using FormEasy';
let formHeading = 'Form submission - FormEasy';
let email = '';
let fields: string[] = [];
let captcha: CaptchaInfo = { type: 'none', data: { secretKey: '' } };

const createJsonResponse = (content: object): GoogleAppsScript.Content.TextOutput =>
    ContentService.createTextOutput(JSON.stringify(content)).setMimeType(ContentService.MimeType.JSON);

function setSheet(name: string) {
    sheetName = name;
}

function setEmail(id: string) {
    email = id;
}

function setSubject(subject: string) {
    emailSubject = subject;
}

function setFormHeading(heading: string) {
    formHeading = heading;
}

function setFields(...fieldsArr: string[]) {
    fields = [...fieldsArr];

    const length = fields.length;
    const sheetNameEmpty = sheetName === '';

    const sheet = sheetNameEmpty ?
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet() :
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    if (sheet === null) {
        return createJsonResponse({
            status: 'error',
            message: sheetNameEmpty ? 'No spreadsheet found.' : `Spreadsheet ${sheetName} not found.`
        });
    }

    const firstRow = sheet.getDataRange().getValues()[0];

    // Checking if columns in sheet and fields are matching
    if (firstRow.toString() !== '') {
        if (
            firstRow[0].toLowerCase() === fields[0].toLowerCase() &&
            firstRow[firstRow.length - 2].toString().toLowerCase() === fields[length - 1].toString().toLowerCase()
        ) {
            return;
        }

        if (firstRow.length > length + 1) {
            sheet.getRange(1, 1, 1, 30).clearContent(); // Clearing up to 30 columns
        }
    }

    const formatFirstLetter = (str: string): string => str[0].toUpperCase() + str.slice(1);

    for (let idx = 0; idx < length; idx++) {
        sheet.getRange(1, idx + 1).setValue(formatFirstLetter(fields[idx]));
        if (idx === length - 1) {
            sheet.getRange(1, idx + 2).setValue('Date');
        }
    }
}

function setRecaptcha(secretKey: string) {
    captcha = { type: 'recaptcha_v2', data: { secretKey } };
}

function action(req: ActionRequest): object {
    let { postData: { contents } } = req;
    let jsonData;

    try {
        jsonData = JSON.parse(contents);
    } catch (err) {
        return createJsonResponse({
            status: 'error',
            message: 'Invalid JSON format',
        });
    }

    switch (captcha.type) {
        case 'recaptcha_v2': {
            const siteKey = jsonData['gCaptchaResponse'];

            if (!siteKey) {
                return createJsonResponse({
                    status: 'error',
                    message: "reCAPTCHA verification under key 'gCaptchaResponse' is required."
                });
            }

            const captchaResponse = UrlFetchApp.fetch('https://www.google.com/recaptcha/api/siteverify', {
                method: 'post',
                payload: {
                    response: siteKey,
                    secret: captcha.data.secretKey,
                }
            });

            const captchaJson = JSON.parse(captchaResponse.getContentText());

            if (!captchaJson.success) {
                return createJsonResponse({
                    status: 'error',
                    message: 'Please tick the box to verify you are not a robot.'
                });
            }

            break;
        }
    }

    let logSheet;

    const allSheets = SpreadsheetApp.getActiveSpreadsheet()
        .getSheets()
        .map(s => s.getName());

    const sheetNameEmpty = sheetName === '';

    if (!sheetNameEmpty) {
        const sheetExists = allSheets.includes(sheetName);

        if (!sheetExists) {
            return createJsonResponse({
                status: 'error',
                message: 'Invalid sheet name',
            });
        }

        logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    } else {
        logSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    }

    if (logSheet === null) {
        return createJsonResponse({
            status: 'error',
            message: sheetNameEmpty ? 'No spreadsheet found.' : `Spreadsheet ${sheetName} not found.`
        });
    }

    if (fields.length < 1) {
        setFields('name', 'email', 'message');
    }

    const length = fields.length;

    const now = new Date();
    const date =
        now.toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
        }) +
        ' ' +
        now.toLocaleTimeString('en-US');

    // Inserting a row after the first row
    logSheet.insertRowAfter(1);

    // Filling the latest data in the second row
    for (let idx = 0; idx < length; idx++) {
        logSheet.getRange(2, idx + 1).setValue(jsonData[fields[idx]]);

        if (idx === length - 1) {
            logSheet.getRange(2, idx + 2).setValue(date);
        }
    }

    const emailData = fields.reduce((a, c) => ({ ...a, [c]: jsonData[c] }), {});
    const htmlBody = HtmlService.createTemplateFromFile('EmailTemplate');
    htmlBody.data = emailData;
    htmlBody.formHeading = formHeading;

    const emailBody = htmlBody.evaluate().getContent();

    if (email) {
        MailApp.sendEmail({
            to: email,
            subject: emailSubject,
            htmlBody: emailBody,
            replyTo: jsonData.email,
        });
    }

    return createJsonResponse({
        status: 'OK',
        message: 'Data logged successfully',
    });
}
