type Properties = {
    sheetName: string,
    emailSubject: string,
    email: string,
    fields: string[],
    captcha: CaptchaInfo
};

type CaptchaInfo = {
    type: 'cloudflare_turnstile' | 'recaptcha_v2' | 'recaptcha_v3' | 'none',
    data: {
        secretKey: string
    }
};

const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

const createJsonResponse = (content: object): GoogleAppsScript.Content.TextOutput =>
    ContentService.createTextOutput(JSON.stringify(content)).setMimeType(ContentService.MimeType.JSON);

const capitalizeFirstLetter = (str: string): string => str[0].toUpperCase() + str.slice(1);

const formatEmail = (data: { [key: string]: any }) =>
    Object.entries(data)
        .map(([heading, resp]) => `${capitalizeFirstLetter(heading)}:\n${resp}`)
        .join('\n\n');

function setFields(sheet: GoogleAppsScript.Spreadsheet.Sheet, fields: string[]) {
    const firstRow = sheet.getRange(1, fields.length).getValues()[0];

    // Checking if columns in sheet and fields are matching
    if (firstRow.toString() !== '') {
        if (
            firstRow[0].toLowerCase() === fields[0].toLowerCase() &&
            firstRow[firstRow.length - 2].toString().toLowerCase() ===
                fields[fields.length - 1].toString().toLowerCase()
        ) {
            return;
        }

        if (firstRow.length > fields.length + 1) {
            sheet.getRange(1, 1, 1, 30).clearContent(); // Clearing up to 30 columns
        }
    }

    sheet
        .getRange(1, 1, 1, fields.length + 1)
        .setValues([
            [...fields.map(field => capitalizeFirstLetter(field)), 'Date']
        ]);
}

function action(
    req: GoogleAppsScript.Events.DoPost, {
        sheetName = '',
        emailSubject = 'New Form Submission',
        email = '',
        fields = ['name', 'email', 'message'],
        captcha = { type: 'none', data: { secretKey: '' } }
    }: Properties
): GoogleAppsScript.Content.TextOutput {
    let { postData: { contents } } = req;
    let jsonData: { [key: string]: any };

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

    const logSheet = sheetName !== '' ?
        activeSpreadsheet.getSheetByName(sheetName) :
        activeSpreadsheet.getActiveSheet();

    if (logSheet === null) {
        return createJsonResponse({
            status: 'error',
            message: 'No sheet found.'
        });
    }

    setFields(logSheet, fields);

    const now = new Date();
    const date =
        now.toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
        }) +
        ' ' +
        now.toLocaleTimeString('en-US');

    logSheet
        .insertRowAfter(1)
        .getRange(2, 1, 1, fields.length + 1)
        .setValues([
            [...fields.map(field => jsonData[field]), date]
        ]);

    if (email !== '') {
        MailApp.sendEmail({
            to: email,
            subject: emailSubject,
            body: formatEmail(fields.reduce((a, c) => ({ ...a, [c]: jsonData[c] }), {})),
            replyTo: jsonData.email,
        });
    }

    return createJsonResponse({
        status: 'OK',
        message: 'Submission logged successfully',
    });
}
