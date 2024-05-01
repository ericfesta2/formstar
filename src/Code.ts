type Properties = {
    sheetName: string,
    emailSubject: string,
    email: string,
    fields: string[],
    captcha: CaptchaInfo
};

type CaptchaInfo = {
    type: 'cloudflare_turnstile' | 'recaptcha' | 'none',
    data: {
        secretKey: string,
        reCaptchaV3MinScore?: number
    }
};

const CaptchaDetails: { [key in Exclude<CaptchaInfo['type'], 'none'> ] : { tokenKey: string, endpoint: string } } = {
    'cloudflare_turnstile': {
        tokenKey: 'cf-turnstile-response',
        endpoint: 'https://challenges.cloudflare.com/turnstile/v0/siteverify'
    },
    'recaptcha': {
        tokenKey: 'g-recaptcha-response',
        endpoint: 'https://www.google.com/recaptcha/api/siteverify'
    }
};

const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

const createJsonResponse = (content: object): GoogleAppsScript.Content.TextOutput =>
    ContentService.createTextOutput(JSON.stringify(content)).setMimeType(ContentService.MimeType.JSON);

const capitalizeFirstLetter = (str: string): string => str[0].toUpperCase() + str.slice(1);

const formatEmail = (data: { [key: string]: any }): string =>
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
    const { postData: { contents } } = req;
    let jsonData: { [key: string]: any };

    try {
        jsonData = JSON.parse(contents);
    } catch (err) {
        return createJsonResponse({
            status: 'error',
            message: 'Invalid JSON format'
        });
    }

    if (captcha.type !== 'none') {
        const captchaDetail = CaptchaDetails[captcha.type];
        const token = jsonData[captchaDetail.tokenKey];

        if (!token) {
            return createJsonResponse({
                status: 'error',
                message: 'No CAPTCHA token received in request'
            });
        }

        const captchaResponse = JSON.parse(
            UrlFetchApp.fetch(captchaDetail.endpoint, {
                method: 'post',
                payload: {
                    response: token,
                    secret: captcha.data.secretKey
                }
            }).getContentText()
        );

        if (
            !captchaResponse.success || (
                typeof captchaResponse['score'] === 'number' && typeof captcha.data.reCaptchaV3MinScore === 'number' &&
                captchaResponse['score'] < captcha.data.reCaptchaV3MinScore
            )
        ) {
            return createJsonResponse({
                status: 'error',
                message: 'CAPTCHA challenge failed.'
            });
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
