# Formstar

Formstar is a serverless form backend library for [Google Apps Script](https://script.google.com) that makes handling form submissions from static websites easy. Formstar is a fork of the [FormEasy](https://github.com/Basharath/FormEasy) project with support for Cloudflare Turnstile CAPTCHAs, enhanced support for reCAPTCHA v3, TypeScript source code, and various other optimizations.

Google Apps Script ID: `1-0U5H6zqVyAnYjBekxDYvx8PoWdSHjf6L6m_GKMK4wE96-Z-9PVtyemi`

## Setup

1. Create a new Google Sheets file within Google Drive - this is where your form data gets stored
2. In the Sheets menu bar, click Extensions > Apps Script
3. In the left sidebar of the Apps Script project, click on the `+` button next to `Libraries`
4. Add the above `Google Apps Script ID` above, click on the `Look up` button, and select the latest version
   (either the latest numeric version for a stable release or `HEAD (Development mode)` to always receive
   the latest changes as soon as possible), then click on the `Add` button.

Your new Apps Script project can now use the `Formstar` JavaScript object.

## Usage

Clear the contents of the Apps Script editor, and add the below function, replacing `<SETTINGS>`
with any combination of the options specified below.

```js
const doPost = req =>
  Formstar.action(req, {
    <SETTINGS>
  });

```

### Configuration

| Setting    | Description | Default Value |
| ---------- | ----------- | ------------- |
| `sheetName`  | The name of the sheet within the target spreadsheet to which form submissions should be written. If empty, uses the default sheet within the target spreadsheet. | `''` |
| `emailSubject` | The subject line of emails sent to the address specified by the `email` setting. Has no effect if `email` is not configured or set to an empty string. | `'New Form Submission'` |
| `fields` | The header fields of the target sheet that records form submissions. If not already present in the sheet, Formstar will write them to the first row of the sheet. | `['name', 'email', 'message']` |
| `captcha` | Settings for configuring recommended spam protection. reCAPTCHA v2, reCAPTCHA v3, and Cloudflare Turnstile (types `recaptcha` for either v2 or v3, and `cloudflare_turnstile`) are supported. For reCAPTCHA v3, you may set the numeric `data.reCaptchaV3MinScore` attribute to the minimum reCAPTCHA v3 score for which submissions should be accepted. Note that you must configure a reCAPTCHA or Turnstile project that supplies site and secret keys to use this capability. | `<NONE>` |

**Example**

```js
const doPost = req =>
  Formstar.action(req, {
    emailSubject: 'New Form Submission from My Site',
    email: 'youremail@domain.com',
    fields: ['name', 'email', 'phone', 'birthday', 'message'],
    captcha: {
      type: 'cloudflare_turnstile',
      data: {
        secretKey: '<TURNSTILE_SECRET_KEY_HERE>'
      }
    }
  });
```

### Deployment
After adding the above function click the `Deploy` button at top right corner and select **New deployment** and select type to `Web app` from the gear icon.

Select the below options:

- Description (optional),
- Execute as `Me (<YOUR EMAIL>)`
- Who has access `Anyone`

Click on the `Deploy` button (authorize the script if you haven't already), and you will get a URL under `Web app`.
Copy that URL and use it as the endpoint for submitting POST requests from your frontend form(s).

Note: You don't need to repeat this process every time you make changes if you want to use the same web app URL. Select **Manage deployments** and update the version to keep the same URL.
