import { fetchCopilotChats } from './api';
import { injectExportButton } from './button';
import { getAccessToken, getMsalIds } from './token';

const APP_TAG = "[Copilot Exporter]";
console.log(`${APP_TAG} Userscript initalized.`)

const inject = () => injectExportButton(
  () => {
    console.log(`${APP_TAG} Export button clicked.`);
    console.log(`${APP_TAG} Getting MSAL ids...`);
    const msalIds = getMsalIds()
    console.log(`${APP_TAG} Getting access token...`);
    getAccessToken(msalIds).then(token => fetchCopilotChats(token, msalIds.localAccountId, msalIds.tenantId)).then(results => console.log(results));
  }
);


if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', inject);
  } else {
    inject();
}