import { injectExportButton } from './button';
import { getAccessToken } from './token';

const APP_TAG = "[Copilot Exporter]";
console.log(`${APP_TAG} Userscript initalized.`)

const inject = () => injectExportButton(
  () => {
    console.log(`${APP_TAG} Export button clicked.`);
    console.log(`${APP_TAG} Getting access token...`);
    getAccessToken().then(token => console.log(token));
  }
);


if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', inject);
  } else {
    inject();
}