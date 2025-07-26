import { injectExportButton } from './button';
import { showExportModal } from './modal';

export const APP_TAG = "[Copilot Exporter]";
console.log(`${APP_TAG} Userscript initalized.`)

const inject = () => injectExportButton(
    () => {
        console.log(`${APP_TAG} Export button clicked.`);
        showExportModal();
    }
);

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', inject);
} else {
    inject();
}