import { showExportModal } from './modal';

export const APP_TAG = "[Copilot Exporter]";
console.log(`${APP_TAG} Userscript initalized.`)

// I can't be bothered to keep updating the original Fluent UI button every time Microsoft decides that
// the Microsoft 365 Dashboard is getting old.

// "Export SVG Vector" by Leonid Tsvetkov, CC Attribution License,
// obtained via https://www.svgrepo.com/svg/458671/export
const EXPORT_SVG = `<svg width="100%" height="100%" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
<path d="M12 5L11.2929 4.29289L12 3.58579L12.7071 4.29289L12 5ZM13 14C13 14.5523 12.5523 15 12 15C11.4477 15 11 14.5523 11 14L13 14ZM6.29289 9.29289L11.2929 4.29289L12.7071 5.70711L7.70711 10.7071L6.29289 9.29289ZM12.7071 4.29289L17.7071 9.29289L16.2929 10.7071L11.2929 5.70711L12.7071 4.29289ZM13 5L13 14L11 14L11 5L13 5Z" fill="#33363F"/>
<path d="M5 16L5 17C5 18.1046 5.89543 19 7 19L17 19C18.1046 19 19 18.1046 19 17V16" stroke="#33363F" stroke-width="2"/>
</svg>`

const BUTTON_ID = "export-menu-button";

// create own floating button perpentually taped to the bottom right
const inject = () => {
    if (document.getElementById(BUTTON_ID)) return;

    const btn = document.createElement("button");
    const svgEl = new DOMParser().parseFromString(EXPORT_SVG, "image/svg+xml").documentElement; // parse and load export svg string
    const svg = document.importNode(svgEl, true); // import doc element into current document

    btn.id = BUTTON_ID;

	btn.style.width = "3em";
	btn.style.height = "3em";
    btn.style.bottom = "16px";
    btn.style.right = "16px";
    btn.style.cursor = "pointer";
    btn.style.position = "fixed";
    btn.append(svg);

    btn.addEventListener("click", showExportModal);
    document.body.appendChild(btn);
}

GM.registerMenuCommand('Open export menu', showExportModal);

if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", inject);
} else {
    inject();
}