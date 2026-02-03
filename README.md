# M365 Copilot Exporter
An exporter for conversations in the Microsoft 365 Copilot Chat integrated into the [Microsoft 365 dashboard](https://m365.cloud.microsoft/chat/).

Exporter design and functionality based on [@Pionxzh](https://github.com/pionxzh)'s [ChatGPT Exporter](https://github.com/pionxzh/chatgpt-exporter). This userscript is not 1:1 in functionality, though!

<div style="display:flex">
  <img width="45%" src="https://github.com/user-attachments/assets/267c86c5-ed41-4989-924f-ae1f1aafb8c7" alt="Copilot Exporter's modal">
  <img width="45%" src="https://github.com/user-attachments/assets/4033fe5a-9ef7-40cc-bf2f-00f1a25f6886" alt="Copilot Exporter in action exporting 400+ conversations">
</div>

> [!CAUTION]
> This tool is maintained on a "when I feel like it" basis. Use at your own risk!

## Features
- View list of Copilot conversations.
- Bulk export Copilot conversations as its raw, internal JSON format.
- Bulk deletion of Copilot conversations.

Compared to ChatGPT Exporter, it does **not** support:
- Exporting in formats OTHER than the official JSON. If you want to do this, I suppose you could convert Copilot's JSON to ChatGPT's, then shove it into ChatGPT Exporter.

## Limitations
The exporter can only show, at maximum, the latest 500 conversations. This is a limit imposed by the API endpoint used to get the list of chats. You'll need to delete some chats if you want to access anything beyond the latest 500 conversations.

## How to install
Install a userscript manager extension into your browser, such as Greasemonkey or Tampermonkey. I personally recommend [Greasemonkey](https://addons.mozilla.org/en-US/firefox/addon/greasemonkey/) (and I have only used this script on it, so your mileage may vary with others)!

### Greasyfork
Once you have a userscript manager, you can download this script on Greasyfork using the link below and clicking "Install this script".
https://greasyfork.org/en/scripts/543763-m365-copilot-exporter

### Manual
Copy or import [`dist/copilot-exporter.user.js`](https://github.com/ganyuke/copilot-exporter/blob/master/dist/copilot-exporter.user.js) into a new script in your desired userscript manager.

## How to use
1. Navigate to your [M365 dashboard](https://m365.cloud.microsoft/chat/).
2. Open the export menu by either:
   - Clicking on the floating export button in the bottom right.
   - Opening the Greasemonkey command menu and selecting "Open export menu".
3. Select the conversations you want to export.
   - If the conversation(s) you want to export are further down, alter the maximum number of conversations shown and refetch.
4. Click export once you've selected the conversations you want to export.

## How to build
For those seeking to maintain this:
1. Clone this repository: `git clone https://github.com/ganyuke/copilot-exporter`.
2. Open the directory: `cd copilot-exporter`.
1. Get [`pnpm`](https://pnpm.io/installation).
2. Run `pnpm build`. The newly-built userscript should be in `dist/`.
