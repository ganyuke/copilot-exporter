# M365 Copilot Exporter
An exporter for the Copilot Chat integrated into the [M365 dashboard](https://m365.cloud.microsoft/chat/).

Exporter design and functionality based on [@Pionxzh](https://github.com/pionxzh)'s [ChatGPT Exporter](https://github.com/pionxzh/chatgpt-exporter). This userscript is not 1:1 in functionality, though!

## Features
- View list of Copilot conversations.
- Bulk export Copilot conversations as its raw, internal JSON format.

Compared to ChatGPT Exporter, it does **not** support:
- Bulk deletion (the "Delete" button in my modal is a lie!!!).
- Exporting in formats OTHER than the official JSON.
  - If you want to do this, I suppose you could convert Copilot's JSON to ChatGPT's, then shove it into ChatGPT Exporter.

## How to use
1. Copy or import [`dist/copilot-exporter.user.js`](https://github.com/ganyuke/copilot-exporter/blob/master/dist/copilot-exporter.user.js) into your desired userscript manager. I recommend [Greasemonkey](https://addons.mozilla.org/en-US/firefox/addon/greasemonkey/)!
2. Navigate to your [M365 dashboard](https://m365.cloud.microsoft/chat/).
3. Open the sidebar (if not expanded), open the "Conversations" fold (if not expanded), then click the new "Export conversations" button. A modal should appear.
4. Select the conversations you want to export.
  - If the conversation(s) you want to export are further down, alter the maximum number of conversations shown and refetch.
5. Click export once you've selected the conversations you want to export.

## How to build
1. Clone this repository: `git clone https://github.com/ganyuke/copilot-exporter`.
2. Open the directory: `cd copilot-exporter`.
1. Get [`pnpm`](https://pnpm.io/installation).
2. Run `pnpm build`. The newly-built userscript should be in `dist/`.