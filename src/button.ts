function waitForElement(selector: string, timeout = 10000): Promise<HTMLElement> {
  return new Promise((resolve, reject) => {
    const found = document.querySelector<HTMLElement>(selector);
    if (found) return resolve(found);

    const observer = new MutationObserver(() => {
      const el = document.querySelector<HTMLElement>(selector);
      if (el) {
        observer.disconnect();
        resolve(el);
      }
    });

    observer.observe(document.body, { childList: true, subtree: true });

    setTimeout(() => {
      observer.disconnect();
      reject(new Error(`Element not found: ${selector}`));
    }, timeout);
  });
}

export async function injectExportButton(callback: Function) {
  const baseSelector = 'button#all-history';
  try {
    const baseBtn = await waitForElement(baseSelector);

    // Clone deeply
    const exportBtn = baseBtn.cloneNode(true) as HTMLButtonElement;

    // Update ID and aria/label text
    exportBtn.id = 'export-conversations';
    exportBtn.setAttribute('aria-label', 'Export conversations');
    exportBtn.value = 'export-conversations';

    const span = exportBtn.querySelector('span');
    if (span) {
      span.textContent = 'Export conversations';
      span.setAttribute('aria-label', 'Export conversations');
    }

    // Hook your click logic
    exportBtn.addEventListener('click', () => {
      callback();
      // TODO: Trigger export logic here
    });

    // Insert after the original button
    baseBtn.parentElement?.insertBefore(exportBtn, baseBtn.nextSibling);
  } catch (err) {
    console.error('[Copilot Exporter] Failed to insert export button:', err);
  }
}