// WebCommando v1.1 - ajoute un raccourci clavier et supprime les alertes bloquantes.

function copyHtml(tabId) {
  chrome.scripting.executeScript({
    target: { tabId },
    func: () => {
      try {
        const html = document.documentElement.outerHTML;
        const signature = `<!-- ScrapCommando SIGNED url=${location.href} ts=${new Date().toISOString()} -->\n`;
        const payload = signature + html;

        const copyViaClipboardApi = async () => {
          try {
            await navigator.clipboard.writeText(payload);
          } catch (e) {
            // Fallback execCommand si l'API Clipboard échoue
            const ta = document.createElement('textarea');
            ta.value = payload;
            ta.setAttribute('readonly','');
            ta.style.position='fixed';
            ta.style.top='-1000px';
            ta.style.left='-1000px';
            document.body.appendChild(ta);
            ta.focus();
            ta.select();
            try { document.execCommand('copy'); } finally { document.body.removeChild(ta); }
          }
        };

        copyViaClipboardApi();
      } catch (e) {
        // Pas d'alert bloquante ici; au pire on logge dans la console.
        console.error('WebCommando copy error:', e);
      }
    }
  });
}

// Clic sur l'icône
chrome.action.onClicked.addListener((tab) => { if (tab && tab.id) copyHtml(tab.id); });

// Raccourci clavier (Ctrl+Shift+U par défaut)
chrome.commands.onCommand.addListener((command) => {
  if (command === 'copy-html') {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      if (tabs && tabs[0]) copyHtml(tabs[0].id);
    });
  }
});
