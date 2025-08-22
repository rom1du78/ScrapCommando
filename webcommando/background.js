chrome.action.onClicked.addListener((tab) => {
  chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: () => {
      navigator.clipboard.writeText(document.documentElement.outerHTML)
        .then(() => {
          alert("✔️ HTML copié dans le presse-papier !");
        })
        .catch(err => {
          alert("❌ Échec de la copie : " + err);
        });
    }
  });
});
