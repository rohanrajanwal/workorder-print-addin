// Runs at document_start (before any page scripts).
// Injects both scripts into the PAGE context so they have access to
// window.geotab, XMLHttpRequest.prototype, etc.
function injectScript(filename, onload) {
  const s = document.createElement('script');
  s.src = chrome.runtime.getURL(filename);
  if (onload) s.onload = onload;
  (document.head || document.documentElement).appendChild(s);
}

// printWorkOrder.js must load first — it defines window.geotab.customButtons.printWorkOrder.
// injector.js loads second — it injects the Print button and calls the handler on click.
injectScript('printWorkOrder.js', () => injectScript('injector.js'));
