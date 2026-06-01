(function () {
  'use strict';

  var PRINTER_SVG = [
    '<svg style="width:14px;height:14px;flex-shrink:0" viewBox="0 0 24 24"',
    ' fill="none" stroke="currentColor" stroke-width="2"',
    ' stroke-linecap="round" stroke-linejoin="round">',
    '<polyline points="6 9 6 2 18 2 18 9"/>',
    '<path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/>',
    '<rect x="6" y="14" width="12" height="8"/>',
    '</svg>'
  ].join('');

  function applyStyle() {
    var btn = document.querySelector('button[aria-label="Print"]');
    if (!btn || btn.dataset.woStyled) return false;

    // Match the Save button style: blue, icon + label
    btn.classList.remove('zen-button--tertiary-black');
    btn.classList.add('zen-button--primary', 'zen-caption', 'zen-text-icon-button');

    // Add printer icon and wrap text to match the Delete/Save button structure
    btn.innerHTML = PRINTER_SVG + '<span class="zen-caption__content">Print</span>';

    // Mark as styled so we don't re-apply on multiple calls
    btn.dataset.woStyled = '1';
    return true;
  }

  // The button may not be in the DOM yet when changePageState fires —
  // poll briefly until it appears.
  if (!applyStyle()) {
    var tries = 0;
    var timer = setInterval(function () {
      if (applyStyle() || ++tries >= 30) {
        clearInterval(timer);
      }
    }, 100);
  }
})();
