(function () {
  'use strict';

  const BTN_ID = 'wo-print-btn';
  // Extract database from URL path (e.g. https://my.geotab.com/demo_buildtesting16/#...)
  const DB = window.location.pathname.split('/').filter(Boolean)[0] || '';
  let capturedAuth = null;

  // 1. Intercept XHR to capture the Bearer auth token from MyGeotab's own requests
  const origSetHeader = XMLHttpRequest.prototype.setRequestHeader;
  XMLHttpRequest.prototype.setRequestHeader = function (k, v) {
    if (k === 'authorization' && v && v.startsWith('Bearer')) {
      capturedAuth = decodeURIComponent(v);
    }
    return origSetHeader.apply(this, arguments);
  };

  // 2. Lightweight Geotab API wrapper using fetch + captured auth token
  function makeApi() {
    function post(method, params) {
      return fetch('https://my.geotab.com/apiv1', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'authorization': capturedAuth || '',
          'database': DB,
          'X-Application-Name': 'MyGeotab',
        },
        body: JSON.stringify({ method, params }),
      })
        .then(r => r.json())
        .then(d => {
          if (d.error) throw new Error(d.error.message || 'API error');
          return d.result;
        });
    }

    return {
      call: (method, params) => post(method, params),
      multiCall: calls =>
        post('ExecuteMultiCall', {
          calls: calls.map(c => ({ method: c[0], params: c[1] })),
        }),
      getSession: () => Promise.resolve(null),
    };
  }

  function getWoId() {
    const m = location.hash.match(/id:([^,&]+)/);
    return m ? m[1] : null;
  }

  function isWoDetailPage() {
    return location.hash.includes('maintenanceWorkOrderDetails');
  }

  function removeButton() {
    const el = document.getElementById(BTN_ID);
    if (el) el.remove();
  }

  // printWorkOrder.js is already loaded (content.js ensures it loads first).
  // This function is synchronous — no async script loading needed.
  function injectButton() {
    if (document.getElementById(BTN_ID)) return;
    if (!isWoDetailPage()) return;

    const section = document.querySelector('.zen-main-header__adaptive-section');
    if (!section) return;

    const btn = document.createElement('button');
    btn.id = BTN_ID;
    btn.title = 'Print Work Order';
    btn.className = 'zen-button zen-button--secondary zen-caption zen-text-icon-button zen-header-button';

    btn.innerHTML = `
      <svg style="width:14px;height:14px;flex-shrink:0" viewBox="0 0 24 24" fill="none"
           stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <polyline points="6 9 6 2 18 2 18 9"/>
        <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/>
        <rect x="6" y="14" width="12" height="8"/>
      </svg>
      <span class="zen-caption__content">Print</span>`;

    btn.addEventListener('click', () => {
      const woId = getWoId();
      if (!woId) {
        alert('No work order ID found in the URL. Please open a specific work order.');
        return;
      }

      const handler =
        window.geotab &&
        window.geotab.customButtons &&
        window.geotab.customButtons.printWorkOrder;
      if (!handler) {
        alert('Print handler not ready. Please refresh the page and try again.');
        return;
      }

      handler({}, makeApi(), { entity: { id: woId } });
    });

    // Insert before the Delete button
    const deleteBtn = section.querySelector('#maintenanceWorkOrderRemoveButton');
    if (deleteBtn) {
      section.insertBefore(btn, deleteBtn);
    } else {
      section.prepend(btn);
    }
  }

  // 3. React to SPA hash changes
  let lastHash = '';
  function checkPage() {
    if (location.hash === lastHash) return;
    lastHash = location.hash;

    if (isWoDetailPage()) {
      // Wait for MyGeotab to finish rendering its toolbar before injecting
      setTimeout(injectButton, 800);
    } else {
      removeButton();
    }
  }

  window.addEventListener('hashchange', checkPage);

  // 4. MutationObserver handles cases where MyGeotab re-renders the header
  // (e.g. after loading data, switching between work orders)
  let lastCheck = 0;
  new MutationObserver(() => {
    const now = Date.now();
    if (now - lastCheck < 500) return;
    lastCheck = now;
    if (isWoDetailPage() && !document.getElementById(BTN_ID)) {
      injectButton();
    }
  }).observe(document.body, { childList: true, subtree: true });

  // 5. Initial check in case the page loads directly on a WO detail URL
  checkPage();
})();
