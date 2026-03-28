(function () {
  'use strict';

  const PRINT_SCRIPT_URL = 'https://rohanrajanwal.github.io/workorder-print-addin/printWorkOrder.js';
  const BTN_ID = 'wo-print-btn';
  // Extract database from URL path (e.g. https://my.geotab.com/demo_buildtesting16/#...)
  const DB = window.location.pathname.split('/').filter(Boolean)[0] || '';
  let scriptLoaded = false;
  let capturedAuth = null;

  // 1. Intercept XHR to capture the Bearer auth token
  const origSetHeader = XMLHttpRequest.prototype.setRequestHeader;
  XMLHttpRequest.prototype.setRequestHeader = function (k, v) {
    if (k === 'authorization' && v && v.startsWith('Bearer')) {
      capturedAuth = decodeURIComponent(v);
    }
    return origSetHeader.apply(this, arguments);
  };

  // 2. Lightweight Geotab API wrapper using fetch + captured auth
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

  // 3. Load printWorkOrder.js once into the main page context
  function loadPrintScript() {
    if (scriptLoaded) return Promise.resolve();
    return new Promise(resolve => {
      const s = document.createElement('script');
      s.src = PRINT_SCRIPT_URL;
      s.onload = () => { scriptLoaded = true; resolve(); };
      s.onerror = () => resolve();
      document.head.appendChild(s);
    });
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

  async function injectButton() {
    if (document.getElementById(BTN_ID)) return;
    if (!isWoDetailPage()) return;

    const section = document.querySelector('.zen-main-header__adaptive-section');
    if (!section) return;

    await loadPrintScript();

    // Re-check after async load
    if (!isWoDetailPage() || document.getElementById(BTN_ID)) return;

    const btn = document.createElement('button');
    btn.id = BTN_ID;
    btn.title = 'Print Work Order';
    btn.style.cssText = [
      'background:#fff',
      'border:1px solid #ccc',
      'color:#333',
      'display:inline-flex',
      'align-items:center',
      'gap:6px',
      'margin-right:8px',
      'padding:5px 12px',
      'border-radius:4px',
      'font-size:13px',
      'font-weight:500',
      'cursor:pointer',
      'line-height:1',
      'height:32px',
    ].join(';');

    btn.innerHTML = `
      <svg style="width:14px;height:14px;flex-shrink:0" viewBox="0 0 24 24" fill="none"
           stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <polyline points="6 9 6 2 18 2 18 9"/>
        <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/>
        <rect x="6" y="14" width="12" height="8"/>
      </svg>
      Print Work Order`;

    btn.addEventListener('mouseenter', () => { btn.style.background = '#f5f5f5'; });
    btn.addEventListener('mouseleave', () => { btn.style.background = '#fff'; });

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
        alert('Print handler not ready. Please wait a moment and try again.');
        return;
      }

      handler({}, makeApi(), { entity: { id: woId } });
    });

    // Insert before the Delete button to match the UI order
    const deleteBtn = section.querySelector('.zen-button--destructive');
    if (deleteBtn) {
      section.insertBefore(btn, deleteBtn);
    } else {
      section.prepend(btn);
    }
  }

  // 4. React to SPA hash changes
  let lastHash = '';
  function checkPage() {
    if (location.hash === lastHash) return;
    lastHash = location.hash;

    if (isWoDetailPage()) {
      setTimeout(injectButton, 800); // wait for MyGeotab to render its own toolbar
    } else {
      removeButton();
    }
  }

  window.addEventListener('hashchange', checkPage);

  // 5. MutationObserver to handle MyGeotab re-rendering the header
  let lastCheck = 0;
  new MutationObserver(() => {
    const now = Date.now();
    if (now - lastCheck < 500) return;
    lastCheck = now;
    if (isWoDetailPage() && !document.getElementById(BTN_ID)) {
      injectButton();
    }
  }).observe(document.body, { childList: true, subtree: true });

  // 6. Initial check on script load
  checkPage();
})();
