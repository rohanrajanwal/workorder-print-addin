/* ========================================
   Print Work Order — MyGeotab Button Add-in
   Generates a print-friendly Repair Order
   from Work Order detail page data.
   ======================================== */

(function () {
  'use strict';

  // ---- Status & Priority Maps ----
  const STATUS_LABELS = {
    Pending: 'Pending',
    InProgress: 'In Progress',
    Completed: 'Completed',
    Closed: 'Closed',
    Deferred: 'Deferred'
  };

  const PRIORITY_LABELS = {
    Low: 'Low',
    Medium: 'Medium',
    High: 'High',
    Critical: 'Critical'
  };

  // ---- Shop Info (localStorage override) ----
  const STORAGE_KEY = 'wo-print-shop-info';

  function getShopInfoOverride() {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      if (stored) return JSON.parse(stored);
    } catch (e) { /* ignore */ }
    return null;
  }

  function saveShopInfoOverride(info) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(info));
  }

  // Fetch shop info from the MyGeotab API: CompanyDetails + logged-in User
  async function fetchShopInfoFromApi(api) {
    const info = { name: '', address: '', phone: '' };
    try {
      // Fetch CompanyDetails and the logged-in user's record in parallel
      const [companyResults, sessionInfo] = await Promise.all([
        api.call('Get', { typeName: 'CompanyDetails' }).catch(() => []),
        (api.getSession ? api.getSession() : Promise.resolve(null))
      ]);

      const company = companyResults && companyResults[0];
      if (company) {
        info.name = company.companyName || '';
        info.phone = company.phoneNumber || '';
      }

      // Try to get the logged-in user's companyAddress for the address field
      if (sessionInfo && sessionInfo.userName) {
        try {
          const users = await api.call('Get', {
            typeName: 'User',
            search: { name: sessionInfo.userName }
          });
          const user = users && users[0];
          if (user) {
            if (!info.name && user.companyName) info.name = user.companyName;
            if (user.companyAddress) info.address = user.companyAddress;
            if (!info.phone && user.phoneNumber) info.phone = user.phoneNumber;
          }
        } catch (e) { /* user fetch failed, continue with what we have */ }
      }
    } catch (e) {
      console.warn('[WO Print] Failed to fetch company info from API:', e);
    }
    return info;
  }

  // Resolve shop info: API first, then localStorage override, then prompt if still empty
  async function resolveShopInfo(api) {
    // 1. Check localStorage override first (user previously edited)
    const override = getShopInfoOverride();
    if (override && override.name) return override;

    // 2. Try the API
    const apiInfo = await fetchShopInfoFromApi(api);
    if (apiInfo.name) {
      // Cache it so we don't re-fetch every time
      saveShopInfoOverride(apiInfo);
      return apiInfo;
    }

    // 3. Nothing found — prompt the user
    const prompted = await promptShopInfo();
    return prompted; // null if cancelled
  }

  function promptShopInfo() {
    return new Promise((resolve) => {
      const existing = getShopInfoOverride() || {};

      const overlay = document.createElement('div');
      overlay.id = 'wo-settings-overlay';
      overlay.innerHTML = `
        <style>
          #wo-settings-overlay {
            position: fixed; inset: 0; z-index: 999999;
            background: rgba(0,0,0,0.5);
            display: flex; align-items: center; justify-content: center;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
          }
          #wo-settings-overlay .wo-modal {
            background: #fff; border-radius: 8px; padding: 28px 32px;
            width: 420px; max-width: 90vw; box-shadow: 0 8px 32px rgba(0,0,0,0.25);
          }
          #wo-settings-overlay h2 {
            margin: 0 0 4px; font-size: 18px; color: #1a1a1a;
          }
          #wo-settings-overlay .wo-subtitle {
            margin: 0 0 20px; font-size: 13px; color: #666;
          }
          #wo-settings-overlay label {
            display: block; font-size: 13px; font-weight: 600;
            color: #333; margin-bottom: 4px;
          }
          #wo-settings-overlay input {
            width: 100%; box-sizing: border-box; padding: 8px 10px;
            border: 1px solid #ccc; border-radius: 4px; font-size: 14px;
            margin-bottom: 14px;
          }
          #wo-settings-overlay input:focus {
            outline: none; border-color: #2D7FF9; box-shadow: 0 0 0 2px rgba(45,127,249,0.2);
          }
          #wo-settings-overlay .wo-btn-row {
            display: flex; justify-content: flex-end; gap: 10px; margin-top: 8px;
          }
          #wo-settings-overlay button {
            padding: 8px 20px; border-radius: 4px; font-size: 14px;
            cursor: pointer; border: 1px solid #ccc; background: #fff; color: #333;
          }
          #wo-settings-overlay button.primary {
            background: #2D7FF9; color: #fff; border-color: #2D7FF9;
          }
          #wo-settings-overlay button.primary:hover { background: #1a6ce5; }
          #wo-settings-overlay button:not(.primary):hover { background: #f5f5f5; }
        </style>
        <div class="wo-modal">
          <h2>Shop Information</h2>
          <p class="wo-subtitle">This appears on the printed Repair Order header. You can change it anytime.</p>
          <label for="wo-shop-name">Company / Shop Name</label>
          <input id="wo-shop-name" type="text" placeholder="e.g. Geotab Fleet Services" value="${existing.name || ''}">
          <label for="wo-shop-address">Address</label>
          <input id="wo-shop-address" type="text" placeholder="e.g. 2440 Winston Park Dr, Oakville ON" value="${existing.address || ''}">
          <label for="wo-shop-phone">Phone</label>
          <input id="wo-shop-phone" type="text" placeholder="e.g. (905) 555-1234" value="${existing.phone || ''}">
          <div class="wo-btn-row">
            <button id="wo-settings-cancel">Cancel</button>
            <button id="wo-settings-save" class="primary">Save & Print</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);

      document.getElementById('wo-settings-cancel').addEventListener('click', () => {
        overlay.remove();
        resolve(null);
      });

      document.getElementById('wo-settings-save').addEventListener('click', () => {
        const info = {
          name: document.getElementById('wo-shop-name').value.trim(),
          address: document.getElementById('wo-shop-address').value.trim(),
          phone: document.getElementById('wo-shop-phone').value.trim()
        };
        saveShopInfoOverride(info);
        overlay.remove();
        resolve(info);
      });

      setTimeout(() => document.getElementById('wo-shop-name').focus(), 50);
    });
  }

  // ---- Format Helpers ----
  function fmtCurrency(val) {
    const num = parseFloat(val) || 0;
    return '$' + num.toFixed(2);
  }

  function fmtDate(isoStr) {
    if (!isoStr) return missing('date missing');
    const d = new Date(isoStr);
    return d.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
  }

  function fmtOdometer(val) {
    if (val == null) return missing('odometer missing');
    return Math.round(val).toLocaleString() + ' km';
  }

  function fmtEngineHours(val) {
    if (val == null) return missing('hours missing');
    return parseFloat(val).toFixed(1) + ' hrs';
  }

  function escHtml(str) {
    if (!str) return '';
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  // Grey placeholder for missing data
  function missing(label) {
    return `<span class="ro-missing">(${label})</span>`;
  }

  // ---- Build Repair Order HTML ----
  function buildRepairOrderHTML(wo, jobs, device, openedByUser, assignedToUser, shopInfo) {
    const yearMakeModel = [device.year, device.make, device.model].filter(Boolean).join(' ') || device.name || missing('vehicle missing');
    const statusLabel = STATUS_LABELS[wo.status] || wo.status || missing('status missing');
    const priorityLabel = PRIORITY_LABELS[wo.priority] || wo.priority || missing('priority missing');
    const openedBy = openedByUser ? (openedByUser.firstName + ' ' + openedByUser.lastName) : missing('user missing');
    const assignedTo = assignedToUser ? (assignedToUser.firstName + ' ' + assignedToUser.lastName) : missing('unassigned');
    const printDate = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });

    // Build job rows — dual-column layout: Parts (left) + Labor (right)
    // Each row: # | Part Description | Qty | Sale | Extended | Labor Description | Extended
    let jobRows = '';
    jobs.forEach((job, i) => {
      const hasParts = job.partsCost != null && parseFloat(job.partsCost) > 0;
      const hasLabor = job.laborCost != null && parseFloat(job.laborCost) > 0;
      const partsExt = parseFloat(job.partsCost) || 0;
      const laborExt = parseFloat(job.laborCost) || 0;
      const notes = job.notes ? `<div class="ro-job-notes">Notes: ${escHtml(job.notes)}</div>` : '';
      jobRows += `
        <tr>
          <td class="ro-col-num">${i + 1}</td>
          <td class="ro-col-partdesc">
            <div class="ro-job-name">${escHtml(job.name) || missing('name missing')}</div>
          </td>
          <td class="ro-col-qty">${hasParts ? '1' : ''}</td>
          <td class="ro-col-money">${hasParts ? fmtCurrency(partsExt) : missing('cost missing')}</td>
          <td class="ro-col-money">${hasParts ? fmtCurrency(partsExt) : ''}</td>
          <td class="ro-col-labordesc">
            <div class="ro-job-detail">${escHtml(job.description) || missing('description missing')}</div>
            ${notes}
          </td>
          <td class="ro-col-money">${hasLabor ? fmtCurrency(laborExt) : missing('cost missing')}</td>
        </tr>
      `;
    });

    if (jobs.length === 0) {
      jobRows = '<tr><td colspan="7" style="text-align:center;color:#999;padding:16px;">No jobs on this work order</td></tr>';
    }

    // Cost summary
    const laborTotal = jobs.reduce((s, j) => s + (parseFloat(j.laborCost) || 0), 0);
    const partsTotal = jobs.reduce((s, j) => s + (parseFloat(j.partsCost) || 0), 0);
    const otherTotal = jobs.reduce((s, j) => s + (parseFloat(j.otherCost) || 0), 0);
    const taxTotal = parseFloat(wo.taxCost) || 0;
    const subTotal = laborTotal + partsTotal + otherTotal;
    const grandTotal = parseFloat(wo.totalCost) || (subTotal + taxTotal);

    return `
      <div class="ro-page">
        <!-- Header -->
        <div class="ro-header">
          <div class="ro-header-left">
            <div class="ro-company">${escHtml(shopInfo.name) || 'Company Name'}</div>
            ${shopInfo.address ? `<div class="ro-shop-detail">${escHtml(shopInfo.address)}</div>` : ''}
            ${shopInfo.phone ? `<div class="ro-shop-detail">Phone: ${escHtml(shopInfo.phone)}</div>` : ''}
          </div>
          <div class="ro-header-right">
            <div class="ro-invoice-box">
              <div class="ro-ro-label">INVOICE</div>
              <div class="ro-ro-number">${escHtml(wo.reference || wo.id)}</div>
            </div>
            <div class="ro-print-date">Print Date: ${printDate}</div>
          </div>
        </div>

        <!-- Vehicle & WO Info -->
        <div class="ro-section">
          <div class="ro-vehicle-header">
            <div class="ro-vehicle-title">${escHtml(yearMakeModel)}</div>
          </div>
          <div class="ro-grid-3">
            <div class="ro-field"><span class="ro-label">VIN:</span> ${escHtml(device.vehicleIdentificationNumber) || missing('VIN missing')}</div>
            <div class="ro-field"><span class="ro-label">Lic #:</span> ${escHtml(device.licensePlate) || missing('plate missing')}</div>
            <div class="ro-field"><span class="ro-label">Unit #:</span> ${escHtml(device.name) || missing('unit missing')}</div>
            <div class="ro-field"><span class="ro-label">Odometer In:</span> ${fmtOdometer(device.odometer)}</div>
            <div class="ro-field"><span class="ro-label">Engine Hrs:</span> ${fmtEngineHours(device.engineHours)}</div>
            <div class="ro-field"><span class="ro-label">Date Opened:</span> ${fmtDate(wo.dateOpened)}</div>
            <div class="ro-field"><span class="ro-label">Status:</span> ${statusLabel}</div>
            <div class="ro-field"><span class="ro-label">Priority:</span> ${priorityLabel}</div>
            <div class="ro-field"><span class="ro-label">Assigned:</span> ${escHtml(assignedTo)}</div>
          </div>
          <div class="ro-field ro-field-full"><span class="ro-label">Opened by:</span> ${escHtml(openedBy)}</div>
        </div>

        <!-- Jobs Table — dual Parts + Labor -->
        <div class="ro-section ro-section-jobs">
          <table class="ro-table">
            <thead>
              <tr>
                <th class="ro-th-num">#</th>
                <th class="ro-th-partdesc">Part Description</th>
                <th class="ro-th-qty">Qty</th>
                <th class="ro-th-money">Sale</th>
                <th class="ro-th-money">Extended</th>
                <th class="ro-th-labordesc">Labor Description</th>
                <th class="ro-th-money">Extended</th>
              </tr>
            </thead>
            <tbody>
              ${jobRows}
            </tbody>
          </table>
        </div>

        <!-- WO Notes -->
        ${wo.notes ? `
        <div class="ro-section">
          <div class="ro-section-title">WORK ORDER NOTES</div>
          <div class="ro-notes-text">${escHtml(wo.notes)}</div>
        </div>
        ` : ''}

        <!-- Footer: Technician + Cost Summary side by side -->
        <div class="ro-footer-row">
          <!-- Sign-off (left) -->
          <div class="ro-signoff-col">
            <div class="ro-section ro-signoff">
              <div class="ro-section-title">COMPLETION SIGN-OFF</div>
              <div class="ro-checklist">
                <div class="ro-check-item">&#9744; Parts Received</div>
                <div class="ro-check-item">&#9744; Work Completed</div>
                <div class="ro-check-item">&#9744; QC Approved</div>
                <div class="ro-check-item">&#9744; Vehicle Returned</div>
              </div>
              <div class="ro-sig-lines">
                <div class="ro-sig-line"><span class="ro-label">Technician:</span> <span class="ro-underline"></span></div>
                <div class="ro-sig-line"><span class="ro-label">Supervisor:</span> <span class="ro-underline"></span></div>
                <div class="ro-sig-line"><span class="ro-label">Notes:</span> <span class="ro-underline"></span></div>
                <div class="ro-sig-line"><span class="ro-underline ro-underline-full"></span></div>
              </div>
            </div>
          </div>

          <!-- Cost summary (right) -->
          <div class="ro-cost-col">
            <div class="ro-section ro-cost-summary">
              <div class="ro-cost-row"><span>Labor:</span><span>${fmtCurrency(laborTotal)}</span></div>
              <div class="ro-cost-row"><span>Parts:</span><span>${fmtCurrency(partsTotal)}</span></div>
              ${otherTotal > 0 ? `<div class="ro-cost-row"><span>Sublet:</span><span>${fmtCurrency(otherTotal)}</span></div>` : `<div class="ro-cost-row"><span>Sublet:</span><span>$0.00</span></div>`}
              <div class="ro-cost-row ro-cost-sub"><span>Sub:</span><span>${fmtCurrency(subTotal)}</span></div>
              <div class="ro-cost-row"><span>Tax:</span><span>${fmtCurrency(taxTotal)}</span></div>
              <div class="ro-cost-row ro-cost-total"><span>Total:</span><span>${fmtCurrency(grandTotal)}</span></div>
            </div>
          </div>
        </div>

        <!-- Authorization -->
        <div class="ro-section ro-auth">
          <div class="ro-auth-text">I hereby authorize the above repair work to be done along with the necessary material and hereby grant you and/or your employees permission to operate the car or truck herein described on street, highways or elsewhere for the purpose to testing and/or inspection.</div>
          <div class="ro-sig-grid-bottom">
            <div class="ro-sig-line"><span class="ro-label">Signature:</span> <span class="ro-underline"></span></div>
            <div class="ro-sig-line"><span class="ro-label">Date:</span> <span class="ro-underline"></span></div>
            <div class="ro-sig-line"><span class="ro-label">Time:</span> <span class="ro-underline"></span></div>
          </div>
        </div>
      </div>
    `;
  }

  // ---- Print CSS ----
  // Shared styles used in both @media print and screen preview
  const SHARED_STYLES = `
      .ro-page {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif;
        font-size: 11px;
        color: #000;
        line-height: 1.4;
      }

      /* Header */
      .ro-header {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        border-bottom: 2px solid #000;
        padding-bottom: 10px;
        margin-bottom: 12px;
      }
      .ro-company { font-size: 18px; font-weight: 700; }
      .ro-shop-detail { font-size: 11px; color: #444; }
      .ro-header-right { text-align: right; }
      .ro-invoice-box {
        border: 2px solid #000; padding: 4px 14px;
        display: inline-block; text-align: center; margin-bottom: 4px;
      }
      .ro-ro-label { font-size: 14px; font-weight: 700; letter-spacing: 1px; }
      .ro-ro-number { font-size: 16px; font-weight: 700; }
      .ro-print-date { font-size: 10px; color: #666; margin-top: 2px; }

      /* Sections */
      .ro-section {
        border: 1px solid #ccc; border-radius: 3px;
        padding: 8px 10px; margin-bottom: 10px;
        page-break-inside: avoid;
      }
      .ro-section-title {
        font-size: 11px; font-weight: 700; text-transform: uppercase;
        letter-spacing: 0.5px; border-bottom: 1px solid #ddd;
        padding-bottom: 4px; margin-bottom: 6px;
      }

      /* Vehicle header */
      .ro-vehicle-header { margin-bottom: 6px; }
      .ro-vehicle-title { font-size: 13px; font-weight: 700; }

      /* Field grid */
      .ro-grid-3 {
        display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 4px 16px;
      }
      .ro-field { font-size: 11px; padding: 2px 0; }
      .ro-field-full { margin-top: 4px; }
      .ro-label { font-weight: 600; color: #333; }
      .ro-missing { color: #bbb; font-style: italic; font-size: 10px; }

      /* Jobs table — dual parts + labor */
      .ro-section-jobs { padding: 0; }
      .ro-table { width: 100%; border-collapse: collapse; }
      .ro-table thead tr {
        background: #f0f0f0 !important;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }
      .ro-table th {
        font-size: 9px; font-weight: 700; text-transform: uppercase;
        text-align: left; padding: 5px 6px; border-bottom: 1px solid #999;
      }
      .ro-th-num { width: 24px; text-align: center; }
      .ro-th-partdesc { }
      .ro-th-qty { width: 32px; text-align: center; }
      .ro-th-money { width: 62px; text-align: right; }
      .ro-th-labordesc { border-left: 2px solid #ccc; padding-left: 8px; }

      .ro-table td {
        padding: 6px 6px; border-bottom: 1px solid #e0e0e0; vertical-align: top;
        font-size: 11px;
      }
      .ro-col-num { text-align: center; font-weight: 600; width: 24px; }
      .ro-col-partdesc { }
      .ro-col-qty { text-align: center; width: 32px; }
      .ro-col-money { text-align: right; font-family: 'Consolas', 'Courier New', monospace; width: 62px; }
      .ro-col-labordesc { border-left: 2px solid #eee; padding-left: 8px; }
      .ro-job-name { font-weight: 600; }
      .ro-job-detail { font-size: 10px; color: #444; }
      .ro-job-notes { font-size: 10px; color: #666; font-style: italic; margin-top: 3px; }

      /* WO Notes */
      .ro-notes-text { font-size: 11px; white-space: pre-wrap; }

      /* Footer row: sign-off left, cost right */
      .ro-footer-row {
        display: flex; gap: 12px; align-items: flex-start;
      }
      .ro-signoff-col { flex: 1; }
      .ro-cost-col { flex: 0 0 220px; }

      /* Cost summary */
      .ro-cost-summary { }
      .ro-cost-row {
        display: flex; justify-content: space-between; padding: 3px 0;
        font-size: 11px; font-family: 'Consolas', 'Courier New', monospace;
      }
      .ro-cost-sub {
        border-top: 1px solid #999; margin-top: 2px; padding-top: 4px;
      }
      .ro-cost-total {
        border-top: 2px solid #000; margin-top: 2px; padding-top: 4px;
        font-weight: 700; font-size: 13px;
      }

      /* Sign-off */
      .ro-signoff { page-break-inside: avoid; }
      .ro-checklist {
        display: grid; grid-template-columns: 1fr 1fr; gap: 4px 24px; margin-bottom: 10px;
      }
      .ro-check-item { font-size: 11px; }
      .ro-sig-lines { display: flex; flex-direction: column; gap: 10px; }
      .ro-sig-line {
        display: flex; align-items: flex-end; gap: 6px; font-size: 11px;
      }
      .ro-underline {
        flex: 1; border-bottom: 1px solid #000; min-width: 80px; height: 16px;
      }
      .ro-underline-full { width: 100%; }

      /* Authorization */
      .ro-auth { }
      .ro-auth-text { font-size: 9px; color: #444; margin-bottom: 8px; line-height: 1.3; }
      .ro-sig-grid-bottom {
        display: grid; grid-template-columns: 2fr 1fr 1fr; gap: 10px;
      }
  `;

  const PRINT_CSS = `
    /* === Shared RO styles (apply in both screen + print) === */
    ${SHARED_STYLES}

    /* === Screen: hide overlay normally === */
    #wo-print-overlay { display: none; }

    /* === Screen preview (dev harness) === */
    #wo-print-overlay.wo-preview {
      display: block !important;
      max-width: 8.5in; margin: 20px auto; padding: 0.5in;
      background: #fff; box-shadow: 0 2px 16px rgba(0,0,0,0.15);
      border: 1px solid #ddd;
    }

    @media print {
      body > *:not(#wo-print-overlay) { display: none !important; }
      #wo-print-overlay {
        display: block !important; position: absolute;
        top: 0; left: 0; width: 100%;
      }
      @page { size: letter; margin: 0.5in; }
    }
  `;

  // ---- Inject Styles ----
  function ensureStyles() {
    if (document.getElementById('wo-print-styles')) return;
    const style = document.createElement('style');
    style.id = 'wo-print-styles';
    style.textContent = PRINT_CSS;
    document.head.appendChild(style);
  }

  // ---- Print Flow ----
  function triggerPrint(html, previewMode, title) {
    ensureStyles();

    const existing = document.getElementById('wo-print-overlay');
    if (existing) existing.remove();

    if (previewMode) {
      // Dev harness: render on screen
      const overlay = document.createElement('div');
      overlay.id = 'wo-print-overlay';
      overlay.innerHTML = html;
      overlay.classList.add('wo-preview');
      document.body.appendChild(overlay);
      return;
    }

    // Print via a clean new window with @page margin:0 to suppress browser headers/footers
    const printTitle = title || 'Repair Order';
    const printCSS = `
      @page { size: letter; margin: 0; }
      body { margin: 0; padding: 0.5in; }
      ${SHARED_STYLES}
    `;
    const win = window.open('', '_blank');
    if (win) {
      win.document.write(`<!DOCTYPE html><html><head>
        <title>${printTitle}</title>
        <style>${printCSS}</style>
        </head><body>${html}</body></html>`);
      win.document.close();
      win.focus();
      setTimeout(() => {
        win.print();
        win.addEventListener('afterprint', () => win.close(), { once: true });
        setTimeout(() => { try { win.close(); } catch (e) {} }, 120000);
      }, 250);
    } else {
      // Popup blocked — fall back to in-page print with title swap
      const origTitle = document.title;
      document.title = printTitle;
      const overlay = document.createElement('div');
      overlay.id = 'wo-print-overlay';
      overlay.innerHTML = html;
      document.body.appendChild(overlay);
      window.print();
      const cleanup = () => {
        document.title = origTitle;
        const el = document.getElementById('wo-print-overlay');
        if (el) el.remove();
      };
      window.addEventListener('afterprint', cleanup, { once: true });
      setTimeout(cleanup, 60000);
    }
  }

  // ---- Data Fetching ----
  async function fetchWorkOrderData(api, state) {
    const woId = state && state.entity && state.entity.id;
    if (!woId) {
      throw new Error('No Work Order ID found. Please open a Work Order first.');
    }

    // Fetch WO + Jobs in parallel
    const [workOrders, jobs] = await api.multiCall([
      ['Get', { typeName: 'MaintenanceWorkOrder', search: { id: woId } }],
      ['Get', { typeName: 'MaintenanceWorkOrderJob', search: { workOrderSearch: { id: woId } } }]
    ]);

    const wo = workOrders[0];
    if (!wo) {
      throw new Error('Work Order not found (id: ' + woId + ')');
    }

    // Fetch Device + Users
    const secondCalls = [
      ['Get', { typeName: 'Device', search: { id: wo.device.id } }],
      ['Get', { typeName: 'User', search: { id: wo.openedByUser.id } }]
    ];
    if (wo.assignedToUser && wo.assignedToUser.id) {
      secondCalls.push(['Get', { typeName: 'User', search: { id: wo.assignedToUser.id } }]);
    }

    const results = await api.multiCall(secondCalls);
    const device = results[0][0] || {};
    const openedByUser = results[1][0] || null;
    const assignedToUser = results[2] ? results[2][0] || null : null;

    return { wo, jobs, device, openedByUser, assignedToUser };
  }

  // ---- Button Handler (MyGeotab integration) ----
  window.geotab = window.geotab || {};
  window.geotab.customButtons = window.geotab.customButtons || {};

  window.geotab.customButtons.printWorkOrder = async function (event, api, state) {
    try {
      // Resolve shop info: API auto-fill → localStorage → prompt
      const shopInfo = await resolveShopInfo(api);
      if (!shopInfo) return; // User cancelled prompt

      // Fetch data
      const { wo, jobs, device, openedByUser, assignedToUser } = await fetchWorkOrderData(api, state);

      // Build & print
      const html = buildRepairOrderHTML(wo, jobs, device, openedByUser, assignedToUser, shopInfo);
      const printTitle = 'Repair Order — ' + (wo.reference || wo.id);
      triggerPrint(html, false, printTitle);

    } catch (err) {
      console.error('[WO Print] Error:', err);
      alert('Print Work Order Error:\n' + err.message);
    }
  };

  // ---- Dev Harness API ----
  window._woPrint = {
    getShopInfo: getShopInfoOverride,
    saveShopInfo: saveShopInfoOverride,
    promptShopInfo,
    resolveShopInfo,
    fetchShopInfoFromApi,
    buildRepairOrderHTML,
    triggerPrint,
    ensureStyles,
    fetchWorkOrderData,
    resetShopInfo: function () {
      localStorage.removeItem(STORAGE_KEY);
    }
  };
})();
