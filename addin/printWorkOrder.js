/* ========================================
   Print Work Order — MyGeotab Button Add-in
   Generates a print-friendly Work Order
   from the Work Order detail page data.
   ======================================== */

(function () {
  'use strict';

  // ---- Status & Priority Maps ----
  // statusCode is a numeric value from the API
  const STATUS_LABELS = {
    1: 'Pending',
    2: 'In Progress',
    3: 'On Hold',
    4: 'Deferred',
    5: 'Completed',
    6: 'Closed'
  };

  // priorityCode is a numeric value from the API
  const PRIORITY_LABELS = {
    100: 'Critical',
    200: 'High',
    300: 'Medium',
    400: 'Low'
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

  // Fetch shop info from the MyGeotab API: CompanyDetails has name + phone.
  // Address isn't exposed by the API for the button-click context — user can
  // edit/add it via promptShopInfo and it persists in localStorage.
  async function fetchShopInfoFromApi(api) {
    const info = { name: '', address: '', phone: '' };
    try {
      const companyResults = await api.call('Get', { typeName: 'CompanyDetails' }).catch(() => []);
      const company = companyResults && companyResults[0];
      if (company) {
        info.name = company.companyName || '';
        info.phone = company.phoneNumber || '';
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
          <p class="wo-subtitle">This appears on the printed Work Order header. You can change it anytime.</p>
          <label for="wo-shop-name">Company / Shop Name</label>
          <input id="wo-shop-name" type="text" placeholder="e.g. Geotab Fleet Services" value="${escHtml(existing.name || '')}">
          <label for="wo-shop-address">Address</label>
          <input id="wo-shop-address" type="text" placeholder="e.g. 2440 Winston Park Dr, Oakville ON" value="${escHtml(existing.address || '')}">
          <label for="wo-shop-phone">Phone</label>
          <input id="wo-shop-phone" type="text" placeholder="e.g. (905) 555-1234" value="${escHtml(existing.phone || '')}">
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

  // MyGeotab stores odometer as a hex string in meters, e.g. "0000000008246960" → 0x8246960 m → 136605 km.
  function fmtOdometer(hexMeters) {
    if (hexMeters == null || hexMeters === '') return missing('odometer missing');
    const meters = typeof hexMeters === 'string' ? parseInt(hexMeters, 16) : Number(hexMeters);
    if (!isFinite(meters) || meters <= 0) return missing('not recorded');
    return Math.round(meters / 1000).toLocaleString() + ' km';
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
    return `<span class="ro-missing">(${escHtml(label)})</span>`;
  }

  // ---- Build Work Order HTML ----
  function buildRepairOrderHTML(wo, jobs, device, shopInfo) {
    const yearMakeModel = [device.year, device.make, device.model].filter(Boolean).join(' ') || device.name || missing('vehicle missing');
    const statusLabel = STATUS_LABELS[wo.statusCode] || String(wo.statusCode) || missing('status missing');
    const priorityLabel = PRIORITY_LABELS[wo.priorityCode] || String(wo.priorityCode) || missing('priority missing');
    // openedByUser is sometimes a User object, sometimes a string ID, sometimes the sentinel "NoUserId".
    const openedByName = wo.openedByUser && typeof wo.openedByUser === 'object' ? wo.openedByUser.name : null;
    const openedBy = openedByName ? escHtml(openedByName)
      : (wo.openedByUser && wo.openedByUser !== 'NoUserId' ? escHtml(wo.openedByUser) : missing('not recorded'));
    const assignedName = wo.assignedTo && typeof wo.assignedTo === 'object' ? wo.assignedTo.name : wo.assignedTo;
    const assignedTo = assignedName && assignedName !== 'NoUserId' ? escHtml(assignedName) : missing('unassigned');
    const printDate = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });

    // Build job rows — dual-column layout: Parts (left) + Labor (right)
    // Each row: # | Part Description | Qty | Sale | Extended | Labor Description | Extended
    let jobRows = '';
    jobs.forEach((job, i) => {
      const typeName = job.maintenanceType && job.maintenanceType.name;
      const jobName = typeName || job.eventDescription || '';
      const hasParts = job.partsCost != null && parseFloat(job.partsCost) > 0;
      const hasLabor = job.laborCosts != null && parseFloat(job.laborCosts) > 0;
      const partsExt = parseFloat(job.partsCost) || 0;
      const laborExt = parseFloat(job.laborCosts) || 0;
      const laborHrs = job.hours ? job.hours + ' hrs' : '';
      // Only render eventDescription as a Note when it wasn't already used as the job name
      const showNotes = job.eventDescription && typeName;
      const notes = showNotes ? `<div class="ro-job-notes">Notes: ${escHtml(job.eventDescription)}</div>` : '';
      const hasLaborDesc = laborHrs || showNotes;
      jobRows += `
        <tr>
          <td class="ro-col-num">${i + 1}</td>
          <td class="ro-col-partdesc">
            <div class="ro-job-name">${escHtml(jobName) || missing('name missing')}</div>
          </td>
          <td class="ro-col-qty">${hasParts ? '1' : ''}</td>
          <td class="ro-col-money">${hasParts ? fmtCurrency(partsExt) : ''}</td>
          <td class="ro-col-money">${hasParts ? fmtCurrency(partsExt) : ''}</td>
          <td class="ro-col-labordesc">
            ${hasLaborDesc ? `<div class="ro-job-detail">${laborHrs}</div>${notes}` : missing('description missing')}
          </td>
          <td class="ro-col-money">${hasLabor ? fmtCurrency(laborExt) : missing('cost missing')}</td>
        </tr>
      `;
    });

    if (jobs.length === 0) {
      jobRows = '<tr><td colspan="7" style="text-align:center;color:#999;padding:16px;">No jobs on this work order</td></tr>';
    }

    // Cost summary — costs live on jobs, not the WO
    const laborTotal = jobs.reduce((s, j) => s + (parseFloat(j.laborCosts) || 0), 0);
    const partsTotal = jobs.reduce((s, j) => s + (parseFloat(j.partsCost) || 0), 0);
    const otherTotal = jobs.reduce((s, j) => s + (parseFloat(j.shippingCosts) || 0), 0);
    const taxTotal = jobs.reduce((s, j) => s + (parseFloat(j.taxCosts) || 0), 0);
    const subTotal = laborTotal + partsTotal + otherTotal;
    // Only trust the API's totalCosts sum when every job has it populated;
    // otherwise compute from components so partial data doesn't undercount.
    const allHaveTotal = jobs.length > 0 && jobs.every(j => j.totalCosts != null);
    const grandTotal = allHaveTotal
      ? jobs.reduce((s, j) => s + (parseFloat(j.totalCosts) || 0), 0)
      : subTotal + taxTotal;

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
            <div class="ro-field"><span class="ro-label">Odometer In:</span> ${fmtOdometer(wo.odometerReading)}</div>
            <div class="ro-field"><span class="ro-label">Engine Hrs:</span> ${wo.engineHoursReadingInHours > 0 ? fmtEngineHours(wo.engineHoursReadingInHours) : missing('not recorded')}</div>
            <div class="ro-field"><span class="ro-label">Date Opened:</span> ${fmtDate(wo.dateTime)}</div>
            <div class="ro-field"><span class="ro-label">Status:</span> ${statusLabel}</div>
            <div class="ro-field"><span class="ro-label">Priority:</span> ${priorityLabel}</div>
            <div class="ro-field"><span class="ro-label">Assigned:</span> ${assignedTo}</div>${wo.completedDateTime ? `
            <div class="ro-field"><span class="ro-label">Date Completed:</span> ${fmtDate(wo.completedDateTime)}</div>` : ''}
          </div>
          <div class="ro-field ro-field-full"><span class="ro-label">Opened by:</span> ${openedBy}</div>
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
      /* Allow the jobs section to break across pages so rows can start on page 1
         and continue onto page 2, instead of pushing the whole table to page 2. */
      .ro-section-jobs { padding: 0; page-break-inside: auto; break-inside: auto; }
      .ro-table { width: 100%; border-collapse: collapse; }
      .ro-table thead { display: table-header-group; }
      .ro-table tbody tr { page-break-inside: avoid; break-inside: avoid; }
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

    // Print in-page: hide MyGeotab content via @media print, show only the overlay
    const printTitle = title || 'Work Order';
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

  // ---- Data Fetching ----
  async function fetchWorkOrderData(api, state) {
    const woId = state && state.entity && state.entity.id;
    if (!woId) {
      throw new Error('No Work Order ID found. Please open a Work Order first.');
    }

    // Fetch WO + Jobs in parallel.
    // IMPORTANT: use search.workOrderId (flat) — search.workOrder.id and search.workOrderSearch
    // are silently ignored by the API and return every job in the database.
    const [workOrders, jobs] = await api.multiCall([
      ['Get', { typeName: 'MaintenanceWorkOrder', search: { id: woId } }],
      ['Get', { typeName: 'MaintenanceWorkOrderJob', search: { workOrderId: woId } }]
    ]);

    const wo = workOrders[0];
    if (!wo) {
      throw new Error('Work Order not found (id: ' + woId + ')');
    }

    // Fetch Device only — user names are already embedded in the WO object
    const device = (await api.call('Get', { typeName: 'Device', search: { id: wo.device.id } }))[0] || {};

    return { wo, jobs, device };
  }

  // ---- Native API wrapper ----
  // MyGeotab's built-in button-click API is callback-based.
  // This wraps it in promises so the async helpers above work unchanged.
  function makePromiseApi(nativeApi) {
    return {
      call: function (method, params) {
        return new Promise(function (resolve, reject) {
          nativeApi.call(method, params, resolve, function (err) {
            reject(new Error((err && err.message) || String(err)));
          });
        });
      },
      multiCall: function (calls) {
        return new Promise(function (resolve, reject) {
          nativeApi.multiCall(calls, resolve, function (err) {
            reject(new Error((err && err.message) || String(err)));
          });
        });
      }
    };
  }

  // ---- Button Handler (MyGeotab integration) ----
  window.geotab = window.geotab || {};
  window.geotab.customButtons = window.geotab.customButtons || {};

  window.geotab.customButtons.printWorkOrder = async function (event, nativeApi, state) {
    const api = makePromiseApi(nativeApi);

    // Fall back to URL hash for the WO id if state doesn't carry it
    if (!state || !state.entity || !state.entity.id) {
      const m = window.location.hash.match(/id:([^,&]+)/);
      if (m) { state = { entity: { id: m[1] } }; }
    }

    try {
      // Resolve shop info: API auto-fill → localStorage → prompt
      const shopInfo = await resolveShopInfo(api);
      if (!shopInfo) return; // User cancelled prompt

      // Fetch data
      const { wo, jobs, device } = await fetchWorkOrderData(api, state);

      // Build & print
      const html = buildRepairOrderHTML(wo, jobs, device, shopInfo);
      const printTitle = 'Work Order — ' + (wo.reference || wo.id);
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

  // ---- Button Styling ----
  // Applied at load time so no separate changePageState script is needed.
  var PRINTER_SVG_ICON = [
    '<svg style="width:14px;height:14px;flex-shrink:0" viewBox="0 0 24 24"',
    ' fill="none" stroke="currentColor" stroke-width="2"',
    ' stroke-linecap="round" stroke-linejoin="round">',
    '<polyline points="6 9 6 2 18 2 18 9"/>',
    '<path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/>',
    '<rect x="6" y="14" width="12" height="8"/>',
    '</svg>'
  ].join('');

  // Apply styling whenever a Print button exists and either isn't styled
  // or had its inner content wiped by a MyGeotab re-render.
  function applyButtonStyle() {
    var btn = document.querySelector('button[aria-label="Print"]');
    if (!btn) return false;
    var hasIcon = !!btn.querySelector('svg');
    if (btn.dataset.woStyled === '1' && hasIcon) return false;
    btn.classList.remove('zen-button--tertiary-black');
    btn.classList.add('zen-button--primary', 'zen-caption', 'zen-text-icon-button');
    btn.innerHTML = PRINTER_SVG_ICON + '<span class="zen-caption__content">Print</span>';
    btn.dataset.woStyled = '1';
    return true;
  }

  // Watch the document for new Print buttons or inner re-renders, then restyle.
  // Runs for the lifetime of the page — coalesced via rAF so heavy SPA re-renders
  // only trigger one styling check per frame.
  function watchForButton() {
    applyButtonStyle();
    var queued = false;
    var observer = new MutationObserver(function () {
      if (queued) return;
      queued = true;
      requestAnimationFrame(function () { queued = false; applyButtonStyle(); });
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }

  if (document.body) {
    watchForButton();
  } else {
    document.addEventListener('DOMContentLoaded', watchForButton, { once: true });
  }
})();
