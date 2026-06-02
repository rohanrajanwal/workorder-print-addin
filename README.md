# Print Work Order — MyGeotab Add-In

Adds a **Print** button to the MyGeotab Work Order detail page (`maintenanceWorkOrderDetails`). Clicking it generates a formatted Work Order PDF directly from the page — no external tools required.

**Status:** Working POC, installed and tested in `geotabdemo55`. Slated for Fall 2026 dev work.

**JPD:** [MYGJPD-2492 — QOL: Printable Work Orders for Shop Operations](https://geotab.atlassian.net/jira/polaris/projects/MYGJPD/ideas/view/4804807?selectedIssue=MYGJPD-2492)

---

## Repository structure

```
addin/                 Production add-in (this is what gets uploaded)
  config.json          MyGeotab add-in configuration
  printWorkOrder.js    Single JS file — handles fetch, render, print

dev/                   Local dev harness (mock MyGeotab SDK for browser testing)
  dev-harness.html
  mock-sdk.js

docs/                  Documentation
  Print-Work-Order-Technical-Deep-Dive.docx   Full architecture walkthrough
  build_doc.py                                 Regenerates the .docx
  jpd-description.md                           Current JPD description source
```

---

## Installing in a MyGeotab database

1. Open MyGeotab → **System Settings → Add-Ins**
2. Make sure **Allow unverified Add-Ins** is set to **Yes**
3. Click **Add-In**
4. **Configuration** tab — paste the contents of [`addin/config.json`](addin/config.json)
5. **Files** tab — upload [`addin/printWorkOrder.js`](addin/printWorkOrder.js)
6. Click **Done**, then **Save**
7. Refresh any Work Order details page — the **Print** button appears next to Save/Delete

---

## How it works (in 3 lines)

1. MyGeotab calls our function when the Print button is clicked, handing us a pre-authed API client.
2. We fetch the WorkOrder, its jobs (filtered by `workOrderId`), and the Device.
3. We build an HTML overlay and call `window.print()` — the browser handles the rest.

For the full architecture walkthrough, see [`docs/Print-Work-Order-Technical-Deep-Dive.docx`](docs/Print-Work-Order-Technical-Deep-Dive.docx).

---

## Mirrors

- **GitLab (primary, Geotab internal):** [git.geotab.com/rohandeeprajanwal/workorder-print-addin](https://git.geotab.com/rohandeeprajanwal/workorder-print-addin)
- **GitHub (personal archive):** [github.com/rohanrajanwal/workorder-print-addin](https://github.com/rohanrajanwal/workorder-print-addin)

When making changes, push to both:
```bash
git push gitlab main && git push origin main
```
