## What / Problem

"I would like the maintenance management software to include an option to print individual work orders for companies like us who have a maintenance program and conduct our own maintenance and repairs on our fleet vehicles."

## Job-to-be-done

**When** an In-House Maintenance Manager has finalized a Work Order in the system and needs to assign the task, **I want to** generate a clean, readable, physical copy of the Work Order (including all jobs, parts, and notes) with enough room to write in the document while balancing filling in the information automatically, **so I can** hand a tangible document to a shop technician to guide their repairs, allowing the business to run its entire shop operation through our maintenance management software.

## Current Solution (Working as of May 2026)

A custom MyGeotab add-in that adds a **Print** button to the Work Order detail page. Clicking it generates a formatted Work Order PDF directly from MyGeotab — no external tools required.

Installed and tested in `geotabdemo55`. The add-in is a single JavaScript file (~30KB) uploaded via System Settings > Add-Ins. Once installed, the Print button appears automatically on every Work Order details page for all users in that database.

**What the printed output includes:**

- Shop/company header (auto-populated from MyGeotab company settings)
- Invoice number (work order reference), print date
- Vehicle info: name, VIN, licence plate, unit number, odometer, engine hours, date opened, date completed (if applicable), status, priority, assigned technician
- Full job/parts table with part descriptions, quantities, costs, labor descriptions, and costs
- Cost summary: labor, parts, sublet, tax, and total
- Completion sign-off section with technician/supervisor signature lines
- Authorization clause with signature, date, and time fields

**Technical details:**

- Runs entirely in the browser — no server, no database, no backend
- Reads data from MyGeotab API (MaintenanceWorkOrder, MaintenanceWorkOrderJob, Device entities)
- Uses the browser's native print-to-PDF for formatting (no external PDF library)
- Button auto-styles to match MyGeotab's toolbar via MutationObserver

## Where did this idea come from

A direct request from a partner via the Community Hub: [Maintenance Management - Self Sufficient Maintenance](https://community.geotab.com/s/feedback/a40Pd000005UiVNIA0/maintenance-management-self-sufficient-maintenance)

## How will we measure success

- **Quantitative:** Track usage frequency of the Print button among customers with in-house maintenance programs.
- **Qualitative:** Partner/customer confirmation that the printed format meets daily shop operations requirements.
