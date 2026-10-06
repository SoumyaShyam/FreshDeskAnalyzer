# Freshdesk Snapshot Analyzer

## Objective

Analyze dated Freshdesk Excel exports and present the latest known record for each unique ticket. The report is split into groups using the Group value on that ticket's latest record.

## Current Version

`0.0.7`

## Requirements

- Node.js 18 or newer to run the local server and read the configured default folder. A modern browser with JavaScript enabled is also required.
- Chromium-based browsers support manual directory selection; other browsers can use the multi-file picker.
- Internet access on first load for the jQuery and SheetJS browser libraries and the optional Google Fonts stylesheet. The selected workbook contents are processed locally and are not uploaded.
- Excel workbooks named `FreshDesk_CRM_<dd-mm-yyyy>.xlsx` (the reader also accepts `.xls`). The date must be a valid calendar date.
- A worksheet named `Main Sheet` is preferred, case-insensitively. If it does not exist, the first worksheet is analyzed.
- The worksheet must include `Ticket ID`, `Status`, `Subject`, `Created time`, and `Group` columns. Other report columns are optional and display as blank when missing.

## Workflow

1. On Windows, double-click `Start Freshdesk Analyzer.bat`. It starts the local server if needed, reuses it if already running, and opens `http://127.0.0.1:8001`. The app automatically reads exports from `C:\Users\soumya.roy\OneDrive - SHYAM STEEL INDUSTRIES LIMITED\Projects\#Freshdesk`; no in-app folder action is required. Do not open `index.html` directly when using the configured-folder auto-load, because browser file security blocks access to arbitrary local folders. Alternatively, run `node server.js` and open the URL manually.
2. The configured folder loads automatically on startup. Select folder to choose another folder, or Select files to load chosen workbooks. Folder selection reads files directly in the folder, not nested subfolders.
3. The application filters for Freshdesk export filenames, validates each embedded date, and reads the preferred worksheet.
4. It processes dumps oldest-to-newest. For each Ticket ID, the row from the newest dump wins; if a ticket appears more than once on the same dump date, the last row encountered wins. The Group shown is from that winning row.
5. The report shows one tab per group, ordered by ascending age; tickets with no age (including closed tickets) appear last. Search filters the current group's ticket ID, status, subject, assignee, requester name, or email. Export CSV downloads the visible search results for the current group in the same order.

## Report Fields

| Report field | Freshdesk source | Display rule |
| --- | --- | --- |
| Sl. No. | Report row order | Re-numbered from 1 for the selected and filtered group; included in CSV export |
| Ticket ID | `Ticket ID` | Ticket identifier |
| Status | `Status` | Status from the ticket's newest record |
| Subject | `Subject` | Ticket subject |
| Created date-time | `Created time` | Localized date and time |
| Age (days) | `Created time`, `Status` | Whole calendar days from created date to newest successfully read dump date, never below zero; blank for closed tickets |
| ETA | `ETA` | Excel date serials are displayed as dates; other values are preserved |
| Resolved date | `Resolved time` | Date portion only |
| Closed date | `Closed time` | Date portion only |
| Assigned to | `Assignee` | Assignee name |
| Raised by | `Full name` | Requester name |
| Email | `Contact ID` | As provided in the export; if this field contains an internal contact identifier rather than an email address, that value is shown unchanged |

Column matching ignores case, spaces, and punctuation for the listed known aliases. If `Contact ID` does not contain email addresses in the source export, update the alias mapping in `app.js` to use the export's actual email header.

## Use Cases

- Review the latest status of all tickets across a sequence of exports.
- Compare team workloads by switching between group reports.
- Find tickets by identifier, subject, assignee, requester, or contact value.
- Download a group's currently visible results for spreadsheet follow-up.
- Read the standard Freshdesk export folder automatically when the local application starts.
- Continue analyzing exports in browsers that do not offer directory access by selecting the files together.

## User Stories

- As a support manager, I want one latest row per ticket so that older snapshots do not inflate current ticket counts.
- As a team lead, I want a separate view for each group so that I can focus on my team's tickets.
- As an analyst, I want resolved and closed timestamps reduced to dates and ticket age calculated against the dump date so that historical snapshots are interpreted consistently.
- As a user handling customer exports, I want analysis to happen locally so the workbooks are not sent to a service.
- As a user in a browser without folder-picker support, I want to select multiple files and receive the same report.
- As a support analyst, I want each ticket's latest Freshdesk status and a sequential row number so the report is easy to reference and export.

## Privacy and Limitations

- The local Node server reads only the configured export folder and binds to `127.0.0.1`; workbooks are not uploaded to an external service.
- Third-party libraries and fonts are loaded from CDNs. For offline deployment, download and serve those dependencies locally and update the corresponding references in `index.html`.
- Manual browser folder access generally requires a secure context (such as localhost or HTTPS) and user permission. The multi-file picker is the fallback.
- Only files directly inside the selected folder are considered; subfolders are not traversed.
- The report's Email column uses the requested `Contact ID` source field. Verify the Freshdesk export's column semantics and adjust the mapping if the export uses a separate email field.
- Duplicate Ticket IDs on one dump date are resolved by worksheet order, with the later row taking precedence.

## Changelog

### 0.0.1 - 2026-10-06

- Initial browser application for selecting Freshdesk Excel dumps.
- Added filename date validation, `Main Sheet` preference, latest-record-per-ticket selection, group tabs, search, summary metrics, age calculation, and group CSV export.
- Added responsive layout, browser-side processing, usage documentation, and AI prompt log.

### 0.0.2 - 2026-10-06

- Added latest Status and sequential Sl. No. columns to group reports and CSV exports.
- Added a loopback-only Node server that automatically loads exports from the configured Freshdesk folder, with an environment-variable override and manual import fallbacks.
- Validated real Freshdesk exports and formatted numeric Excel ETA dates for readability.
- Updated setup, privacy, workflow, and report-field documentation.

### 0.0.3 - 2026-10-06

- Removed the manual Load default button; configured-folder loading still runs automatically when the app starts.
- Suppressed age for tickets whose latest status is Closed.
- Sorted each group's report and CSV rows by ascending age, with blank ages last and Ticket ID as a stable tie-breaker.

### 0.0.4 - 2026-10-06

- Clarified that the configured Freshdesk folder is read automatically when the app is opened through the local Node server, and documented the direct-file browser security limitation.

### 0.0.5 - 2026-10-06

- Added a Windows one-click launcher that starts or reuses the local server, checks its folder API, and opens the application.
- Updated the quick-start steps to avoid launching the HTML file directly.

### 0.0.6 - 2026-10-06

- Displayed the configured folder path before loading completes.
- Added an in-app explanation when `index.html` is opened directly, since browsers prohibit silent access to arbitrary local folders from `file://` pages.

### 0.0.7 - 2026-10-06

- Moved Age (days) directly after Created date-time in the on-screen report and CSV exports.