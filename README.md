# Belsito Plumbing Email Assistant

A local Windows 11 desktop assistant for prioritizing email follow-up. Start with synthetic sample emails, test each function, and enable read-only Outlook Classic access only after the sample checks pass and you grant permission.

The application is being delivered in stages. The Windows installation and real Outlook mailbox have **not** been tested in this Linux cloud workspace. No changes have been made to your computer.

## What is included

- A **Review My Emails** button and a prioritized report for the last 30 days.
- Categories for customer service/new jobs, estimates/proposals, outstanding invoices/collections, regulatory notices including DOB, DEP and FDNY, and urgency.
- Suggested actions and the wording that triggered each flag.
- Local completion tracking, reopening items, and an **Export to Excel** button that creates an `.xlsx` workbook without requiring Excel to be installed.
- An optional Outlook Classic COM reader for the default Inbox and its subfolders, matching your selected scope.

Rules run locally. They can miss a message or flag an item that is already resolved; review the evidence before acting. Collections flags come from email wording and are not reconciled against QuickBooks. Regulatory flags do not establish that a notice is authentic. Attachments are not read.

## Start with Stage 1

Follow [Windows installation and permission steps](docs/WINDOWS_SETUP.md), then [test one function at a time](docs/ACCEPTANCE_CHECKLIST.md). Stage 1 requires Python 3.11 or newer with Tkinter and uses only Python's standard library. Cloud checks use Python 3.12.14.

After you approve creating the local virtual environment, the Python 3.12 example is:

```bat
py -3.12 -m venv .venv
.venv\Scripts\python.exe -m belsito_mail --sample-check
```

Run `Belsito.cmd` to open the interface. It always starts in **Sample emails (Stage 1)** mode. It does not install software or access Outlook automatically. If your approved Python version differs, change `-3.12` in the setup command to that version.

## Stage 2: permission-controlled Outlook test

After completing the sample review, marking an item completed, and exporting a workbook, confirm **Sample checks passed — enable Stage 2** in the interface. Enabling Stage 2 alone does not read Outlook.

The optional dependency is pinned in [requirements-outlook.txt](requirements-outlook.txt). Install it only after approving that separate change. Open the existing Outlook Classic POP profile yourself, select **Outlook Classic (Stage 2)**, and click **Review My Emails**. Each review asks permission before reading.

The reader attaches to an already-running Outlook Classic instance and reviews the default Inbox and its subfolders for the last 30 days. It does not search every account, PST, archive, Sent Items or Deleted Items folder. It reads up to 5,000 items, traverses up to 1,000 folders, and reads at most 100,000 body characters per message. Warnings identify incomplete coverage; a report with a warning must not be treated as complete.

## Data and safety

The assistant has no email send, delete, move, save, mark-as-read or other mailbox-edit operation, and no QuickBooks integration. It uses no external AI or remote email API and stores no email passwords. Outlook itself may perform its normal background synchronization while open.

Email text stays in application memory during review. Completion status and the sample-test acknowledgement are saved locally at:

```text
%LOCALAPPDATA%\BelsitoPlumbing\EmailAssistant\state.sqlite3
```

That database contains opaque item IDs, source labels, completion timestamps and the acknowledgement; it contains no email subjects, bodies, addresses or credentials. Sample and Outlook completion records are separate. Exported workbooks contain report fields such as subjects, senders, evidence and suggested actions, but not full email bodies. Choose a local, non-synchronized export folder if you want the report to stay on this computer.

## Development validation

Run from the application folder:

```bat
python -m unittest discover -s tests -v
python -m belsito_mail --sample-check
```

The tests use synthetic messages and fake COM objects; they do not access a mailbox. `--sample-check --export sample-report.xlsx` creates a sample workbook at the specified path. A command-line sample check does not enable Stage 2; the interactive acceptance checks and acknowledgement are still required.

The current version passed 62 automated tests in the cloud, including the Tkinter interface on a temporary test display. Independent spreadsheet readers opened the sample workbook. An Excel-format compatibility update is included following a repair warning during the Windows sample test. Re-export and verify the workbook in Microsoft Excel before completing Stage 1. Actual Outlook Classic/POP access still requires the separately approved Stage 2 test.

A packaged executable, automatic scheduled reports, custom rule editor, external AI connection and QuickBooks access are future work, not part of this version.
