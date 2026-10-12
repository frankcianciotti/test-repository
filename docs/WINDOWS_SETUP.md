# Windows setup, one approved step at a time

These steps install a sample-first assistant on your Windows 11 computer. Preparation and automated testing were done in a Linux cloud workspace. Your computer and Outlook mailbox have not been accessed. Complete Stage 1 before proceeding to Stage 2.

## Before making changes

Have the application files, an ordinary user account, and a local folder outside OneDrive or another synchronized location ready. Outlook and QuickBooks need no configuration changes for Stage 1. Do not run an automatic installer or change Outlook security settings.

The following check reads installed Python versions without installing anything. Open Command Prompt and run:

```bat
py -0p
```

If `py` is unavailable, Python may be absent or its launcher may not be installed. The assistant supports Python 3.11 or newer with Tkinter. Cloud validation used Python 3.12.14; Windows execution still needs your test.

## Stage 1: synthetic emails only

### 1. Copy the application to a local folder

**Permission checkpoint:** Approve copying/extracting the application before doing so. This creates an application folder and files; it does not modify Outlook or QuickBooks.

Extract the supplied ZIP into a folder such as `C:\BelsitoEmailAssistant`. `Belsito.cmd`, `belsito_mail`, `tests` and `requirements-outlook.txt` should be directly inside that folder. If the ZIP adds an outer folder, use the folder that actually contains `Belsito.cmd`.

Open Command Prompt and enter that application folder:

```bat
cd /d C:\BelsitoEmailAssistant
```

### 2. Install Python only if required

**Permission checkpoint:** If a suitable Python installation is missing, agree on the Python version and approve its installation before downloading or running an installer. Explain the selected version, installation location and options first.

If a new installation is needed, choose an official Python **3.13** Windows installer from [python.org](https://www.python.org/downloads/windows/) after agreeing on that change. Include Tcl/Tk and the Python launcher. Prefer a per-user installation when available. Python 3.12 security releases, including 3.12.14, may be source-only; do not expect a Windows installer for the exact cloud-tested patch version. The command example below uses an existing Python 3.12 installation. For Python 3.13, replace `-3.12` with `-3.13`.

After installation, reopen Command Prompt, return to the application folder, and run `py -0p` again to verify the version. If Windows requests administrator access or proposes unrelated components, stop and review that change before continuing.

### 3. Create the application's virtual environment

**Permission checkpoint:** Approve creating the `.venv` folder in the application directory. This makes a separate local Python environment. It does not install the Outlook adapter or change your existing Python packages.

Run:

```bat
py -3.12 -m venv .venv
.venv\Scripts\python.exe --version
```

No virtual-environment activation or administrator privileges are required. Stage 1 has no third-party dependency installation.

### 4. Run the synthetic checks

**Permission checkpoint:** Approve running the bundled program and tests. The command-line sample check uses a temporary local state database and workbook and removes those temporary files afterward. It does not start or read Outlook.

```bat
.venv\Scripts\python.exe -m unittest discover -s tests -v
.venv\Scripts\python.exe -m belsito_mail --sample-check
```

The expected outcome is passing tests and a message beginning `Sample check passed`. If a test fails, retain the output and stop before Outlook testing. Do not include real email content in an error report.

### 5. Launch and test the interface

**Permission checkpoint:** Approve the first launch. It creates the application's local state directory and SQLite database under `%LOCALAPPDATA%\BelsitoPlumbing\EmailAssistant`. This stores completion IDs and the sample-test acknowledgement, with no email content or passwords.

Double-click `Belsito.cmd`. The interface opens in **Sample emails (Stage 1)** mode on every launch. Click **Review My Emails**, then follow the [Stage 1 acceptance checklist](ACCEPTANCE_CHECKLIST.md) in order.

Each export asks you to choose a location. Approve the new file by clicking **Save**; Windows asks before overwriting an existing report. Use a local folder outside cloud-sync storage. Generating the `.xlsx` file does not require Excel; opening it in Excel is a separate user action.

If the launcher cannot find `.venv\Scripts\pythonw.exe`, return to step 3. For visible error output, run:

```bat
.venv\Scripts\python.exe -m belsito_mail
```

If Python reports that Tkinter is missing, review the Python installation's Tcl/Tk component before making any installation change.

### 6. Stop and confirm Stage 1 results

Do not access the actual mailbox until you have tested categories, priorities, details, completion/reopening and the Excel report. The **Sample checks passed — enable Stage 2** button becomes available after a sample review, a completion action and an export. Its confirmation asks you to acknowledge your inspection of those results. Enabling Stage 2 does not access Outlook or install anything.

The acknowledgement persists locally. The command-line sample check alone does not grant it.

## Stage 2: optional read-only Outlook test

### 7. Install the Outlook COM adapter

**Permission checkpoint:** Approve installing `pywin32==311` into this application's `.venv`. This downloads the dependency from the configured Python package index. It does not send email content to that index, store an email password, or modify Outlook/QuickBooks.

From the application folder, run:

```bat
.venv\Scripts\python.exe -m pip install -r requirements-outlook.txt
```

If installation fails, read the error before retrying. Do not disable TLS/package verification or run a global/post-install setup script to force it through.

### 8. Open the intended existing Outlook Classic profile

Open **Microsoft Outlook Classic** yourself and confirm that the intended Network Solutions POP mailbox is the default Inbox. The assistant attaches only to an already-running Outlook instance; it does not create an Outlook profile or ask for a mail password. The new Outlook application does not provide this COM interface.

The selected scope is **default Inbox and its subfolders**. Messages moved to folders elsewhere in a PST are outside this scope. Only locally available messages received within the last 30 days are eligible. Outlook itself can synchronize normally while open.

Before reading, inspect a few representative Outlook messages and record their unread/read status, folders and counts for the Stage 2 checklist. Do not change your POP account, default profile, Outlook trust settings or QuickBooks configuration as part of this step.

### 9. Grant permission for one read

Select **Outlook Classic (Stage 2)** and click **Review My Emails**. The application asks: **Permission to read Outlook**. Read the stated scope, then choose **Yes** only when you authorize that review. **No** leaves Outlook unread by the assistant. A new permission prompt appears for every Outlook review.

Outlook may show its own protected-field/security prompt. The assistant does not bypass it. Stop and review the prompt if unexpected; do not weaken Outlook security settings to make the review work.

The reader limits each review to 5,000 items, 1,000 folders and 100,000 body characters per message. It skips attachments. Any access failure or safety-limit warning means coverage may be incomplete. Complete the [Stage 2 acceptance checklist](ACCEPTANCE_CHECKLIST.md) before relying on daily reports.

If the intended Inbox cannot be read, confirm that Outlook Classic is fully open under your usual Windows account and uses the intended existing default profile, then retry after resolving that issue. Opening a different profile or using new Outlook does not establish that the adapter works with your POP mailbox.

## Daily use after both stages pass

Open Outlook Classic, launch `Belsito.cmd`, select Outlook mode, and click **Review My Emails**. Grant read permission, inspect the prioritized report, mark completed follow-ups locally, and export if needed. The report is created on demand; this version does not schedule background reviews or send responses.

The interface shows open items by default. Use **Show completed** to display completed items, or **Reopen selected** to return them to the open list. Exports include completed and open items with their status, irrespective of this display filter.

Email bodies exist only in the running application's memory. The local state database stores opaque completion IDs/source/timestamps and the test acknowledgement. Export files contain report fields including sender, subject and matching evidence but exclude full bodies. Keep those reports in your chosen local folder.

## Removing the assistant later

Removal is a separate change requiring your permission. After closing the app, deleting its application folder removes the program and `.venv`. Separately deleting `%LOCALAPPDATA%\BelsitoPlumbing\EmailAssistant` removes completion history and the test acknowledgement. Exported reports remain where you saved them. No Outlook mailbox or QuickBooks file is part of either removal path.
