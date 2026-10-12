# Acceptance checklist

Test one function at a time. Stop at a failed check and resolve it before proceeding. Automated tests use synthetic data and fake Outlook objects; Windows appearance, Excel compatibility and actual Outlook behavior still require your approval and inspection on your computer.

## Stage 1: sample-only interface

Install and launch only after the permission steps in [Windows setup](WINDOWS_SETUP.md). Keep **Sample emails (Stage 1)** selected. No actual Outlook mailbox is needed for these checks.

### 1. Launch and review

- [ ] `Belsito.cmd` opens the Belsito Plumbing window in sample mode.
- [ ] The button is labeled **Review My Emails**.
- [ ] Clicking it produces a report labeled **Synthetic sample emails** / **SAMPLE**.
- [ ] A selected row displays suggested action, evidence and synthetic message text.

### 2. Categories and exclusion checks

The samples contain 16 messages, including old/future and resolved examples. The normal result is 11 follow-up items. Compare individual examples:

| Sample subject | Expected result |
| --- | --- |
| Emergency: burst pipe at Example Building | Customer service/new job; Urgent; Immediate |
| New job: plumbing renovation | Customer service/new job |
| Follow-up on proposal BP-SAMPLE-14 | Estimate/proposal follow-up; High because older than 7 days |
| Estimate for water heater replacement | Estimate/proposal follow-up |
| Outstanding invoice SAMPLE-1042 | Outstanding invoice/collections; High |
| DOB notice: plumbing inspection | Regulatory notice; High due to near deadline |
| DEP compliance notice | Regulatory notice; High |
| FDNY inspection notice | Regulatory notice; High |
| OSHA compliance notice | Regulatory notice; High |
| Urgent new job and quote request | Customer service/new job; Estimate/proposal follow-up; Urgent; Immediate |
| Service request: routine appointment, not urgent | Customer service/new job; no Urgent category |
| Re: overdue invoice SAMPLE-900 | Excluded: sample says paid in full |
| Re: Emergency leaking pipe | Excluded: sample says repaired and job complete |
| Weekly plumbing industry newsletter | Excluded |
| Urgent service request from last month | Excluded: older than 30 days |
| Service request with a future timestamp | Excluded: future timestamp |

- [ ] The expected categories appear and the excluded examples are absent.
- [ ] Evidence corresponds to wording in the sample, rather than an unexplained decision.
- [ ] Regulatory flags mention DOB, DEP, FDNY and another agency through their samples.

### 3. Priorities and daily report

- [ ] Immediate items appear before High items, and High before Normal.
- [ ] Immediate and High rows have visible highlighting.
- [ ] Open/Immediate/High/Completed counts agree with the displayed results.
- [ ] Suggested actions give a useful starting point for each category.
- [ ] The report displays its review time and clearly identifies the sample source.

The report is a rule-based triage aid. It does not resolve threads, prove balances, authenticate notices or determine actual job status. Deadline recognition currently covers explicitly labeled `YYYY-MM-DD` dates; verify other dates yourself.

### 4. Complete and reopen an item

- [ ] Select one sample follow-up and click **Mark completed**.
- [ ] It disappears from the default open list and Completed increases.
- [ ] Select **Show completed**; the row is shown with status **Completed**.
- [ ] Close and reopen the app, review samples again, and verify that completion persists.
- [ ] Select the completed row and click **Reopen selected**; it returns to Open.
- [ ] Complete a sample item again so the completion action is part of this session's Stage 1 checks.

Completion changes affect only the local assistant database. They do not change an Outlook flag, read/unread status, category or message.

### 5. Excel export

- [ ] Click **Export to Excel** and save a `SAMPLE-...xlsx` report to an approved local folder.
- [ ] Open it in Excel, if available, and confirm that it opens without a repair warning.
- [ ] It contains a title, review time/source, filters and prioritized report rows.
- [ ] Columns include Status, Priority, Categories, Received, Sender, Email, Subject, Suggested action, Reason, Due date and Source.
- [ ] Both completed and open items appear with correct status.
- [ ] Full email bodies are absent from the workbook.
- [ ] A new export after reopening an item reflects its updated status.

Excel is optional for generating the report but useful for this visual acceptance check. If Excel is unavailable, retain the generated sample file and defer the Excel-opening check; do not claim it passed.

### 6. Acknowledge Stage 1, then stop

- [ ] During the current app session, review samples, complete an item and export a report.
- [ ] **Sample checks passed — enable Stage 2** becomes available.
- [ ] Click it only after verifying the preceding results, and read the acknowledgement dialog.
- [ ] Enabling Stage 2 does not read Outlook.

**Stop here.** Do not proceed to actual emails until Stage 1 is satisfactory and you separately approve installing the Stage 2 dependency and reading the mailbox. A passing command-line sample check does not replace your interface test.

## Stage 2: actual Outlook Classic, read-only

Proceed only after the separate setup permissions and Stage 1 acknowledgement. The scope you selected is the intended default **Inbox and its subfolders**, covering messages received in the last 30 days.

### 7. Permission and prerequisites

- [ ] You approved installing the pinned Outlook dependency into `.venv`.
- [ ] Existing Outlook Classic is open with the intended POP profile; new Outlook is not used.
- [ ] You inspect representative unread/read messages and record their original status, folder and folder counts yourself.
- [ ] Select Outlook mode and click **Review My Emails**; a read-permission dialog appears.
- [ ] Choose **No** once and verify that no review runs.
- [ ] Click again and choose **Yes** only when ready to grant permission for that read.

### 8. Compare the report with Outlook

- [ ] The report is clearly labeled **OUTLOOK** / **Outlook Inbox + subfolders**.
- [ ] Manually compare several service requests, estimate follow-ups, collection emails, regulatory messages and urgent messages with their reported categories and priorities.
- [ ] Verify an eligible message in an Inbox subfolder is included when its wording merits a flag.
- [ ] Messages older than 30 days are not included.
- [ ] The assistant reports any incomplete coverage warning; treat that report as partial and resolve the stated cause.

Not every recent message becomes a follow-up item. Compare unflagged messages too so you can identify missed wording. No attachment-only job, invoice or notice can be discovered by this version.

### 9. Inspect mailbox state after review

- [ ] Using Outlook yourself, compare the original messages' read/unread status, folders and counts with the observations from step 7.
- [ ] No email was moved, deleted, sent or edited by the assistant.
- [ ] Mark an Outlook follow-up completed in the assistant, review it with **Show completed**, then reopen it.
- [ ] Inspect that message again in Outlook; the completion action changed only the assistant's status.
- [ ] QuickBooks was not accessed or changed by the assistant.

Normal Outlook background synchronization can receive new messages during the test. Account for that when comparing counts; an unexplained state change requires investigation before relying on the adapter.

### 10. Export and repeat

- [ ] Export an Outlook report to an approved local folder and inspect its status, priorities and evidence.
- [ ] Confirm the report excludes full bodies while recognizing that subjects/senders/evidence are mailbox-derived content.
- [ ] Run another Outlook review; it asks for fresh permission.
- [ ] Close/reopen the app; it starts in sample mode and Outlook completion records persist only in the assistant.

## When a check fails

Do not bypass Outlook security prompts or alter trust settings. If Outlook is not open, open the existing Classic profile and retry. If a dependency is missing, resolve its installation only after approval. If the report warns about unreadable folders, body shortening or the 5,000-item/1,000-folder limits, do not claim complete coverage.

Report the failed step and safe error text without sharing customer email content. A packaged executable, rule customization, scheduling, external AI and QuickBooks integration require later stages and separate approval; they are not included in these acceptance checks.
