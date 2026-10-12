"""Native desktop interface. Sample mode is the default on every launch."""

from __future__ import annotations

from collections import Counter
from datetime import datetime
from pathlib import Path
import queue
import sqlite3
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .classifier import review_messages
from .excel_export import export_report
from .models import FollowUpItem
from .sample_data import sample_messages
from .state import LocalState, default_state_path


class EmailAssistant:
    def __init__(self, root: tk.Tk, state: LocalState):
        self.root = root
        self.state = state
        self.items: list[FollowUpItem] = []
        self.source = "sample"
        self.reviewed_at: datetime | None = None
        self.review_warnings: tuple[str, ...] = ()
        self.busy = False
        self.sample_reviewed = False
        self.sample_completion_tested = False
        self.sample_exported = False
        self.events: queue.Queue = queue.Queue()
        self.mode = tk.StringVar(value="sample")
        self.show_completed = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="Start with synthetic emails. Outlook has not been accessed.")
        self.summary = tk.StringVar(value="No review yet.")
        self._build()
        self.root.protocol("WM_DELETE_WINDOW", self._close)
        self.root.after(100, self._poll)

    def _build(self) -> None:
        self.root.title("Belsito Plumbing — Email Follow-Up")
        self.root.geometry("1180x760")
        self.root.minsize(900, 620)
        style = ttk.Style(self.root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Treeview", rowheight=28, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        frame = ttk.Frame(self.root, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Belsito Plumbing", font=("Segoe UI", 22, "bold")).pack(anchor="w")
        ttk.Label(frame, text="Daily email follow-up • last 30 days • local processing", font=("Segoe UI", 11)).pack(anchor="w", pady=(0, 8))
        ttk.Label(frame, text="Stage 1: review samples, complete an item, and export. Verify the results before enabling Stage 2.", wraplength=1100).pack(anchor="w")
        controls = ttk.Frame(frame)
        controls.pack(fill="x", pady=10)
        self.sample_radio = ttk.Radiobutton(controls, text="Sample emails (Stage 1)", variable=self.mode, value="sample")
        self.sample_radio.pack(side="left")
        self.outlook_radio = ttk.Radiobutton(controls, text="Outlook Classic (Stage 2)", variable=self.mode, value="outlook", command=self._select_outlook)
        self.outlook_radio.pack(side="left", padx=10)
        self.review_button = ttk.Button(controls, text="Review My Emails", command=self._review)
        self.review_button.pack(side="right")
        self.stage_button = ttk.Button(controls, text="Sample checks passed — enable Stage 2", command=self._acknowledge_sample)
        self.stage_button.pack(side="right", padx=8)
        self.stage_button.state(["disabled"])
        ttk.Label(frame, textvariable=self.summary, font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 6))
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill="both", expand=True)
        columns = ("status", "priority", "categories", "received", "sender", "subject")
        self.table = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="extended")
        for name, heading, width in (
            ("status", "Status", 90), ("priority", "Priority", 95),
            ("categories", "Category", 225), ("received", "Received", 110),
            ("sender", "From", 150), ("subject", "Subject", 340),
        ):
            self.table.heading(name, text=heading)
            self.table.column(name, width=width, minwidth=60, stretch=name in ("categories", "subject"))
        vertical = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        horizontal = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.table.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        self.table.tag_configure("Immediate", background="#ffe1df")
        self.table.tag_configure("High", background="#fff0cb")
        self.table.tag_configure("completed", foreground="#666666")
        self.table.bind("<<TreeviewSelect>>", self._detail)
        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=8)
        self.complete_button = ttk.Button(actions, text="Mark completed", command=lambda: self._set_completed(True))
        self.complete_button.pack(side="left")
        self.reopen_button = ttk.Button(actions, text="Reopen selected", command=lambda: self._set_completed(False))
        self.reopen_button.pack(side="left", padx=8)
        ttk.Checkbutton(actions, text="Show completed", variable=self.show_completed, command=self._refresh).pack(side="left", padx=8)
        self.export_button = ttk.Button(actions, text="Export to Excel", command=self._export)
        self.export_button.pack(side="right")
        detail_frame = ttk.LabelFrame(frame, text="Selected item — suggested action, evidence, and local email text", padding=8)
        detail_frame.pack(fill="both")
        self.detail = tk.Text(detail_frame, height=9, wrap="word", font=("Segoe UI", 10), state="disabled")
        scroll = ttk.Scrollbar(detail_frame, command=self.detail.yview)
        self.detail.configure(yscrollcommand=scroll.set)
        self.detail.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        ttk.Label(frame, textvariable=self.status, wraplength=1100).pack(anchor="w", pady=(8, 2))
        ttk.Label(frame, text="Email decisions use rules and need your review. No external AI, email actions, or QuickBooks access.", wraplength=1100).pack(anchor="w")
        self._set_busy(False)

    def _select_outlook(self) -> None:
        if not self.state.sample_test_passed:
            self.mode.set("sample")
            messagebox.showinfo("Finish Stage 1 first", "Review the samples, mark an item completed, export to Excel, then click 'Sample checks passed'. Outlook has not been accessed.", parent=self.root)

    def _acknowledge_sample(self) -> None:
        if not (self.sample_reviewed and self.sample_completion_tested and self.sample_exported):
            return
        if not messagebox.askyesno("Confirm sample testing", "Have you verified the sample categories and priorities, completion tracking, and Excel report?\n\nThis enables the Outlook stage. It does not access Outlook. Each Outlook review will require separate permission.", parent=self.root):
            return
        try:
            self.state.acknowledge_sample_test()
        except (OSError, sqlite3.Error):
            messagebox.showerror("Could not save acknowledgement", "The local settings database could not be updated. Outlook remains disabled.", parent=self.root)
            return
        self.status.set("Sample testing acknowledged. Select Outlook Classic only when you are ready for Stage 2.")
        self.stage_button.state(["disabled"])

    def _review(self) -> None:
        if self.busy:
            return
        selected = self.mode.get()
        if selected == "outlook":
            if not self.state.sample_test_passed:
                self._select_outlook()
                return
            if not messagebox.askyesno("Permission to read Outlook", "Allow this review to read the last 30 days of locally stored messages in the default Outlook Classic Inbox and its subfolders?\n\nOpen Outlook Classic with the intended POP profile first. This app will not send, delete, move, mark as read, or modify messages; it will not access QuickBooks. It will keep processing local.\n\nThis review reads up to 5,000 messages and reports any incomplete coverage. Outlook itself may perform its usual background mail synchronization.", parent=self.root):
                return
        # Deadlines are calendar dates on the user's Windows clock. The reader
        # and classifier independently convert the rolling window to UTC.
        now = datetime.now().astimezone()
        self._set_busy(True)
        self.status.set("Reviewing synthetic sample emails…" if selected == "sample" else "Reading Outlook Inbox and subfolders locally…")
        # Read the database on the GUI thread; COM is initialized in the worker.
        sample_test_passed = self.state.sample_test_passed

        def work() -> None:
            try:
                warnings: tuple[str, ...] = ()
                if selected == "sample":
                    messages = sample_messages(now)
                else:
                    from .outlook_reader import ReadAuthorization, read_outlook_messages
                    result = read_outlook_messages(
                        authorization=ReadAuthorization(sample_test_passed, True),
                        now=now, include_subfolders=True,
                    )
                    messages = result.messages
                    warnings = result.warnings
                items = review_messages(messages, now)
                self.events.put(("success", selected, now, items, len(messages), warnings))
            except Exception as exc:
                # Never log the exception or mailbox content to disk.
                from .outlook_reader import OutlookReadError
                explanation = str(exc) if isinstance(exc, OutlookReadError) else "The review could not finish. No email action was performed. Check the installation and rerun the sample tests."
                self.events.put(("error", explanation))

        threading.Thread(target=work, daemon=True, name="local-mail-review").start()

    def _poll(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                self._set_busy(False)
                if event[0] == "error":
                    self.status.set("Review failed; the previous report has been retained.")
                    messagebox.showerror("Review could not finish", event[1], parent=self.root)
                    continue
                _, self.source, self.reviewed_at, self.items, count, warnings = event
                self.review_warnings = warnings
                if self.source == "sample":
                    self.sample_reviewed = True
                label = "SAMPLE" if self.source == "sample" else "OUTLOOK"
                completed_time = self.reviewed_at.astimezone().strftime("%Y-%m-%d %H:%M")
                self.status.set(f"{label} review at {completed_time}: read {count} messages; identified {len(self.items)} follow-up items." + (" INCOMPLETE COVERAGE — see warning." if warnings else ""))
                self._refresh()
                self._set_busy(False)
                if warnings:
                    messagebox.showwarning("Report has incomplete coverage", "\n".join(warnings) + "\n\nDo not treat this report as a complete mailbox review.", parent=self.root)
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        for widget in (self.review_button, self.sample_radio, self.outlook_radio):
            widget.state(["disabled"] if busy else ["!disabled"])
        for widget in (self.complete_button, self.reopen_button, self.export_button):
            widget.state(["disabled"] if busy or self.reviewed_at is None else ["!disabled"])
        may_acknowledge = (not busy and not self.state.sample_test_passed and self.sample_reviewed and self.sample_completion_tested and self.sample_exported)
        self.stage_button.state(["!disabled"] if may_acknowledge else ["disabled"])

    def _refresh(self) -> None:
        try:
            completed = self.state.completed_ids(self.source)
        except sqlite3.Error:
            messagebox.showerror("Completion status unavailable", "The local completion database could not be read. Restart after checking the app-data folder.", parent=self.root)
            return
        self.table.delete(*self.table.get_children())
        remaining = [item for item in self.items if item.id not in completed]
        counts = Counter(item.priority for item in remaining)
        source_label = "Synthetic sample emails" if self.source == "sample" else "Outlook Inbox + subfolders"
        self.summary.set(f"{source_label}  |  Open: {len(remaining)}  |  Immediate: {counts['Immediate']}  |  High: {counts['High']}  |  Completed: {len(self.items) - len(remaining)}")
        for item in self.items:
            done = item.id in completed
            if done and not self.show_completed.get():
                continue
            self.table.insert("", "end", iid=item.id, values=(
                "Completed" if done else "Open", item.priority, "; ".join(item.categories),
                item.message.received_at.astimezone().strftime("%Y-%m-%d"),
                item.message.sender_name or item.message.sender_email, item.message.subject,
            ), tags=("completed",) if done else (item.priority,))
        self._write_detail("")

    def _detail(self, _event=None) -> None:
        selected = self.table.selection()
        if not selected:
            return
        item = next((item for item in self.items if item.id == selected[0]), None)
        if item is None:
            return
        due = f"\nExplicit deadline: {item.due_date.isoformat()}" if item.due_date else ""
        self._write_detail(
            f"Suggested action: {item.action}{due}\n\nWhy flagged:\n" +
            "\n".join(f"• {reason}" for reason in item.reasons) +
            f"\n\nFrom: {item.message.sender_name} <{item.message.sender_email}>\n"
            f"Subject: {item.message.subject}\n\n{item.message.body}"
        )

    def _write_detail(self, text: str) -> None:
        self.detail.configure(state="normal")
        self.detail.delete("1.0", "end")
        self.detail.insert("1.0", text)
        self.detail.configure(state="disabled")

    def _set_completed(self, completed: bool) -> None:
        if self.busy:
            return
        selected = self.table.selection()
        if not selected:
            messagebox.showinfo("Select an item", "Select a follow-up item first.", parent=self.root)
            return
        try:
            for item_id in selected:
                self.state.set_completed(self.source, item_id, completed)
        except (OSError, sqlite3.Error):
            messagebox.showerror("Could not save status", "The local completion database could not be updated. Outlook was not modified.", parent=self.root)
            return
        if self.source == "sample" and completed:
            self.sample_completion_tested = True
        self._refresh()
        self._set_busy(False)

    def _export(self) -> None:
        if self.busy or self.reviewed_at is None:
            return
        prefix = "SAMPLE-" if self.source == "sample" else ""
        suggested = f"{prefix}Belsito-Follow-Up-{self.reviewed_at.astimezone():%Y-%m-%d}.xlsx"
        target = filedialog.asksaveasfilename(parent=self.root, title="Save local Excel report", initialfile=suggested, defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")], confirmoverwrite=True)
        if not target:
            return
        try:
            source_label = "Synthetic sample emails" if self.source == "sample" else "Outlook Inbox + subfolders"
            if self.review_warnings:
                source_label += " — INCOMPLETE COVERAGE: " + "; ".join(self.review_warnings)
            export_report(self.items, self.state.completed_ids(self.source), Path(target), self.reviewed_at, source_label)
        except (OSError, ValueError, sqlite3.Error):
            messagebox.showerror("Export could not finish", "The report could not be saved. If it is open in Excel, close it and try again. Check that the selected folder is writable.", parent=self.root)
            return
        if self.source == "sample":
            self.sample_exported = True
        self._set_busy(False)
        self.status.set("Excel report saved at your selected location. It contains report metadata, without full email bodies.")

    def _close(self) -> None:
        if self.busy:
            messagebox.showinfo("Review in progress", "Please let this review finish before closing the application.", parent=self.root)
            return
        self.items.clear()
        self.state.close()
        self.root.destroy()


def launch() -> int:
    root = tk.Tk()
    try:
        state = LocalState(default_state_path())
    except (OSError, sqlite3.Error):
        root.withdraw()
        messagebox.showerror("Local storage unavailable", "The assistant could not create its local completion database. No Outlook mailbox was accessed. Check permissions in your local app-data folder.", parent=root)
        root.destroy()
        return 1
    EmailAssistant(root, state)
    root.mainloop()
    return 0
