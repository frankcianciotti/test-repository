"""Launch only the desktop app, or an explicitly requested synthetic smoke test."""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
import tempfile


def main() -> int:
    parser = argparse.ArgumentParser(description="Belsito Plumbing local email assistant")
    parser.add_argument("--sample-check", action="store_true", help="check synthetic data without Outlook")
    parser.add_argument("--export", type=Path, help="save the synthetic report when using --sample-check")
    args = parser.parse_args()
    if args.export and not args.sample_check:
        parser.error("--export requires --sample-check")
    if args.sample_check:
        from .classifier import review_messages
        from .excel_export import export_report
        from .sample_data import sample_messages
        from .state import LocalState

        now = datetime.now().astimezone()
        messages = sample_messages(now)
        items = review_messages(messages, now)
        if not items:
            raise RuntimeError("The synthetic sample produced no follow-up items.")
        with tempfile.TemporaryDirectory(prefix="belsito-sample-") as directory:
            state = LocalState(Path(directory) / "state.sqlite3")
            state.set_completed("sample", items[0].id, True)
            completed = state.completed_ids("sample")
            if items[0].id not in completed or state.completed_ids("outlook"):
                raise RuntimeError("Completion storage did not preserve isolated sample status.")
            target = args.export or Path(directory) / "sample-report.xlsx"
            export_report(items, completed, target, now, "Synthetic sample emails")
            if not target.is_file():
                raise RuntimeError("The sample report was not created.")
            state.close()
        print(f"Sample check passed: {len(messages)} messages, {len(items)} follow-up items; "
              "completion tracking and Excel export passed. Outlook was not accessed.")
        if args.export:
            print(f"Synthetic report: {args.export.resolve()}")
        return 0

    from .ui import launch
    return launch()


if __name__ == "__main__":
    raise SystemExit(main())
