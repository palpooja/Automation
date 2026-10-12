"""Findings M1 (screenshots and page copies kept forever) and M2 (error handler that can crash)."""
import os
import time

from core import store
from engines import ledger_engine as eng


def test_old_files_in_the_log_folder_are_deleted_new_ones_kept(fake_data_dir):
    logs = store.LOG_DIR
    (logs / "page_captures_1000000001_20260101_000000").mkdir(parents=True)
    old = [logs / "error_x.png", logs / "run_20260101.txt",
           logs / "page_captures_1000000001_20260101_000000" / "01_screen.png"]
    new = logs / "run_today.txt"
    for f in old + [new]:
        f.write_text("x")
    long_ago = time.time() - 40 * 86400
    for f in old + [logs / "page_captures_1000000001_20260101_000000"]:
        os.utime(f, (long_ago, long_ago))
    removed = store.prune_logs(days=30)
    assert removed >= 3
    assert new.exists()
    assert not any(f.exists() for f in old)


def test_retention_of_zero_days_keeps_everything(fake_data_dir):
    store.LOG_DIR.mkdir(parents=True)
    f = store.LOG_DIR / "run_old.txt"
    f.write_text("x")
    os.utime(f, (0, 0))
    assert store.prune_logs(days=0) == 0 and f.exists()


class _DeadPage:
    def screenshot(self, **kw):
        raise RuntimeError("Target page, context or browser has been closed")


def test_screenshot_after_the_browser_died_does_not_raise(tmp_path, monkeypatch):
    monkeypatch.setattr(eng, "WORK_DIR", str(tmp_path))
    assert eng.safe_screenshot(_DeadPage(), "error_VAT_1.png") is None


def test_error_screenshots_carry_the_client_tin_and_an_absolute_folder(tmp_path, monkeypatch):
    monkeypatch.setattr(eng, "WORK_DIR", str(tmp_path))
    monkeypatch.setattr(eng, "CURRENT_TIN", "1000000001")
    p = eng.work_file("error_VAT_1.png")
    assert p.parent == tmp_path and "1000000001" in p.name
