"""Finding H8: nothing may be marked 'done' unless it was really downloaded."""
import datetime
import json
import os
import time

from engines import ledger_engine as eng
from engines import returns_pw


def _progress(folder, data):
    (folder / "_progress.json").write_text(json.dumps(data), encoding="utf-8")


def test_month_missing_from_the_ledger_on_an_earlier_day_is_checked_again(tmp_path):
    # written by v10: no date attached
    _progress(tmp_path, {"Value Added Tax": {"SEP-2026": "not in ledger",
                                             "AUG-2026": "no payment link"}})
    assert eng.already_done(tmp_path, "Value Added Tax", "SEP-2026") is False
    assert eng.already_done(tmp_path, "Value Added Tax", "AUG-2026") is False


def test_month_missing_from_the_ledger_today_is_not_checked_twice_on_the_same_day(tmp_path):
    # keeps a same-day re-run (Resume) as fast as before
    eng.mark_progress(tmp_path, "Value Added Tax", "SEP-2026", "not in ledger")
    assert eng.already_done(tmp_path, "Value Added Tax", "SEP-2026") is True


def test_saved_month_stays_done(tmp_path):
    eng.mark_progress(tmp_path, "Value Added Tax", "JUL-2026", "saved")
    assert eng.already_done(tmp_path, "Value Added Tax", "JUL-2026") is True


def test_failed_month_is_retried(tmp_path):
    eng.mark_progress(tmp_path, "Value Added Tax", "JUL-2026", "failed")
    assert eng.already_done(tmp_path, "Value Added Tax", "JUL-2026") is False


def test_level2_saved_on_an_earlier_day_needs_refreshing(tmp_path):
    f = tmp_path / "Value Added Tax_L2.xlsx"
    f.write_bytes(b"x" * 2000)
    yesterday = time.time() - 86400 * 1.5
    os.utime(f, (yesterday, yesterday))
    assert eng.level2_is_current(f) is False
    os.utime(f, None)                                   # saved now
    assert eng.level2_is_current(f) is True
    assert eng.level2_is_current(tmp_path / "missing.xlsx") is False


def test_too_small_return_file_does_not_count_as_already_saved(tmp_path):
    good = tmp_path / "RT0012345678AB 01-07-2025 to 31-07-2025.pdf"
    good.write_bytes(b"%PDF" + b"x" * 5000)
    tiny = tmp_path / "RT0099999999ZZ 01-08-2025 to 31-08-2025.pdf"
    tiny.write_bytes(b"%PDF-1.4 tiny")
    empty_web = tmp_path / "RT0077777777YY 01-09-2025 to 30-09-2025.mhtml"
    empty_web.write_bytes(b"")
    have = returns_pw.existing_downloads(str(tmp_path))
    assert returns_pw.already_saved(have, good.stem, "RT0012345678AB") is True
    assert returns_pw.already_saved(have, tiny.stem, "RT0099999999ZZ") is False
    assert returns_pw.already_saved(have, empty_web.stem, "RT0077777777YY") is False


def test_progress_file_is_written_whole_or_not_at_all(tmp_path, monkeypatch):
    eng.mark_progress(tmp_path, "PAYE", "JUL-2026", "saved")
    assert not list(tmp_path.glob("*.tmp")), "temporary file left behind"
    data = json.loads((tmp_path / "_progress.json").read_text(encoding="utf-8"))
    assert data["PAYE"]["JUL-2026"].startswith("saved")
    _ = datetime  # (date stamps are checked through already_done above)
