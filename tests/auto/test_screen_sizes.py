"""'Make the GUI different-resolution friendly' (user, 6 Oct 2026).  The real window is started in
virtual screens of common sizes; with Windows display scaling 150% a 1920x1080 screen behaves like
1280x720 (here: DPI 144).  Version 10.9 opened at a fixed 1240x820, never smaller than 1000x700."""
import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

PROBE = Path(__file__).with_name("window_probe.py")
SCREENS = [(1366, 768, 96), (1280, 720, 96), (1920, 1080, 144), (1024, 768, 96), (2560, 1440, 96)]


def probe(tmp_path, w, h, dpi, clients=3):
    if not shutil.which("xvfb-run"):
        pytest.skip("needs xvfb-run")
    folder = tmp_path / f"{w}x{h}_{dpi}"
    folder.mkdir()
    r = subprocess.run(["xvfb-run", "-a", "-s", f"-screen 0 {w}x{h}x24 -dpi {dpi}", sys.executable,
                        str(PROBE), str(folder), str(clients)], capture_output=True, text=True, timeout=120)
    line = next((x for x in r.stdout.splitlines() if x.startswith("PROBE")), None)
    assert line, r.stdout[-800:] + r.stderr[-1500:]
    return json.loads(line[5:])


@pytest.mark.parametrize("w,h,dpi", SCREENS)
def test_window_fits_the_screen_and_everything_can_be_reached(tmp_path, w, h, dpi):
    m = probe(tmp_path, w, h, dpi)
    x, y, ww, wh = m["window"]
    assert ww <= w and wh <= h, f"window {ww}x{wh} larger than the screen {w}x{h}"
    assert m["minsize"][0] <= w and m["minsize"][1] <= h, "minimum size larger than the screen"
    for name in ("run", "stop"):
        bx, by, bw, bh, shown = m[name]
        assert shown and bx >= x and bx + bw <= x + ww, f"{name} button off the side of the window"
    assert m["scroll"], "the main area scrolls when the screen is short"
    run_bottom = m["run"][1] + m["run"][3]
    if run_bottom > y + wh:
        assert m["scrollbar_shown"], "button below the window and no scroll bar to reach it"
    if m["content_width"] > m["view_width"] + 1:
        assert m["hscroll_shown"], "an option panel wider than the window and no sideways scroll bar"
    if w >= 1280:
        assert not m["hscroll_shown"], "no sideways scrolling on screens 1280 px wide or more"


def test_text_is_scaled_on_a_scaled_screen(tmp_path):
    normal = probe(tmp_path, 1920, 1080, 96)
    scaled = probe(tmp_path, 1920, 1080, 144)
    assert scaled["run"][3] > normal["run"][3] * 1.2, "buttons and text grow with Windows scaling"


def test_run_list_on_the_left_stays_short_with_many_clients(tmp_path):
    m = probe(tmp_path, 1366, 768, 96, clients=60)
    assert m["queue_rows"] <= 14, f"{m['queue_rows']} rows in the left 'This run' list for 60 clients"
