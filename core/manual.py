"""
core/manual.py - 'Open portal (manual)': log in, then hand the browser to the
person and keep it open until they press 'Close browser'.

While it is open:
  * record with Chrome's own Recorder (F12 > Recorder) - the login is already
    done, so the recording does not contain the password;
  * press 'Save this page for Claude' on any screen: every frame's HTML as the
    browser currently shows it + a screenshot + the frame list are saved, and
    on closing everything is zipped into the log folder.
Nothing is clicked or downloaded by the program in this mode.
"""
import datetime
import re
import zipfile

from core.runner import Stopped


def _safe(name):
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", name)[:60] or "frame"


def capture(page, folder, n):
    """Save the current screen: every frame's live HTML, a screenshot, the frame list."""
    folder.mkdir(parents=True, exist_ok=True)
    saved = []
    lines = []
    for i, fr in enumerate(page.frames):
        url = fr.url or ""
        lines.append(f"frame {i}: name={fr.name!r} url={url}")
        if url.startswith(("about:", "chrome-error:")):
            continue
        try:
            html = fr.evaluate("() => document.documentElement.outerHTML")
        except Exception as e:
            lines.append(f"   (could not read: {type(e).__name__})")
            continue
        tail = _safe(url.split("?")[0].rstrip("/").split("/")[-1]).replace(".html", "").replace(".htm", "")
        p = folder / f"{n:02d}_frame{i}_{tail}.html"
        p.write_text(html, encoding="utf-8")
        saved.append(p)
    (folder / f"{n:02d}_frames.txt").write_text("\n".join(lines), encoding="utf-8")
    try:
        page.screenshot(path=str(folder / f"{n:02d}_screen.png"), full_page=True)
    except Exception:
        pass
    return saved


def run_manual(ctx, client, open_session):
    stamp = f"{datetime.datetime.now():%Y%m%d_%H%M%S}"
    folder = ctx.log_path.parent / f"page_captures_{client['login']}_{stamp}"
    n = 0
    with open_session(ctx, client) as page:
        # The person's browser should behave like normal Chrome: the 'faster loading' filter only
        # answers while the program works with the browser, which it does not do here (6 Oct 2026:
        # nothing clickable after login).  The waits below also keep the browser serviced.
        try:
            page.context.unroute("**/*")
        except Exception:
            pass

        def keep_serviced():
            try:
                page.wait_for_timeout(50)
            except Exception:
                pass
            return False
        ctx.log("The browser is yours. Use the portal normally. To record your steps, press F12 "
                "in the browser, open 'Recorder' and start a recording. On any screen you want "
                "me to see, press 'Save this page for Claude'.", "ok")
        while True:
            ans = ctx.ask("The browser is yours - nothing is automated now. Press 'Save this page "
                          "for Claude' on each screen you want to keep, then 'Close browser'.",
                          "manual", ("", ""), poll=keep_serviced)
            if ans == "capture":
                n += 1
                saved = capture(page, folder, n)
                ctx.log(f"Saved screen {n}: {len(saved)} frame(s) + screenshot.", "ok")
                continue
            break
    if n:
        zpath = folder.with_suffix(".zip")
        with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
            for f in sorted(folder.iterdir()):
                z.write(f, f.name)
        ctx.log(f"Page captures: {zpath}  - send this file to Claude. It contains what was on "
                "screen (client data), but no password.", "ok")
        return zpath
    if ctx.stop_event.is_set():
        raise Stopped()
    return None
