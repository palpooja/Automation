"""Speed changes must not change results (the user's rule: faster only with 100% the same output)."""
import time

import pytest

from engines import returns_pw

pytestmark = pytest.mark.browser

PAGE = """<html><body><script>
setTimeout(() => document.body.insertAdjacentHTML('beforeend',
  '<input type="button" id="printid" value="Print">'), {print_ms});
if ({sched_ms} >= 0) setTimeout(() => document.body.insertAdjacentHTML('beforeend',
  '<input type="button" id="printWthSchid" value="Print with Schedules">'), {sched_ms});
</script></body></html>"""


def _wait(page, print_ms, sched_ms):
    page.set_content(PAGE.format(print_ms=print_ms, sched_ms=sched_ms))
    t0 = time.monotonic()
    loaded, _, sched = returns_pw.wait_for_print_buttons(page, 20, lambda: None, lambda *_: None)
    return loaded, sched is not None, time.monotonic() - t0


def test_schedules_button_three_seconds_after_print_is_still_used(page):
    # the old loop could catch this (it watched 2.5-3.5 s after Print); the new one always does
    loaded, has_sched, _ = _wait(page, 300, 3300)
    assert loaded and has_sched


def test_both_buttons_together_are_taken_without_the_grace_wait(page):
    loaded, has_sched, took = _wait(page, 300, 300)
    assert loaded and has_sched and took < 1.5


def test_print_only_return_waits_the_full_grace_then_prints_plain(page):
    loaded, has_sched, took = _wait(page, 300, -1)
    assert loaded and not has_sched
    assert took >= returns_pw.SCHED_GRACE_S


def test_return_that_never_loads_is_not_saved(page):
    page.set_content("<html><body>loading...</body></html>")
    loaded, _, sched = returns_pw.wait_for_print_buttons(page, 2, lambda: None, lambda *_: None)
    assert loaded is False and sched is None
