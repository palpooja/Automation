"""Finding H7: only the expected 'generate Tax Credit Certificate' confirmation may be accepted."""
import pytest

from engines import paye_engine, wht_engine

PAGE = """<html><body><script>
window.answers = [];
function ask(t) { window.answers.push(confirm(t)); }
</script></body></html>"""


def _ask(page, text):
    page.evaluate("t => ask(t)", text)
    return page.evaluate("() => window.answers[window.answers.length - 1]")


@pytest.mark.browser
@pytest.mark.parametrize("make_bot", [
    lambda pg: wht_engine.WhtBot(pg, lambda *_: None, lambda: None),
    lambda pg: paye_engine.PayeBot(pg, lambda *_: None, lambda: None),
], ids=["WHT", "PAYE"])
def test_unexpected_confirmation_is_refused_and_expected_one_accepted_only_while_generating(page, make_bot):
    page.set_content(PAGE)
    bot = make_bot(page)
    with bot.expecting_generate_confirm():
        assert _ask(page, "Are you sure you want to generate Tax Credit Certificate?") is True
        assert _ask(page, "Do you want to delete this certificate?") is False
    assert _ask(page, "Do you want to generate Tax Credit Certificate?") is False   # not generating now
    assert any("delete" in m.lower() for m in bot.dialogs)


def test_generate_tcc_options_are_off_unless_ticked():
    from modules.wht_receivable import WhtReceivableModule
    assert WhtReceivableModule.DEFAULT_GENERATE_WHT_TCC is False
    assert WhtReceivableModule.DEFAULT_GENERATE_PAYE_TCC is False
