"""User, 8 Oct 2026: 'why are we not downloading IT-WHT?' - the Withholding Tax Return was supported
(Headquarter -> Both box, Print with Schedules) but not ticked by default, so it was never downloaded."""
from engines import returns_common as rc


def test_withholding_tax_return_is_downloaded_by_default():
    assert "Withholding Tax Return" in rc.DEFAULT_TICKED
    assert "Withholding Tax Return" in rc.ALL_CATEGORIES
    assert "WTHDRT" in rc.BRANCH_CODES, "Headquarter -> Both is set for it, like VAT and PAYE"


def test_every_default_is_a_portal_return_type():
    assert rc.DEFAULT_TICKED <= set(rc.ALL_CATEGORIES)
