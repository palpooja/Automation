"""Return PDFs must be machine-readable and complete (user, 4 Oct 2026: 'the old program made neat,
machine-readable PDFs').  Checked with the same tools a reader or a data-extraction program would use."""
import subprocess

import pytest
from pypdf import PdfReader

from engines import returns_pw

pytestmark = pytest.mark.browser


def _pdf(page, url, tmp_path, name):
    page.goto(url)
    page.wait_for_timeout(900)
    page.evaluate(returns_pw.UNBOX_JS)
    out = str(tmp_path / f"{name}.pdf")
    returns_pw._print_to_pdf(page, out)
    text = subprocess.run(["pdftotext", "-layout", out, "-"], capture_output=True, text=True).stdout \
        if subprocess.run(["which", "pdftotext"], capture_output=True).returncode == 0 else \
        " ".join(p.extract_text() or "" for p in PdfReader(out).pages)
    return out, text


def test_every_column_of_a_wide_schedule_is_in_the_pdf(page, mock_url, tmp_path):
    out, text = _pdf(page, f"{mock_url}/Portal/wide_return.html", tmp_path, "wide")
    words = text.split()
    missing = [i for i in range(1, 25) if f"{i * 1000:,}.00" not in words]   # one value per column
    assert missing == [], f"columns missing from the PDF: {missing}"
    assert "LASTLINE" in text


def test_a_normal_return_stays_a4_portrait_at_full_size(page, mock_url, tmp_path):
    out, _ = _pdf(page, f"{mock_url}/Portal/return.html?id=R-VAT-01", tmp_path, "normal")
    box = PdfReader(out).pages[0].mediabox
    assert (round(float(box.width)), round(float(box.height))) == (596, 842)     # A4 portrait


def test_the_print_buttons_are_not_printed(page, mock_url, tmp_path):
    _, text = _pdf(page, f"{mock_url}/Portal/return.html?id=R-VAT-01", tmp_path, "buttons")
    assert "Print with Schedules" not in text and "Print" not in text.split()
    assert "Return R-VAT-01" in text


def test_text_is_real_text_with_embedded_fonts_not_pictures(page, mock_url, tmp_path):
    out, text = _pdf(page, f"{mock_url}/Portal/return.html?id=R-VAT-01", tmp_path, "readable")
    reader = PdfReader(out)
    assert sum(len(p.images) for p in reader.pages) == 0
    assert "Wide content" in text
