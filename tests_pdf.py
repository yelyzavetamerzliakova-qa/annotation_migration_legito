"""Offline tests for PDF ingestion; makes NO Gemini API calls."""
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from docx import Document
from pdf_ingest import OCRTranscription, load_source_document
from migrate import SuggestedField, Extraction, convert_document

HERE = Path(__file__).parent


def test_docx_still_supported():
    data = (HERE / "samples" / "example_nda.docx").read_bytes()
    doc, info = load_source_document(data, "test.docx", "unused")
    assert info["input_format"] == "DOCX"
    assert any(p.text for p in doc.paragraphs)


def test_selectable_pdf():
    pdf_data = (HERE / "samples" / "example_contract.pdf").read_bytes()
    with patch("pdf_ingest.ocr_pdf_pages", side_effect=AssertionError("OCR must not be called")):
        doc, info = load_source_document(pdf_data, "original.pdf", "unused")
    assert info["input_format"] == "PDF"
    assert info["pdf_pages"] == 1
    assert info["ocr_pages"] == 0
    assert any("Northstar Labs" in p.text for p in doc.paragraphs)
    paras = list(doc.paragraphs)
    i = next(i for i, p in enumerate(paras) if "Northstar Labs" in p.text)
    proposals = Extraction(fields=[SuggestedField(
        paragraph_id=i, source_text="Northstar Labs",
        field_type="TextInput", label="Party A"
    )])
    accepted, skipped = convert_document(doc, proposals)
    assert len(accepted) == 1 and not skipped
    assert any("[TextInput: Party A]" in p.text for p in doc.paragraphs)
    buffer = BytesIO()
    doc.save(buffer)
    assert Document(BytesIO(buffer.getvalue()))


def test_scanned_pdf_mocked_ocr():
    pdf_data = (HERE / "samples" / "example_scanned_contract.pdf").read_bytes()
    mocked_ocr = OCRTranscription(pages=[
        "SERVICE AGREEMENT\nThis Agreement is between Northstar Labs and Riverside Solutions.\nEffective: 1 June 2027."
    ], review_notes=["Scan looks slightly faint"])
    with patch("pdf_ingest.ocr_pdf_pages", return_value=mocked_ocr) as mocked:
        doc, info = load_source_document(pdf_data, "scanned.pdf", "unused")
    assert mocked.called
    assert info["ocr_pages"] == 1
    assert "Scan looks slightly faint" in info["review_warnings"]
    assert any("Northstar Labs" in p.text for p in doc.paragraphs)


def test_invalid_pdf_rejected():
    try:
        load_source_document(b"not a pdf", "bad.pdf", "unused")
    except ValueError as exc:
        assert "valid PDF" in str(exc)
    else:
        raise AssertionError("Expected ValueError")


if __name__ == "__main__":
    test_docx_still_supported()
    test_selectable_pdf()
    test_scanned_pdf_mocked_ocr()
    test_invalid_pdf_rejected()
    print("PASS: DOCX, selectable PDF, scanned PDF (mocked OCR), annotated Word output, invalid PDF")
