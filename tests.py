"""Local validation; does not call Gemini API."""
from docx import Document
from migrate import Extraction, SuggestedField, convert_document


def run_test():
    d = Document()
    p = d.add_paragraph()
    p.add_run("Agreement between ")
    p.add_run("Northstar Labs s.r.o.").bold = True
    p.add_run(" and Riverbend Systems a.s. on 4 November 2026.")
    original_other = "Confidential information remains protected for 12 months."
    d.add_paragraph(original_other)
    result = Extraction(fields=[
        SuggestedField(paragraph_id=0, source_text="Northstar Labs s.r.o.", field_type="TextInput", label="Disclosing party"),
        SuggestedField(paragraph_id=0, source_text="4 November 2026", field_type="Date", label="Agreement date"),
        SuggestedField(paragraph_id=0, source_text="Unknown Acme", field_type="TextInput", label="Should skip"),
    ])
    accepted, skipped = convert_document(d, result)
    assert len(accepted) == 2
    assert len(skipped) == 1
    assert d.paragraphs[0].text == "Agreement between [TextInput: Disclosing party] and Riverbend Systems a.s. on [Date]."
    assert d.paragraphs[1].text == original_other
    assert all(not ("[TextInput:" in r.text and "[Date]" in r.text) for r in d.paragraphs[0].runs)
    print("PASS: annotations inserted, other legal text preserved, bad proposal skipped")


if __name__ == "__main__":
    run_test()
