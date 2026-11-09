# Legito AI Migration — DOCX + PDF OCR upgrade

An upgrade for the **existing** Streamlit + Python + Gemini academic prototype.

## What is new?

- Accept `.docx` **and** `.pdf` uploads (up to 10 MB).
- PDF input up to **10 pages**.
- For searchable PDFs, `pypdf` extracts selectable text locally.
- For scanned/image-only pages (or almost-empty pages), Gemini Vision OCR transcribes only those pages.
- A reconstructed editable Word document goes through the **same existing** Gemini field-identification and Legito-annotation code.
- The user downloads `*_legito.docx` and a JSON review report including how many pages needed OCR.
- No direct Legito API import. Human review remains necessary.

**Important limitation:** Unlike the original DOCX workflow, **PDF → Word reconstruction is not layout-preserving**. Text flow, headers, signatures, table structure, images, columns, stamps, and page breaks may change or be lost. OCR can misread legal text. Do not rely on this output for real legal documents without careful verification. Use fictional files for demos.

## Files to update in your EXISTING repo

Replace these three existing files, and add one new file:

1. `streamlit_app.py` (replace)
2. `migrate.py` (replace)
3. `requirements.txt` (replace)
4. `pdf_ingest.py` (**add**) — PDF + OCR logic

`tests_pdf.py` and the `samples` directory are optional but recommended for testing.

**Do not replace or upload `.env`, or `.streamlit/secrets.toml`.** Do not copy real API keys into any source file. The `.env.example` in this package has only a dummy placeholder.

## Local test on your Mac

```bash
cd ~/Desktop/legito_gemini_starter
source ~/.venv/bin/activate
python -m pip install -r requirements.txt
python tests.py
python tests_pdf.py
python -m streamlit run streamlit_app.py
```

Upload `samples/example_contract.pdf` (searchable) or `samples/example_scanned_contract.pdf` (image-only OCR) or your existing `samples/example_nda.docx`.

The `tests_pdf.py` scan test **mocks OCR**: it verifies the conversion pipeline but does **not** prove a live Gemini OCR API call. For that, upload the scanned sample through your running Streamlit app using a currently available Gemini model.

The CLI also supports PDFs now:

```bash
python migrate.py samples/example_contract.pdf --model YOUR_AVAILABLE_GEMINI_MODEL
python migrate.py samples/example_scanned_contract.pdf --model YOUR_AVAILABLE_GEMINI_MODEL
```

## Update GitHub

```bash
cd ~/Desktop/legito_gemini_starter
git switch streamlit-prototype
git status --short
# Ensure your real .env is not tracked. Check before staging.
git add streamlit_app.py migrate.py pdf_ingest.py requirements.txt tests_pdf.py samples/
git diff --cached --name-only
# Review staged files: NEVER include .env or real credentials.
git commit -m "Add PDF upload and Gemini OCR support"
git push origin streamlit-prototype
```

Streamlit Community Cloud normally redeploys changes pushed to the deployed branch. Your Gemini key stays in the hosting **Secrets** settings; no need to add it to GitHub again. If Streamlit fails to import `pypdf`, verify you pushed the new `requirements.txt`.

## Workflow design

```text
DOCX → python-docx parsing -----------------┐
                                          ├→ Gemini field extraction → annotated DOCX → manual Legito import
PDF  → pypdf selectable text ──┐           │
       Gemini Vision OCR if scan └→ rebuild Word ┘
```

PDF OCR and annotation can involve **two Gemini API requests**, and failures such as 429/503 are possible. Models available to your API key may change; use the exact working model ID shown in Google AI Studio. PDF content is sent to Google's Gemini service for OCR when needed, and extracted contract text is sent for annotation.
