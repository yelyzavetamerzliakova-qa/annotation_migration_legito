# Legito + Gemini — beginner prototype

The workflow: **fictional Word contract (.docx) -> Python -> Gemini API -> annotated Word (.docx) -> manual Legito import and review**.

## 1. Requirements

- Python 3.11 or later.
- A Gemini API key from https://aistudio.google.com/api-keys
- Permission to import templates in a Legito workspace.
- Use **invented / non-confidential contracts only** in this prototype.

## 2. Open a terminal in this folder

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
cp .env.example .env
```

### Windows PowerShell

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
Copy-Item .env.example .env
```

If PowerShell blocks activation, you can instead use `.venv\Scripts\python.exe` in place of `python` for commands below.

## 3. Add your API key

Open `.env` in a text editor. Replace `PASTE_YOUR_OWN_API_KEY_HERE` with your Gemini API key; do not include quotes. Never share `.env` or commit it to GitHub. The file is ignored by `.gitignore`.

## 4. Run it

```bash
python migrate.py samples/example_nda.docx
```

To choose an explicitly supported Gemini model from your Google AI Studio project:

```bash
python migrate.py samples/example_nda.docx --model gemini-flash-latest
```

Note: Model IDs and access can change. If Google returns a model-not-found or access error, select an available model in AI Studio and pass its exact model ID using `--model`.

## 5. Check the output

- `outputs/example_nda_legito.docx`: Word file with Legito import annotations such as `[TextInput: Client name]`, `[Date]` and `[Money]`.
- `outputs/example_nda_review.json`: proposed field mapping, skipped fields, review notes and elapsed time.
- Compare the entire output to the original and fix every mistaken proposed field before import. Treat this as an unapproved draft.

## 6. Import into Legito

1. Open your Legito Workspace with rights to edit templates.
2. Open or create a Template Suite and launch **Template Editor**.
3. Choose **Import from Word**, select the generated DOCX, and review import settings.
4. Test the resulting TextInput / Date / Money elements and correct errors in Template Editor.
5. For a fair comparison against Legito native AI, test the original DOCX separately with **Apply AI conversion** and measure its time / correction effort.

Official annotations reference: https://www.legito.com/knowledge-base/import-from-word/

## 7. What is and is not automated

Automated: Word text read, Gemini JSON suggestions, validated field insertion, generated DOCX and review report.

Not automated: direct Legito API import, actual document review, condition/branch creation, OCR of scanned PDFs, bulk processing.

For complex Word files this prototype is incomplete: text in headers/footers, textboxes, images, embedded content, hyperlinks/track changes and nested tables is not fully supported. It rejects ambiguous/unsupported replacements rather than silently modifying them. Paragraph formatting is mostly preserved but **must be reviewed** before using the output.

Crucial business-case baseline: Legito itself already offers AI-powered conversion of Word templates. Compare manual vs Legito native AI vs our Gemini prototype before claiming any productivity advantage.
