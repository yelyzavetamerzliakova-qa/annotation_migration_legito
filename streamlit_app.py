"""Legito-inspired UI for the existing Word -> Gemini -> annotated DOCX demo.

Keeps migrate.py's extraction, replacement validation, and document writing
unchanged. Legito import and legal review are manual.
"""

from __future__ import annotations

from hashlib import sha256
from io import BytesIO
import json
import os
from pathlib import Path
import time

import streamlit as st
from docx import Document
from dotenv import load_dotenv

from migrate import ask_gemini, convert_document, paragraphs_in_document

BASE_DIR = Path(__file__).resolve().parent
MAX_DOC_BYTES = 10 * 1024 * 1024
DEFAULT_MODEL = "gemini-3.5-flash-lite"  # Keep the working app's default.

st.set_page_config(
    page_title="Document Migration | AI Prototype",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

load_dotenv(BASE_DIR / ".env")
try:
    cloud_key = st.secrets.get("GEMINI_API_KEY", "")
except Exception:
    cloud_key = ""
if cloud_key:
    os.environ["GEMINI_API_KEY"] = str(cloud_key)

# Palette and proportions inspired by Legito's light, green-accented website.
# Brand-inspired prototype, not an official Legito product or endorsed app.
st.markdown(
    """
    <style>
    :root {
        --legito-green: #78b66b;
        --legito-green-deep: #487f43;
        --ink: #303940;
        --muted: #68727b;
        --line: #e4e9e4;
        --surface: #ffffff;
        --canvas: #f8faf8;
    }

    html, body, [class*="css"], [data-testid="stAppViewContainer"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
        color: var(--ink);
    }
    [data-testid="stAppViewContainer"] {
        background: var(--canvas);
    }
    [data-testid="stHeader"] {
        background: transparent;
    }
    #MainMenu, footer { visibility: hidden; }
    .block-container {
        max-width: 1120px !important;
        padding-top: 1.25rem !important;
        padding-bottom: 3.3rem !important;
    }
    [data-testid="stHorizontalBlock"] { align-items: start; }
    .app-nav {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 7px 0 25px;
        border-bottom: 1px solid var(--line);
        margin-bottom: 46px;
        gap: 14px;
    }
    .brand-group { display:flex; align-items:center; gap: 11px; }
    .brand-icon {
        width: 35px; height: 35px; border: 2px solid #96b990;
        border-radius: 8px; display:flex; align-items:center; justify-content:center;
        background: white;
    }
    .brand-icon svg { display:block; width:21px; height:21px; }
    .brand-name { color: #31393f; font-size: 22px; font-weight: 760; letter-spacing: -.8px; line-height: 1; }
    .brand-suffix { border-left: 1px solid #dce3db; margin-left: 4px; padding-left: 15px; font-weight: 550; color: #71806f; font-size: 12px; }
    .app-pill { color: #66835f; background: #edf5eb; font-size: 11px; font-weight: 750;
        border: 1px solid #dcebd7; border-radius: 24px; padding: 7px 12px; white-space:nowrap;
        letter-spacing:.5px; }
    .overline { color: var(--legito-green-deep); text-transform:uppercase; font-weight: 760;
        letter-spacing: 1.9px; font-size: 11px; margin-bottom: 14px; }
    .hero-headline { font-weight: 770; font-size: clamp(35px,4vw,57px); letter-spacing:-2px;
        line-height: 1.12; color:#30393e; margin: 0 0 15px; }
    .hero-headline em { color: var(--legito-green-deep); font-style:normal; }
    .hero-subtitle { color: #747e80; font-size: 17px; line-height: 1.8; margin: 0 0 24px; max-width:590px; }
    .hero-benefits { display:flex; gap: 9px; flex-wrap:wrap; margin-bottom: 12px; }
    .hero-chip { display:inline-block; border:1px solid #dde6da; background:#fff;
        color: #66716b; border-radius: 8px; font-size:12px; padding: 7px 11px; font-weight:550; }
    .workflow-card { border: 1px solid #e1e7df; background: #fff; border-radius: 17px;
        padding: 26px 27px 26px; box-shadow:0 12px 35px rgba(45,65,42,.035); margin: 8px 0 16px; }
    .workflow-heading { font-size:12px; letter-spacing:1px; color:#6a7b69; font-weight:770;
        text-transform:uppercase; margin-bottom: 22px; }
    .workflow-step { display:flex; align-items:start; gap:14px; padding-bottom:17px; }
    .workflow-step:last-child { padding-bottom: 0; }
    .step-number { width: 28px; height: 28px; flex:0 0 28px; border-radius: 9px;
        background: #eaf5e7; color: #54834a; display:flex; align-items:center;
        justify-content:center; font-size: 12px; font-weight:800; }
    .step-text { color:#36413a; font-weight:700; font-size:14px; margin:1px 0 3px; }
    .step-desc { color:#818c81; font-size:12px; line-height:1.55; }
    .section-divider { border:none; height:1px; background:var(--line); margin:35px 0 28px; }
    .section-title { font-size:24px; font-weight:740; color:var(--ink); letter-spacing:-.55px;
        margin:0 0 6px; }
    .section-intro { color: #77827a; margin:0 0 21px; font-size:14px; }
    [data-testid="stFileUploader"] > label { font-weight: 650; color:#34423a; }
    [data-testid="stFileUploaderDropzone"] {
        border: 1.5px dashed #a9c9a4 !important;
        border-radius: 14px !important;
        background: #ffffff !important;
        padding: 24px 12px !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        border-color: #7cac72 !important;
        background: #f7fbf5 !important;
    }
    [data-testid="stFileUploaderDropzone"] button {
        border:1px solid #a7caa0 !important; background:#eff8ed !important;
        color:#3d773a !important; border-radius: 8px !important;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] span { color: #64766a !important; }
    div[data-testid="stButton"] button[kind="primary"],
    div[data-testid="stDownloadButton"] button[kind="primary"] {
        background: #79b96a !important; color:#ffffff !important;
        border: 1px solid #79b96a !important; border-radius: 9px !important;
        min-height: 48px !important; font-weight:720 !important;
        box-shadow:none !important;
    }
    div[data-testid="stButton"] button[kind="primary"]:hover,
    div[data-testid="stDownloadButton"] button[kind="primary"]:hover {
        background: #599e4e !important; border-color: #599e4e !important;
    }
    div[data-testid="stDownloadButton"] button {
        border-radius:9px !important; font-weight:660 !important;
        min-height:48px !important;
    }
    div[data-testid="stButton"] button:disabled { opacity: .6; }
    [data-testid="stExpander"] { border: 1px solid var(--line) !important;
        background: #ffffff; border-radius: 12px !important; }
    [data-testid="stMetric"] { padding: 10px 0 16px; }
    [data-testid="stMetricValue"] { color:#4c8546; font-size:30px; }
    .note-card { background: #f0f6ee; border: 1px solid #dae8d6; border-radius: 12px;
        padding: 18px 19px; margin:20px 0 0; font-size:13px; color:#597057; line-height:1.75; }
    .note-card strong { color: #416e40; }
    .success-head { background: #eff8ec; border:1px solid #d3e7cf; border-radius:14px;
        padding:19px 23px; color:#3e7540; font-size: 17px; font-weight: 750;
        margin: 0 0 14px; }
    .success-head span { display:block; font-size:13px; color:#658369;
        font-weight:450; margin-top:5px; }
    .footer-note { border-top:1px solid var(--line); margin-top:48px; padding-top:20px;
        color:#8c948c; font-size:12px; line-height:1.6; }
    @media (max-width: 680px) {
      .block-container { padding: 15px 16px 35px !important; }
      .brand-suffix { display:none; }
      .app-nav { margin-bottom:25px; }
      .hero-headline { font-size: 34px; letter-spacing:-1px; }
      .hero-subtitle {font-size:15px;}
      .workflow-card {padding:18px;}
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="app-nav">
        <div class="brand-group">
            <div class="brand-icon" aria-hidden="true">
                <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M4 3.8L18.8 12L4 20.2V3.8Z" stroke="#78B66B" stroke-width="2.4" stroke-linejoin="round"/>
                  <path d="M10 8L18.8 12L10 16" stroke="#426E42" stroke-width="1.7" stroke-linecap="round"/>
                </svg>
            </div>
            <span class="brand-name">Legito</span>
            <span class="brand-suffix">AI Migration</span>
        </div>
        <div class="app-pill">INDEPENDENT PROTOTYPE</div>
    </div>
    """,
    unsafe_allow_html=True,
)

hero_left, hero_right = st.columns([1.45, 1], gap="large")
with hero_left:
    st.markdown(
        """
        <div class="overline">AI-POWERED DOCUMENT AUTOMATION</div>
        <h1 class="hero-headline">From old contracts<br>to <em>smart templates.</em></h1>
        <p class="hero-subtitle">Give your legacy Word documents a head start.
        Our prototype detects reusable fields with Gemini and prepares a Word file
        for review and import into Legito.</p>
        <div class="hero-benefits">
          <span class="hero-chip">Word → Gemini → Legito</span>
          <span class="hero-chip">Keeps original wording</span>
          <span class="hero-chip">Human review required</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
with hero_right:
    st.markdown(
        """
        <div class="workflow-card">
          <div class="workflow-heading">HOW IT WORKS</div>
          <div class="workflow-step"><div class="step-number">01</div>
            <div><div class="step-text">Upload Word</div>
              <div class="step-desc">Choose a fictional legacy contract in DOCX format.</div></div></div>
          <div class="workflow-step"><div class="step-number">02</div>
            <div><div class="step-text">Identify reusable fields</div>
              <div class="step-desc">Gemini proposes names, dates and amounts.</div></div></div>
          <div class="workflow-step"><div class="step-number">03</div>
            <div><div class="step-text">Download and review</div>
              <div class="step-desc">Get the annotated Word file, then import it into Legito manually.</div></div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-title">Convert a document</div>'
    '<div class="section-intro">Select a DOCX file, choose a model, and generate an annotated version.</div>',
    unsafe_allow_html=True,
)

form_col, details_col = st.columns([1.65, .85], gap="large")
with form_col:
    uploaded = st.file_uploader(
        "Word document (.docx)",
        type=["docx"],
        help="Maximum 10 MB. Use fictional documents only for this demonstration.",
    )
    with st.expander("Advanced settings", expanded=False):
        model = st.text_input(
            "Gemini model ID",
            value=DEFAULT_MODEL,
            help="Keep the model ID that worked for your Gemini API key.",
        )
    if uploaded is not None:
        st.caption(f"Selected: {uploaded.name} · {uploaded.size / 1024:.1f} KB")
    api_available = bool(os.getenv("GEMINI_API_KEY"))
    if not api_available:
        st.error("Gemini API key missing. Set GEMINI_API_KEY in your .env file or Streamlit app secrets.")
    convert_clicked = st.button(
        "Generate annotated Word →",
        type="primary",
        use_container_width=True,
        disabled=(uploaded is None or not model.strip() or not api_available),
    )

with details_col:
    st.markdown(
        """
        <div class="note-card">
          <strong>Designed for human review</strong><br>
          This tool identifies candidate fields without approving legal text.
          You must check the output before importing it into Legito.
          <br><br>
          <strong>Supported in this demo</strong><br>
          DOCX up to 10 MB · text in body paragraphs and tables · TextInput, Date and Money fields.
        </div>
        """,
        unsafe_allow_html=True,
    )

if convert_clicked and uploaded is not None:
    if uploaded.size > MAX_DOC_BYTES:
        st.error("For this demo, DOCX files must be 10 MB or smaller.")
    else:
        try:
            with st.spinner("Gemini is analyzing your contract and preparing annotations…"):
                started = time.perf_counter()
                data = uploaded.getvalue()
                input_hash = sha256(data).hexdigest()
                doc = Document(BytesIO(data))
                extraction = ask_gemini(paragraphs_in_document(doc), model.strip())
                accepted, skipped = convert_document(doc, extraction)
                output = BytesIO()
                doc.save(output)
                stem = Path(uploaded.name).stem
                report = {
                    "source": uploaded.name,
                    "model": model.strip(),
                    "proposed_fields": len(extraction.fields),
                    "inserted_fields": len(accepted),
                    "accepted": accepted,
                    "skipped": skipped,
                    "gemini_review_notes": extraction.review_notes,
                    "elapsed_seconds": round(time.perf_counter() - started, 2),
                    "status": "READY_FOR_HUMAN_REVIEW_NOT_APPROVED",
                }
                st.session_state["conversion"] = {
                    "source_hash": input_hash,
                    "source": uploaded.name,
                    "docx": output.getvalue(),
                    "report": report,
                    "docx_name": f"{stem}_legito.docx",
                    "report_name": f"{stem}_review.json",
                }
        except Exception as exc:
            st.session_state.pop("conversion", None)
            error_text = str(exc)
            if "503" in error_text:
                st.error("Gemini is temporarily overloaded (503). Try again or choose another available model.")
            elif "429" in error_text:
                st.error("Gemini rate limit reached (429). Check your quota.")
            elif isinstance(exc, ValueError):
                st.error(f"Document could not be processed: {exc}")
            else:
                st.error(f"Conversion failed ({type(exc).__name__}). Check the model, API key and DOCX file.")

conversion = st.session_state.get("conversion")
if conversion and uploaded is not None:
    if sha256(uploaded.getvalue()).hexdigest() == conversion["source_hash"]:
        report = conversion["report"]
        st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="success-head">✓ Your annotated document is ready'
            '<span>AI suggestions have been applied to a new Word document. Review before import.</span></div>',
            unsafe_allow_html=True,
        )
        metric_one, metric_two, metric_three = st.columns(3)
        metric_one.metric("Fields inserted", report["inserted_fields"])
        metric_two.metric("Proposed by AI", report["proposed_fields"])
        metric_three.metric("Processing time", f'{report["elapsed_seconds"]:.1f}s')
        download_doc, download_json = st.columns([1.25, 1], gap="medium")
        with download_doc:
            st.download_button(
                "↓ Download annotated DOCX",
                data=conversion["docx"],
                file_name=conversion["docx_name"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True,
            )
        with download_json:
            st.download_button(
                "↓ Download review report",
                data=json.dumps(report, ensure_ascii=False, indent=2).encode("utf-8"),
                file_name=conversion["report_name"],
                mime="application/json",
                use_container_width=True,
            )
        with st.expander("View field proposals and review notes"):
            st.json(report, expanded=False)

st.markdown(
    """
    <div class="footer-note">Document Migration · Independent academic prototype inspired by Legito's visual style.
    Not affiliated with or endorsed by Legito. Please use fictional documents only.
    Generated annotations require human verification; Legito import is manual.</div>
    """,
    unsafe_allow_html=True,
)
