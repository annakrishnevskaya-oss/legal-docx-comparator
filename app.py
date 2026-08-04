from __future__ import annotations

from hashlib import sha256
from pathlib import Path

import streamlit as st

from comparator import process_docx


st.set_page_config(
    page_title="Сравнение редакций",
    page_icon="🟧",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
        :root {
            --orange: #ff7a00;
            --orange-dark: #ea6c00;
            --orange-soft: #fff5ec;
            --green: #73b82d;
            --green-dark: #5a9820;
            --text: #252525;
            --muted: #777a76;
            --line: #e8e9e6;
            --background: #fbfaf7;
        }

        html, body, [class*="css"] {
            font-family: Inter, "Segoe UI", Arial, sans-serif;
        }

        .stApp {
            background:
                radial-gradient(circle at 7% 8%, rgba(255, 122, 0, 0.07), transparent 24%),
                radial-gradient(circle at 94% 92%, rgba(115, 184, 45, 0.07), transparent 22%),
                var(--background);
            color: var(--text);
        }

        header[data-testid="stHeader"] {
            background: transparent;
        }

        #MainMenu,
        footer,
        div[data-testid="stDecoration"] {
            display: none;
        }

        .block-container {
            max-width: 720px;
            padding-top: 8.5rem;
            padding-bottom: 5rem;
        }

        .page-title {
            margin: 0;
            text-align: center;
            color: var(--text);
            font-size: clamp(34px, 6vw, 50px);
            line-height: 1.08;
            font-weight: 900;
            letter-spacing: -1.8px;
        }

        .page-title span {
            color: var(--orange);
        }

        .page-caption {
            margin: 14px 0 34px;
            text-align: center;
            color: var(--muted);
            font-size: 15px;
        }

        .upload-label {
            margin: 0 0 10px 3px;
            color: var(--text);
            font-size: 15px;
            font-weight: 800;
        }

        div[data-testid="stFileUploader"] {
            padding: 14px;
            border: 1px solid var(--line);
            border-radius: 22px;
            background: #ffffff;
            box-shadow: 0 14px 36px rgba(35, 35, 35, 0.07);
        }

        div[data-testid="stFileUploaderDropzone"] {
            min-height: 210px;
            border: 2px dashed rgba(255, 122, 0, 0.52);
            border-radius: 17px;
            background: linear-gradient(135deg, #fffaf5, #fffefd);
            transition: border-color 0.18s ease, background 0.18s ease;
        }

        div[data-testid="stFileUploaderDropzone"]:hover {
            border-color: var(--orange);
            background: var(--orange-soft);
        }

        div[data-testid="stFileUploaderDropzone"] button {
            min-height: 43px;
            padding: 0 22px;
            border: 0 !important;
            border-radius: 12px !important;
            color: #ffffff !important;
            background: var(--orange) !important;
            font-weight: 850 !important;
            box-shadow: 0 9px 20px rgba(255, 122, 0, 0.22) !important;
        }

        div[data-testid="stFileUploaderDropzone"] button:hover {
            background: var(--orange-dark) !important;
        }

        div[data-testid="stFileUploaderFile"] {
            border-radius: 13px;
            border-color: var(--line);
        }

        div[data-testid="stSpinner"] {
            justify-content: center;
            margin-top: 22px;
        }

        div[data-testid="stAlert"] {
            margin-top: 22px;
            border-radius: 14px;
        }

        .ready-label {
            margin: 28px 0 11px;
            text-align: center;
            color: var(--green-dark);
            font-size: 14px;
            font-weight: 800;
        }

        div.stDownloadButton > button {
            width: 100%;
            min-height: 54px;
            border: 0;
            border-radius: 14px;
            color: #ffffff;
            background: linear-gradient(135deg, var(--green), #8acb48);
            font-size: 16px;
            font-weight: 900;
            box-shadow: 0 12px 25px rgba(115, 184, 45, 0.24);
            transition: transform 0.18s ease, box-shadow 0.18s ease;
        }

        div.stDownloadButton > button:hover {
            color: #ffffff;
            background: linear-gradient(135deg, var(--green-dark), var(--green));
            transform: translateY(-1px);
            box-shadow: 0 15px 30px rgba(115, 184, 45, 0.29);
        }

        div.stDownloadButton > button:focus {
            color: #ffffff;
        }

        @media (max-width: 700px) {
            .block-container {
                padding-top: 5rem;
                padding-left: 1rem;
                padding-right: 1rem;
            }

            .page-title {
                font-size: 38px;
                letter-spacing: -1.2px;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <h1 class="page-title">Сравнение <span>редакций</span></h1>
    <p class="page-caption">Загрузите редактируемый документ DOCX</p>
    <div class="upload-label">Загрузить файл</div>
    """,
    unsafe_allow_html=True,
)

uploaded_file = st.file_uploader(
    "Загрузить файл",
    type=["docx"],
    accept_multiple_files=False,
    max_upload_size=200,
    label_visibility="collapsed",
)

if uploaded_file is None:
    for key in (
        "processed_fingerprint",
        "comparison_result",
        "comparison_filename",
        "processing_error",
    ):
        st.session_state.pop(key, None)
else:
    source = uploaded_file.getvalue()
    fingerprint = sha256(source).hexdigest()

    if st.session_state.get("processed_fingerprint") != fingerprint:
        st.session_state.pop("comparison_result", None)
        st.session_state.pop("comparison_filename", None)
        st.session_state.pop("processing_error", None)

        try:
            with st.spinner("Сравниваем редакции…"):
                result_bytes, _stats = process_docx(source)

            source_name = Path(uploaded_file.name)
            st.session_state["comparison_result"] = result_bytes
            st.session_state["comparison_filename"] = (
                f"{source_name.stem}_с_выделенными_различиями.docx"
            )
        except Exception as exc:
            st.session_state["processing_error"] = str(exc)

        st.session_state["processed_fingerprint"] = fingerprint

if st.session_state.get("processing_error"):
    st.error(st.session_state["processing_error"], icon="⚠️")

if "comparison_result" in st.session_state:
    st.markdown(
        '<div class="ready-label">Файл обработан и готов</div>',
        unsafe_allow_html=True,
    )
    st.download_button(
        "Скачать готовый файл",
        data=st.session_state["comparison_result"],
        file_name=st.session_state["comparison_filename"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
        on_click="ignore",
    )
