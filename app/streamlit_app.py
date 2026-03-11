import streamlit as st
import os
import sys
import tempfile

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.filters import read_filter_questions
from src.pipeline import run_transpose_job

st.set_page_config(page_title="Valuation Transposer", layout="wide")

SETTINGS_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "configs", "settings.json"))

if "step" not in st.session_state:
    st.session_state.step = 1

if "answers" not in st.session_state:
    st.session_state.answers = {}

if "input_file_path" not in st.session_state:
    st.session_state.input_file_path = None

st.title("Valuation Model → Output Transposer")

st.markdown("### 1) Upload input valuation model")
uploaded_file = st.file_uploader("Upload .xlsm / .xlsx", type=["xlsm", "xlsx"])

if uploaded_file is not None:
    suffix = os.path.splitext(uploaded_file.name)[1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.read())
    tmp.close()

    st.session_state.input_file_path = tmp.name
    st.session_state.step = 2
    st.success("File uploaded successfully.")

if st.session_state.step >= 2:
    st.markdown("### 2) Filters")

    try:
        filters = read_filter_questions()
    except Exception as e:
        st.error(f"Could not load filter definitions: {e}")
        st.stop()

    answers = {}

    for item in filters:
        question_key = item.idx
        input_key = f"filter_{item.idx}"
        question_label = item.title
        required = item.required
        options = getattr(item, "options", [])
        qtype = getattr(item, "kind", "text")

        if qtype == "dropdown":
            if options:
                select_options = [""] + options
                selected = st.selectbox(
                    f"{question_label}{' *' if required else ''}",
                    select_options,
                    key=input_key
                )
                answers[question_key] = selected or None
            else:
                answers[question_key] = st.text_input(
                    f"{question_label}{' *' if required else ''}",
                    key=input_key,
                    help="No dropdown options are configured yet for this filter."
                ) or None

        elif qtype in ("yesno", "yes_no", "binary"):
            answers[question_key] = st.radio(
                f"{question_label}{' *' if required else ''}",
                ["", "Yes", "No"],
                horizontal=True,
                key=input_key,
                format_func=lambda value: "Select..." if value == "" else value,
            ) or None
        else:
            answers[question_key] = st.text_input(
                f"{question_label}{' *' if required else ''}",
                key=input_key
            ) or None

    st.session_state.answers = answers

    if st.button("Run Transposer"):
        try:
            run_transpose_job(
                input_path=st.session_state.input_file_path,
                settings_path=SETTINGS_PATH,
                answers=st.session_state.answers,
            )
            st.success("Output created successfully.")
        except Exception as e:
            st.error(f"Error during processing: {e}")
