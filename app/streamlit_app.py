import streamlit as st
import os
import sys
import tempfile

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.filters import read_filter_questions
from src.pipeline import run_transpose_job

st.set_page_config(page_title="Valuation Transposer", layout="wide")

FILTERS_PATH = r"C:\Users\giuli\OneDrive\Val-Transposer-Output\Filters\Filters.xlsx"
OUTPUT_PATH = r"C:\Users\giuli\OneDrive\Val-Transposer-Output\Output sheet.xlsx"

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
        filters = read_filter_questions(
            FILTERS_PATH,
            "Filters",  # or "filters" depending on sheet name
            5,          # question_row
            8,          # required_row
            "B",        # start_col  (column B)
            "Q",        # end_col    (column Q, 16 questions)
            10          # dropdown_start_row
        )
        st.write("DEBUG filter item type:", type(filters[0]))
        st.write("DEBUG fields:", dir(filters[0]))
    except Exception as e:
        st.error(f"Could not read Filters.xlsx: {e}")
        st.stop()

    answers = {}

    for item in filters:
        question_key = item.key
        question_label = item.label
        required = item.required
        options = getattr(item, "options", [])
        qtype = getattr(item, "qtype", getattr(item, "type", "dropdown"))

        if qtype == "dropdown":
            if not options:
                options = ["No options found"]

            answers[question_key] = st.selectbox(
                f"{question_label}{' *' if required else ''}",
                options,
                key=question_key
            )

        elif qtype in ("yesno", "yes_no", "binary"):
            answers[question_key] = st.radio(
                f"{question_label}{' *' if required else ''}",
                ["Yes", "No"],
                horizontal=True,
                key=question_key
            )

    st.session_state.answers = answers

    if st.button("Run Transposer"):
        try:
            run_transpose_job(
                input_file=st.session_state.input_file_path,
                filters=st.session_state.answers,
                output_path=OUTPUT_PATH
            )
            st.success("Output created successfully.")
        except Exception as e:
            st.error(f"Error during processing: {e}")