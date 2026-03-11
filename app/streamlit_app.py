import streamlit as st
import os
import sys
import tempfile
from io import BytesIO
import pandas as pd

try:
    from st_combobox import st_combobox
except ImportError:
    st_combobox = None

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src import filters as filters_module
from src.pipeline import build_preview_result, run_transpose_job

st.set_page_config(page_title="Valuation Transposer", layout="wide")

SETTINGS_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "configs", "settings.json"))

if "step" not in st.session_state:
    st.session_state.step = 1

if "answers" not in st.session_state:
    st.session_state.answers = {}

if "input_file_path" not in st.session_state:
    st.session_state.input_file_path = None

if "validation_errors" not in st.session_state:
    st.session_state.validation_errors = []

if "preview_result" not in st.session_state:
    st.session_state.preview_result = None

if "preview_answers" not in st.session_state:
    st.session_state.preview_answers = None


def _build_preview_excel_bytes(preview_df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        preview_df.to_excel(writer, index=False, sheet_name="Preview")
    return output.getvalue()


def _make_search_function(options):
    normalized_options = [str(option) for option in options]

    def search_function(searchterm: str):
        query = (searchterm or "").strip()
        if not query:
            return normalized_options[:100]

        lowered_query = query.casefold()
        startswith_matches = [
            option for option in normalized_options
            if option.casefold().startswith(lowered_query)
        ]
        contains_matches = [
            option for option in normalized_options
            if lowered_query in option.casefold()
            and option not in startswith_matches
        ]

        matches = startswith_matches + contains_matches
        if all(option.casefold() != lowered_query for option in matches):
            matches = [query] + matches

        return matches[:100]

    return search_function


def _resolve_combobox_value(component_value, input_key, options):
    normalized_options = {str(option).casefold(): str(option) for option in options}
    state = st.session_state.get(input_key, {})

    candidates = []
    for value in (
        component_value,
        state.get("result") if isinstance(state, dict) else None,
        state.get("search") if isinstance(state, dict) else None,
    ):
        cleaned = (value or "").strip()
        if cleaned:
            candidates.append(cleaned)

    for candidate in candidates:
        matched_option = normalized_options.get(candidate.casefold())
        if matched_option is not None:
            return matched_option

    if not candidates:
        return None

    return max(candidates, key=len)


def _render_dropdown_combobox(options, input_key, question_label):
    if st_combobox is None:
        return st.text_input(
            "Enter value",
            key=input_key,
            label_visibility="collapsed",
            placeholder=f"Type an existing or new {question_label}",
            help=f"{len(options)} saved options are available for this filter. New values will be added when you run the transposer.",
        ) or None

    try:
        component_value = st_combobox(
            _make_search_function(options),
            key=input_key,
            label=None,
            placeholder=f"Type an existing or new {question_label}",
            blank_search_value="",
            default_options=options[:100],
        )
        return _resolve_combobox_value(component_value, input_key, options)
    except IndexError:
        state = st.session_state.get(input_key, {})
        fallback_value = None
        if isinstance(state, dict):
            fallback_value = _resolve_combobox_value(state.get("result"), input_key, options)

        # Reset the component state so the field can recover on the next rerun.
        st.session_state[input_key] = {
            "result": fallback_value,
            "search": fallback_value or "",
            "options_js": [],
        }
        st.warning(f"{question_label} autocomplete reset after a component error. Please select the value again.")
        return fallback_value or None

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
    st.markdown("<p style='color: red; margin-top: -0.5rem;'>* necessary</p>", unsafe_allow_html=True)

    try:
        filters = filters_module.read_filter_questions()
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

        label_col, input_col = st.columns([1.2, 2.2], gap="small")
        with label_col:
            st.markdown(f"**{question_label}**{' *' if required else ''}")

        with input_col:
            if qtype == "fixed_dropdown":
                answers[question_key] = st.selectbox(
                    "Select value",
                    [""] + options,
                    key=input_key,
                    label_visibility="collapsed",
                ) or None

            elif qtype == "dropdown":
                answers[question_key] = _render_dropdown_combobox(options, input_key, question_label)

            elif qtype in ("yesno", "yes_no", "binary"):
                if input_key not in st.session_state or st.session_state[input_key] not in ("Yes", "No", None):
                    st.session_state[input_key] = None
                answers[question_key] = st.radio(
                    "Select value",
                    ["Yes", "No"],
                    index=None,
                    horizontal=True,
                    key=input_key,
                    label_visibility="collapsed",
                ) or None
            elif qtype == "date":
                answers[question_key] = st.text_input(
                    "Enter year",
                    key=input_key,
                    label_visibility="collapsed",
                    placeholder="YYYY",
                ) or None
            elif qtype == "number":
                answers[question_key] = st.text_input(
                    "Enter number",
                    key=input_key,
                    label_visibility="collapsed",
                    placeholder="Numbers only",
                ) or None
            else:
                answers[question_key] = st.text_input(
                    "Enter value",
                    key=input_key,
                    label_visibility="collapsed",
                ) or None

    st.session_state.answers = answers
    if st.session_state.preview_answers != st.session_state.answers:
        st.session_state.preview_result = None
        st.session_state.preview_answers = None

    if st.session_state.validation_errors:
        st.error("Please review the following filters: " + ", ".join(st.session_state.validation_errors))

    if st.button("Preview Transposer Output"):
        try:
            validate_answers = getattr(filters_module, "validate_answers", None)
            if callable(validate_answers):
                is_valid, missing_filters = validate_answers(filters, st.session_state.answers)
                st.session_state.validation_errors = missing_filters
                if not is_valid:
                    st.rerun()
            else:
                st.session_state.validation_errors = []

            preview_result = build_preview_result(
                input_path=st.session_state.input_file_path,
                settings_path=SETTINGS_PATH,
                answers=st.session_state.answers,
            )
            st.session_state.preview_result = {
                "headers": preview_result.headers,
                "rows": preview_result.rows,
                "qualifying_columns": preview_result.qualifying_columns,
                "warnings": preview_result.warnings,
            }
            st.session_state.preview_answers = dict(st.session_state.answers)
            st.session_state.validation_errors = []
        except Exception as e:
            st.error(f"Error during processing: {e}. Please make sure the document is closed and try again.")

    if st.session_state.preview_result is not None:
        preview_result = st.session_state.preview_result
        st.markdown("### 3) Preview Output")

        for warning in preview_result.get("warnings", []):
            st.warning(warning)

        if preview_result["rows"]:
            preview_df = pd.DataFrame(preview_result["rows"], columns=preview_result["headers"])
            st.dataframe(preview_df, use_container_width=True, hide_index=True, height=420)
            st.download_button(
                "Download Preview as Excel",
                data=_build_preview_excel_bytes(preview_df),
                file_name="preview_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("No qualifying projection years were found for the current file and filters.")

        if st.button("Upload Preview to GitHub Output"):
            try:
                run_result = run_transpose_job(
                    input_path=st.session_state.input_file_path,
                    settings_path=SETTINGS_PATH,
                    answers=st.session_state.preview_answers or st.session_state.answers,
                )
                st.success(f"Uploaded {run_result.rows_written} row(s) to the output workbook.")
            except Exception as e:
                st.error(f"Error during processing: {e}. Please make sure the document is closed and try again.")
