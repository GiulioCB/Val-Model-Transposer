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

FILTER_DISPLAY_NAMES = {
    "PropertyType": "Property Type",
    "MeetingSpace (SQM)": "Meeting Space",
    "HealthClub": "Health Club",
    "FoodOutlets": "Food Outlets",
    "BeverageOutlets": "Beverage Outlets",
}

st.markdown(
    """
    <style>
    .st-key-reset_filters_button button,
    .st-key-reset_filters_button button:focus,
    .st-key-reset_filters_button button:active,
    .st-key-reset_filters_button button:visited {
        color: #b00020 !important;
        border: 1px solid #b00020 !important;
        background: #ffffff !important;
        transition: background-color 180ms ease, color 180ms ease, border-color 180ms ease !important;
    }

    .st-key-reset_filters_button button p,
    .st-key-reset_filters_button button:focus p,
    .st-key-reset_filters_button button:active p,
    .st-key-reset_filters_button button:visited p {
        color: #b00020 !important;
    }

    .st-key-reset_filters_button button:hover {
        background: #000000 !important;
        color: #b00020 !important;
        border-color: #b00020 !important;
    }

    .st-key-reset_filters_button button:hover p {
        color: #b00020 !important;
    }

    .st-key-preview_output_button button,
    .st-key-preview_output_button button:focus,
    .st-key-preview_output_button button:active,
    .st-key-preview_output_button button:visited {
        color: #000000 !important;
        border: 1px solid #000000 !important;
        background: #ffffff !important;
        transition: background-color 180ms ease, color 180ms ease, border-color 180ms ease !important;
    }

    .st-key-preview_output_button button p,
    .st-key-preview_output_button button:focus p,
    .st-key-preview_output_button button:active p,
    .st-key-preview_output_button button:visited p {
        color: #000000 !important;
    }

    .st-key-preview_output_button button:hover {
        background: #0f8f3d !important;
        color: #ffffff !important;
        border-color: #0f8f3d !important;
    }

    .st-key-preview_output_button button:hover p {
        color: #ffffff !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

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

if "preview_dirty" not in st.session_state:
    st.session_state.preview_dirty = False

if "filter_reset_version" not in st.session_state:
    st.session_state.filter_reset_version = 0

if "hotel_status" not in st.session_state:
    st.session_state.hotel_status = "New Hotel"

if "uploaded_file_name" not in st.session_state:
    st.session_state.uploaded_file_name = None

if "upload_reset_version" not in st.session_state:
    st.session_state.upload_reset_version = 0


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

        remaining_matches = [
            option for option in normalized_options
            if option not in startswith_matches and option not in contains_matches
        ]

        matches = startswith_matches + contains_matches + remaining_matches
        if all(option.casefold() != lowered_query for option in matches):
            matches = [query] + matches

        deduped_matches = []
        seen = set()
        for option in matches:
            key = option.casefold()
            if key in seen:
                continue
            seen.add(key)
            deduped_matches.append(option)

        return deduped_matches[:100]

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
            help=f"{len(options)} configured options are available for this filter.",
            autocomplete="off",
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

        st.session_state[input_key] = {
            "result": fallback_value,
            "search": fallback_value or "",
            "options_js": [],
        }
        st.warning(f"{question_label} autocomplete reset after a component error. Please select the value again.")
        return fallback_value or None


def _reset_filter_state(filters):
    current_reset_version = st.session_state.filter_reset_version
    for item in filters:
        input_key = f"filter_{item.idx}_{current_reset_version}"
        st.session_state.pop(input_key, None)
        st.session_state.pop(f"{input_key}_react", None)

    st.session_state.answers = {}
    st.session_state.validation_errors = []
    st.session_state.preview_result = None
    st.session_state.preview_answers = None
    st.session_state.preview_dirty = False
    st.session_state.filter_reset_version += 1


def _reset_upload_state():
    st.session_state.step = 1
    st.session_state.answers = {}
    st.session_state.input_file_path = None
    st.session_state.uploaded_file_name = None
    st.session_state.validation_errors = []
    st.session_state.preview_result = None
    st.session_state.preview_answers = None
    st.session_state.preview_dirty = False
    st.session_state.hotel_status = "New Hotel"
    st.session_state.filter_reset_version += 1
    st.session_state.upload_reset_version += 1


def _format_validation_errors(validation_errors):
    formatted_errors = []
    for error in validation_errors:
        text = str(error).strip()
        if ": " in text and text.startswith("Filter "):
            _, remainder = text.split(": ", 1)
            text = remainder.strip()
        for internal_name, display_name in FILTER_DISPLAY_NAMES.items():
            text = text.replace(internal_name, display_name)
        formatted_errors.append(text)
    return formatted_errors


def _get_display_label(question_label):
    return FILTER_DISPLAY_NAMES.get(question_label, question_label)


def _get_default_hotel_status() -> str:
    # Placeholder for the future database match against BA/BB coordinates.
    # For now every uploaded workbook defaults to "New Hotel".
    return "New Hotel"

st.title("Benchmarking Database Uploader")

upload_title_col, upload_reset_col = st.columns([6, 1.5], vertical_alignment="center")
with upload_title_col:
    st.markdown("### 1) Upload Valuation Model (excel)")
with upload_reset_col:
    reset_upload_clicked = st.button(
        "Reset Upload",
        use_container_width=True,
        key="reset_upload_button",
    )

st.markdown(
    "<p style='margin: -0.35rem 0 0.6rem; color: #5f6368; font-size: 0.92rem;'>Please upload only one valuation model per upload cycle</p>",
    unsafe_allow_html=True,
)
if reset_upload_clicked:
    _reset_upload_state()
    st.rerun()

uploaded_file = st.file_uploader(
    "Upload .xlsm / .xlsx",
    type=["xlsm", "xlsx"],
    key=f"uploaded_file_{st.session_state.upload_reset_version}",
)

if uploaded_file is not None:
    suffix = os.path.splitext(uploaded_file.name)[1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.read())
    tmp.close()

    st.session_state.input_file_path = tmp.name
    st.session_state.step = 2
    if st.session_state.uploaded_file_name != uploaded_file.name:
        st.session_state.hotel_status = _get_default_hotel_status()
        st.session_state.uploaded_file_name = uploaded_file.name

if st.session_state.step >= 2:
    st.markdown(
        "<p style='margin: 0.2rem 0 0.6rem; color: #1c8b4d; font-size: 0.95rem;'>Valuation model uploaded successfully.</p>",
        unsafe_allow_html=True,
    )

    st.radio(
        "Hotel Status",
        ["New Hotel", "Existing Hotel, new data"],
        key="hotel_status",
        horizontal=True,
        label_visibility="collapsed",
    )

    title_col, reset_col = st.columns([6, 1.5], vertical_alignment="center")
    with title_col:
        st.markdown("### 2) Filters")
    with reset_col:
        reset_filters_clicked = st.button(
            "Reset Filters",
            use_container_width=True,
            key="reset_filters_button",
            type="secondary",
        )

    try:
        filters = filters_module.read_filter_questions()
    except Exception as e:
        st.error(f"Could not load filter definitions: {e}")
        st.stop()

    if reset_filters_clicked:
        _reset_filter_state(filters)
        st.rerun()

    answers = {}

    for item in filters:
        question_key = item.idx
        input_key = f"filter_{item.idx}_{st.session_state.filter_reset_version}"
        question_label = item.title
        display_label = _get_display_label(question_label)
        required = item.required
        options = getattr(item, "options", [])
        qtype = getattr(item, "kind", "text")

        label_col, input_col = st.columns([1.2, 2.2], gap="small")
        with label_col:
            required_marker = " <span style='color: #b00020;'>*</span>" if required else ""
            st.markdown(f"**{display_label}**{required_marker}", unsafe_allow_html=True)
            if qtype == "dropdown":
                st.markdown(
                    "<p style='margin: -0.65rem 0 0; line-height: 1; color: #8a8f98; font-size: 0.76rem; white-space: nowrap;'>New value: Press \"Enter\" to save it</p>",
                    unsafe_allow_html=True,
                )

        with input_col:
            if qtype == "fixed_dropdown":
                answers[question_key] = st.selectbox(
                    "Select value",
                    [""] + options,
                    key=input_key,
                    label_visibility="collapsed",
                ) or None

            elif qtype == "dropdown":
                answers[question_key] = _render_dropdown_combobox(options, input_key, display_label)

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
                    autocomplete="off",
                ) or None
            elif qtype == "number":
                answers[question_key] = st.text_input(
                    "Enter number",
                    key=input_key,
                    label_visibility="collapsed",
                    placeholder="Numbers only",
                    autocomplete="off",
                ) or None
            else:
                answers[question_key] = st.text_input(
                    "Enter value",
                    key=input_key,
                    label_visibility="collapsed",
                    autocomplete="off",
                ) or None

    st.markdown("<p style='color: #b00020; margin: 0.25rem 0 0.75rem;'>* necessary</p>", unsafe_allow_html=True)

    st.session_state.answers = answers
    if st.session_state.preview_answers != st.session_state.answers:
        st.session_state.preview_dirty = st.session_state.preview_result is not None
    else:
        st.session_state.preview_dirty = False

    if st.session_state.validation_errors:
        formatted_errors = _format_validation_errors(st.session_state.validation_errors)
        st.markdown(
            (
                "<p style='margin: 0 0 0.75rem; color: #b00020; font-size: 0.98rem;'>"
                "Please review these filters: "
                f"{', '.join(formatted_errors)}"
                "</p>"
            ),
            unsafe_allow_html=True,
        )

    preview_button_label = (
        "Update Transposer Output"
        if st.session_state.preview_result is not None and st.session_state.preview_dirty
        else "Preview Transposer Output"
    )
    preview_clicked = st.button(
        preview_button_label,
        key="preview_output_button",
        type="secondary",
    )
    if preview_clicked:
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
            st.session_state.preview_dirty = False
            st.session_state.validation_errors = []
        except Exception as e:
            st.error(f"Error during processing: {e}. Please make sure the document is closed and try again.")

    if st.session_state.preview_result is not None:
        preview_result = st.session_state.preview_result
        st.markdown("### 3) Preview Output")

        if st.session_state.preview_dirty:
            st.info("The preview below is based on the previous filter selection. Click 'Update Transposer Output' to refresh it.")

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

        if st.button("Upload Preview to GitHub Output", disabled=st.session_state.preview_dirty):
            try:
                run_result = run_transpose_job(
                    input_path=st.session_state.input_file_path,
                    settings_path=SETTINGS_PATH,
                    answers=st.session_state.preview_answers or st.session_state.answers,
                )
                st.success(f"Uploaded {run_result.rows_written} row(s) to the output workbook.")
            except Exception as e:
                st.error(f"Error during processing: {e}. Please make sure the document is closed and try again.")
