import re
from datetime import datetime, date

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="MSISDN Format Converter",
    page_icon="📱",
    layout="wide",
)

st.markdown(
    """
    <style>
    .main {background: linear-gradient(180deg, #f7f9fc 0%, #ffffff 55%);}
    .block-container {padding-top: 2rem; max-width: 1200px;}
    .hero {padding: 1.25rem 1.5rem; border-radius: 16px;
           background: linear-gradient(110deg,#102a43,#176b87);
           color: white; margin-bottom: 1rem;}
    .hero h1 {color: white; margin-bottom: .25rem;}
    .hero p {color: #e6f4f7; margin-bottom: 0;}
    </style>
    <div class="hero">
      <h1>📱 MSISDN Format Converter</h1>
      <p>Upload an Excel file, validate MSISDN / IMSI / date-time values,
      and generate pipe-delimited output.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.write(
    "Upload an Excel workbook (.xlsx) containing columns for "
    "**MSISDN**, **IMSI**, and **DATE/TIME**."
)
st.info(
    "This app formats data locally in Python. An LLM or API key is not needed."
)


def clean_text(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    # Avoid identifiers being rendered as e.g. 12345.0 by Excel.
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


def find_column(columns, candidates):
    normalized = {
        re.sub(r"[^a-z0-9]", "", str(column).lower()): column
        for column in columns
    }
    for candidate in candidates:
        key = re.sub(r"[^a-z0-9]", "", candidate.lower())
        if key in normalized:
            return normalized[key]
    return None


def format_datetime(value):
    if pd.isna(value) or str(value).strip() == "":
        raise ValueError("date/time is empty")

    if isinstance(value, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(value).strftime("%Y%m%d%H%M%S")

    raw = str(value).strip()
    digits = re.sub(r"\D", "", raw)

    if len(digits) == 14:
        return digits
    if len(digits) == 8:
        # If only a date is supplied, use midnight.
        return digits + "000000"

    parsed = pd.to_datetime(raw, errors="coerce", dayfirst=False)
    if pd.isna(parsed):
        parsed = pd.to_datetime(raw, errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        raise ValueError(f"could not parse date/time: {raw}")

    return parsed.strftime("%Y%m%d%H%M%S")


def build_record(msisdn, imsi, dt_string):
    # Preserve the exact record layout requested.
    return (
        f"{msisdn}|{imsi}|{msisdn}_abc|"
        f"{msisdn}_{dt_string}|{msisdn}_{dt_string}||"
        f"{msisdn}_{dt_string}||||||||||"
    )


with st.expander("View expected input and output format", expanded=True):
    st.markdown("**Required input columns**")
    st.code(
        "MSISDN        IMSI              DATETIME\n"
        "3486672316    410060497914995   2026-08-27 16:23:55",
        language="text",
    )
    st.markdown("**Output header and example record**")
    st.code(
        "1|hash|\n"
        "3486672316|410060497914995|3486672316_abc|"
        "3486672316_20260827162355|3486672316_20260827162355||"
        "3486672316_20260827162355||||||||||",
        language="text",
    )
    st.caption(
        "The first line is now `<count>|hash|`, where count is the number "
        "of successfully generated records."
    )

uploaded = st.file_uploader("Choose Excel file", type=["xlsx", "xls"])

if uploaded is not None:
    try:
        workbook = pd.ExcelFile(uploaded)
        sheet_name = st.selectbox("Select worksheet", workbook.sheet_names)
        source_df = pd.read_excel(uploaded, sheet_name=sheet_name, dtype=object)
        source_df.columns = [str(column).strip() for column in source_df.columns]

        msisdn_col = find_column(
            source_df.columns, ["MSISDN", "Mobile Number", "Phone Number"]
        )
        imsi_col = find_column(
            source_df.columns, ["IMSI", "Subscriber IMSI"]
        )
        datetime_col = find_column(
            source_df.columns,
            [
                "MSISDN_DATETIME",
                "DATETIME",
                "DATE TIME",
                "DATE/TIME",
                "DATE",
                "TIMESTAMP",
            ],
        )

        st.subheader("Input preview")
        st.dataframe(source_df.head(20), use_container_width=True)

        st.subheader("Map your columns")
        columns = list(source_df.columns)
        if not columns:
            st.error("The selected worksheet does not contain any columns.")
            st.stop()

        left, middle, right = st.columns(3)
        with left:
            msisdn_col = st.selectbox(
                "MSISDN column",
                columns,
                index=columns.index(msisdn_col)
                if msisdn_col in columns
                else 0,
            )
        with middle:
            imsi_col = st.selectbox(
                "IMSI column",
                columns,
                index=columns.index(imsi_col)
                if imsi_col in columns
                else min(1, len(columns) - 1),
            )
        with right:
            datetime_col = st.selectbox(
                "Date/time column",
                columns,
                index=columns.index(datetime_col)
                if datetime_col in columns
                else min(2, len(columns) - 1),
            )

        if st.button("Convert file", type="primary", use_container_width=True):
            output_rows = []
            errors = []

            for excel_row, (_, row) in enumerate(source_df.iterrows(), start=2):
                msisdn = clean_text(row[msisdn_col])
                imsi = clean_text(row[imsi_col])

                try:
                    if not msisdn:
                        raise ValueError("MSISDN is empty")
                    if not imsi:
                        raise ValueError("IMSI is empty")
                    if not re.fullmatch(r"\d+", msisdn):
                        raise ValueError("MSISDN should contain digits only")
                    if not re.fullmatch(r"\d+", imsi):
                        raise ValueError("IMSI should contain digits only")

                    dt_string = format_datetime(row[datetime_col])
                    output_rows.append(build_record(msisdn, imsi, dt_string))
                except Exception as exc:
                    errors.append(f"Excel row {excel_row}: {exc}")

            st.session_state["converted_lines"] = output_rows
            st.session_state["conversion_errors"] = errors

        if "converted_lines" in st.session_state:
            lines = st.session_state["converted_lines"]
            errors = st.session_state.get("conversion_errors", [])

            first_col, second_col = st.columns(2)
            first_col.metric("Rows converted", len(lines))
            second_col.metric("Rows with errors", len(errors))

            if lines:
                # New header: count equals the number of generated data records.
                output_header = f"{len(lines)}|hash|"
                full_output = output_header + "\n" + "\n".join(lines) + "\n"

                st.subheader("Generated output")
                st.caption(
                    f"Header: {len(lines)}|hash| — count includes only "
                    "successfully generated records."
                )
                st.code(full_output.rstrip(), language="text")

                st.download_button(
                    "⬇️ Download output TXT",
                    data=full_output.encode("utf-8"),
                    file_name="msisdn_output.txt",
                    mime="text/plain",
                    type="primary",
                    use_container_width=True,
                )

                # CSV download also contains the requested header as its first line.
                st.download_button(
                    "⬇️ Download output CSV",
                    data=full_output.encode("utf-8"),
                    file_name="msisdn_output.csv",
                    mime="text/plain",
                    use_container_width=True,
                )
            else:
                st.warning(
                    "No valid records were generated, so the count header "
                    "has not been added. Please review the row errors below."
                )

            if errors:
                with st.expander("Review rows that were skipped"):
                    st.write(
                        "Rows with missing or invalid values are not included "
                        "in the output count."
                    )
                    for error in errors[:200]:
                        st.write(f"- {error}")
                    if len(errors) > 200:
                        st.caption(
                            f"Only the first 200 of {len(errors)} errors are shown."
                        )

    except Exception as exc:
        st.error(f"Could not read this workbook: {exc}")
        st.caption(
            "Please upload a valid Excel workbook. If using an old .xls file, "
            "try saving it as .xlsx first."
        )

st.divider()
st.caption(
    "Tip: keep MSISDN and IMSI columns as text in Excel so leading zeros "
    "are preserved."
)
