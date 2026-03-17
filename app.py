import math
import os
import re
import zipfile
from io import BytesIO

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Excel File Splitter", page_icon="📄", layout="centered")


def safe_name(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r'[\\/*?:"<>|]+', "_", name)
    name = re.sub(r"\s+", "_", name)
    return name[:100] if name else "output"


def split_dataframe(df: pd.DataFrame, mode: str, value: int):
    total_rows = len(df)

    if total_rows == 0:
        return []

    if mode == "files":
        num_files = value
        rows_per_file = math.ceil(total_rows / num_files)
    elif mode == "rows":
        rows_per_file = value
        num_files = math.ceil(total_rows / rows_per_file)
    else:
        raise ValueError("Invalid split mode")

    chunks = []
    for i in range(num_files):
        start = i * rows_per_file
        end = start + rows_per_file
        chunk = df.iloc[start:end]
        if not chunk.empty:
            chunks.append((i + 1, chunk))

    return chunks


def build_zip(chunks, output_format: str, base_name: str, sheet_name: str) -> BytesIO:
    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for part_num, chunk in chunks:
            file_base = f"{base_name}_part_{part_num}"

            if output_format == "xlsx":
                temp_buffer = BytesIO()
                with pd.ExcelWriter(temp_buffer, engine="openpyxl") as writer:
                    chunk.to_excel(
                        writer,
                        index=False,
                        sheet_name=sheet_name[:31] if sheet_name else "Sheet1",
                    )
                temp_buffer.seek(0)
                zip_file.writestr(f"{file_base}.xlsx", temp_buffer.read())

            else:
                csv_data = chunk.to_csv(index=False)
                zip_file.writestr(f"{file_base}.csv", csv_data)

    zip_buffer.seek(0)
    return zip_buffer


st.title("Excel File Splitter")
st.write(
    "Upload an Excel file and split it into multiple files while keeping the header row in each file."
)

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        file_size_mb = uploaded_file.size / (1024 * 1024)
        st.info(f"Uploaded file size: {file_size_mb:.2f} MB")

        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names

        if not sheet_names:
            st.error("No sheets found in the uploaded Excel file.")
            st.stop()

        selected_sheet = st.selectbox("Select sheet", sheet_names)

        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

        st.success(f"Loaded sheet '{selected_sheet}' successfully.")
        st.write(f"Total rows: **{len(df)}**")
        st.write(f"Total columns: **{len(df.columns)}**")

        split_option = st.radio(
            "Choose split method",
            ["Number of files", "Rows per file"],
        )

        if split_option == "Number of files":
            split_value = st.number_input(
                "Enter number of files",
                min_value=1,
                value=2,
                step=1,
            )
            mode = "files"
        else:
            split_value = st.number_input(
                "Enter rows per file",
                min_value=1,
                value=1000,
                step=1,
            )
            mode = "rows"

        output_format = st.selectbox("Output format", ["xlsx", "csv"])

        file_prefix = st.text_input(
            "Output file prefix",
            value=safe_name(os.path.splitext(uploaded_file.name)[0]),
        )

        show_preview = st.checkbox("Show preview")
        if show_preview:
            st.dataframe(df.head(20), use_container_width=True)

        if st.button("Split File"):
            with st.spinner("Splitting file..."):
                chunks = split_dataframe(df, mode, int(split_value))

                if not chunks:
                    st.warning("No data found to split.")
                else:
                    zip_buffer = build_zip(
                        chunks=chunks,
                        output_format=output_format,
                        base_name=safe_name(file_prefix),
                        sheet_name=selected_sheet,
                    )

                    st.success(f"Successfully split into {len(chunks)} files.")

                    st.download_button(
                        label="Download ZIP",
                        data=zip_buffer,
                        file_name=f"{safe_name(file_prefix)}_split.zip",
                        mime="application/zip",
                    )

    except Exception as e:
        st.error(f"Error processing file: {e}")
