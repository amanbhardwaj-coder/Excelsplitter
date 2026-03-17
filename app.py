import streamlit as st
import pandas as pd
import math
import zipfile
import tempfile
import os
from io import BytesIO

st.set_page_config(page_title="Excel Splitter", page_icon="📄", layout="centered")

st.title("Excel File Splitter")
st.write("Upload an Excel file and split it into multiple files while keeping the header row in each file.")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

split_option = st.radio(
    "Choose split method",
    ["Number of files", "Rows per file"]
)

output_format = st.selectbox("Output format", ["xlsx", "csv"])

def split_dataframe(df, mode, value):
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

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"File loaded successfully. Total rows: {len(df)}")

        if split_option == "Number of files":
            num_files = st.number_input("Enter number of files", min_value=1, value=2, step=1)
            mode = "files"
            value = int(num_files)
        else:
            rows_per_file = st.number_input("Enter rows per file", min_value=1, value=1000, step=1)
            mode = "rows"
            value = int(rows_per_file)

        preview = st.checkbox("Show preview of uploaded data")
        if preview:
            st.dataframe(df.head(20), use_container_width=True)

        if st.button("Split File"):
            chunks = split_dataframe(df, mode, value)

            if not chunks:
                st.warning("No data found to split.")
            else:
                zip_buffer = BytesIO()

                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for part_num, chunk in chunks:
                        if output_format == "xlsx":
                            temp_buffer = BytesIO()
                            with pd.ExcelWriter(temp_buffer, engine="openpyxl") as writer:
                                chunk.to_excel(writer, index=False, sheet_name="Sheet1")
                            zip_file.writestr(f"part_{part_num}.xlsx", temp_buffer.getvalue())
                        else:
                            csv_data = chunk.to_csv(index=False)
                            zip_file.writestr(f"part_{part_num}.csv", csv_data)

                zip_buffer.seek(0)

                st.success(f"Successfully split into {len(chunks)} files.")
                st.download_button(
                    label="Download ZIP",
                    data=zip_buffer,
                    file_name="split_files.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"Error processing file: {e}")
