import streamlit as st
from openpyxl import load_workbook
from copy import copy
from io import BytesIO

st.title("Excel Dosyalarını Birleştir")

uploaded_files = st.file_uploader(
    "Excel dosyalarını yükleyin (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
)

has_header = st.checkbox("Dosyalarda başlık satırı var", value=True)

if uploaded_files and st.button("Birleştir"):
    base_data = BytesIO(uploaded_files[0].getvalue())
    merged_wb = load_workbook(base_data)
    merged_ws = merged_wb.active

    for file in uploaded_files[1:]:
        file_data = BytesIO(file.getvalue())
        wb = load_workbook(file_data)
        ws = wb.active

        start = 2 if has_header else 1
        for row in ws.iter_rows(min_row=start, values_only=False):
            target_row = merged_ws.max_row + 1
            for cell in row:
                new_cell = merged_ws.cell(row=target_row, column=cell.col_idx, value=cell.value)
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.border = copy(cell.border)
                    new_cell.fill = copy(cell.fill)
                    new_cell.number_format = cell.number_format
                    new_cell.protection = copy(cell.protection)
                    new_cell.alignment = copy(cell.alignment)

    output = BytesIO()
    merged_wb.save(output)
    output.seek(0)

    st.download_button(
        label="Birleştirilmiş Dosyayı İndir",
        data=output,
        file_name="merged.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
