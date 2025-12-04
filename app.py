import streamlit as st
import tempfile
import os

from assign_groups_xu_ly_het_mot_lan_tu_file_Input_ALL import (
    read_input,
    validate,
    assign,
    write_assigned_to_same_file,
)

st.title("Phân nhóm khách hàng từ file Excel")

st.write("Upload file Excel có các sheet: Customers, GroupName, GroupSize. Ứng dụng sẽ thêm sheet 'Assigned' vào file.")

uploaded_file = st.file_uploader("Chọn file Excel đầu vào (.xlsx)", type=["xlsx"])

seed = st.number_input("Seed ngẫu nhiên (tuỳ chọn, để mặc định cũng được)", value=42, step=1)

if uploaded_file is not None:
    # Lưu file tạm
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_input.write(uploaded_file.read())
    temp_input.close()
    input_path = temp_input.name

    if st.button("Chạy phân nhóm khách hàng"):
        try:
            # Gọi lại các hàm trong file gốc
            customers, groups, groupsize = read_input(input_path)
            customers, groups, groupsize = validate(customers, groups, groupsize)
            assigned = assign(customers, groups, groupsize, seed=int(seed))
            write_assigned_to_same_file(input_path, assigned, sheet_name="Assigned")

            # Đọc lại file sau khi đã ghi sheet Assigned để cho tải xuống
            with open(input_path, "rb") as f:
                data = f.read()

            st.success(f"Đã phân nhóm cho {len(assigned)} khách hàng và ghi vào sheet 'Assigned'.")
            st.download_button(
                label="Tải file Excel đã phân nhóm",
                data=data,
                file_name="output_assigned.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Có lỗi xảy ra: {e}")
else:
    st.info("Hãy upload một file Excel để bắt đầu.")
