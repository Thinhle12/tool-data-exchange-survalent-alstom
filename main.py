import os
import pandas as pd
from pyxlsb import open_workbook

def process_files(input_folder, output_folder, user_input):
    # Tìm file .xlsb và .csv trong thư mục input
    xlsb_file = None
    csv_file = None

    for file in os.listdir(input_folder):
        if file.endswith(".xlsb"):
            xlsb_file = os.path.join(input_folder, file)
        elif file.endswith(".csv"):
            csv_file = os.path.join(input_folder, file)

    # Kiểm tra sự tồn tại của các file
    if not xlsb_file:
        print("Không tìm thấy file .xlsb trong thư mục input.")
        return
    if not csv_file:
        print("Không tìm thấy file .csv trong thư mục input.")
        return

    # Đọc file .csv
    df_csv = pd.read_csv(csv_file)
    updated_rows = []

    # Đọc file .xlsb
    with open_workbook(xlsb_file) as wb:
        for sheet_name in wb.sheets:
            with wb.get_sheet(sheet_name) as sheet:
                for row in sheet.rows():
                    # Đọc dữ liệu từ các cột D, H, và I trong file .xlsb
                    row_data = [r.v for r in row]
                    if len(row_data) > 8:  # Kiểm tra đủ số cột
                        xlsb_col_d = row_data[3]  # Cột D
                        xlsb_col_h = row_data[7]  # Cột H
                        xlsb_col_i = row_data[8]  # Cột I

                        # Điều kiện 1: Cột H trong file .xlsb phải là "UC" + tên ngăn + "K1"
                        if xlsb_col_h == f"UC{user_input}K1":
                            # Lọc file .csv để kiểm tra điều kiện
                            for index, csv_row in df_csv.iterrows():
                                csv_col_f = csv_row["F"]
                                csv_col_h = csv_row["H"]

                                # Điều kiện 2: Kiểm tra dữ liệu trước dấu "," trong cột F của file .csv
                                if "," in csv_col_f:
                                    csv_prefix, csv_suffix = csv_col_f.split(",", 1)
                                    if csv_prefix == f"PCPTHO_RC{user_input}CM_K1" and csv_suffix == str(xlsb_col_i):
                                        # Thỏa mãn điều kiện, copy dữ liệu từ xlsb_col_d vào csv_col_h
                                        df_csv.at[index, "H"] = xlsb_col_d
                                        updated_rows.append(index)

    # Xác định tên file đầu ra
    csv_filename = os.path.basename(csv_file)  # Lấy tên file gốc
    output_csv_name = os.path.splitext(csv_filename)[0] + "_done.csv"  # Thêm hậu tố "_done"
    output_csv = os.path.join(output_folder, output_csv_name)

    # Ghi file .csv kết quả
    os.makedirs(output_folder, exist_ok=True)
    df_csv.to_csv(output_csv, index=False)
    print(f"Dữ liệu đã được cập nhật cho {len(updated_rows)} dòng và lưu vào '{output_csv}'.")

# Nhập thông tin từ người dùng
input_folder = "input"
output_folder = "output"
user_input = input("Nhập tên ngăn: ")

# Gọi hàm xử lý
process_files(input_folder, output_folder, user_input)
