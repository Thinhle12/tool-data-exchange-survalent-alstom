import os
import pandas as pd
import pyxlsb  # Thư viện để đọc file .xlsb

# Hàm làm sạch dữ liệu trong cột "PID"
def clean_csv(input_folder):
    csv_files = [f for f in os.listdir(input_folder) if f.endswith(".csv")]

    if not csv_files:
        print("Không tìm thấy file .csv nào trong thư mục input.")
        return

    for csv_file in csv_files:
        input_path = os.path.join(input_folder, csv_file)
        df = pd.read_csv(input_path)

        if "PID" not in df.columns:
            print(f"File {csv_file} không có cột 'PID'. Bỏ qua file này.")
            continue

        # Xử lý làm sạch dựa trên tên file
        if csv_file.startswith("ANALOG"):
            df["PID"] = df["PID"].str.replace(":A$", "", regex=True)
        elif csv_file.startswith("CONTROL"):
            df["PID"] = df["PID"].str.replace(":S$", "", regex=True)
        elif csv_file.startswith("STATUS"):
            df["PID"] = df["PID"].str.replace(":S$", "", regex=True)

        # Lưu file đã làm sạch
        df.to_csv(input_path, index=False)
        print(f"Đã làm sạch dữ liệu trong file: {csv_file}")

# Hàm xử lý cột "PID" và tạo các cột "Device" và "Point"
def process_csv(input_folder):
    csv_files = [f for f in os.listdir(input_folder) if f.endswith(".csv")]

    if not csv_files:
        print("Không tìm thấy file .csv nào trong thư mục input.")
        return

    for csv_file in csv_files:
        input_path = os.path.join(input_folder, csv_file)
        df = pd.read_csv(input_path)

        if "PID" not in df.columns:
            print(f"File {csv_file} không có cột 'PID'. Bỏ qua file này.")
            continue

        # Tạo cột "Device"
        def generate_device(value):
            if pd.isna(value) or "," not in value:
                return None
            prefix, _ = value.split(",", 1)
            if prefix.startswith("PCPTHO_RC") and "CM_" in prefix:
                parts = prefix.split("CM_")
                if len(parts) == 2:
                    code = parts[1]
                    mapping = {
                        "COMMON": "CM",
                        "K1": "71",
                        "K2": "73",
                        "K3": "75",
                        "K4": "77",
                        "K5": "79",
                        "K6": "81",
                        "T1": "31",
                        "T2": "33",
                        "T3": "35",
                    }
                    if code in mapping:
                        return f"UC{parts[0][9:]}{mapping[code]}"
            return None

        # Tạo cột "Point"
        def generate_point(value):
            if pd.isna(value) or "," not in value:
                return None
            _, after_comma = value.split(",", 1)
            return after_comma.strip()

        # Áp dụng hàm xử lý cho cột "PID"
        df["Device"] = df["PID"].apply(generate_device)
        df["Point"] = df["PID"].apply(generate_point)

        # Chèn cột "Device" và "Point" vào vị trí mong muốn
        pid_index = df.columns.get_loc("PID")  # Lấy vị trí của cột "PID"
        cols = df.columns.tolist()  # Lấy danh sách cột hiện tại

        # Đưa "Device" vào sau "PID"
        cols.insert(pid_index + 1, cols.pop(cols.index("Device")))

        # Đưa "Point" vào sau "Device"
        cols.insert(pid_index + 2, cols.pop(cols.index("Point")))

        df = df[cols]  # Đặt lại thứ tự cột

        # Lưu file kết quả vào thư mục input với tên mới
        output_filename = os.path.splitext(csv_file)[0] + "_DaThemCot.csv"
        output_path = os.path.join(input_folder, output_filename)
        df.to_csv(output_path, index=False)
        print(f"Đã xử lý và lưu file: {output_filename}")

# Hàm xử lý file _DaThemCot.csv và file .xlsb
def process_csv2(input_folder, output_folder):
    csv_files = [f for f in os.listdir(input_folder) if f.endswith("_DaThemCot.csv")]
    xlsb_files = [f for f in os.listdir(input_folder) if f.endswith(".xlsb")]

    if not csv_files or not xlsb_files:
        print("Không tìm thấy file _DaThemCot.csv hoặc file .xlsb nào trong thư mục input.")
        return

    for csv_file in csv_files:
        csv_path = os.path.join(input_folder, csv_file)
        df_csv = pd.read_csv(csv_path)

        if "Device" not in df_csv.columns or "Point" not in df_csv.columns:
            print(f"File {csv_file} không có cột 'Device' hoặc 'Point'.")
            continue

        # Đọc file .xlsb và xử lý
        xlsb_file = xlsb_files[0]
        xlsb_path = os.path.join(input_folder, xlsb_file)
        is_status_file = csv_file.startswith("STATUS")  # Kiểm tra nếu file .csv bắt đầu bằng "STATUS"

        with pyxlsb.open_workbook(xlsb_path) as wb:
            for sheet_name in wb.sheets:
                # Bỏ qua các sheet "TC" và "TC1" nếu file .csv bắt đầu bằng "STATUS"
                if is_status_file and sheet_name in ["TC", "TC1"]:
                    print(f"Bỏ qua sheet: {sheet_name} vì file .csv bắt đầu bằng 'STATUS'.")
                    continue

                with wb.get_sheet(sheet_name) as sheet:
                    for row in sheet.rows():
                        h_value = row[7].v if len(row) > 7 else None  # Cột H
                        i_value = row[8].v if len(row) > 8 else None  # Cột I
                        d_value = row[3].v if len(row) > 3 else None  # Cột D

                        if pd.isna(h_value) or pd.isna(i_value):
                            continue

                        # Tìm hàng tương ứng trong file _DaThemCot.csv
                        mask = (df_csv["Device"] == h_value) & (df_csv["Point"] == i_value)

                        # Đảm bảo cột "ADDRESS" tồn tại và có kiểu dữ liệu là chuỗi
                        if "ADDRESS" not in df_csv.columns:
                            df_csv["ADDRESS"] = ""  # Tạo cột "ADDRESS" nếu chưa có
                        else:
                            df_csv["ADDRESS"] = df_csv["ADDRESS"].astype(str)

                        # Xử lý giá trị d_value trước khi gán
                        if isinstance(d_value, float) and d_value.is_integer():
                            d_value = int(d_value)  # Loại bỏ phần thập phân nếu giá trị là số nguyên
                        elif isinstance(d_value, float):
                            d_value = str(d_value)  # Chuyển thành chuỗi nếu giá trị là số thực
                        elif d_value is not None:
                            d_value = str(d_value)  # Chuyển thành chuỗi nếu không phải là None

                        # Gán giá trị vào cột "ADDRESS"
                        df_csv.loc[mask, "ADDRESS"] = d_value

        # Lưu file kết quả
        output_filename = os.path.splitext(csv_file)[0] + "_Done.csv"
        output_path = os.path.join(output_folder, output_filename)
        os.makedirs(output_folder, exist_ok=True)
        df_csv.to_csv(output_path, index=False)
        print(f"Đã xử lý và lưu file: {output_filename}")

# Thư mục đầu vào và đầu ra
input_folder = "input"
output_folder = "output"

# Gọi hàm làm sạch và xử lý
clean_csv(input_folder)
process_csv(input_folder)
process_csv2(input_folder, output_folder)
