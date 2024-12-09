import os
import pandas as pd

# Hàm xử lý cột "PID" và tạo các cột "Device" và "Point"
def process_csv(input_folder, output_folder):
    # Lấy danh sách file .csv trong thư mục input
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
                        return f"UC{parts[0][9:]}{mapping[code]}"  # **** là phần giữa "PCPTHO_RC" và "CM_"
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

        # Lưu file kết quả
        output_filename = os.path.splitext(csv_file)[0] + "_DaXuLy.csv"
        output_path = os.path.join(output_folder, output_filename)
        os.makedirs(output_folder, exist_ok=True)
        df.to_csv(output_path, index=False)
        print(f"Đã xử lý và lưu file: {output_filename}")

# Thư mục đầu vào và đầu ra
input_folder = "input"
output_folder = "output"

# Gọi hàm xử lý
process_csv(input_folder, output_folder)
