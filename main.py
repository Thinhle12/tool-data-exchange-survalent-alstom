import os
import pandas as pd
import pyxlsb
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import time

# Hàm hiển thị log từng dòng
def log_message(message):
    log_output.insert(tk.END, message + "\n")
    log_output.see(tk.END)
    root.update_idletasks()
    time.sleep(0.1)

# Hàm làm sạch dữ liệu trong cột "PID"
def clean_csv(input_folder):
    log_message("Bắt đầu làm sạch dữ liệu...")
    csv_files = [f for f in os.listdir(input_folder) if f.lower().endswith(".csv")]

    if not csv_files:
        log_message("Không tìm thấy file .csv nào trong thư mục input.")
        return

    for csv_file in csv_files:
        input_path = os.path.join(input_folder, csv_file)
        df = pd.read_csv(input_path)

        if "PID" not in df.columns:
            log_message(f"File {csv_file} không có cột 'PID'. Bỏ qua.")
            continue

        # Xử lý làm sạch dựa trên tên file
        if csv_file.upper().startswith("ANALOG"):
            df["PID"] = df["PID"].str.replace(":A$", "", regex=True)
        elif csv_file.upper().startswith("CONTROL"):
            df["PID"] = df["PID"].str.replace(":S$", "", regex=True)
        elif csv_file.upper().startswith("STATUS"):
            df["PID"] = df["PID"].str.replace(":S$", "", regex=True)

        # Lưu file đã làm sạch
        df.to_csv(input_path, index=False)
        log_message(f"Đã làm sạch dữ liệu trong file: {csv_file}")

# Hàm xử lý cột "PID" và tạo các cột "Device" và "Point"
def process_csv(input_folder):
    log_message("Bắt đầu xử lý cột 'PID'...")
    csv_files = [f for f in os.listdir(input_folder) if f.lower().endswith(".csv")]

    if not csv_files:
        log_message("Không tìm thấy file .csv nào trong thư mục input.")
        return

    for csv_file in csv_files:
        input_path = os.path.join(input_folder, csv_file)
        df = pd.read_csv(input_path)

        if "PID" not in df.columns:
            log_message(f"File {csv_file} không có cột 'PID'. Bỏ qua.")
            continue

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

        def generate_point(value):
            if pd.isna(value) or "," not in value:
                return None
            _, after_comma = value.split(",", 1)
            return after_comma.strip()

        df["Device"] = df["PID"].apply(generate_device)
        df["Point"] = df["PID"].apply(generate_point)

        pid_index = df.columns.get_loc("PID")
        cols = df.columns.tolist()
        cols.insert(pid_index + 1, cols.pop(cols.index("Device")))
        cols.insert(pid_index + 2, cols.pop(cols.index("Point")))
        df = df[cols]

        output_filename = os.path.splitext(csv_file)[0] + "_DaThemCot.csv"
        output_path = os.path.join(input_folder, output_filename)
        df.to_csv(output_path, index=False)
        log_message(f"Đã xử lý và lưu file: {output_filename}")

# Hàm xử lý file _DaThemCot.csv và file .xlsb
def process_csv2(input_folder, output_folder):
    log_message("Bắt đầu xử lý file _DaThemCot.csv...")
    csv_files = [f for f in os.listdir(input_folder) if f.lower().endswith("_dathemcot.csv")]
    xlsb_files = [f for f in os.listdir(input_folder) if f.lower().endswith(".xlsb")]

    if not csv_files or not xlsb_files:
        log_message("Không tìm thấy file _DaThemCot.csv hoặc file .xlsb nào.")
        return

    for csv_file in csv_files:
        csv_path = os.path.join(input_folder, csv_file)
        df_csv = pd.read_csv(csv_path)

        if "Device" not in df_csv.columns or "Point" not in df_csv.columns:
            log_message(f"File {csv_file} không có cột 'Device' hoặc 'Point'.")
            continue

        xlsb_file = xlsb_files[0]
        xlsb_path = os.path.join(input_folder, xlsb_file)
        is_status_file = csv_file.upper().startswith("STATUS")

        with pyxlsb.open_workbook(xlsb_path) as wb:
            for sheet_name in wb.sheets:
                if is_status_file and sheet_name in ["TC", "TC1"]:
                    log_message(f"Bỏ qua sheet: {sheet_name}.")
                    continue

                with wb.get_sheet(sheet_name) as sheet:
                    for row in sheet.rows():
                        h_value = row[7].v if len(row) > 7 else None
                        i_value = row[8].v if len(row) > 8 else None
                        d_value = row[3].v if len(row) > 3 else None

                        if pd.isna(h_value) or pd.isna(i_value):
                            continue

                        mask = (df_csv["Device"] == h_value) & (df_csv["Point"] == i_value)

                        if "ADDRESS" not in df_csv.columns:
                            df_csv["ADDRESS"] = ""
                        else:
                            df_csv["ADDRESS"] = df_csv["ADDRESS"].astype(str)

                        if isinstance(d_value, float) and d_value.is_integer():
                            d_value = int(d_value)
                        elif isinstance(d_value, float):
                            d_value = str(d_value)
                        elif d_value is not None:
                            d_value = str(d_value)

                        df_csv.loc[mask, "ADDRESS"] = d_value

        # Loại bỏ đuôi ".0" trong cột ADDRESS
        if "ADDRESS" in df_csv.columns:
            df_csv["ADDRESS"] = df_csv["ADDRESS"].apply(
                lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (float, int)) and x == int(x) else x
            )

        output_filename = os.path.splitext(csv_file)[0] + "_Done.csv"
        output_path = os.path.join(output_folder, output_filename)
        os.makedirs(output_folder, exist_ok=True)
        df_csv.to_csv(output_path, index=False)
        log_message(f"Đã xử lý và lưu file: {output_filename}")

# Hàm xóa cột "Device" và "Point" khỏi các file _Done.csv
def remove_columns_from_done(output_folder):
    log_message("Bắt đầu xóa cột 'Device' và 'Point'...")
    done_files = [f for f in os.listdir(output_folder) if f.lower().endswith("_done.csv")]

    if not done_files:
        log_message("Không tìm thấy file _Done.csv nào trong thư mục output.")
        return

    for done_file in done_files:
        file_path = os.path.join(output_folder, done_file)
        df = pd.read_csv(file_path)

        # Loại bỏ đuôi ".0" trong cột ADDRESS
        if "ADDRESS" in df.columns:
            df["ADDRESS"] = df["ADDRESS"].apply(
                lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (float, int)) and x == int(x) else x
            )

        if "Device" in df.columns and "Point" in df.columns:
            df.drop(columns=["Device", "Point"], inplace=True)
            df.to_csv(file_path, index=False)
            log_message(f"Đã xóa cột 'Device' và 'Point' khỏi file: {done_file}")

# Hàm chạy chương trình
def run_program():
    input_folder = "input"
    output_folder = "output"
    steps = 4
    progress = 0

    clean_csv(input_folder)
    progress += 1
    progress_var.set(int((progress / steps) * 100))

    process_csv(input_folder)
    progress += 1
    progress_var.set(int((progress / steps) * 100))

    process_csv2(input_folder, output_folder)
    progress += 1
    progress_var.set(int((progress / steps) * 100))

    remove_columns_from_done(output_folder)
    progress += 1
    progress_var.set(int((progress / steps) * 100))

    log_message("Chương trình đã hoàn tất.")

# Tạo giao diện GUI
root = tk.Tk()
root.title("Tool copy data exchange")
root.iconbitmap("main.ico")

# Cấu hình lưới chính của root để các thành phần co giãn
root.rowconfigure(0, weight=1)
root.columnconfigure(0, weight=1)

frame = ttk.Frame(root, padding=10)
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Cấu hình lưới trong frame
frame.rowconfigure(0, weight=8)  # Ô log chiếm nhiều không gian
frame.rowconfigure(1, weight=1)  # Progress bar co giãn theo chiều dọc nhỏ
frame.rowconfigure(2, weight=1)  # Nút RUN
frame.rowconfigure(3, weight=1)  # Footer
frame.columnconfigure(0, weight=1)

# Log Output
log_output = ScrolledText(frame, wrap=tk.WORD, height=20, width=80)
log_output.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

# Progress Bar
progress_var = tk.IntVar()
progress_bar = ttk.Progressbar(frame, variable=progress_var, maximum=100)
progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)

# Run Button
run_button = ttk.Button(frame, text="RUN", command=run_program)
run_button.grid(row=2, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))

# Footer
footer = ttk.Label(frame, text="Thinhlh", anchor="center")
footer.grid(row=3, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))

root.mainloop()
