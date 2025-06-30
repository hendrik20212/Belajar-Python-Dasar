# C:\\Users\\Administrator\\Downloads
import os
import time
import threading
import requests
import pyexcel as pe
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv, set_key
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import tkinter as tk
from tkinter import scrolledtext
from tkinter import ttk
from tabulate import tabulate

# ==== Firebase ====
import firebase_admin
from firebase_admin import credentials, db

FIREBASE_CRED_FILE = "serviceAccountKey.json"  # Pastikan file ini ada di folder yang sama!
FIREBASE_DB_URL = "https://outobund-default-rtdb.asia-southeast1.firebasedatabase.app/"  # GANTI!

def init_firebase():
    if not firebase_admin._apps:
        cred = credentials.Certificate(FIREBASE_CRED_FILE)
        firebase_admin.initialize_app(cred, {
            'databaseURL': FIREBASE_DB_URL
        })

def get_user_profile(user_id):
    try:
        init_firebase()
        user_ref = db.reference(f'users/{user_id}')
        data = user_ref.get()
        if data:
            nama = data.get('Name', '-')       # <-- perhatikan huruf besar!
            posisi = data.get('Position', '-')
            shift = data.get('Shift', '-')
            return nama, posisi, shift
        else:
            log(f"User ID {user_id} tidak ditemukan di Database.", level="ERROR")
            return '-', '-', '-'
    except Exception as e:
        log(f"Firebase error: {e}", level="ERROR")
        return '-', '-', '-'

# ==== STATISTIK ====
class BotStats:
    def __init__(self):
        self.success_count = 0
        self.error_count = 0
        self.start_time = None
        self.last_run_time = None

    def record_success(self):
        self.success_count += 1
        self.last_run_time = time.time()

    def record_error(self):
        self.error_count += 1
        self.last_run_time = time.time()

    def uptime(self):
        if self.start_time is None:
            return "00:00:00"
        elapsed = int(time.time() - self.start_time)
        h = elapsed // 3600
        m = (elapsed % 3600) // 60
        s = elapsed % 60
        return f"{h:02}:{m:02}:{s:02}"

    def last_run(self):
        if self.last_run_time is None:
            return "Belum pernah dijalankan"
        dt = datetime.fromtimestamp(self.last_run_time)
        return dt.strftime("%Y-%m-%d %H:%M:%S")

stats = BotStats()

# ==== GUI Tkinter ====
root = tk.Tk()
root.title("Z-Logix Bot Uploader")
root.configure(bg="#1e1e1e")
root.geometry("900x600")

# --- FRAME LAYOUT UTAMA ---
top_frame = tk.Frame(root, bg="#1e1e1e")
top_frame.pack(fill="x", padx=10, pady=(10,0))

# Credentials Frame
cred_frame = tk.LabelFrame(top_frame, text="Login Credentials", bg="#1e1e1e", fg="white",
                           font=("Segoe UI", 10, "bold"), bd=2, relief=tk.RIDGE)
cred_frame.grid(row=0, column=0, padx=20, pady=0, sticky="nw")

# Form credentials pakai grid
tk.Label(cred_frame, text="wh-system User:", bg="#1e1e1e", fg="white", anchor="e", width=20).grid(row=0, column=0, sticky="e", padx=8, pady=2)
zlogix_user_entry = tk.Entry(cred_frame, width=18)
zlogix_user_entry.grid(row=0, column=1, padx=6, pady=2)
tk.Label(cred_frame, text="wh-system Pass:", bg="#1e1e1e", fg="white", anchor="e", width=20).grid(row=0, column=2, sticky="e", padx=8, pady=2)
zlogix_pass_entry = tk.Entry(cred_frame, width=18, show="*")
zlogix_pass_entry.grid(row=0, column=3, padx=6, pady=2)

tk.Label(cred_frame, text="WebApp User:", bg="#1e1e1e", fg="white", anchor="e", width=20).grid(row=1, column=0, sticky="e", padx=8, pady=2)
webapp_user_entry = tk.Entry(cred_frame, width=18)
webapp_user_entry.grid(row=1, column=1, padx=6, pady=2)
tk.Label(cred_frame, text="WebApp Pass:", bg="#1e1e1e", fg="white", anchor="e", width=20).grid(row=1, column=2, sticky="e", padx=8, pady=2)
webapp_pass_entry = tk.Entry(cred_frame, width=18, show="*")
webapp_pass_entry.grid(row=1, column=3, padx=6, pady=2)

# User Info Frame
profile_frame = tk.LabelFrame(top_frame, text="User Info", bg="#1e1e1e", fg="white", font=("Segoe UI", 10, "bold"), bd=2, relief=tk.RIDGE)
profile_frame.grid(row=0, column=1, padx=40, pady=0, sticky="ne")

# Label-value dua kolom, titik dua sejajar
tk.Label(profile_frame, text="Nama", bg="#1e1e1e", fg="white", anchor="e", width=7).grid(row=0, column=0, sticky="e", padx=(10,0), pady=2)
tk.Label(profile_frame, text=":", bg="#1e1e1e", fg="white", anchor="center", width=1).grid(row=0, column=1, sticky="e", pady=2)
nama_label = tk.Label(profile_frame, text="-", bg="#1e1e1e", fg="white", anchor="w", width=20)
nama_label.grid(row=0, column=2, sticky="w", padx=(2,10), pady=2)

tk.Label(profile_frame, text="Posisi", bg="#1e1e1e", fg="white", anchor="e", width=7).grid(row=1, column=0, sticky="e", padx=(10,0), pady=2)
tk.Label(profile_frame, text=":", bg="#1e1e1e", fg="white", anchor="center", width=1).grid(row=1, column=1, sticky="e", pady=2)
posisi_label = tk.Label(profile_frame, text="-", bg="#1e1e1e", fg="white", anchor="w", width=20)
posisi_label.grid(row=1, column=2, sticky="w", padx=(2,10), pady=2)

tk.Label(profile_frame, text="Shift", bg="#1e1e1e", fg="white", anchor="e", width=7).grid(row=2, column=0, sticky="e", padx=(10,0), pady=2)
tk.Label(profile_frame, text=":", bg="#1e1e1e", fg="white", anchor="center", width=1).grid(row=2, column=1, sticky="e", pady=2)
shift_label = tk.Label(profile_frame, text="-", bg="#1e1e1e", fg="white", anchor="w", width=20)
shift_label.grid(row=2, column=2, sticky="w", padx=(2,10), pady=2)

# Tombol dan progress bar (di bawah/top_frame)
action_frame = tk.Frame(root, bg="#1e1e1e")
action_frame.pack(fill="x", pady=(5,0))
run_button = ttk.Button(action_frame, text="Jalankan Bot")
run_button.grid(row=0, column=0, padx=20, sticky="w")
progress_bar = ttk.Progressbar(action_frame, mode='indeterminate', length=250)
progress_bar.grid(row=0, column=1, padx=10, sticky="w")

# LOG BOX tetap di bawah pakai pack
log_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, state=tk.DISABLED, bg="#252526", fg="white", font=("Consolas", 10))
log_text.pack(expand=True, fill='both', padx=10, pady=10)

# ==== LOGGING ====
def log(msg, level="INFO"):
    log_text.tag_config("info", foreground="#00ffff")
    log_text.tag_config("success", foreground="#00ff00")
    log_text.tag_config("error", foreground="#ff3333")
    log_text.tag_config("warning", foreground="#ffcc00")
    log_text.tag_config("separator", foreground="#ffffff")

    timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    icon = "ℹ️"
    lower_msg = msg.lower()

    if level == "SUCCESS":
        icon = "✅"
        tag = "success"
    elif level == "ERROR":
        icon = "❌"
        tag = "error"
    elif level == "WARNING":
        icon = "⚠️"
        tag = "warning"
    else:
        tag = "info"

    if "login" in lower_msg:
        icon = "🔐"
    elif "membuka halaman" in lower_msg or "navigasi" in lower_msg:
        icon = "🧭"
    elif "mengambil parameter" in lower_msg or "download" in lower_msg:
        icon = "📥"
    elif "upload file" in lower_msg or "unggah" in lower_msg or "diunggah ke web app" in lower_msg:
        icon = "⬆️"
    elif "file dihapus" in lower_msg or "menghapus file" in lower_msg:
        icon = "🗑️"
    elif "menunggu" in lower_msg:
        icon = "⏳"
    elif "proses selesai" in lower_msg:
        icon = "🎉"

    log_text.config(state=tk.NORMAL)
    log_text.insert(tk.END, f"{timestamp} [{level}] {icon} {msg}\n", tag)

    if level == "SUCCESS" and "proses selesai" in lower_msg:
        log_text.insert(tk.END, "\n==============================================================================================\n\n", "separator")
        log_stats()
    log_text.yview(tk.END)
    log_text.config(state=tk.DISABLED)

def log_stats():
    sukses = str(stats.success_count)
    error = str(stats.error_count)
    uptime = stats.uptime()
    headers = ["Sukses", "Error", "Uptime"]
    data = [sukses, error, uptime]

    table_str = tabulate([data], headers, tablefmt="grid", stralign="center", numalign="center")
    lines = table_str.splitlines()
    width = len(lines[0])

    log_text.tag_config("stat_line", foreground="#ff66cc")
    log_text.tag_config("stat_sukses", foreground="#00ff00")
    log_text.tag_config("stat_error", foreground="#ff3333")
    log_text.tag_config("stat_uptime", foreground="#c586ff")
    log_text.tag_config("stat_title", foreground="#ffffff")

    log_text.config(state=tk.NORMAL)

    title = " Statistic "
    title_pad = (width - len(title)) // 2
    log_text.insert(tk.END, lines[0] + "\n", "stat_line")
    log_text.insert(tk.END, "|" + " " * title_pad + title + " " * (width - len(title) - title_pad - 2) + "|\n", "stat_title")
    log_text.insert(tk.END, lines[0] + "\n", "stat_line")

    header_line = None
    data_line = None
    header_found = False
    for line in lines:
        if "|" in line and not header_found:
            header_line = line
            header_found = True
        elif "|" in line and header_found:
            data_line = line
            break

    header_parts = header_line.split("|")
    log_text.insert(tk.END, header_parts[0] + "|", "stat_line")
    log_text.insert(tk.END, header_parts[1], "stat_sukses")
    log_text.insert(tk.END, "|", "stat_line")
    log_text.insert(tk.END, header_parts[2], "stat_error")
    log_text.insert(tk.END, "|", "stat_line")
    log_text.insert(tk.END, header_parts[3], "stat_uptime")
    log_text.insert(tk.END, "|" + "\n", "stat_line")

    found_middle_line = False
    for l in lines:
        if "+" in l and found_middle_line is False:
            found_middle_line = True
            continue
        if "+" in l and found_middle_line is True:
            log_text.insert(tk.END, l + "\n", "stat_line")
            break

    data_parts = data_line.split("|")
    log_text.insert(tk.END, data_parts[0] + "|", "stat_line")
    log_text.insert(tk.END, data_parts[1], "stat_sukses")
    log_text.insert(tk.END, "|", "stat_line")
    log_text.insert(tk.END, data_parts[2], "stat_error")
    log_text.insert(tk.END, "|", "stat_line")
    log_text.insert(tk.END, data_parts[3], "stat_uptime")
    log_text.insert(tk.END, "|" + "\n", "stat_line")

    log_text.insert(tk.END, lines[-1] + "\n", "stat_line")
    log_text.insert(tk.END, "\n==============================================================================================\n\n", "separator")
    log_text.config(state=tk.DISABLED)

# ==== ENV & KONFIGURASI ====
def load_config():
    load_dotenv()
    config = {
        "ZLOGIX_URL": os.getenv("ZLOGIX_URL"),
        "ZLOGIX_USERNAME": os.getenv("ZLOGIX_USERNAME"),
        "ZLOGIX_PASSWORD": os.getenv("ZLOGIX_PASSWORD"),
        "WEBAPP_URL": os.getenv("WEBAPP_URL"),
        "WEBAPP_USERID": os.getenv("WEBAPP_USERID"),
        "WEBAPP_PASSWORD": os.getenv("WEBAPP_PASSWORD"),
        "DOWNLOAD_DIR": r"C:\\Users\\Rizki Aldiansyah\\Downloads"
    }
    for key, value in config.items():
        if not value:
            raise ValueError(f"Konfigurasi env `{key}` belum diisi")
    return config

def save_user_to_env(zlogix_user, zlogix_pass, webapp_user, webapp_pass):
    env_path = os.path.join(os.getcwd(), '.env')
    set_key(env_path, "ZLOGIX_USERNAME", zlogix_user)
    set_key(env_path, "ZLOGIX_PASSWORD", zlogix_pass)
    set_key(env_path, "WEBAPP_USERID", webapp_user)
    set_key(env_path, "WEBAPP_PASSWORD", webapp_pass)
    load_dotenv(override=True)

def get_chrome_driver(download_dir):
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--incognito")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    import tempfile
    import getpass
    user_temp_dir = os.path.join(tempfile.gettempdir(), f"zlogix_chrome_temp_{getpass.getuser()}")
    os.makedirs(user_temp_dir, exist_ok=True)
    options.add_argument(f'--user-data-dir={user_temp_dir}')
    options.add_argument(f'--disk-cache-dir={user_temp_dir}')

    options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True
    })
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def cleanup_chrome_temp():
    import shutil
    import getpass
    import tempfile
    user_temp_dir = os.path.join(tempfile.gettempdir(), f"zlogix_chrome_temp_{getpass.getuser()}")
    if os.path.exists(user_temp_dir):
        try:
            shutil.rmtree(user_temp_dir)
            print(f"Folder temporary Chrome telah dihapus: {user_temp_dir}")
        except Exception as e:
            print(f"Gagal menghapus folder temp: {e}")

def schedule_cleanup_chrome_temp(interval_seconds=86400):
    def loop():
        while True:
            cleanup_chrome_temp()
            time.sleep(interval_seconds)
    threading.Thread(target=loop, daemon=True).start()

def zlogix_login(driver, wait, url, username, password):
    log("Membuka halaman login Z-Logix...")
    driver.get(url)
    wait.until(EC.presence_of_element_located((By.ID, "txtUserName"))).send_keys(username)
    driver.find_element(By.ID, "txtPassword").send_keys(password)
    driver.find_element(By.ID, "btnSubmit").click()
    log("Login berhasil ke Z-Logix.")

def zlogix_navigate_and_download(driver, wait, download_dir):
    log("Navigasi ke halaman Outbound Reference...")
    wait.until(EC.element_to_be_clickable((By.ID, "ctl00_TreeView1n26"))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.ID, "ctl00_TreeView1t27"))).click()
    time.sleep(3)
    log("Mengambil parameter POST...")
    viewstate = driver.find_element(By.ID, "__VIEWSTATE").get_attribute("value")
    viewstategen = driver.find_element(By.ID, "__VIEWSTATEGENERATOR").get_attribute("value")
    selenium_cookies = driver.get_cookies()
    session = requests.Session()
    for cookie in selenium_cookies:
        session.cookies.set(cookie['name'], cookie['value'])
    download_url = driver.current_url
    payload = {
        '__EVENTTARGET': 'ctl00$ContentPlaceHolder1$btnDownload',
        '__EVENTARGUMENT': '',
        '__VIEWSTATE': viewstate,
        '__VIEWSTATEGENERATOR': viewstategen
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    response = session.post(download_url, data=payload, headers=headers)
    file_path = os.path.join(download_dir, "Outbound_Reference.xls")
    with open(file_path, "wb") as f:
        f.write(response.content)
    log(f"Data source berhasil di-download: {file_path}")

    xlsx_file_path = file_path.replace('.xls', '.xlsx')
    try:
        pe.save_book_as(file_name=file_path, dest_file_name=xlsx_file_path)
        log(f"File berhasil dikonversi ke: {xlsx_file_path}")
    except Exception as e:
        log(f"Gagal konversi file xls ke xlsx: {e}", level="ERROR")

    return file_path

def parse_and_upload_to_firebase(xlsx_file_path, db):
    from datetime import timedelta, timezone
    WIB = timezone(timedelta(hours=7))
    log("Memulai parsing & upload ke Firebase...")
    init_firebase()
    try:
        import math
        df_all = pd.read_excel(xlsx_file_path, header=None)
        header_row_idx = None
        for i, row in df_all.iterrows():
            if any(str(cell).strip().lower() == "job no" for cell in row):
                header_row_idx = i
                break
        if header_row_idx is None:
            log("Header 'Job No' tidak ditemukan di file Excel.", level="ERROR")
            return

        df = pd.read_excel(xlsx_file_path, header=header_row_idx)
        field_map = {
            "jobNo": ["Job No", "Job No."],
            "deliveryDate": ["ETD", "Delivery Date"],
            "deliveryNote": ["Delivery Note No.", "Delivery Note"],
            "remark": ["Ref No.", "Remark"],
            "status": ["Status"],
            "qty": ["BCNo", "Plan Qty"]
        }
        for opts in field_map.values():
            if not any(col in df.columns for col in opts):
                log(f"Kolom '{opts[0]}' tidak ditemukan di file Excel.", level="ERROR")
                return

        def get_val(row, opts):
            for col in opts:
                if col in row:
                    val = row[col]
                    # Khusus remark, jika kosong/nan/None, return ""
                    if col.lower() in ['remark', 'ref no.', 'ref no']:
                        if pd.isna(val) or val is None or (isinstance(val, float) and math.isnan(val)):
                            return ""
                        return str(val).strip()
                    # Umum
                    if pd.isna(val) or val is None:
                        return ""
                    return str(val).strip()
            return ""

        # --- Ambil seluruh jobNo lama dari database untuk sinkronisasi status ---
        jobs_db_ref = db.reference("PhxOutboundJobs")
        jobs_db_snapshot = jobs_db_ref.get() or {}
        jobNo_db_list = list(jobs_db_snapshot.keys())

        jobNo_new_list = []
        upload_count = 0
        error_count = 0

        for idx, row in df.iterrows():
            jobNo = get_val(row, field_map["jobNo"])
            if not jobNo or any(c in jobNo for c in '.#$[]'):
                continue
            jobNo_new_list.append(jobNo)

            existing_job_ref = db.reference(f'PhxOutboundJobs/{jobNo}')
            existing_job_data = existing_job_ref.get() or {}

            delivery_date_raw = get_val(row, field_map["deliveryDate"])
            delivery_date = ""
            try:
                if delivery_date_raw and delivery_date_raw.replace('.', '', 1).isdigit():
                    date_float = float(delivery_date_raw)
                    date = datetime(1899, 12, 30) + pd.to_timedelta(date_float, unit='D')
                    delivery_date = date.strftime("%d-%b-%Y")
                else:
                    date = pd.to_datetime(delivery_date_raw, errors='coerce')
                    if pd.notna(date):
                        delivery_date = date.strftime("%d-%b-%Y")
                    else:
                        delivery_date = delivery_date_raw
            except Exception:
                delivery_date = delivery_date_raw

            job = {
                "jobNo": jobNo,
                "deliveryDate": delivery_date,
                "deliveryNote": get_val(row, field_map["deliveryNote"]),
                "remark": get_val(row, field_map["remark"]),
                "status": get_val(row, field_map["status"]),
                "qty": get_val(row, field_map["qty"]),
                "team": existing_job_data.get("team", ""),
                "jobType": existing_job_data.get("jobType", ""),
                "shift": existing_job_data.get("shift", ""),
                "teamName": existing_job_data.get("teamName", "")
            }

            # finishAt: HH:MM WIB, hanya di-set sekali
            finish_statuses = {"packed", "loaded", "completed"}
            status_now = job["status"].strip().lower()
            prev_finishAt = existing_job_data.get("finishAt", "")
            if prev_finishAt:
                job["finishAt"] = prev_finishAt
            elif status_now in finish_statuses:
                job["finishAt"] = datetime.now(WIB).strftime("%H:%M")

            try:
                existing_job_ref.set(job)
                upload_count += 1
            except Exception as e:
                error_count += 1
                log(f"Gagal upload job {jobNo}: {e}", level="ERROR")

        # --- Sinkronisasi: Set status "Completed" untuk job lama yang tidak ada di file baru ---
        update_missing_count = 0
        update_missing_error = 0
        jobNo_new_set = set(jobNo_new_list)
        for jobNo in jobNo_db_list:
            if jobNo not in jobNo_new_set:
                try:
                    jobs_db_ref.child(jobNo).update({"status": "Completed"})
                    update_missing_count += 1
                except Exception as e:
                    update_missing_error += 1
                    log(f"Gagal update status 'Completed' untuk jobNo {jobNo}: {e}", level="ERROR")

        log(
            f"Parsing & upload ke Firebase selesai. Berhasil: {upload_count}, Gagal: {error_count}. "
            f"Job lama diubah status: {update_missing_count}, Error update: {update_missing_error}",
            level="SUCCESS"
        )
    except Exception as e:
        log(f"Gagal parsing/upload ke Firebase: {e}", level="ERROR")

def delete_file(file_path):
    try:
        os.remove(file_path)
        log(f"File dihapus dari lokal: {file_path}")
    except Exception as e:
        log(f"Gagal menghapus file: {e}", level="WARNING")

def update_profile(user_id):
    nama, posisi, shift = get_user_profile(user_id)
    nama_label.config(text=nama)
    posisi_label.config(text=posisi)
    shift_label.config(text=shift)

def on_zlogix_user_change(*args):
    user_id = zlogix_user_entry.get()
    if user_id.strip().isdigit():
        update_profile(user_id)
    else:
        nama_label.config(text="-")
        posisi_label.config(text="-")
        shift_label.config(text="-")

zlogix_user_entry_var = tk.StringVar()
zlogix_user_entry.config(textvariable=zlogix_user_entry_var)
zlogix_user_entry_var.trace_add('write', lambda *a: on_zlogix_user_change())

def run_bot():
    if stats.start_time is None:
        stats.start_time = time.time()
    progress_bar.start()
    try:
        config = load_config()
        driver = get_chrome_driver(config["DOWNLOAD_DIR"])
        wait = WebDriverWait(driver, 15)
        zlogix_login(driver, wait, config["ZLOGIX_URL"], config["ZLOGIX_USERNAME"], config["ZLOGIX_PASSWORD"])
        file_path = zlogix_navigate_and_download(driver, wait, config["DOWNLOAD_DIR"])
        xlsx_file_path = file_path.replace('.xls', '.xlsx')
        parse_and_upload_to_firebase(xlsx_file_path, db)
        delete_file(file_path)
        if os.path.exists(xlsx_file_path):
            delete_file(xlsx_file_path)
        driver.quit()
        stats.record_success()
        log("Proses selesai.", level="SUCCESS")
    except Exception as e:
        stats.record_error()
        log("Terjadi kesalahan selama proses bot.", level="ERROR")
        log(f"Detail: {e}", level="ERROR")
    finally:
        progress_bar.stop()

def schedule_run(interval_minutes=2):
    def loop():
        countdown = interval_minutes * 60
        while True:
            log("Menunggu interval selanjutnya...")
            for remaining in range(countdown, 0, -1):
                timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
                log_text.config(state=tk.NORMAL)
                last_line = log_text.get("end-2l", "end-1l")
                if '⏱️' in last_line:
                    log_text.delete("end-2l", "end-1l")
                log_text.insert(tk.END, f"{timestamp}   ⏱️ {remaining} detik tersisa untuk eksekusi selanjutnya...\n")
                log_text.yview(tk.END)
                log_text.config(state=tk.DISABLED)
                time.sleep(1)
            log("Menjalankan bot otomatis...")
            run_bot()
    threading.Thread(target=loop, daemon=True).start()

def handle_run_and_schedule():
    zlogix_user = zlogix_user_entry.get()
    zlogix_pass = zlogix_pass_entry.get()
    webapp_user = webapp_user_entry.get()
    webapp_pass = webapp_pass_entry.get()

    if not all([zlogix_user, zlogix_pass, webapp_user, webapp_pass]):
        log("Semua kolom user dan password harus diisi!", level="ERROR")
        return

    try:
        save_user_to_env(zlogix_user, zlogix_pass, webapp_user, webapp_pass)
        log("Data user & password berhasil disimpan di .env")
    except Exception as e:
        log(f"Gagal menyimpan data ke .env: {e}", level="ERROR")
        return

    threading.Thread(target=lambda: [run_bot(), schedule_run()], daemon=True).start()

run_button.config(command=handle_run_and_schedule)
schedule_cleanup_chrome_temp()
root.mainloop()
