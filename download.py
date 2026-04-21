from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time
import os
import datetime
import shutil
import pandas as pd
from openpyxl import load_workbook

# =====================================
# SETUP
# =====================================

bulan_map = {
    "01": "JANUARI", "02": "FEBRUARI", "03": "MARET",
    "04": "APRIL",   "05": "MEI",      "06": "JUNI",
    "07": "JULI",    "08": "AGUSTUS",  "09": "SEPTEMBER",
    "10": "OKTOBER", "11": "NOVEMBER", "12": "DESEMBER"
}

bulan      = datetime.datetime.now().strftime("%m")
nama_bulan = bulan_map[bulan]

download_root = os.path.join(r"D:\Automatis Monitoring ATM", nama_bulan)
os.makedirs(download_root, exist_ok=True)

# =====================================
# LOG
# =====================================

def tulis_log(pesan):
    tanggal  = datetime.datetime.now().strftime("%Y-%m-%d")
    waktu    = datetime.datetime.now().strftime("%H:%M:%S")
    log_path = os.path.join(download_root, f"log_{tanggal}.txt")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{waktu}] {pesan}\n")
    print(pesan)

# =====================================
# CREATE DRIVER
# =====================================

def create_driver():
    options = webdriver.ChromeOptions()

    # Ganti NamaUser sesuai komputer kamu (cek di chrome://version)
    profile_path = r"C:\Users\NamaUser\AppData\Local\Google\Chrome\User Data"
    options.add_argument(f"--user-data-dir={profile_path}")
    options.add_argument("--profile-directory=Default")

    prefs = {
        "download.default_directory": download_root,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-gpu")

    driver = webdriver.Chrome(options=options)
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": download_root}
    )

    return driver

# =====================================
# CEK JAM BOLEH LOGIN (08:00 - 22:00)
# server sistem mati di luar jam ini,
# tapi session yang sudah login tetap bisa jalan
# =====================================

def boleh_login():
    now = datetime.datetime.now().time()
    return datetime.time(8, 0) <= now <= datetime.time(22, 0)

# =====================================
# TUNGGU SAMPAI JAM 08:00
# =====================================

def tunggu_jam_buka():
    tulis_log("🌙 Session mati di jam tidur sistem, menunggu jam 08:00...")
    while True:
        if boleh_login():
            tulis_log("☀️ Jam sudah 08:00, siap login...")
            break
        time.sleep(60)

# =====================================
# LOGIN
# =====================================

def login(driver, wait):
    driver.get("http://172.100.10.24/rekan_brks")

    username = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text']")))
    password = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']")))

    username.clear()
    password.clear()
    username.send_keys("BRK020920")
    password.send_keys("z1z1z1z1")

    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Log In')]"))).click()
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'Dashboard')]")))

    tulis_log("✅ Login berhasil")

# =====================================
# CEK SESSION
# =====================================

def is_logged_in(driver):
    try:
        driver.find_element(By.XPATH, "//*[contains(text(),'Dashboard')]")
        return True
    except:
        return False

# =====================================
# SPLIT VENDOR
# =====================================

def split_vendor(file_path):
    df = pd.read_excel(file_path)
    df = df[["ID ATM", "Merk ATM", "Lokasi ATM", "Vendor", "Limit", "Sisa Saldo"]]
    df = df.sort_values(by="Sisa Saldo")

    vendor_map = {
        "Pekanbaru":      "PEKANBARU",
        "Batam":          "BATAM",
        "Dumai":          "DUMAI",
        "Tanjung Pinang": "TANJUNG PINANG"
    }

    base = os.path.dirname(file_path)
    name = os.path.basename(file_path)

    for key, folder in vendor_map.items():
        df_v = df[df["Vendor"].str.contains(key, case=False, na=False)]
        if df_v.empty:
            continue

        path = os.path.join(base, folder)
        os.makedirs(path, exist_ok=True)

        save = os.path.join(path, name)
        df_v.to_excel(save, index=False)

        wb = load_workbook(save)
        ws = wb.active
        ws.auto_filter.ref = ws.dimensions
        wb.save(save)

        tulis_log(f"✔ Vendor split: {folder}")

# =====================================
# BERSIHKAN FILE XLSX LAMA
# =====================================

def bersihkan_folder():
    for f in os.listdir(download_root):
        if f.endswith(".xlsx"):
            os.remove(os.path.join(download_root, f))

# =====================================
# TUNGGU DOWNLOAD SELESAI
# =====================================

def wait_download(timeout=120):
    start = time.time()
    while True:
        if time.time() - start > timeout:
            raise TimeoutError("❌ Download melebihi batas waktu 120 detik")

        files = os.listdir(download_root)
        if any(f.endswith(".crdownload") for f in files):
            time.sleep(1)
            continue
        if any(f.endswith(".xlsx") for f in files):
            return True

        time.sleep(1)

# =====================================
# DOWNLOAD
# =====================================

def download_file(driver, wait):
    jam = datetime.datetime.now().strftime("%H.%M")
    tulis_log(f"\n⬇️ Download jam: {jam}")

    # 1. Pergi ke Dashboard dulu (refresh navigasi)
    dashboard = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'Dashboard')]"))
    )
    driver.execute_script("arguments[0].click();", dashboard)
    time.sleep(2)

    # 2. Balik ke Notif Pengisian → data realtime refresh
    notif = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'Notif Pengisian')]"))
    )
    driver.execute_script("arguments[0].click();", notif)
    time.sleep(3)

    # 3. Bersihkan file lama sebelum export
    bersihkan_folder()

    # 4. Klik Export
    export = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Export to Vendor')]"))
    )
    driver.execute_script("arguments[0].click();", export)
    tulis_log("📥 Export ditekan")

    # 5. Tunggu download selesai
    wait_download()

    files = [f for f in os.listdir(download_root) if f.endswith(".xlsx")]
    if not files:
        tulis_log("❌ File tidak ditemukan setelah download")
        return None

    latest = max(
        [os.path.join(download_root, f) for f in files],
        key=os.path.getctime
    )

    # 6. Pindah ke folder tanggal dengan nama jam
    tanggal  = datetime.datetime.now().strftime("%Y-%m-%d")
    folder   = os.path.join(download_root, tanggal)
    os.makedirs(folder, exist_ok=True)

    new_name = f"Monitoring Saldo ATM BRKS {jam}.xlsx"
    new_path = os.path.join(folder, new_name)

    if os.path.exists(new_path):
        os.remove(new_path)

    shutil.move(latest, new_path)
    tulis_log(f"✅ Download selesai: {new_path}")

    # 7. Split per vendor
    split_vendor(new_path)

    return new_path

# =====================================
# TUNGGU KE JAM BERIKUTNYA (xx:01)
# =====================================

def wait_next():
    now      = datetime.datetime.now()
    next_run = (now + datetime.timedelta(hours=1)).replace(minute=1, second=0, microsecond=0)
    delay    = (next_run - now).total_seconds()
    tulis_log(f"⏳ Menunggu sampai {next_run.strftime('%H:%M:%S')}")
    time.sleep(delay)

# =====================================
# MAIN
# =====================================

tulis_log("🚀 Automation berjalan...\n")

driver = create_driver()
wait   = WebDriverWait(driver, 30)

# Login awal — tunggu sampai jam buka jika perlu
while True:
    if boleh_login():
        login(driver, wait)
        break
    else:
        tulis_log("⏳ Di luar jam operasional sistem, menunggu jam 08:00...")
        time.sleep(60)

# Loop utama — jalan 24 jam, download selama session aktif
while True:
    try:
        wait_next()

        # Cek session masih aktif atau tidak
        if not is_logged_in(driver):
            tulis_log("⚠️ Session mati...")

            # Kalau jam tidur sistem → tunggu sampai jam 08:00 dulu
            if not boleh_login():
                tunggu_jam_buka()

            # Restart driver dan login ulang
            tulis_log("♻️ Restart driver dan login ulang...")
            try:
                driver.quit()
            except:
                pass

            time.sleep(5)
            driver = create_driver()
            wait   = WebDriverWait(driver, 30)
            login(driver, wait)

        # Download tetap jalan 24 jam selama session aktif
        download_file(driver, wait)

    except Exception as e:
        tulis_log(f"❌ Error: {e}")

        if boleh_login():
            tulis_log("♻️ Restart driver...")
            try:
                driver.quit()
            except:
                pass

            time.sleep(5)
            driver = create_driver()
            wait   = WebDriverWait(driver, 30)
            login(driver, wait)
        else:
            # Jam tidur sistem, jangan coba login
            # Tunggu saja sampai jam 08:00
            tunggu_jam_buka()

            tulis_log("♻️ Restart driver setelah jam buka...")
            try:
                driver.quit()
            except:
                pass

            time.sleep(5)
            driver = create_driver()
            wait   = WebDriverWait(driver, 30)
            login(driver, wait)