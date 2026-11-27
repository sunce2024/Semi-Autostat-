
import os
import time
import numpy as np
import pandas as pd
import streamlit as st
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# ตัวแปร Global สำหรับเก็บ Browser Session
driver = None


# --- ฟังก์ชันช่วยดึงค่าและป้องกัน Error NaN ---
def get_safe_int(df, row, col):
    try:
        val = df.iloc[row, col]
        if pd.isna(val):
            return 0
        return int(float(val))
    except Exception:
        return 0


# --- ฟังก์ชันช่วยดึง Driver หรือเปิดใหม่ถ้ายังไม่มี ---
def get_or_create_driver():
    """เช็กว่ามี Driver ใน Session หรือยัง ถ้าไม่มีหรือถูกปิดไปแล้วจะสร้างใหม่"""
    if "driver" in st.session_state and st.session_state.driver is not None:
        try:
            # ทดสอบว่าเบราว์เซอร์ยังเปิดอยู่จริงไหม
            _ = st.session_state.driver.current_url
            return st.session_state.driver
        except Exception:
            # ถ้าถูกปิดไปแล้ว ให้ตั้งเป็น None เพื่อสร้างใหม่
            st.session_state.driver = None

    # สร้าง Driver ใหม่
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    st.session_state.driver = webdriver.Chrome(options=options)
    return st.session_state.driver


# ==========================================
# ขั้นตอนที่ 1: เปิด Browser และไปหน้า Login
# ==========================================
def step1_open_login():
    """ขั้นตอนที่ 1: เปิดหน้าเว็บ Login"""
    try:
        driver = get_or_create_driver()

        # เช็กว่าอยู่ในระบบแล้วหรือยัง (ไม่อยู่หน้า login.php)
        if "login.php" not in driver.current_url.lower():
            return "✅ เบราว์เซอร์เปิดอยู่และอยู่ในระบบแล้ว! คุณสามารถกดขั้นตอนที่ 2 เพื่อเริ่มกรอกข้อมูลได้เลยครับ"

        driver.get("http://10.1.2.9/court_stat_program/login.php")
        return "🌐 เปิดหน้าเว็บเรียบร้อยแล้ว! กรุณา Login บนหน้าต่าง Chrome ให้เสร็จสิ้น แล้วค่อยกดปุ่มขั้นตอนที่ 2"

    except Exception as e:
        return f"❌ เกิดข้อผิดพลาดในการเปิดหน้าเว็บ: {e}"


# ==========================================
# ขั้นตอนที่ 2: อ่าน Excel และกรอกข้อมูลอัตโนมัติ
# ==========================================
def step2_fill_data():
    """ขั้นตอนที่ 2: กรอกข้อมูล (ตรวจเช็กสถานะ Login แม่นยำขึ้น)"""
    start_time = time.time()  # ⏱️ เริ่มจับเวลา
    try:
        # 1. ดึง Driver ที่มีอยู่ หรือเปิดใหม่ถ้าไม่มี
        driver = get_or_create_driver()

        # 2. ตรวจสอบว่าอยู่หน้า Login หรือไม่
        # ถ้า URL มีคำว่า login.php แสดงว่าผู้ใช้ยังไม่ได้กดยืนยัน Login หรือ Login ไม่ผ่าน
        current_url = driver.current_url.lower()

        if "login.php" in current_url or current_url == "about:blank":
            # สั่งเปิดหน้า Login เผื่อกรณีเบราว์เซอร์ยังไม่ได้เปิดหน้าเว็บ
            if "court_stat_program" not in current_url:
                driver.get("http://10.1.2.9/court_stat_program/login.php")

            return "⚠️ เบราว์เซอร์ยังอยู่ที่หน้า Login! โปรดพิมพ์ Username/Password กดเข้าสู่ระบบบนหน้าต่าง Chrome ให้เรียบร้อยก่อน แล้วค่อยกด 'ขั้นตอนที่ 2' อีกครั้งครับ"
        # --- 1. อ่านข้อมูลจาก Excel ---
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(BASE_DIR, "stat_rpt_month_book.xls")

        print("กำลังอ่านไฟล์ Excel...")
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine="xlrd")
        df_stat = all_sheets["บัญชีคดีสิ่งแวดล้อมรายเดือน"].fillna(0)

        # ดึงค่าตัวแปรจากเซลล์ต่างๆ
        iid1x3_2 = get_safe_int(df_stat, 3, 9)
        iid1x3_4 = get_safe_int(df_stat, 6, 9)
        iid1x3_5 = get_safe_int(df_stat, 8, 9)
        iid1x3_6 = get_safe_int(df_stat, 9, 9)
        iid1x3_7 = get_safe_int(df_stat, 11, 9)
        iid1x3_9 = get_safe_int(df_stat, 14, 9)
        iid1x3_10 = get_safe_int(df_stat, 15, 9)
        iid1x3_11 = get_safe_int(df_stat, 17, 9)
        iid1x3_12 = get_safe_int(df_stat, 19, 9)

        iid1x3_24 = get_safe_int(df_stat, 20, 9)
        iid1x3_15 = get_safe_int(df_stat, 22, 9)
        iid1x3_16 = get_safe_int(df_stat, 23, 9)
        iid1x3_17 = get_safe_int(df_stat, 24, 9)

        iid1x3_25 = get_safe_int(df_stat, 25, 9)
        iid1x3_19 = get_safe_int(df_stat, 5, 22)
        iid1x3_20 = get_safe_int(df_stat, 7, 22)
        iid1x3_21 = get_safe_int(df_stat, 10, 22)
        iid1x3_22 = get_safe_int(df_stat, 12, 22)
        iid1x3_23 = get_safe_int(df_stat, 14, 22)
        
            

        print(
            f"ดึงข้อมูลสำเร็จ -> iid1x3_2: {iid1x3_2}, iid1x3_4: {iid1x3_4}, iid1x3_5 : {iid1x3_5 }, iid1x3_6 {iid1x3_6}"
        )

        # --- 2. เปลี่ยนไปยังหน้าฟอร์ม fid=27 ---
        print("กำลังเปลี่ยนไปหน้าฟอร์ม fid=27...")
        driver.get(
            "http://10.1.2.9/court_stat_program/form_case.php?fid=152"
        )

        # ตั้งเวลารอ Element โหลดสูงสุด 30 วินาที
        wait = WebDriverWait(driver, 30)

        # --- 3. เริ่มกรอกข้อมูลอัตโนมัติ ---
        print("กำลังกรอกข้อมูลอัตโนมัติ...")

        input_1 = wait.until(
            EC.presence_of_element_located((By.NAME, "id1x3_2"))
        )
        input_1.clear()
        input_1.send_keys(str(iid1x3_2))

        input_2 = wait.until(
            EC.presence_of_element_located((By.NAME, "id1x3_4"))
        )
        input_2.clear()
        input_2.send_keys(str(iid1x3_4))

        input_3 = wait.until(
            EC.presence_of_element_located((By.NAME, "id1x3_5"))
        )
        input_3.clear()
        input_3.send_keys(str(iid1x3_5))

        input_4 = wait.until(
            EC.presence_of_element_located((By.NAME, "id1x3_6"))
        )
        input_4.clear()
        input_4.send_keys(str(iid1x3_6))
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_7")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_7))
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_9")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_9))
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_10")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_10))
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_11")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_11))
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_12")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_12))
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_24")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_24))
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_15")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_15))
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_16")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_16))
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_17")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_17))
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_25")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_25))
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_19")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_19))
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_20")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_20))
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_21")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_21))
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_22")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_22))
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_23")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_23))

        end_time = time.time()
        elapsed_time = round(end_time - start_time, 2)
        msg = (
            f"✅ กรอกข้อมูลคดีสิ่งแวดล้อมเสร็จเรียบร้อย!\n"
            f"📊 ผลลัพธ์: (ใช้เวลาป้อนข้อมูล:*** {elapsed_time}*** วินาที)\n"
            f"📌 กรุณาตรวจสอบความถูกต้องบนหน้าเว็บ แล้วกดปุ่มบันทึกด้วยตัวเองครับ"
        )
        print(msg)
        return msg

    except Exception as e:
        error_msg = f"❌ เกิดข้อผิดพลาดขณะกรอกข้อมูล: {e}"
        print(error_msg)
        return error_msg


# ==========================================
# ขั้นตอนที่ 3: ปิด Browser
# ==========================================
def step3_close_browser():
    """ขั้นตอนที่ 3: ปิด Browser และคืนทรัพยากรระบบ"""
    global driver
    closed_success = False

    # 1. พยายามปิดผ่าน Selenium Object
    if driver is not None:
        try:
            driver.quit()
            closed_success = True
        except Exception:
            # หากผู้ใช้ปิดหน้าต่าง Chrome ไปก่อนแล้ว
            pass
        finally:
            driver = None

    # 2. ปิด Process chromedriver ที่อาจค้างอยู่ใน Background (สำหรับ Windows)
    try:
        os.system("taskkill /f /im chromedriver.exe /t >nul 2>&1")
    except Exception:
        pass

    if closed_success:
        return "🔒 ปิด Browser เรียบร้อยแล้ว"
    else:
        return "🔒 เคลียร์และปิด Browser เรียบร้อยแล้ว"