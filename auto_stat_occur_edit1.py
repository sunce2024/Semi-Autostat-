
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
        df_stat = all_sheets["ปริมาณคดี"].fillna(0)

            # ดงเจริญ
        iid1x1x5_1215 = get_safe_int(df_stat, 12, 6)  
        iid1x19x5_1215 = get_safe_int(df_stat, 12, 10)
        iid1x20x5_1215 = get_safe_int(df_stat, 12, 12)
        iid1x8x5_1215 = get_safe_int(df_stat, 12, 13)
        # ตะพานหิน
        iid1x1x5_1206 = get_safe_int(df_stat, 5, 6)  
        iid1x19x5_1206 = get_safe_int(df_stat, 5, 10)
        iid1x20x5_1206 = get_safe_int(df_stat, 5, 12)
        iid1x8x5_1206 = get_safe_int(df_stat, 5, 13)
        # ทับคล้อ
        iid1x1x5_1207 = get_safe_int(df_stat, 9, 6)  
        iid1x19x5_1207 = get_safe_int(df_stat, 9, 10)
        iid1x20x5_1207 = get_safe_int(df_stat, 9, 12)
        iid1x8x5_1207 = get_safe_int(df_stat, 9, 13)
        # บางมูลนาก
        iid1x1x5_1209 = get_safe_int(df_stat, 6, 6)  
        iid1x19x5_1209 = get_safe_int(df_stat, 6, 10)
        iid1x20x5_1209 = get_safe_int(df_stat, 6, 12)
        iid1x8x5_1209 = get_safe_int(df_stat, 6, 13)
        # บึงนาราง
        iid1x1x5_1216 = get_safe_int(df_stat, 11, 6)  
        iid1x19x5_1216 = get_safe_int(df_stat, 11, 10)
        iid1x20x5_1216 = get_safe_int(df_stat, 11, 12)
        iid1x8x5_1216 = get_safe_int(df_stat, 11, 13)
        # โพทะเล
        iid1x1x5_1210 = get_safe_int(df_stat, 7, 6)  
        iid1x19x5_1210 = get_safe_int(df_stat, 7, 10)
        iid1x20x5_1210 = get_safe_int(df_stat, 7, 12)
        iid1x8x5_1210 = get_safe_int(df_stat, 7, 13)
        # โพธิ์ประทับช้าง
        iid1x1x5_1212 = get_safe_int(df_stat, 4, 6)  
        iid1x19x5_1212 = get_safe_int(df_stat, 4, 10)
        iid1x20x5_1212 = get_safe_int(df_stat, 4, 12)
        iid1x8x5_1212 = get_safe_int(df_stat, 4, 13)
        # เมืองพิจิตร
        iid1x1x5_1205 = get_safe_int(df_stat, 2, 6)  
        iid1x19x5_1205 = get_safe_int(df_stat, 2, 10)
        iid1x20x5_1205 = get_safe_int(df_stat, 2, 12)
        iid1x8x5_1205 = get_safe_int(df_stat, 2, 13)
        # วชิรบารมี
        iid1x1x5_1213 = get_safe_int(df_stat, 13, 6)  
        iid1x19x5_1213 = get_safe_int(df_stat, 13, 10)
        iid1x20x5_1213 = get_safe_int(df_stat, 13, 12)
        iid1x8x5_1213 = get_safe_int(df_stat, 13, 13)
        # วังทรายพูน
        iid1x1x5_1208 = get_safe_int(df_stat, 3, 6)  
        iid1x19x5_1208 = get_safe_int(df_stat, 3, 10)
        iid1x20x5_1208 = get_safe_int(df_stat, 3, 12)
        iid1x8x5_1208 = get_safe_int(df_stat, 3, 13)
        # สาวเหล็ก
        iid1x1x5_1214 = get_safe_int(df_stat, 10, 6)  
        iid1x19x5_1214 = get_safe_int(df_stat, 10, 10)
        iid1x20x5_1214 = get_safe_int(df_stat, 10, 12)
        iid1x8x5_1214 = get_safe_int(df_stat, 10, 13)
        # สามง่าม
        iid1x1x5_1211 = get_safe_int(df_stat, 8, 6)  
        iid1x19x5_1211 = get_safe_int(df_stat, 8, 10)
        iid1x20x5_1211 = get_safe_int(df_stat, 8, 12)
        iid1x8x5_1211 = get_safe_int(df_stat, 8, 13)

        

        # --- 2. เปลี่ยนไปยังหน้าฟอร์ม fid=27 ---
        print("กำลังเปลี่ยนไปหน้าฟอร์ม fid=31...")
        driver.get(
            "http://10.1.2.9/court_stat_program/form_location.php?fid=31"
        )

        # ตั้งเวลารอ Element โหลดสูงสุด 30 วินาที
        wait = WebDriverWait(driver, 30)

        # --- 3. เริ่มกรอกข้อมูลอัตโนมัติ ---
        print("กำลังกรอกข้อมูลอัตโนมัติ...")

        # ดงเจริญ
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x1x5_1215")))
        input_1.clear()
        input_1.send_keys(str(iid1x1x5_1215))
        input_2 = driver.find_element(By.NAME, "id1x19x5_1215")
        input_2.clear()
        input_2.send_keys(str(iid1x19x5_1215))
        input_2 = driver.find_element(By.NAME, "id1x20x5_1215")
        input_2.clear()
        input_2.send_keys(str(iid1x20x5_1215))
        input_2 = driver.find_element(By.NAME, "id1x8x5_1215")
        input_2.clear()
        input_2.send_keys(str(iid1x8x5_1215))

        # ตะพานหิน
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x1x5_1206")))
        input_1.clear()
        input_1.send_keys(str(iid1x1x5_1206))
        input_2 = driver.find_element(By.NAME, "id1x19x5_1206")
        input_2.clear()
        input_2.send_keys(str(iid1x19x5_1206))
        input_2 = driver.find_element(By.NAME, "id1x20x5_1206")
        input_2.clear()
        input_2.send_keys(str(iid1x20x5_1206))
        input_2 = driver.find_element(By.NAME, "id1x8x5_1206")
        input_2.clear()
        input_2.send_keys(str(iid1x8x5_1206))

        # ทับคล้อ
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x1x5_1207")))
        input_1.clear()
        input_1.send_keys(str(iid1x1x5_1207))
        input_2 = driver.find_element(By.NAME, "id1x19x5_1207")
        input_2.clear()
        input_2.send_keys(str(iid1x19x5_1207))
        input_2 = driver.find_element(By.NAME, "id1x20x5_1207")
        input_2.clear()
        input_2.send_keys(str(iid1x20x5_1207))
        input_2 = driver.find_element(By.NAME, "id1x8x5_1207")
        input_2.clear()
        input_2.send_keys(str(iid1x8x5_1207))

        # บางมูลนาก
        input_2 = driver.find_element(By.NAME, "id1x1x5_1209")
        input_2.clear()
        input_2.send_keys(str(iid1x1x5_1209))    
        input_2 = driver.find_element(By.NAME, "id1x19x5_1209")
        input_2.clear()
        input_2.send_keys(str(iid1x19x5_1209))
        input_2 = driver.find_element(By.NAME, "id1x20x5_1209")
        input_2.clear()
        input_2.send_keys(str(iid1x20x5_1209))
        input_2 = driver.find_element(By.NAME, "id1x8x5_1209")
        input_2.clear()
        input_2.send_keys(str(iid1x8x5_1209))

        # บึงนาราง
        input_2 = driver.find_element(By.NAME, "id1x1x5_1216")
        input_2.clear()
        input_2.send_keys(str(iid1x1x5_1216))    
        input_2 = driver.find_element(By.NAME, "id1x19x5_1216")
        input_2.clear()
        input_2.send_keys(str(iid1x19x5_1216))
        input_2 = driver.find_element(By.NAME, "id1x20x5_1216")
        input_2.clear()
        input_2.send_keys(str(iid1x20x5_1216))
        input_2 = driver.find_element(By.NAME, "id1x8x5_1216")
        input_2.clear()
        input_2.send_keys(str(iid1x8x5_1216))

        # โพธิ์ทะเล
        input_2 = driver.find_element(By.NAME, "id1x1x5_1210")
        input_2.clear()
        input_2.send_keys(str(iid1x1x5_1210))    
        input_2 = driver.find_element(By.NAME, "id1x19x5_1210")
        input_2.clear()
        input_2.send_keys(str(iid1x19x5_1210))
        input_2 = driver.find_element(By.NAME, "id1x20x5_1210")
        input_2.clear()
        input_2.send_keys(str(iid1x20x5_1210))
        input_2 = driver.find_element(By.NAME, "id1x8x5_1210")
        input_2.clear()
        input_2.send_keys(str(iid1x8x5_1210))

        # โพธิ์ประทับช้าง
        input_2 = driver.find_element(By.NAME, "id1x1x5_1212")
        input_2.clear()
        input_2.send_keys(str(iid1x1x5_1212))    
        input_2 = driver.find_element(By.NAME, "id1x19x5_1212")
        input_2.clear()
        input_2.send_keys(str(iid1x19x5_1212))
        input_2 = driver.find_element(By.NAME, "id1x20x5_1212")
        input_2.clear()
        input_2.send_keys(str(iid1x20x5_1212))
        input_2 = driver.find_element(By.NAME, "id1x8x5_1212")
        input_2.clear()
        input_2.send_keys(str(iid1x8x5_1212))

        # เมืองพิจิตร
        input_2 = driver.find_element(By.NAME, "id1x1x5_1205")
        input_2.clear()
        input_2.send_keys(str(iid1x1x5_1205))    
        input_2 = driver.find_element(By.NAME, "id1x19x5_1205")
        input_2.clear()
        input_2.send_keys(str(iid1x19x5_1205))
        input_2 = driver.find_element(By.NAME, "id1x20x5_1205")
        input_2.clear()
        input_2.send_keys(str(iid1x20x5_1205))
        input_2 = driver.find_element(By.NAME, "id1x8x5_1205")
        input_2.clear()
        input_2.send_keys(str(iid1x8x5_1205))

        # วชิรบารมี
        input_2 = driver.find_element(By.NAME, "id1x1x5_1213")
        input_2.clear()
        input_2.send_keys(str(iid1x1x5_1213))    
        input_2 = driver.find_element(By.NAME, "id1x19x5_1213")
        input_2.clear()
        input_2.send_keys(str(iid1x19x5_1213))
        input_2 = driver.find_element(By.NAME, "id1x20x5_1213")
        input_2.clear()
        input_2.send_keys(str(iid1x20x5_1213))
        input_2 = driver.find_element(By.NAME, "id1x8x5_1213")
        input_2.clear()
        input_2.send_keys(str(iid1x8x5_1213))

        # วังทรายพูน
        input_2 = driver.find_element(By.NAME, "id1x1x5_1208")
        input_2.clear()
        input_2.send_keys(str(iid1x1x5_1208))    
        input_2 = driver.find_element(By.NAME, "id1x19x5_1208")
        input_2.clear()
        input_2.send_keys(str(iid1x19x5_1208))
        input_2 = driver.find_element(By.NAME, "id1x20x5_1208")
        input_2.clear()
        input_2.send_keys(str(iid1x20x5_1208))
        input_2 = driver.find_element(By.NAME, "id1x8x5_1208")
        input_2.clear()
        input_2.send_keys(str(iid1x8x5_1208))

        # สากเหล็ก
        input_2 = driver.find_element(By.NAME, "id1x1x5_1214")
        input_2.clear()
        input_2.send_keys(str(iid1x1x5_1214))    
        input_2 = driver.find_element(By.NAME, "id1x19x5_1214")
        input_2.clear()
        input_2.send_keys(str(iid1x19x5_1214))
        input_2 = driver.find_element(By.NAME, "id1x20x5_1214")
        input_2.clear()
        input_2.send_keys(str(iid1x20x5_1214))
        input_2 = driver.find_element(By.NAME, "id1x8x5_1214")
        input_2.clear()
        input_2.send_keys(str(iid1x8x5_1214))

        # สามง่าม
        input_2 = driver.find_element(By.NAME, "id1x1x5_1211")
        input_2.clear()
        input_2.send_keys(str(iid1x1x5_1211))    
        input_2 = driver.find_element(By.NAME, "id1x19x5_1211")
        input_2.clear()
        input_2.send_keys(str(iid1x19x5_1211))
        input_2 = driver.find_element(By.NAME, "id1x20x5_1211")
        input_2.clear()
        input_2.send_keys(str(iid1x20x5_1211))
        input_2 = driver.find_element(By.NAME, "id1x8x5_1211")
        input_2.clear()
        input_2.send_keys(str(iid1x8x5_1211))

        end_time = time.time()
        elapsed_time = round(end_time - start_time, 2)
        msg = (
            f"✅ กรอกข้อมูลคดีเกิดเสร็จเรียบร้อย!\n"
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
    """ขั้นตอนที่ 3: ปิด Browser เมื่อทำงานเสร็จสิ้น"""
    global driver
    if driver:
        driver.quit()
        driver = None
        return "🔒 ปิด Browser เรียบร้อยแล้ว"
    return "⚠️ ไม่มี Browser เปิดูอยู่"