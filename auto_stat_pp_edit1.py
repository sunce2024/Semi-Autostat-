
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
        current_url = driver.current_url.lower()

        if "login.php" in current_url or current_url == "about:blank":
            # สั่งเปิดหน้า Login เผื่อกรณีเบราว์เซอร์ยังไม่ได้เปิดหน้าเว็บ
            if "court_stat_program" not in current_url:
                driver.get("http://10.1.2.9/court_stat_program/login.php")

            # ✅ คืนค่าแจ้งเตือนเฉพาะตอนที่อยู่หน้า Login จริงๆ เท่านั้น
            return "⚠️ เบราว์เซอร์ยังอยู่ที่หน้า Login! โปรดพิมพ์ Username/Password กดเข้าสู่ระบบบนหน้าต่าง Chrome ให้เรียบร้อยก่อน แล้วค่อยกด 'ขั้นตอนที่ 2' อีกครั้งครับ"
        # --- 1. อ่านข้อมูลจาก Excel ---
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(BASE_DIR, "stat_rpt_month_book.xls")

        print("กำลังอ่านไฟล์ Excel...")
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine="xlrd")
        df_stat = all_sheets["บัญชีคดีผู้บริโภครายเดือน"].fillna(0)

        #-------------- แถวผู้บริโภคเป็นโจทก์---------------------
    #หัวข้อซื้อขาย
        iid1x7_4 = get_safe_int(df_stat, 5, 14)  
        iid1x7_5 = get_safe_int(df_stat, 6, 14)
        iid1x7_6 = get_safe_int(df_stat, 7, 14)  
        iid1x7_1001 = get_safe_int(df_stat, 8, 14) 
        iid1x7_1002 = get_safe_int(df_stat, 9, 14) 
        iid1x7_1003 = get_safe_int(df_stat, 10, 14) 
    #หัวข้อเช่าทรัพย์
        iid1x7_1005 = get_safe_int(df_stat, 12, 14) 
        iid1x7_1006 = get_safe_int(df_stat, 13, 14) 
        iid1x7_2107 = get_safe_int(df_stat, 14, 14) 
        iid1x7_2108 = get_safe_int(df_stat, 15, 14) 
        iid1x7_2109 = get_safe_int(df_stat, 16, 14)
        iid1x7_2110 = get_safe_int(df_stat, 17, 14)
        iid1x7_2111 = get_safe_int(df_stat, 18, 14)
    #หัวข้อเช่าซื้อ
        iid1x7_2113 = get_safe_int(df_stat, 20, 14)
        iid1x7_2114 = get_safe_int(df_stat, 21, 14)
        iid1x7_2115 = get_safe_int(df_stat, 22, 14)
        iid1x7_2116 = get_safe_int(df_stat, 23, 14)
    #หัวข้อรับขน
        iid1x7_2118 = get_safe_int(df_stat, 25, 14)
        iid1x7_2119 = get_safe_int(df_stat, 26, 14)
        iid1x7_2120 = get_safe_int(df_stat, 27, 14)
    #หัวข้อสินเชื่อกู้ยืม
        iid1x7_2122 = get_safe_int(df_stat, 29, 14)
        iid1x7_2123 = get_safe_int(df_stat, 30, 14)
        iid1x7_2124 = get_safe_int(df_stat, 31, 14)
        iid1x7_2125 = get_safe_int(df_stat, 32, 14)
    #หัวข้อกองทุน
        iid1x7_2158 = get_safe_int(df_stat, 35, 14)
    # กรอ ไม่มีใน EXCEL iid1x7_2159 = get_safe_int(df_stat, 36, 7)
    # ไม่ระบุกองทุน ไม่มีใน EXCEL iid1x7_2160 = get_safe_int(df_stat, 37, 7)
        iid1x7_2128 = get_safe_int(df_stat, 37, 14)   
        iid1x7_2129 = get_safe_int(df_stat, 38, 14) 
    #หัวข้อสินค้าไม่ปลอดภัย
        iid1x7_2131 = get_safe_int(df_stat, 40, 14) 
        iid1x7_2132 = get_safe_int(df_stat, 41, 14) 
        iid1x7_2133 = get_safe_int(df_stat, 42, 14) 
        iid1x7_2134 = get_safe_int(df_stat, 43, 14) 
        iid1x7_2135 = get_safe_int(df_stat, 44, 14) 
        iid1x7_2136 = get_safe_int(df_stat, 45, 14) 
        iid1x7_2137 = get_safe_int(df_stat, 46, 14) 
        iid1x7_2138 = get_safe_int(df_stat, 47, 14) 
        iid1x7_2139 = get_safe_int(df_stat, 48, 14) 
    #หัวข้อสาธารณูปโภค
        iid1x7_2141 = get_safe_int(df_stat, 50, 14) 
        iid1x7_2142 = get_safe_int(df_stat, 51, 14) 
        iid1x7_2143 = get_safe_int(df_stat, 52, 14)
    #หัวข้อความงาม
        iid1x7_2145 = get_safe_int(df_stat, 54, 14)
        iid1x7_2146 = get_safe_int(df_stat, 55, 14)
        iid1x7_2147 = get_safe_int(df_stat, 56, 14)
        iid1x7_2148 = get_safe_int(df_stat, 57, 14)
    #หัวข้อบริการท่องเที่ยว
        iid1x7_2150 = get_safe_int(df_stat, 64, 14)
        iid1x7_2151 = get_safe_int(df_stat, 65, 14)
        iid1x7_2152 = get_safe_int(df_stat, 66, 14)
        iid1x7_2153 = get_safe_int(df_stat, 65, 14)
        iid1x7_2154 = get_safe_int(df_stat, 66, 14)
        iid1x7_2155 = get_safe_int(df_stat, 67, 14)
        iid1x7_2156 = get_safe_int(df_stat, 68, 14)
        iid1x7_2157 = get_safe_int(df_stat, 69, 14)

    #-------------- แถวผู้ประกอบธุรกิจเป็นโจทก์---------------------

    #หัวข้อซื้อขาย
        iid1x8_4 = get_safe_int(df_stat, 5, 23)  
        iid1x8_5 = get_safe_int(df_stat, 6, 23)
        iid1x8_6 = get_safe_int(df_stat, 7, 23)  
        iid1x8_1001 = get_safe_int(df_stat, 8, 23) 
        iid1x8_1002 = get_safe_int(df_stat, 9, 23) 
        iid1x8_1003 = get_safe_int(df_stat, 10, 23) 
    #หัวข้อเช่าทรัพย์
        iid1x8_1005 = get_safe_int(df_stat, 12, 23) 
        iid1x8_1006 = get_safe_int(df_stat, 13, 23) 
        iid1x8_2107 = get_safe_int(df_stat, 14, 23) 
        iid1x8_2108 = get_safe_int(df_stat, 15, 23) 
        iid1x8_2109 = get_safe_int(df_stat, 16, 23)
        iid1x8_2110 = get_safe_int(df_stat, 17, 23)
        iid1x8_2111 = get_safe_int(df_stat, 18, 23)
    #หัวข้อเช่าซื้อ
        iid1x8_2113 = get_safe_int(df_stat, 20, 23)
        iid1x8_2114 = get_safe_int(df_stat, 21, 23)
        iid1x8_2115 = get_safe_int(df_stat, 22, 23)
        iid1x8_2116 = get_safe_int(df_stat, 23, 23)
    #หัวข้อรับขน
        iid1x8_2118 = get_safe_int(df_stat, 25, 23)
        iid1x8_2119 = get_safe_int(df_stat, 26, 23)
        iid1x8_2120 = get_safe_int(df_stat, 27, 23)
    #หัวข้อสินเชื่อกู้ยืม
        iid1x8_2122 = get_safe_int(df_stat, 28, 23)
        iid1x8_2123 = get_safe_int(df_stat, 30, 23)
        iid1x8_2124 = get_safe_int(df_stat, 31, 23)
        iid1x8_2125 = get_safe_int(df_stat, 29, 23)
    #หัวข้อกองทุน
        iid1x8_2158 = get_safe_int(df_stat, 35, 23)
    # กรอ ไม่มีใน EXCEL iid1x8_2159 = get_safe_int(df_stat, 36, 11)
    # ไม่ระบุกองทุน ไม่มีใน EXCEL iid1x8_2160 = get_safe_int(df_stat, 37, 11)
        iid1x8_2128 = get_safe_int(df_stat, 37, 23)   
        iid1x8_2129 = get_safe_int(df_stat, 38, 23) 
    #หัวข้อสินค้าไม่ปลอดภัย
        iid1x8_2131 = get_safe_int(df_stat, 40, 23) 
        iid1x8_2132 = get_safe_int(df_stat, 41, 23) 
        iid1x8_2133 = get_safe_int(df_stat, 42, 23) 
        iid1x8_2134 = get_safe_int(df_stat, 43, 23) 
        iid1x8_2135 = get_safe_int(df_stat, 44, 23) 
        iid1x8_2136 = get_safe_int(df_stat, 45, 23) 
        iid1x8_2137 = get_safe_int(df_stat, 46, 23) 
        iid1x8_2138 = get_safe_int(df_stat, 47, 23) 
        iid1x8_2139 = get_safe_int(df_stat, 48, 23) 
    #หัวข้อสาธารณูปโภค
        iid1x8_2141 = get_safe_int(df_stat, 50, 23) 
        iid1x8_2142 = get_safe_int(df_stat, 51, 23) 
        iid1x8_2143 = get_safe_int(df_stat, 52, 23)
    #หัวข้อความงาม
        iid1x8_2145 = get_safe_int(df_stat, 54, 23)
        iid1x8_2146 = get_safe_int(df_stat, 55, 23)
        iid1x8_2147 = get_safe_int(df_stat, 56, 23)
        iid1x8_2148 = get_safe_int(df_stat, 57, 23)
    #หัวข้อบริการท่องเที่ยว
        iid1x8_2150 = get_safe_int(df_stat, 64, 23)
        iid1x8_2151 = get_safe_int(df_stat, 65, 23)
        iid1x8_2152 = get_safe_int(df_stat, 66, 23)
        iid1x8_2153 = get_safe_int(df_stat, 65, 23)
        iid1x8_2154 = get_safe_int(df_stat, 66, 23)
        iid1x8_2155 = get_safe_int(df_stat, 67, 23)
        iid1x8_2156 = get_safe_int(df_stat, 68, 23)
        iid1x8_2157 = get_safe_int(df_stat, 69, 23)


    #-------------- แถวผู้มีอำนาจฟ้องคดีแทนเป็นโจทก์---------------------


    #หัวข้อซื้อขาย
        iid1x9_4 = get_safe_int(df_stat, 5, 27)  
        iid1x9_5 = get_safe_int(df_stat, 6, 27) 
        iid1x9_6 = get_safe_int(df_stat, 7, 27)  
        iid1x9_1001 = get_safe_int(df_stat, 8, 27) 
        iid1x9_1002 = get_safe_int(df_stat, 9, 27) 
        iid1x9_1003 = get_safe_int(df_stat, 10, 27) 
    #หัวข้อเช่าทรัพย์
        iid1x9_1005 = get_safe_int(df_stat, 12, 27) 
        iid1x9_1006 = get_safe_int(df_stat, 13, 27) 
        iid1x9_2107 = get_safe_int(df_stat, 14, 27) 
        iid1x9_2108 = get_safe_int(df_stat, 15, 27) 
        iid1x9_2109 = get_safe_int(df_stat, 16, 27) 
        iid1x9_2110 = get_safe_int(df_stat, 17, 27) 
        iid1x9_2111 = get_safe_int(df_stat, 18, 27) 
    #หัวข้อเช่าซื้อ
        iid1x9_2113 = get_safe_int(df_stat, 20, 27) 
        iid1x9_2114 = get_safe_int(df_stat, 21, 27) 
        iid1x9_2115 = get_safe_int(df_stat, 22, 27) 
        iid1x9_2116 = get_safe_int(df_stat, 23, 27) 
    #หัวข้อรับขน
        iid1x9_2118 = get_safe_int(df_stat, 25, 27) 
        iid1x9_2119 = get_safe_int(df_stat, 26, 27) 
        iid1x9_2120 = get_safe_int(df_stat, 27, 27) 
    #หัวข้อสินเชื่อกู้ยืม
        iid1x9_2122 = get_safe_int(df_stat, 29, 27) 
        iid1x9_2123 = get_safe_int(df_stat, 30, 27) 
        iid1x9_2124 = get_safe_int(df_stat, 31, 27) 
        iid1x9_2125 = get_safe_int(df_stat, 32, 27) 
    #หัวข้อกองทุน
        iid1x9_2158 = get_safe_int(df_stat, 35, 27) 
    # กรอ ไม่มีใน EXCEL iid1x9_2159 = get_safe_int(df_stat, 36, 16) 
    # ไม่ระบุกองทุน ไม่มีใน EXCEL iid1x9_2160 = get_safe_int(df_stat, 37, 16) 
        iid1x9_2128 = get_safe_int(df_stat, 37, 27)   
        iid1x9_2129 = get_safe_int(df_stat, 38, 27) 
    #หัวข้อสินค้าไม่ปลอดภัย
        iid1x9_2131 = get_safe_int(df_stat, 40, 27) 
        iid1x9_2132 = get_safe_int(df_stat, 41, 27) 
        iid1x9_2133 = get_safe_int(df_stat, 42, 27) 
        iid1x9_2134 = get_safe_int(df_stat, 43, 27) 
        iid1x9_2135 = get_safe_int(df_stat, 44, 27) 
        iid1x9_2136 = get_safe_int(df_stat, 45, 27) 
        iid1x9_2137 = get_safe_int(df_stat, 46, 27) 
        iid1x9_2138 = get_safe_int(df_stat, 47, 27) 
        iid1x9_2139 = get_safe_int(df_stat, 48, 27) 
    #หัวข้อสาธารณูปโภค
        iid1x9_2141 = get_safe_int(df_stat, 50, 27) 
        iid1x9_2142 = get_safe_int(df_stat, 51, 27) 
        iid1x9_2143 = get_safe_int(df_stat, 52, 27) 
    #หัวข้อความงาม
        iid1x9_2145 = get_safe_int(df_stat, 54, 27) 
        iid1x9_2146 = get_safe_int(df_stat, 55, 27) 
        iid1x9_2147 = get_safe_int(df_stat, 56, 27) 
        iid1x9_2148 = get_safe_int(df_stat, 57, 27) 
    #หัวข้อบริการท่องเที่ยว
        iid1x9_2150 = get_safe_int(df_stat, 64, 27) 
        iid1x9_2151 = get_safe_int(df_stat, 65, 27) 
        iid1x9_2152 = get_safe_int(df_stat, 66, 27) 
        iid1x9_2153 = get_safe_int(df_stat, 65, 27) 
        iid1x9_2154 = get_safe_int(df_stat, 66, 27) 
        iid1x9_2155 = get_safe_int(df_stat, 67, 27) 
        iid1x9_2156 = get_safe_int(df_stat, 68, 27) 
        iid1x9_2157 = get_safe_int(df_stat, 69, 27) 


        

        # --- 2. เปลี่ยนไปยังหน้าฟอร์ม fid=27 ---
        print("กำลังเปลี่ยนไปหน้าฟอร์ม fid=103...")
        driver.get(
            "http://10.1.2.9/court_stat_program/form_case.php?fid=103"
        )

        # ตั้งเวลารอ Element โหลดสูงสุด 30 วินาที
        wait = WebDriverWait(driver, 30)

        # --- 3. เริ่มกรอกข้อมูลอัตโนมัติ ---
        print("กำลังกรอกข้อมูลอัตโนมัติ...")

        # ---------------กรอกคดีผู้บริโภค ผู้บริโภคเป็นโจทก์------------------
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x7_4")))
        input_1.clear()
        input_1.send_keys(str(iid1x7_4 ))

        input_2 = driver.find_element(By.NAME, "id1x7_5")
        input_2.clear()
        input_2.send_keys(str(iid1x7_5))
        input_2 = driver.find_element(By.NAME, "id1x7_6")
        input_2.clear()
        input_2.send_keys(str(iid1x7_6))

        input_2 = driver.find_element(By.NAME, "id1x7_1001")
        input_2.clear()
        input_2.send_keys(str(iid1x7_1001))
        input_2 = driver.find_element(By.NAME, "id1x7_1002")
        input_2.clear()
        input_2.send_keys(str(iid1x7_1002))
        input_2 = driver.find_element(By.NAME, "id1x7_1003")
        input_2.clear()
        input_2.send_keys(str(iid1x7_1003))

        input_2 = driver.find_element(By.NAME, "id1x7_1005")
        input_2.clear()
        input_2.send_keys(str(iid1x7_1005))
        input_2 = driver.find_element(By.NAME, "id1x7_1006")
        input_2.clear()
        input_2.send_keys(str(iid1x7_1006))
        input_2 = driver.find_element(By.NAME, "id1x7_2107")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2107))
        input_2 = driver.find_element(By.NAME, "id1x7_2108")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2108))
        input_2 = driver.find_element(By.NAME, "id1x7_2109")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2109))
        input_2 = driver.find_element(By.NAME, "id1x7_2110")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2110))
        input_2 = driver.find_element(By.NAME, "id1x7_2111")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2111))

        input_2 = driver.find_element(By.NAME, "id1x7_2113")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2113))
        input_2 = driver.find_element(By.NAME, "id1x7_2114")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2114))
        input_2 = driver.find_element(By.NAME, "id1x7_2115")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2115))
        input_2 = driver.find_element(By.NAME, "id1x7_2116")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2116))

        input_2 = driver.find_element(By.NAME, "id1x7_2118")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2118))
        input_2 = driver.find_element(By.NAME, "id1x7_2119")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2119))
        input_2 = driver.find_element(By.NAME, "id1x7_2120")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2120))

        input_2 = driver.find_element(By.NAME, "id1x7_2122")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2122))
        input_2 = driver.find_element(By.NAME, "id1x7_2123")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2123))
        input_2 = driver.find_element(By.NAME, "id1x7_2124")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2124))
        input_2 = driver.find_element(By.NAME, "id1x7_2125")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2125))

        input_2 = driver.find_element(By.NAME, "id1x7_2158")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2158))
    # input_2 = driver.find_element(By.NAME, "id1x7_2159")
    #   input_2.clear()
    #   input_2.send_keys(str(iid1x7_2159))
    #  input_2 = driver.find_element(By.NAME, "id1x7_2160")
    #  input_2.clear()
    # input_2.send_keys(str(iid1x7_2160))
        input_2 = driver.find_element(By.NAME, "id1x7_2128")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2128))
        input_2 = driver.find_element(By.NAME, "id1x7_2129")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2129))

        input_2 = driver.find_element(By.NAME, "id1x7_2131")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2131))
        input_2 = driver.find_element(By.NAME, "id1x7_2132")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2132))
        input_2 = driver.find_element(By.NAME, "id1x7_2133")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2133))
        input_2 = driver.find_element(By.NAME, "id1x7_2134")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2134))
        input_2 = driver.find_element(By.NAME, "id1x7_2135")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2135))
        input_2 = driver.find_element(By.NAME, "id1x7_2136")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2136))
        input_2 = driver.find_element(By.NAME, "id1x7_2137")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2137))
        input_2 = driver.find_element(By.NAME, "id1x7_2138")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2138))
        input_2 = driver.find_element(By.NAME, "id1x7_2139")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2139))

        input_2 = driver.find_element(By.NAME, "id1x7_2141")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2141))
        input_2 = driver.find_element(By.NAME, "id1x7_2142")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2142))
        input_2 = driver.find_element(By.NAME, "id1x7_2143")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2143))

        input_2 = driver.find_element(By.NAME, "id1x7_2145")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2145))
        input_2 = driver.find_element(By.NAME, "id1x7_2146")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2146))
        input_2 = driver.find_element(By.NAME, "id1x7_2147")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2147))
        input_2 = driver.find_element(By.NAME, "id1x7_2148")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2148))
        input_2 = driver.find_element(By.NAME, "id1x7_2150")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2150))
        input_2 = driver.find_element(By.NAME, "id1x7_2151")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2151))
        input_2 = driver.find_element(By.NAME, "id1x7_2152")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2152))
        input_2 = driver.find_element(By.NAME, "id1x7_2153")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2153))
        input_2 = driver.find_element(By.NAME, "id1x7_2154")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2154))
        input_2 = driver.find_element(By.NAME, "id1x7_2155")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2155))
        input_2 = driver.find_element(By.NAME, "id1x7_2156")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2156))
        input_2 = driver.find_element(By.NAME, "id1x7_2157")
        input_2.clear()
        input_2.send_keys(str(iid1x7_2157))

    # ---------------กรอกคดีผู้บริโภค ผู้ประกอบการเป็นโจทก์------------------  

        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x8_4")))
        input_1.clear()
        input_1.send_keys(str(iid1x8_4 ))

        input_2 = driver.find_element(By.NAME, "id1x8_5")
        input_2.clear()
        input_2.send_keys(str(iid1x8_5))
        input_2 = driver.find_element(By.NAME, "id1x8_6")
        input_2.clear()
        input_2.send_keys(str(iid1x8_6))

        input_2 = driver.find_element(By.NAME, "id1x8_1001")
        input_2.clear()
        input_2.send_keys(str(iid1x8_1001))
        input_2 = driver.find_element(By.NAME, "id1x8_1002")
        input_2.clear()
        input_2.send_keys(str(iid1x8_1002))
        input_2 = driver.find_element(By.NAME, "id1x8_1003")
        input_2.clear()
        input_2.send_keys(str(iid1x8_1003))

        input_2 = driver.find_element(By.NAME, "id1x8_1005")
        input_2.clear()
        input_2.send_keys(str(iid1x8_1005))
        input_2 = driver.find_element(By.NAME, "id1x8_1006")
        input_2.clear()
        input_2.send_keys(str(iid1x8_1006))
        input_2 = driver.find_element(By.NAME, "id1x8_2107")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2107))
        input_2 = driver.find_element(By.NAME, "id1x8_2108")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2108))
        input_2 = driver.find_element(By.NAME, "id1x8_2109")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2109))
        input_2 = driver.find_element(By.NAME, "id1x8_2110")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2110))
        input_2 = driver.find_element(By.NAME, "id1x8_2111")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2111))

        input_2 = driver.find_element(By.NAME, "id1x8_2113")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2113))
        input_2 = driver.find_element(By.NAME, "id1x8_2114")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2114))
        input_2 = driver.find_element(By.NAME, "id1x8_2115")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2115))
        input_2 = driver.find_element(By.NAME, "id1x8_2116")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2116))

        input_2 = driver.find_element(By.NAME, "id1x8_2118")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2118))
        input_2 = driver.find_element(By.NAME, "id1x8_2119")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2119))
        input_2 = driver.find_element(By.NAME, "id1x8_2120")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2120))

        input_2 = driver.find_element(By.NAME, "id1x8_2122")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2122))
        input_2 = driver.find_element(By.NAME, "id1x8_2123")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2123))
        input_2 = driver.find_element(By.NAME, "id1x8_2124")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2124))
        input_2 = driver.find_element(By.NAME, "id1x8_2125")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2125))

        input_2 = driver.find_element(By.NAME, "id1x8_2158")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2158))
    # input_2 = driver.find_element(By.NAME, "id1x8_2159")
    #   input_2.clear()
    #   input_2.send_keys(str(iid1x8_2159))
    #  input_2 = driver.find_element(By.NAME, "id1x8_2160")
    #  input_2.clear()
    # input_2.send_keys(str(iid1x8_2160))
        input_2 = driver.find_element(By.NAME, "id1x8_2128")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2128))
        input_2 = driver.find_element(By.NAME, "id1x8_2129")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2129))

        input_2 = driver.find_element(By.NAME, "id1x8_2131")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2131))
        input_2 = driver.find_element(By.NAME, "id1x8_2132")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2132))
        input_2 = driver.find_element(By.NAME, "id1x8_2133")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2133))
        input_2 = driver.find_element(By.NAME, "id1x8_2134")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2134))
        input_2 = driver.find_element(By.NAME, "id1x8_2135")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2135))
        input_2 = driver.find_element(By.NAME, "id1x8_2136")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2136))
        input_2 = driver.find_element(By.NAME, "id1x8_2137")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2137))
        input_2 = driver.find_element(By.NAME, "id1x8_2138")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2138))
        input_2 = driver.find_element(By.NAME, "id1x8_2139")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2139))

        input_2 = driver.find_element(By.NAME, "id1x8_2141")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2141))
        input_2 = driver.find_element(By.NAME, "id1x8_2142")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2142))
        input_2 = driver.find_element(By.NAME, "id1x8_2143")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2143))

        input_2 = driver.find_element(By.NAME, "id1x8_2145")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2145))
        input_2 = driver.find_element(By.NAME, "id1x8_2146")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2146))
        input_2 = driver.find_element(By.NAME, "id1x8_2147")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2147))
        input_2 = driver.find_element(By.NAME, "id1x8_2148")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2148))
        input_2 = driver.find_element(By.NAME, "id1x8_2150")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2150))
        input_2 = driver.find_element(By.NAME, "id1x8_2151")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2151))
        input_2 = driver.find_element(By.NAME, "id1x8_2152")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2152))
        input_2 = driver.find_element(By.NAME, "id1x8_2153")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2153))
        input_2 = driver.find_element(By.NAME, "id1x8_2154")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2154))
        input_2 = driver.find_element(By.NAME, "id1x8_2155")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2155))
        input_2 = driver.find_element(By.NAME, "id1x8_2156")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2156))
        input_2 = driver.find_element(By.NAME, "id1x8_2157")
        input_2.clear()
        input_2.send_keys(str(iid1x8_2157))




    # ---------------กรอกคดีผู้บริโภค ผู้รับมอบอำนาจเป็นโจทก์------------------  

        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x9_4")))
        input_1.clear()
        input_1.send_keys(str(iid1x9_4 ))

        input_2 = driver.find_element(By.NAME, "id1x9_5")
        input_2.clear()
        input_2.send_keys(str(iid1x9_5))
        input_2 = driver.find_element(By.NAME, "id1x9_6")
        input_2.clear()
        input_2.send_keys(str(iid1x9_6))

        input_2 = driver.find_element(By.NAME, "id1x9_1001")
        input_2.clear()
        input_2.send_keys(str(iid1x9_1001))
        input_2 = driver.find_element(By.NAME, "id1x9_1002")
        input_2.clear()
        input_2.send_keys(str(iid1x9_1002))
        input_2 = driver.find_element(By.NAME, "id1x9_1003")
        input_2.clear()
        input_2.send_keys(str(iid1x9_1003))

        input_2 = driver.find_element(By.NAME, "id1x9_1005")
        input_2.clear()
        input_2.send_keys(str(iid1x9_1005))
        input_2 = driver.find_element(By.NAME, "id1x9_1006")
        input_2.clear()
        input_2.send_keys(str(iid1x9_1006))
        input_2 = driver.find_element(By.NAME, "id1x9_2107")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2107))
        input_2 = driver.find_element(By.NAME, "id1x9_2108")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2108))
        input_2 = driver.find_element(By.NAME, "id1x9_2109")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2109))
        input_2 = driver.find_element(By.NAME, "id1x9_2110")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2110))
        input_2 = driver.find_element(By.NAME, "id1x9_2111")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2111))

        input_2 = driver.find_element(By.NAME, "id1x9_2113")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2113))
        input_2 = driver.find_element(By.NAME, "id1x9_2114")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2114))
        input_2 = driver.find_element(By.NAME, "id1x9_2115")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2115))
        input_2 = driver.find_element(By.NAME, "id1x9_2116")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2116))

        input_2 = driver.find_element(By.NAME, "id1x9_2118")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2118))
        input_2 = driver.find_element(By.NAME, "id1x9_2119")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2119))
        input_2 = driver.find_element(By.NAME, "id1x9_2120")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2120))

        input_2 = driver.find_element(By.NAME, "id1x9_2122")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2122))
        input_2 = driver.find_element(By.NAME, "id1x9_2123")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2123))
        input_2 = driver.find_element(By.NAME, "id1x9_2124")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2124))
        input_2 = driver.find_element(By.NAME, "id1x9_2125")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2125))

        input_2 = driver.find_element(By.NAME, "id1x9_2158")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2158))
    # input_2 = driver.find_element(By.NAME, "id1x9_2159")
    #   input_2.clear()
    #   input_2.send_keys(str(iid1x9_2159))
    #  input_2 = driver.find_element(By.NAME, "id1x9_2160")
    #  input_2.clear()
    # input_2.send_keys(str(iid1x9_2160))
        input_2 = driver.find_element(By.NAME, "id1x9_2128")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2128))
        input_2 = driver.find_element(By.NAME, "id1x9_2129")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2129))

        input_2 = driver.find_element(By.NAME, "id1x9_2131")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2131))
        input_2 = driver.find_element(By.NAME, "id1x9_2132")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2132))
        input_2 = driver.find_element(By.NAME, "id1x9_2133")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2133))
        input_2 = driver.find_element(By.NAME, "id1x9_2134")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2134))
        input_2 = driver.find_element(By.NAME, "id1x9_2135")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2135))
        input_2 = driver.find_element(By.NAME, "id1x9_2136")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2136))
        input_2 = driver.find_element(By.NAME, "id1x9_2137")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2137))
        input_2 = driver.find_element(By.NAME, "id1x9_2138")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2138))
        input_2 = driver.find_element(By.NAME, "id1x9_2139")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2139))

        input_2 = driver.find_element(By.NAME, "id1x9_2141")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2141))
        input_2 = driver.find_element(By.NAME, "id1x9_2142")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2142))
        input_2 = driver.find_element(By.NAME, "id1x9_2143")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2143))

        input_2 = driver.find_element(By.NAME, "id1x9_2145")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2145))
        input_2 = driver.find_element(By.NAME, "id1x9_2146")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2146))
        input_2 = driver.find_element(By.NAME, "id1x9_2147")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2147))
        input_2 = driver.find_element(By.NAME, "id1x9_2148")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2148))
        input_2 = driver.find_element(By.NAME, "id1x9_2150")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2150))
        input_2 = driver.find_element(By.NAME, "id1x9_2151")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2151))
        input_2 = driver.find_element(By.NAME, "id1x9_2152")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2152))
        input_2 = driver.find_element(By.NAME, "id1x9_2153")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2153))
        input_2 = driver.find_element(By.NAME, "id1x9_2154")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2154))
        input_2 = driver.find_element(By.NAME, "id1x9_2155")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2155))
        input_2 = driver.find_element(By.NAME, "id1x9_2156")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2156))
        input_2 = driver.find_element(By.NAME, "id1x9_2157")
        input_2.clear()
        input_2.send_keys(str(iid1x9_2157))
        end_time = time.time()
        elapsed_time = round(end_time - start_time, 2)


        msg = (
            f"✅ กรอกข้อมูลคดีผู้บริโภคเสร็จเรียบร้อย!\n"
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