
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
        file_path = os.path.join(BASE_DIR, "stat_rpt_month_crime.xls")

        print("กำลังอ่านไฟล์ Excel...")
        all_sheets = pd.read_excel(file_path, sheet_name=None, engine="xlrd")
        df_stat = all_sheets["บัญชีรายเดือน"].fillna(0)

        #iid1x3_3 = get_safe_int(df_stat, 3, 8)  
        iid1x3_3 = get_safe_int(df_stat, 4, 8)
        iid1x3_4 = get_safe_int(df_stat, 5, 8)  
        iid1x3_5 = get_safe_int(df_stat, 6, 8) 
        #iid1x3_6 = get_safe_int(df_stat, 7, 8)
        iid1x3_7 = get_safe_int(df_stat, 14, 11)
        iid1x3_8 = get_safe_int(df_stat, 9, 8)    
    # iid1x3_9 = get_safe_int(df_stat, 10, 8)
        iid1x3_10 = get_safe_int(df_stat, 11, 8)
        iid1x3_11 = get_safe_int(df_stat, 12, 8)
        iid1x3_12 = get_safe_int(df_stat, 13, 8)
        iid1x3_13 = get_safe_int(df_stat, 14, 8)
        iid1x3_14 = get_safe_int(df_stat, 15, 8)
        iid1x3_15 = get_safe_int(df_stat, 16, 8)
        iid1x3_16 = get_safe_int(df_stat, 17, 8)
        iid1x3_17 = get_safe_int(df_stat, 18, 8)
    # iid1x3_18 = get_safe_int(df_stat, 19, 8)
        iid1x3_19 = get_safe_int(df_stat, 39, 11)
        iid1x3_20 = get_safe_int(df_stat, 21, 8)
        iid1x3_21 = get_safe_int(df_stat, 22, 8)
        iid1x3_22 = get_safe_int(df_stat, 23, 8)
        iid1x3_23 = get_safe_int(df_stat, 24, 8)
    #  iid1x3_24 = get_safe_int(df_stat, 25, 8)
        iid1x3_25 = get_safe_int(df_stat, 26, 8)
        iid1x3_26 = get_safe_int(df_stat, 26, 8)
        iid1x3_27= get_safe_int(df_stat, 57, 11)
        iid1x3_28 = get_safe_int(df_stat, 30, 8)
        iid1x3_30 = get_safe_int(df_stat, 32, 8)
        iid1x3_31 = get_safe_int(df_stat, 65, 11)
        iid1x3_32 = get_safe_int(df_stat, 39, 8)
        iid1x3_33= get_safe_int(df_stat, 5, 27)
        iid1x3_34 = get_safe_int(df_stat, 43, 8)
        iid1x3_35 = get_safe_int(df_stat, 46, 8)

        iid1x3_37 = get_safe_int(df_stat, 48, 8)
        iid1x3_38 = get_safe_int(df_stat, 53, 8)
        iid1x3_39 = get_safe_int(df_stat, 55, 8)
        iid1x3_40 = get_safe_int(df_stat, 57, 8)
        iid1x3_41 = get_safe_int(df_stat, 59, 8)
        iid1x3_42 = get_safe_int(df_stat, 63, 8)

        #iid1x3_43 = get_safe_int(df_stat, 17, 24)
        iid1x3_44 = get_safe_int(df_stat, 65, 8)
        iid1x3_45 = get_safe_int(df_stat, 3, 27)
        iid1x3_46 = get_safe_int(df_stat, 5, 27)

        iid1x3_48 = get_safe_int(df_stat, 7, 27)
        iid1x3_49 = get_safe_int(df_stat, 9, 27)
        iid1x3_50 = get_safe_int(df_stat, 13, 27)
        iid1x3_52 = get_safe_int(df_stat, 15, 27)
        iid1x3_53 = get_safe_int(df_stat, 17, 27)
        iid1x3_54 = get_safe_int(df_stat, 19, 27)
        iid1x3_55 = get_safe_int(df_stat, 21, 27)
        iid1x3_56 = get_safe_int(df_stat, 23, 27)
        iid1x3_57 = get_safe_int(df_stat, 27, 27)
        iid1x3_58 = get_safe_int(df_stat, 29, 27)
        iid1x3_59 = get_safe_int(df_stat, 31, 27)
        iid1x3_60 = get_safe_int(df_stat, 35, 27)
        iid1x3_61 = get_safe_int(df_stat, 36, 27)
        iid1x3_62 = get_safe_int(df_stat, 38, 27)
        iid1x3_63 = get_safe_int(df_stat, 42, 27)

        # แท๊ปใหม่ id2x3
        iid2x3_65 = get_safe_int(df_stat, 44, 27)
        iid2x3_66 = get_safe_int(df_stat, 45, 27)
        iid2x3_67 = get_safe_int(df_stat, 47, 27)
        iid2x3_68 = get_safe_int(df_stat, 49, 27)
        iid2x3_69 = get_safe_int(df_stat, 50, 27)
        iid2x3_70 = get_safe_int(df_stat, 52, 27)
        iid2x3_71 = get_safe_int(df_stat, 54, 27)
        iid2x3_72 = get_safe_int(df_stat, 56, 27)
        iid2x3_73 = get_safe_int(df_stat, 58, 27)
        iid2x3_74 = get_safe_int(df_stat, 60, 27)
        iid2x3_75 = get_safe_int(df_stat, 62, 27)

        # ความผิดประมวลกฎหมายอื่นๆ

        iid2x3_76 = get_safe_int(df_stat, 69, 11)
        iid2x3_77 = get_safe_int(df_stat, 70, 11)
        iid2x3_78 = get_safe_int(df_stat, 71, 11)
        iid2x3_79 = get_safe_int(df_stat, 72, 11)

        iid2x3_90 = get_safe_int(df_stat, 73, 11)
        iid2x3_91 = get_safe_int(df_stat, 74, 11)
        iid2x3_92 = get_safe_int(df_stat, 75, 11)
        iid2x3_243 = get_safe_int(df_stat, 76, 11)
        iid2x3_94 = get_safe_int(df_stat, 77, 11)
        iid2x3_244 = get_safe_int(df_stat, 78, 11)
        iid2x3_245 = get_safe_int(df_stat, 79, 11)
        iid2x3_246 = get_safe_int(df_stat, 80, 11)
        iid2x3_97 = get_safe_int(df_stat, 81, 11)
        iid2x3_98 = get_safe_int(df_stat, 82, 11)
        iid2x3_99 = get_safe_int(df_stat, 83, 11)
        iid2x3_247 = get_safe_int(df_stat, 84, 11)
        iid2x3_248= get_safe_int(df_stat, 85, 11)
        iid2x3_101 = get_safe_int(df_stat, 86, 11)
        iid2x3_249 = get_safe_int(df_stat, 87, 11)
        iid2x3_251 = get_safe_int(df_stat, 88, 11)
        iid2x3_104 = get_safe_int(df_stat, 89, 11)
        iid2x3_208 = get_safe_int(df_stat, 90, 11)
        iid2x3_211 = get_safe_int(df_stat, 91, 11)
        iid2x3_209 = get_safe_int(df_stat, 92, 11)

        iid2x3_210 = get_safe_int(df_stat, 94, 11)
        iid2x3_212 = get_safe_int(df_stat, 96, 11)
        iid2x3_106 = get_safe_int(df_stat, 98, 11)
        iid2x3_107 = get_safe_int(df_stat, 100, 11)
        iid2x3_108 = get_safe_int(df_stat, 102, 11)
        iid2x3_109 = get_safe_int(df_stat, 104, 11)
        iid2x3_110 = get_safe_int(df_stat, 106, 11)
        iid2x3_111 = get_safe_int(df_stat, 108, 11)
        iid2x3_112 = get_safe_int(df_stat, 110, 11)
        iid2x3_113 = get_safe_int(df_stat, 112, 11)
        iid2x3_114 = get_safe_int(df_stat, 114, 11)
        iid2x3_252 = get_safe_int(df_stat, 116, 11)
        iid2x3_250 = get_safe_int(df_stat, 118, 11)
        iid2x3_116 = get_safe_int(df_stat, 120, 11)
        iid2x3_117 = get_safe_int(df_stat, 69, 29)
        iid2x3_118 = get_safe_int(df_stat, 70, 29)
        iid2x3_119 = get_safe_int(df_stat, 71, 29)

        iid2x3_213 = get_safe_int(df_stat, 72, 29)
        iid2x3_216 = get_safe_int(df_stat, 73, 29)
        iid2x3_217 = get_safe_int(df_stat, 74, 29)
        iid2x3_218 = get_safe_int(df_stat, 75, 29)
        iid2x3_219 = get_safe_int(df_stat, 76, 29)
        iid2x3_220 = get_safe_int(df_stat, 77, 29)
        iid2x3_221 = get_safe_int(df_stat, 78, 29)
        iid2x3_222 = get_safe_int(df_stat, 79, 29)
        iid2x3_223 = get_safe_int(df_stat, 80, 29)
        iid2x3_224 = get_safe_int(df_stat, 81, 29)
        iid2x3_225 = get_safe_int(df_stat, 82, 29)
        iid2x3_226 = get_safe_int(df_stat, 83, 29)
        iid2x3_227 = get_safe_int(df_stat, 84, 29)
        iid2x3_228 = get_safe_int(df_stat, 85, 29)
        iid2x3_229 = get_safe_int(df_stat, 86, 29)
        iid2x3_230 = get_safe_int(df_stat, 87, 29)
        iid2x3_231 = get_safe_int(df_stat, 88, 29)
        iid2x3_232 = get_safe_int(df_stat, 89, 29)
        iid2x3_233 = get_safe_int(df_stat, 90, 29)
        iid2x3_234 = get_safe_int(df_stat, 91, 29)
        iid2x3_235 = get_safe_int(df_stat, 93, 29)
        iid2x3_236 = get_safe_int(df_stat, 95, 29)
        iid2x3_237 = get_safe_int(df_stat, 97, 29)
        iid2x3_238 = get_safe_int(df_stat, 99, 29)
        iid2x3_239 = get_safe_int(df_stat, 101, 29)
        iid2x3_240 = get_safe_int(df_stat, 103, 29)
        iid2x3_241 = get_safe_int(df_stat, 105, 29)
        

        # แท๊ปใหม่ id3x3 คดึอาญา
        iid3x3_121 = get_safe_int(df_stat, 33, 9)
        iid3x3_122 = get_safe_int(df_stat, 34, 9)
        iid3x3_123 = get_safe_int(df_stat, 35, 9)




        print(f"ดึงข้อมูลสำเร็จ -> คดีอาญา: {iid3x3_122}, ทดสอบ: {iid3x3_123} ")
        # --- 2. เปลี่ยนไปยังหน้าฟอร์ม fid=27 ---
        print("กำลังเปลี่ยนไปหน้าฟอร์ม fid=30...")
        driver.get(
            "http://10.1.2.9/court_stat_program/form_case.php?fid=30"
        )

        # ตั้งเวลารอ Element โหลดสูงสุด 30 วินาที
        wait = WebDriverWait(driver, 30)

        # --- 3. เริ่มกรอกข้อมูลอัตโนมัติ ---
        print("กำลังกรอกข้อมูลอัตโนมัติ...")

        # ข้อหาคดีอาญา
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_3")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_3 ))

        
        input_2 = driver.find_element(By.NAME, "id1x3_4")
        input_2.clear()
        input_2.send_keys(str(iid1x3_4))
        input_2 = driver.find_element(By.NAME, "id1x3_5")
        input_2.clear()
        input_2.send_keys(str(iid1x3_5))

        input_2 = driver.find_element(By.NAME, "id1x3_7")
        input_2.clear()
        input_2.send_keys(str(iid1x3_7))
        input_2 = driver.find_element(By.NAME, "id1x3_8")
        input_2.clear()
        input_2.send_keys(str(iid1x3_8))
        
        
        input_2 = driver.find_element(By.NAME, "id1x3_10")
        input_2.clear()
        input_2.send_keys(str(iid1x3_10))
        input_2 = driver.find_element(By.NAME, "id1x3_11")
        input_2.clear()
        input_2.send_keys(str(iid1x3_11))
        input_2 = driver.find_element(By.NAME, "id1x3_12")
        input_2.clear()
        input_2.send_keys(str(iid1x3_12))
        
        input_2 = driver.find_element(By.NAME, "id1x3_14")
        input_2.clear()
        input_2.send_keys(str(iid1x3_14))
        input_2 = driver.find_element(By.NAME, "id1x3_15")
        input_2.clear()
        input_2.send_keys(str(iid1x3_16))
        input_2 = driver.find_element(By.NAME, "id1x3_16")
        input_2.clear()
        input_2.send_keys(str(iid1x3_16))
        input_2 = driver.find_element(By.NAME, "id1x3_17")
        input_2.clear()
        input_2.send_keys(str(iid1x3_17))
        
        input_2 = driver.find_element(By.NAME, "id1x3_19")
        input_2.clear()
        input_2.send_keys(str(iid1x3_19))
        input_2 = driver.find_element(By.NAME, "id1x3_20")
        input_2.clear()
        input_2.send_keys(str(iid1x3_20))
        input_2 = driver.find_element(By.NAME, "id1x3_21")
        input_2.clear()
        input_2.send_keys(str(iid1x3_21))
        input_2 = driver.find_element(By.NAME, "id1x3_22")
        input_2.clear()
        input_2.send_keys(str(iid1x3_22))
        input_2 = driver.find_element(By.NAME, "id1x3_23")
        input_2.clear()
        input_2.send_keys(str(iid1x3_23))
        
        input_2 = driver.find_element(By.NAME, "id1x3_25")
        input_2.clear()
        input_2.send_keys(str(iid1x3_25))
        input_2 = driver.find_element(By.NAME, "id1x3_26")
        input_2.clear()
        input_2.send_keys(str(iid1x3_26))
        input_2 = driver.find_element(By.NAME, "id1x3_27")
        input_2.clear()
        input_2.send_keys(str(iid1x3_27))
        input_2 = driver.find_element(By.NAME, "id1x3_28")
        input_2.clear()
        input_2.send_keys(str(iid1x3_28))
        input_2 = driver.find_element(By.NAME, "id1x3_30")
        input_2.clear()
        input_2.send_keys(str(iid1x3_30))
        input_2 = driver.find_element(By.NAME, "id1x3_31")
        input_2.clear()
        input_2.send_keys(str(iid1x3_31))
        input_2 = driver.find_element(By.NAME, "id1x3_32")
        input_2.clear()
        input_2.send_keys(str(iid1x3_32))
        input_2 = driver.find_element(By.NAME, "id1x3_33")
        input_2.clear()
        input_2.send_keys(str(iid1x3_33))
        input_2 = driver.find_element(By.NAME, "id1x3_34")
        input_2.clear()
        input_2.send_keys(str(iid1x3_34))
        input_2 = driver.find_element(By.NAME, "id1x3_35")
        input_2.clear()
        input_2.send_keys(str(iid1x3_35))

        input_2 = driver.find_element(By.NAME, "id1x3_37")
        input_2.clear()
        input_2.send_keys(str(iid1x3_37))
        input_2 = driver.find_element(By.NAME, "id1x3_38")
        input_2.clear()
        input_2.send_keys(str(iid1x3_38))
        input_2 = driver.find_element(By.NAME, "id1x3_39")
        input_2.clear()
        input_2.send_keys(str(iid1x3_39))
        input_2 = driver.find_element(By.NAME, "id1x3_40")
        input_2.clear()
        input_2.send_keys(str(iid1x3_40))
        input_2 = driver.find_element(By.NAME, "id1x3_41")
        input_2.clear()
        input_2.send_keys(str(iid1x3_41))
        input_2 = driver.find_element(By.NAME, "id1x3_42")
        input_2.clear()
        input_2.send_keys(str(iid1x3_42))
        
        input_2 = driver.find_element(By.NAME, "id1x3_44")
        input_2.clear()
        input_2.send_keys(str(iid1x3_44))
        input_2 = driver.find_element(By.NAME, "id1x3_45")
        input_2.clear()
        input_2.send_keys(str(iid1x3_45))
        input_2 = driver.find_element(By.NAME, "id1x3_46")
        input_2.clear()
        input_2.send_keys(str(iid1x3_46))
        
        input_2 = driver.find_element(By.NAME, "id1x3_48")
        input_2.clear()
        input_2.send_keys(str(iid1x3_48))
        input_2 = driver.find_element(By.NAME, "id1x3_49")
        input_2.clear()
        input_2.send_keys(str(iid1x3_49))
        input_2 = driver.find_element(By.NAME, "id1x3_50")
        input_2.clear()
        input_2.send_keys(str(iid1x3_50))

        input_2 = driver.find_element(By.NAME, "id1x3_52")
        input_2.clear()
        input_2.send_keys(str(iid1x3_52))
        input_2 = driver.find_element(By.NAME, "id1x3_53")
        input_2.clear()
        input_2.send_keys(str(iid1x3_53))
        input_2 = driver.find_element(By.NAME, "id1x3_54")
        input_2.clear()
        input_2.send_keys(str(iid1x3_54))
        input_2 = driver.find_element(By.NAME, "id1x3_55")
        input_2.clear()
        input_2.send_keys(str(iid1x3_55))
        input_2 = driver.find_element(By.NAME, "id1x3_56")
        input_2.clear()
        input_2.send_keys(str(iid1x3_56))
        input_2 = driver.find_element(By.NAME, "id1x3_57")
        input_2.clear()
        input_2.send_keys(str(iid1x3_57))
        input_2 = driver.find_element(By.NAME, "id1x3_58")
        input_2.clear()
        input_2.send_keys(str(iid1x3_58))
        input_2 = driver.find_element(By.NAME, "id1x3_59")
        input_2.clear()
        input_2.send_keys(str(iid1x3_59))
        input_2 = driver.find_element(By.NAME, "id1x3_60")
        input_2.clear()
        input_2.send_keys(str(iid1x3_60))
        input_2 = driver.find_element(By.NAME, "id1x3_61")
        input_2.clear()
        input_2.send_keys(str(iid1x3_61))
        input_2 = driver.find_element(By.NAME, "id1x3_62")
        input_2.clear()
        input_2.send_keys(str(iid1x3_62))
        input_2 = driver.find_element(By.NAME, "id1x3_63")
        input_2.clear()
        input_2.send_keys(str(iid1x3_63))


        # ----------------------------------------------------
                #  สลับไป Tab View 1
                # ----------------------------------------------------
        try:
            tab1_button = driver.find_element(By.XPATH, "//a[@href='#view1']")
            tab1_button.click()
            time.sleep(0.5)
        except Exception:
            pass  # ข้ามหากคลิกไม่ได้

        input_2 = driver.find_element(By.NAME, "id2x3_65")
        input_2.clear()
        input_2.send_keys(str(iid2x3_65))
        input_2 = driver.find_element(By.NAME, "id2x3_66")
        input_2.clear()
        input_2.send_keys(str(iid2x3_66))
        input_2 = driver.find_element(By.NAME, "id2x3_67")
        input_2.clear()
        input_2.send_keys(str(iid2x3_67))
        input_2 = driver.find_element(By.NAME, "id2x3_68")
        input_2.clear()
        input_2.send_keys(str(iid2x3_68))
        input_2 = driver.find_element(By.NAME, "id2x3_242")
        input_2.clear()
        input_2.send_keys(str(iid2x3_69))
        input_2 = driver.find_element(By.NAME, "id2x3_70")
        input_2.clear()
        input_2.send_keys(str(iid2x3_70))
    # input_2 = driver.find_element(By.NAME, "id2x3_71")
    # input_2.clear()
    # input_2.send_keys(str(iid2x3_71))
        input_2 = driver.find_element(By.NAME, "id2x3_72")
        input_2.clear()
        input_2.send_keys(str(iid2x3_72))
        input_2 = driver.find_element(By.NAME, "id2x3_73")
        input_2.clear()
        input_2.send_keys(str(iid2x3_73))
        input_2 = driver.find_element(By.NAME, "id2x3_74")
        input_2.clear()
        input_2.send_keys(str(iid2x3_74))
        input_2 = driver.find_element(By.NAME, "id2x3_75")
        input_2.clear()
        input_2.send_keys(str(iid2x3_75))
        input_2 = driver.find_element(By.NAME, "id2x3_76")
        input_2.clear()
        input_2.send_keys(str(iid2x3_76))
        input_2 = driver.find_element(By.NAME, "id2x3_77")
        input_2.clear()
        input_2.send_keys(str(iid2x3_77))
        input_2 = driver.find_element(By.NAME, "id2x3_78")
        input_2.clear()
        input_2.send_keys(str(iid2x3_78))

        input_2 = driver.find_element(By.NAME, "id2x3_90")
        input_2.clear()
        input_2.send_keys(str(iid2x3_90))
        input_2 = driver.find_element(By.NAME, "id2x3_91")
        input_2.clear()
        input_2.send_keys(str(iid2x3_91))
        input_2 = driver.find_element(By.NAME, "id2x3_92")
        input_2.clear()
        input_2.send_keys(str(iid2x3_92))
        input_2 = driver.find_element(By.NAME, "id2x3_243")
        input_2.clear()
        input_2.send_keys(str(iid2x3_243))
        input_2 = driver.find_element(By.NAME, "id2x3_94")
        input_2.clear()
        input_2.send_keys(str(iid2x3_94))
        input_2 = driver.find_element(By.NAME, "id2x3_244")
        input_2.clear()
        input_2.send_keys(str(iid2x3_244))
        input_2 = driver.find_element(By.NAME, "id2x3_245")
        input_2.clear()
        input_2.send_keys(str(iid2x3_245))
        input_2 = driver.find_element(By.NAME, "id2x3_246")
        input_2.clear()
        input_2.send_keys(str(iid2x3_246))
        input_2 = driver.find_element(By.NAME, "id2x3_97")
        input_2.clear()
        input_2.send_keys(str(iid2x3_97))
        input_2 = driver.find_element(By.NAME, "id2x3_98")
        input_2.clear()
        input_2.send_keys(str(iid2x3_98))
        input_2 = driver.find_element(By.NAME, "id2x3_99")
        input_2.clear()
        input_2.send_keys(str(iid2x3_99))
        input_2 = driver.find_element(By.NAME, "id2x3_247")
        input_2.clear()
        input_2.send_keys(str(iid2x3_247))
        input_2 = driver.find_element(By.NAME, "id2x3_248")
        input_2.clear()
        input_2.send_keys(str(iid2x3_248))
        input_2 = driver.find_element(By.NAME, "id2x3_101")
        input_2.clear()
        input_2.send_keys(str(iid2x3_101))
        input_2 = driver.find_element(By.NAME, "id2x3_249")
        input_2.clear()
        input_2.send_keys(str(iid2x3_249))
        input_2 = driver.find_element(By.NAME, "id2x3_251")
        input_2.clear()
        input_2.send_keys(str(iid2x3_251))
        input_2 = driver.find_element(By.NAME, "id2x3_104")
        input_2.clear()
        input_2.send_keys(str(iid2x3_104))
        input_2 = driver.find_element(By.NAME, "id2x3_208")
        input_2.clear()
        input_2.send_keys(str(iid2x3_208))
        input_2 = driver.find_element(By.NAME, "id2x3_211")
        input_2.clear()
        input_2.send_keys(str(iid2x3_211))
        input_2 = driver.find_element(By.NAME, "id2x3_209")
        input_2.clear()
        input_2.send_keys(str(iid2x3_209))
        input_2 = driver.find_element(By.NAME, "id2x3_210")
        input_2.clear()
        input_2.send_keys(str(iid2x3_210))
        input_2 = driver.find_element(By.NAME, "id2x3_212")
        input_2.clear()
        input_2.send_keys(str(iid2x3_212))
        input_2 = driver.find_element(By.NAME, "id2x3_106")
        input_2.clear()
        input_2.send_keys(str(iid2x3_106))
        input_2 = driver.find_element(By.NAME, "id2x3_107")
        input_2.clear()
        input_2.send_keys(str(iid2x3_107))
        input_2 = driver.find_element(By.NAME, "id2x3_108")
        input_2.clear()
        input_2.send_keys(str(iid2x3_108))
        input_2 = driver.find_element(By.NAME, "id2x3_109")
        input_2.clear()
        input_2.send_keys(str(iid2x3_109))
        input_2 = driver.find_element(By.NAME, "id2x3_110")
        input_2.clear()
        input_2.send_keys(str(iid2x3_110))
        input_2 = driver.find_element(By.NAME, "id2x3_111")
        input_2.clear()
        input_2.send_keys(str(iid2x3_111))
        input_2 = driver.find_element(By.NAME, "id2x3_112")
        input_2.clear()
        input_2.send_keys(str(iid2x3_112))
        input_2 = driver.find_element(By.NAME, "id2x3_113")
        input_2.clear()
        input_2.send_keys(str(iid2x3_113))
        input_2 = driver.find_element(By.NAME, "id2x3_114")
        input_2.clear()
        input_2.send_keys(str(iid2x3_114))
        input_2 = driver.find_element(By.NAME, "id2x3_252")
        input_2.clear()
        input_2.send_keys(str(iid2x3_252))
        input_2 = driver.find_element(By.NAME, "id2x3_250")
        input_2.clear()
        input_2.send_keys(str(iid2x3_250))
        input_2 = driver.find_element(By.NAME, "id2x3_116")
        input_2.clear()
        input_2.send_keys(str(iid2x3_116))
        input_2 = driver.find_element(By.NAME, "id2x3_117")
        input_2.clear()
        input_2.send_keys(str(iid2x3_117))
        input_2 = driver.find_element(By.NAME, "id2x3_118")
        input_2.clear()
        input_2.send_keys(str(iid2x3_118))
        input_2 = driver.find_element(By.NAME, "id2x3_119")
        input_2.clear()
        input_2.send_keys(str(iid2x3_119))
        input_2 = driver.find_element(By.NAME, "id2x3_213")
        input_2.clear()
        input_2.send_keys(str(iid2x3_213))
        input_2 = driver.find_element(By.NAME, "id2x3_216")
        input_2.clear()
        input_2.send_keys(str(iid2x3_216))
        input_2 = driver.find_element(By.NAME, "id2x3_217")
        input_2.clear()
        input_2.send_keys(str(iid2x3_217))
        input_2 = driver.find_element(By.NAME, "id2x3_218")
        input_2.clear()
        input_2.send_keys(str(iid2x3_218))
        input_2 = driver.find_element(By.NAME, "id2x3_219")
        input_2.clear()
        input_2.send_keys(str(iid2x3_219))
        input_2 = driver.find_element(By.NAME, "id2x3_220")
        input_2.clear()
        input_2.send_keys(str(iid2x3_220))
        input_2 = driver.find_element(By.NAME, "id2x3_221")
        input_2.clear()
        input_2.send_keys(str(iid2x3_221))
        input_2 = driver.find_element(By.NAME, "id2x3_222")
        input_2.clear()
        input_2.send_keys(str(iid2x3_222))
        input_2 = driver.find_element(By.NAME, "id2x3_223")
        input_2.clear()
        input_2.send_keys(str(iid2x3_223))
        input_2 = driver.find_element(By.NAME, "id2x3_224")
        input_2.clear()
        input_2.send_keys(str(iid2x3_224))
        input_2 = driver.find_element(By.NAME, "id2x3_225")
        input_2.clear()
        input_2.send_keys(str(iid2x3_225))
        input_2 = driver.find_element(By.NAME, "id2x3_226")
        input_2.clear()
        input_2.send_keys(str(iid2x3_226))
        input_2 = driver.find_element(By.NAME, "id2x3_227")
        input_2.clear()
        input_2.send_keys(str(iid2x3_227))
        input_2 = driver.find_element(By.NAME, "id2x3_228")
        input_2.clear()
        input_2.send_keys(str(iid2x3_228))
        input_2 = driver.find_element(By.NAME, "id2x3_229")
        input_2.clear()
        input_2.send_keys(str(iid2x3_229))
        input_2 = driver.find_element(By.NAME, "id2x3_230")
        input_2.clear()
        input_2.send_keys(str(iid2x3_230))
        input_2 = driver.find_element(By.NAME, "id2x3_231")
        input_2.clear()
        input_2.send_keys(str(iid2x3_231))
        input_2 = driver.find_element(By.NAME, "id2x3_232")
        input_2.clear()
        input_2.send_keys(str(iid2x3_232))
        input_2 = driver.find_element(By.NAME, "id2x3_233")
        input_2.clear()
        input_2.send_keys(str(iid2x3_233))
        input_2 = driver.find_element(By.NAME, "id2x3_234")
        input_2.clear()
        input_2.send_keys(str(iid2x3_234))
        input_2 = driver.find_element(By.NAME, "id2x3_235")
        input_2.clear()
        input_2.send_keys(str(iid2x3_235))
        input_2 = driver.find_element(By.NAME, "id2x3_236")
        input_2.clear()
        input_2.send_keys(str(iid2x3_236))
        input_2 = driver.find_element(By.NAME, "id2x3_237")
        input_2.clear()
        input_2.send_keys(str(iid2x3_238))
        input_2 = driver.find_element(By.NAME, "id2x3_239")
        input_2.clear()
        input_2.send_keys(str(iid2x3_239))
        input_2 = driver.find_element(By.NAME, "id2x3_240")
        input_2.clear()
        input_2.send_keys(str(iid2x3_240))
        input_2 = driver.find_element(By.NAME, "id2x3_241")
        input_2.clear()
        input_2.send_keys(str(iid2x3_241))
        
                # ----------------------------------------------------
                #  สลับไป Tab View 2
                # ----------------------------------------------------
        try:
            tab1_button = driver.find_element(By.XPATH, "//a[@href='#view2']")
            tab1_button.click()
            time.sleep(0.5)
        except Exception:
            pass  # ข้ามหากคลิกไม่ได้
    # tab ใหม่ http  ใส่link driver.get("http://10.1.2.9/court_stat_program/login.php")
        input_2 = driver.find_element(By.NAME, "id3x3_121")
        input_2.clear()
        input_2.send_keys(str(iid3x3_121))
        input_2 = driver.find_element(By.NAME, "id3x3_122")
        input_2.clear()
        input_2.send_keys(str(iid3x3_122))
        input_2 = driver.find_element(By.NAME, "id3x3_123")
        input_2.clear()
        input_2.send_keys(str(iid3x3_123))

        end_time = time.time()
        elapsed_time = round(end_time - start_time, 2)
                        # ----------------------------------------------------
                #  สลับไป Tab View 3
                # ----------------------------------------------------
        try:
            tab1_button = driver.find_element(By.XPATH, "//a[@href='#view3']")
            tab1_button.click()
            time.sleep(0.5)
        except Exception:
            pass  # ข้ามหากคลิกไม่ได้


        msg = (
            f"✅ กรอกข้อมูลคดีเกิดอาญาเสร็จเรียบร้อย!\n"
            f"📊 ผลลัพธ์: (ใช้เวลาป้อนข้อมูล:** {elapsed_time}** วินาที)\n"
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