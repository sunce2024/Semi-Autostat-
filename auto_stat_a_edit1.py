
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
        df_stat = all_sheets["สถิติรายเดือน"].fillna(0)

        aa7918 = get_safe_int(df_stat, 5, 34) 
    #พิจารณาอาญาไม่เกิน 1 เดือน ใช้ยอดรวม
        aa7921 = get_safe_int(df_stat, 10, 13)
        aa7922 = get_safe_int(df_stat, 10, 20)
        aa7923 = get_safe_int(df_stat, 10, 29)
        aa7924 = get_safe_int(df_stat, 10, 40)
        aa7925 = get_safe_int(df_stat, 10, 53)
        aa7926 = get_safe_int(df_stat, 10, 71)
        aa7927 = get_safe_int(df_stat, 10, 80)
        aa7928 = get_safe_int(df_stat, 10, 90)
        aa7929 = get_safe_int(df_stat, 10, 99)
        #การจำนวนยื่นอุทธรณ์
        aa01 = get_safe_int(df_stat, 18, 24)
        aa02 = get_safe_int(df_stat, 23, 24)
        aa17261 = aa01+aa02
        #จำนวนไม่อนุญาตอุทธรณ์
        aa03 = get_safe_int(df_stat, 18, 43)
        aa04 = get_safe_int(df_stat, 23, 43)
        aa17263 = aa03+aa04
        #จำนวนส่งอุทธรณ์
        aa05 = get_safe_int(df_stat, 18, 68)
        aa06 = get_safe_int(df_stat, 23, 68)
        aa17265 = aa05+aa06
        #การจำนวนยื่นฎีกา
        aa07 = get_safe_int(df_stat, 19, 24)
        aa08 = get_safe_int(df_stat, 24, 24)
        aa17262 = aa07+aa08
        #จำนวนไม่อนุญาตฎีกา
        aa09 = get_safe_int(df_stat, 19, 43)
        aa10 = get_safe_int(df_stat, 24, 43)
        aa17264 = aa09+aa10
        #จำนวนส่งฎีกา
        aa11 = get_safe_int(df_stat, 19, 68)
        aa12 = get_safe_int(df_stat, 24, 68)
        aa17266 = aa11+aa12
        #ใต่สวนชันสูตรพลิกศพ
        aa9875 = get_safe_int(df_stat, 28, 18)
        aa9876 = get_safe_int(df_stat, 29, 18)
        #ใต่สวนมูลฟ้อง
        aa13 = get_safe_int(df_stat, 52, 32)
        aa14 = get_safe_int(df_stat, 52, 50)
        aa7989 = aa13+aa14
        aa15 = get_safe_int(df_stat, 52, 73)
        aa16 = get_safe_int(df_stat, 52, 89)
        aa7992 = aa15+aa16
        #ผัดฟ้องฝากขัง 
        aa17 = get_safe_int(df_stat, 50, 32)
        aa18 = get_safe_int(df_stat, 50, 50)
        aa19 = get_safe_int(df_stat, 51, 32)
        aa20 = get_safe_int(df_stat, 51, 50)
        aa22062 = aa17+aa18+aa19+aa20
        aa21 = get_safe_int(df_stat, 50, 73)
        aa22 = get_safe_int(df_stat, 50, 89)
        aa23 = get_safe_int(df_stat, 51, 73)
        aa24 = get_safe_int(df_stat, 51, 89)
        aa22063 = aa21+aa22+aa23+aa24
        #ระยะเวลาค้าง
        aa7995 = get_safe_int(df_stat, 52, 9)
        aa7996 = get_safe_int(df_stat, 52, 18)
        aa7997 = get_safe_int(df_stat, 52, 26)
        aa7998 = get_safe_int(df_stat, 52, 35)
        aa24956 = get_safe_int(df_stat, 52, 45)
        aa8000 = get_safe_int(df_stat, 52, 67)
        aa8001 = get_safe_int(df_stat, 52, 77)
        aa8002 = get_safe_int(df_stat, 52, 88)
        aa8003 = get_safe_int(df_stat, 52, 97)
    #ข้อ10.ระยะเวลาค้างในชั้นใต่สวนมูลฟ้อง ยังไม่มีในรายงาน บน Excel
    #ข้อ 11.ระยะเวลาค้างในชั้นใต่สวนมูลฟ้อง นับแต่วันศาลสอบคำให้การ ยังไม่มีในรายงาน บน Excel
    #ข้อ 12.ระยะเวลาค้างพิจารณาคดีอาญาที่จำเลยต้องขัง นับแต่วันศาลสอบคำให้การ ยังไม่มีในรายงาน บน Excel
    #ข้อ 13.จำนวนจำเลยต้องขังระหว่างพิจารณา นับแต่วันศาลสอบคำให้การ ยังไม่มีในรายงาน บน Excel
    #ข้อ 14.ขั้นตอนคดีค้าง
        aa25 = get_safe_int(df_stat, 55, 60)
        aa26 = get_safe_int(df_stat, 55, 87)
        aa17714 = aa25+aa26
        #aa27 = get_safe_int(df_stat, 56, 59)
       #aa28 = get_safe_int(df_stat, 56, 85)
        #aa17715 = aa27+aa28
        aa29 = get_safe_int(df_stat, 56, 60)
        aa30 = get_safe_int(df_stat, 56, 87)
        aa17716 = aa29+aa30
        aa31 = get_safe_int(df_stat, 58, 60)
        aa32 = get_safe_int(df_stat, 58, 87)
        aa17718 = aa31+aa32
        aa33 = get_safe_int(df_stat, 59, 60)
        aa34 = get_safe_int(df_stat, 59, 87)
        aa17719 = aa33+aa34
        aa35 = get_safe_int(df_stat, 60, 60)
        aa36 = get_safe_int(df_stat, 60, 87)
        aa17720 = aa35+aa36
        aa37 = get_safe_int(df_stat, 61, 60)
        aa38 = get_safe_int(df_stat, 61, 87)
        aa17721 = aa37+aa38
        aa39 = get_safe_int(df_stat, 62, 60)
        aa40 = get_safe_int(df_stat, 62, 87)
        aa17722 = aa39+aa40
        aa41 = get_safe_int(df_stat, 63, 60)
        aa42 = get_safe_int(df_stat, 63, 87)
        aa17723 = aa41+aa42
        aa43 = get_safe_int(df_stat, 64, 60)
        aa44 = get_safe_int(df_stat, 64, 87)
        aa17724 = aa43+aa44
        aa45 = get_safe_int(df_stat, 65, 60)
        aa46 = get_safe_int(df_stat, 65, 87)
        aa17725 = aa45+aa46
        aa47 = get_safe_int(df_stat, 66, 60)
        aa48 = get_safe_int(df_stat, 66, 87)
        aa17726 = aa45+aa46
        aa49 = get_safe_int(df_stat, 67, 60)
        aa50 = get_safe_int(df_stat, 67, 87)
        aa17727 = aa49+aa50
        aa51 = get_safe_int(df_stat, 68, 60)
        aa52 = get_safe_int(df_stat, 68, 87)
        aa17728 = aa51+aa52
    #ข้อ 15.คดีที่ศาลแต่งตั้งทนาย
        aa53 = get_safe_int(df_stat, 80, 7)
        aa54 = get_safe_int(df_stat, 80, 16)
        aa8054 = aa53+aa54
        aa55 = get_safe_int(df_stat, 81, 7)
        aa56 = get_safe_int(df_stat, 81, 16)
        aa8055 = aa55+aa56
        aa57 = get_safe_int(df_stat, 83, 7)
        aa58 = get_safe_int(df_stat, 83, 16)
        aa8053 = aa57+aa58
        aa59 = get_safe_int(df_stat, 84, 7)
        aa60 = get_safe_int(df_stat, 84, 16)
        aa8056 = aa59+aa60
        aa61 = get_safe_int(df_stat, 80, 23)
        aa62 = get_safe_int(df_stat, 80, 37)
        aa8058 = aa61+aa62
        aa63 = get_safe_int(df_stat, 81, 23)
        aa64 = get_safe_int(df_stat, 81, 37)
        aa8059 = aa63+aa64
        aa65 = get_safe_int(df_stat, 83, 23)
        aa66 = get_safe_int(df_stat, 83, 37)
        aa8057 = aa65+aa66
        aa67 = get_safe_int(df_stat, 84, 23)
        aa68 = get_safe_int(df_stat, 84, 37)
        aa8060 = aa67+aa68
        aa69 = get_safe_int(df_stat, 80, 51)
        aa70 = get_safe_int(df_stat, 80, 67)
        aa8062 = aa69+aa70
        aa71 = get_safe_int(df_stat, 81, 51)
        aa72 = get_safe_int(df_stat, 81, 67)
        aa8063 = aa71+aa72
        aa73 = get_safe_int(df_stat, 83, 51)
        aa74 = get_safe_int(df_stat, 83, 67)
        aa8061 = aa73+aa74
        aa75 = get_safe_int(df_stat, 78, 61)
        aa76 = get_safe_int(df_stat, 78, 78)
        aa8064 = aa75+aa76
    #ข้อ 16.คดีที่ศาลจ่ายเงินรางวัลทนาย
    #ข้อ 17.คดีที่ศาลจ่ายค่าตรวจวิทยาศาสตร์
    #ข้อ 18.คดีที่ศาลจ่ายค่าปวยการ
    #ข้อ 19.คดีที่ศาลจ่ายค่านักจิตวิทยา
    #ข้อ 20.คดีที่ศาลจ่ายค่าที่พักล่าม
    #ข้อ 21.หมายจับ-ค้น
    #ข้อ 22.ผลหมายจับ-ค้น
    #ข้อ 23.การปล่อยตัวชั่วคราว
        aa8123 = get_safe_int(df_stat, 130, 21)
        aa8124 = get_safe_int(df_stat, 131, 21)
        aa8125 = get_safe_int(df_stat, 132, 21)
        aa8126 = get_safe_int(df_stat, 130, 34)
        aa8127 = get_safe_int(df_stat, 131, 34)
        aa8128 = get_safe_int(df_stat, 132, 34)
        aa15738 = get_safe_int(df_stat, 130, 46)
        aa15739 = get_safe_int(df_stat, 131, 46)
        aa15740 = get_safe_int(df_stat, 132, 46)
        aa15741 = get_safe_int(df_stat, 130, 65)
        aa15742 = get_safe_int(df_stat, 131, 65)
        aa15743 = get_safe_int(df_stat, 132, 65)
        aa8129 = get_safe_int(df_stat, 130, 80)
        aa8130 = get_safe_int(df_stat, 131, 80)
        aa8131 = get_safe_int(df_stat, 132, 80)
        aa8132 = get_safe_int(df_stat, 130, 92)
        aa8133 = get_safe_int(df_stat, 131, 92)
        aa8134 = get_safe_int(df_stat, 132, 92)
    #ข้อ 24ข้อมูลบังคับคดีนายประกัน 
    # ข้อ 25 คดียาเสพติด 
    # ผลิต
        aa8143 = get_safe_int(df_stat, 147, 15)
        aa8144 = get_safe_int(df_stat, 148, 15)
        aa8145 = get_safe_int(df_stat, 149, 15)
        aa8146 = get_safe_int(df_stat, 150, 15)
        aa8147 = get_safe_int(df_stat, 151, 15)
        aa8148 = get_safe_int(df_stat, 152, 15)
        aa8149 = get_safe_int(df_stat, 153, 15)
        aa8150 = get_safe_int(df_stat, 154, 15)
        aa8151 = get_safe_int(df_stat, 155, 15)
    # จำหน่าย
        aa8153 = get_safe_int(df_stat, 147, 23)
        aa8154 = get_safe_int(df_stat, 148, 23)
        aa8155 = get_safe_int(df_stat, 149, 23)
        aa8156 = get_safe_int(df_stat, 150, 23)
        aa8157 = get_safe_int(df_stat, 151, 23)
        aa8158 = get_safe_int(df_stat, 152, 23)
        aa8159 = get_safe_int(df_stat, 153, 23)
        aa8160 = get_safe_int(df_stat, 154, 23)
        aa8161 = get_safe_int(df_stat, 155, 23)
    # ครอบครอง
        aa8163 = get_safe_int(df_stat, 147, 36)
        aa8164 = get_safe_int(df_stat, 148, 36)
        aa8165 = get_safe_int(df_stat, 149, 36)
        aa8166 = get_safe_int(df_stat, 150, 36)
        aa8167 = get_safe_int(df_stat, 151, 36)
        aa8168 = get_safe_int(df_stat, 152, 36)
        aa8169 = get_safe_int(df_stat, 153, 36)
        aa8170 = get_safe_int(df_stat, 154, 36)
        aa8171 = get_safe_int(df_stat, 155, 36)
    # เสพ
        aa8173 = get_safe_int(df_stat, 147, 51)
        aa8174 = get_safe_int(df_stat, 148, 51)
        aa8175 = get_safe_int(df_stat, 149, 51)
        aa8176 = get_safe_int(df_stat, 150, 51)
        aa8177 = get_safe_int(df_stat, 151, 51)
        aa8178 = get_safe_int(df_stat, 152, 51)
        aa8179 = get_safe_int(df_stat, 153, 51)
        aa8180 = get_safe_int(df_stat, 154, 51)
        aa8181 = get_safe_int(df_stat, 155, 51)
    # ครอบครองเพื่อเสพ
        aa8183 = get_safe_int(df_stat, 147, 71)
        aa8184 = get_safe_int(df_stat, 148, 71)
        aa8185 = get_safe_int(df_stat, 149, 71)
        aa8186 = get_safe_int(df_stat, 150, 71)
        aa8187 = get_safe_int(df_stat, 151, 71)
        aa8188 = get_safe_int(df_stat, 152, 71)
        aa8189 = get_safe_int(df_stat, 153, 71)
        aa8190 = get_safe_int(df_stat, 154, 71)
        aa8191 = get_safe_int(df_stat, 155, 71)
    # ครอบครองจำหน่าย
        aa8193 = get_safe_int(df_stat, 147, 84)
        aa8194 = get_safe_int(df_stat, 148, 84)
        aa8195 = get_safe_int(df_stat, 149, 84)
        aa8196 = get_safe_int(df_stat, 150, 84)
        aa8197 = get_safe_int(df_stat, 151, 84)
        aa8198 = get_safe_int(df_stat, 152, 84)
        aa8199 = get_safe_int(df_stat, 153, 84)
        aa8200 = get_safe_int(df_stat, 154, 84)
        aa8201 = get_safe_int(df_stat, 155, 84)
    # อื่นๆ
        aa8203 = get_safe_int(df_stat, 147, 98)
        aa8204 = get_safe_int(df_stat, 148, 98)
        aa8205 = get_safe_int(df_stat, 149, 98)
        aa8206 = get_safe_int(df_stat, 150, 98)
        aa8207 = get_safe_int(df_stat, 151, 98)
        aa8208 = get_safe_int(df_stat, 152, 98)
        aa8209 = get_safe_int(df_stat, 153, 98)
        aa8210 = get_safe_int(df_stat, 154, 98)
        aa8211 = get_safe_int(df_stat, 155, 98)
    #คดีฟื้นฟู
        aa17749 = get_safe_int(df_stat, 159, 45)
    #จำหน่ายคดีชั่วคราว
        aa17751 = get_safe_int(df_stat, 160, 16)
        print(
            f"ดึงข้อมูลสำเร็จ -> จำหน่ายคดีชั่วคราว: {aa17751}")

        # --- 2. เปลี่ยนไปยังหน้าฟอร์ม fid=27 ---
        print("กำลังเปลี่ยนไปหน้าฟอร์ม fid=29...")
        driver.get( "http://10.1.2.9/court_stat_program/form_statistic.php?fid=29")

        # ตั้งเวลารอ Element โหลดสูงสุด 30 วินาที
        wait = WebDriverWait(driver, 30)

        # --- 3. เริ่มกรอกข้อมูลอัตโนมัติ ---
        print("กำลังกรอกข้อมูลอัตโนมัติ...")

        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "a7918")))
        input_1.clear()
        input_1.send_keys(str(aa7918))

        # ช่องระยะเวลาพิจารณาคดีอาญา
        input_2 = driver.find_element(By.NAME, "a7921")
        input_2.clear()
        input_2.send_keys(str(aa7921))
        input_2 = driver.find_element(By.NAME, "a7922")
        input_2.clear()
        input_2.send_keys(str(aa7922))
        input_2 = driver.find_element(By.NAME, "a7923")
        input_2.clear()
        input_2.send_keys(str(aa7923))
        input_2 = driver.find_element(By.NAME, "a7924")
        input_2.clear()
        input_2.send_keys(str(aa7924))
        input_2 = driver.find_element(By.NAME, "a7925")
        input_2.clear()
        input_2.send_keys(str(aa7925))
        input_2 = driver.find_element(By.NAME, "a7926")
        input_2.clear()
        input_2.send_keys(str(aa7926))
        input_2 = driver.find_element(By.NAME, "a7927")
        input_2.clear()
        input_2.send_keys(str(aa7927))
        input_2 = driver.find_element(By.NAME, "a7928")
        input_2.clear()
        input_2.send_keys(str(aa7928))
        input_2 = driver.find_element(By.NAME, "a7929")
        input_2.clear()
        input_2.send_keys(str(aa7929))
        # การอุทธรณ์
        input_2 = driver.find_element(By.NAME, "a17261")
        input_2.clear()
        input_2.send_keys(str(aa17261))
        input_2 = driver.find_element(By.NAME, "a17263")
        input_2.clear()
        input_2.send_keys(str(aa17263))
        input_2 = driver.find_element(By.NAME, "a17265")
        input_2.clear()
        input_2.send_keys(str(aa17265))
        # การฎีกา
        input_2 = driver.find_element(By.NAME, "a17262")
        input_2.clear()
        input_2.send_keys(str(aa17262))
        input_2 = driver.find_element(By.NAME, "a17264")
        input_2.clear()
        input_2.send_keys(str(aa17264))
        input_2 = driver.find_element(By.NAME, "a17266")
        input_2.clear()
        input_2.send_keys(str(aa17266))
        #ใต่สวนชันสูตรพลิกศพ
        input_2 = driver.find_element(By.NAME, "a9875")
        input_2.clear()
        input_2.send_keys(str(aa9875))
        input_2 = driver.find_element(By.NAME, "a9876")
        input_2.clear()
        input_2.send_keys(str(aa9876))
        #การละเมิดอำนาจศาล
    # input_2 = driver.find_element(By.NAME, "a7943")
    #   input_2.clear()
    #  input_2.send_keys(str(aa7943))
    #   input_2 = driver.find_element(By.NAME, "a7944")
    #    input_2.clear()
    #    input_2.send_keys(str(aa7944))
    #   input_2 = driver.find_element(By.NAME, "a7945")
    #    input_2.clear()
    #    input_2.send_keys(str(aa7945))
    #    input_2 = driver.find_element(By.NAME, "a7946")
    #   input_2.clear()
    #    input_2.send_keys(str(aa7946))
    #    input_2 = driver.find_element(By.NAME, "a7947")
    #    input_2.clear()
    #    input_2.send_keys(str(aa7947))
    #    input_2 = driver.find_element(By.NAME, "a7948")
    #   input_2.clear()
    #   input_2.send_keys(str(aa7948))
    #   input_2 = driver.find_element(By.NAME, "a7949")
    #   input_2.clear()
    #   input_2.send_keys(str(aa7949))
    #   input_2 = driver.find_element(By.NAME, "a7950")
    #   input_2.clear()
    #   input_2.send_keys(str(aa7950))
    #   input_2 = driver.find_element(By.NAME, "a7951")
    #   input_2.clear()
    #   input_2.send_keys(str(aa7951))
    #   input_2 = driver.find_element(By.NAME, "a7952")
    #   input_2.clear()
    #   input_2.send_keys(str(aa7952))
        #การไต่สวนมูลฟ้อง
        input_2 = driver.find_element(By.NAME, "a7989")
        input_2.clear()
        input_2.send_keys(str(aa7989))
        input_2 = driver.find_element(By.NAME, "a7992")
        input_2.clear()
        input_2.send_keys(str(aa7992))
        #ผัดฟ้อง
        input_2 = driver.find_element(By.NAME, "a22062")
        input_2.clear()
        input_2.send_keys(str(aa22062))
        input_2 = driver.find_element(By.NAME, "a22063")
        input_2.clear()
        input_2.send_keys(str(aa22063))
    #ระยะเวลาค้าง
        input_2 = driver.find_element(By.NAME, "a7995")
        input_2.clear()
        input_2.send_keys(str(aa7995))
        input_2 = driver.find_element(By.NAME, "a7996")
        input_2.clear()
        input_2.send_keys(str(aa7996))
        input_2 = driver.find_element(By.NAME, "a7997")
        input_2.clear()
        input_2.send_keys(str(aa7997))
        input_2 = driver.find_element(By.NAME, "a7998")
        input_2.clear()
        input_2.send_keys(str(aa7998))
        input_2 = driver.find_element(By.NAME, "a24956")
        input_2.clear()
        input_2.send_keys(str(aa24956))
        input_2 = driver.find_element(By.NAME, "a8000")
        input_2.clear()
        input_2.send_keys(str(aa8000))
        input_2 = driver.find_element(By.NAME, "a8001")
        input_2.clear()
        input_2.send_keys(str(aa8001))
        input_2 = driver.find_element(By.NAME, "a8002")
        input_2.clear()
        input_2.send_keys(str(aa8002))
        input_2 = driver.find_element(By.NAME, "a8003")
        input_2.clear()
        input_2.send_keys(str(aa8003))

    #ข้อ10.ระยะเวลาค้างในชั้นใต่สวนมูลฟ้อง ยังไม่มีในรายงาน บน Excel
    #ข้อ 11.ระยะเวลาค้างในชั้นใต่สวนมูลฟ้อง นับแต่วันศาลสอบคำให้การ ยังไม่มีในรายงาน บน Excel
    #ข้อ 12.ระยะเวลาค้างพิจารณาคดีอาญาที่จำเลยต้องขัง นับแต่วันศาลสอบคำให้การ ยังไม่มีในรายงาน บน Excel
    #ข้อ 13.จำนวนจำเลยต้องขังระหว่างพิจารณา นับแต่วันศาลสอบคำให้การ ยังไม่มีในรายงาน บน Excel
    #ข้อ 14.ขั้นตอนคดีค้าง
        input_2 = driver.find_element(By.NAME, "a17714")
        input_2.clear()
        input_2.send_keys(str(aa17714))
        #input_2 = driver.find_element(By.NAME, "a17715")
       # input_2.clear()
       # input_2.send_keys(str(aa17715))
        input_2 = driver.find_element(By.NAME, "a17716")
        input_2.clear()
        input_2.send_keys(str(aa17716))
    # input_2 = driver.find_element(By.NAME, "a17717")
    # input_2.clear()
    # input_2.send_keys(str(aa17717))
        input_2 = driver.find_element(By.NAME, "a17718")
        input_2.clear()
        input_2.send_keys(str(aa17718))
        input_2 = driver.find_element(By.NAME, "a17719")
        input_2.clear()
        input_2.send_keys(str(aa17719))
        input_2 = driver.find_element(By.NAME, "a17720")
        input_2.clear()
        input_2.send_keys(str(aa17720))
        input_2 = driver.find_element(By.NAME, "a17721")
        input_2.clear()
        input_2.send_keys(str(aa17721))
        input_2 = driver.find_element(By.NAME, "a17722")
        input_2.clear()
        input_2.send_keys(str(aa17722))
        input_2 = driver.find_element(By.NAME, "a17723")
        input_2.clear()
        input_2.send_keys(str(aa17723))
        input_2 = driver.find_element(By.NAME, "a17724")
        input_2.clear()
        input_2.send_keys(str(aa17724))
        input_2 = driver.find_element(By.NAME, "a17725")
        input_2.clear()
        input_2.send_keys(str(aa17725))
        input_2 = driver.find_element(By.NAME, "a17726")
        input_2.clear()
        input_2.send_keys(str(aa17726))
        input_2 = driver.find_element(By.NAME, "a17727")
        input_2.clear()
        input_2.send_keys(str(aa17727))
        input_2 = driver.find_element(By.NAME, "a17728")
        input_2.clear()
        input_2.send_keys(str(aa17728))
    # เงินรางวัลทนาย 
        input_2 = driver.find_element(By.NAME, "a8054")
        input_2.clear()
        input_2.send_keys(str(aa8054))
        input_2 = driver.find_element(By.NAME, "a8055")
        input_2.clear()
        input_2.send_keys(str(aa8055))
        input_2 = driver.find_element(By.NAME, "a8053")
        input_2.clear()
        input_2.send_keys(str(aa8053))
        input_2 = driver.find_element(By.NAME, "a8056")
        input_2.clear()
        input_2.send_keys(str(aa8056))
        #ประเภทที่ 2 
        input_2 = driver.find_element(By.NAME, "a8058")
        input_2.clear()
        input_2.send_keys(str(aa8058))
        input_2 = driver.find_element(By.NAME, "a8059")
        input_2.clear()
        input_2.send_keys(str(aa8059))
        input_2 = driver.find_element(By.NAME, "a8057")
        input_2.clear()
        input_2.send_keys(str(aa8057))
        input_2 = driver.find_element(By.NAME, "a8060")
        input_2.clear()
        input_2.send_keys(str(aa8060))
        #ประเภทที่ 3
        input_2 = driver.find_element(By.NAME, "a8062")
        input_2.clear()
        input_2.send_keys(str(aa8062))
        input_2 = driver.find_element(By.NAME, "a8063")
        input_2.clear()
        input_2.send_keys(str(aa8063))
        input_2 = driver.find_element(By.NAME, "a8061")
        input_2.clear()
        input_2.send_keys(str(aa8061))
        input_2 = driver.find_element(By.NAME, "a8064")
        input_2.clear()
        input_2.send_keys(str(aa8064))
        #ข้อ 16.คดีที่ศาลจ่ายเงินรางวัลทนาย
    #ข้อ 17.คดีที่ศาลจ่ายค่าตรวจวิทยาศาสตร์
    #ข้อ 18.คดีที่ศาลจ่ายค่าปวยการ
    #ข้อ 19.คดีที่ศาลจ่ายค่านักจิตวิทยา
    #ข้อ 20.คดีที่ศาลจ่ายค่าที่พักล่าม
    #ข้อ 21.หมายจับ-ค้น
    #ข้อ 22.ผลหมายจับ-ค้น
    #ข้อ 23.การปล่อยตัวชั่วคราว
        input_2 = driver.find_element(By.NAME, "a8123")
        input_2.clear()
        input_2.send_keys(str(aa8123))
        input_2 = driver.find_element(By.NAME, "a8124")
        input_2.clear()
        input_2.send_keys(str(aa8124))
        input_2 = driver.find_element(By.NAME, "a8125")
        input_2.clear()
        input_2.send_keys(str(aa8125))
        input_2 = driver.find_element(By.NAME, "a8126")
        input_2.clear()
        input_2.send_keys(str(aa8126))
        input_2 = driver.find_element(By.NAME, "a8127")
        input_2.clear()
        input_2.send_keys(str(aa8127))
        input_2 = driver.find_element(By.NAME, "a8128")
        input_2.clear()
        input_2.send_keys(str(aa8128))
        input_2 = driver.find_element(By.NAME, "a15738")
        input_2.clear()
        input_2.send_keys(str(aa15738))
        input_2 = driver.find_element(By.NAME, "a15739")
        input_2.clear()
        input_2.send_keys(str(aa15739))
        input_2 = driver.find_element(By.NAME, "a15740")
        input_2.clear()
        input_2.send_keys(str(aa15740))
        input_2 = driver.find_element(By.NAME, "a15741")
        input_2.clear()
        input_2.send_keys(str(aa15741))
        input_2 = driver.find_element(By.NAME, "a15742")
        input_2.clear()
        input_2.send_keys(str(aa15742))
        input_2 = driver.find_element(By.NAME, "a15743")
        input_2.clear()
        input_2.send_keys(str(aa15743))
        input_2 = driver.find_element(By.NAME, "a8129")
        input_2.clear()
        input_2.send_keys(str(aa8129))
        input_2 = driver.find_element(By.NAME, "a8130")
        input_2.clear()
        input_2.send_keys(str(aa8130))
        input_2 = driver.find_element(By.NAME, "a8131")
        input_2.clear()
        input_2.send_keys(str(aa8131))
        input_2 = driver.find_element(By.NAME, "a8132")
        input_2.clear()
        input_2.send_keys(str(aa8132))
        input_2 = driver.find_element(By.NAME, "a8133")
        input_2.clear()
        input_2.send_keys(str(aa8133))
        input_2 = driver.find_element(By.NAME, "a8134")
        input_2.clear()
        input_2.send_keys(str(aa8134))

        #ข้อ 24ข้อมูลบังคับคดีนายประกัน 
    # ข้อ 25 คดียาเสพติด
        input_2 = driver.find_element(By.NAME, "a8143")
        input_2.clear()
        input_2.send_keys(str(aa8143))
        input_2 = driver.find_element(By.NAME, "a8144")
        input_2.clear()
        input_2.send_keys(str(aa8144))
        input_2 = driver.find_element(By.NAME, "a8145")
        input_2.clear()
        input_2.send_keys(str(aa8145))
        input_2 = driver.find_element(By.NAME, "a8146")
        input_2.clear()
        input_2.send_keys(str(aa8146))
        input_2 = driver.find_element(By.NAME, "a8147")
        input_2.clear()
        input_2.send_keys(str(aa8147))
        input_2 = driver.find_element(By.NAME, "a8148")
        input_2.clear()
        input_2.send_keys(str(aa8148))
        input_2 = driver.find_element(By.NAME, "a8149")
        input_2.clear()
        input_2.send_keys(str(aa8149))
        input_2 = driver.find_element(By.NAME, "a8150")
        input_2.clear()
        input_2.send_keys(str(aa8150))
        input_2 = driver.find_element(By.NAME, "a8151")
        input_2.clear()
        input_2.send_keys(str(aa8151))
        #จำหน่าย
        input_2 = driver.find_element(By.NAME, "a8153")
        input_2.clear()
        input_2.send_keys(str(aa8153))
        input_2 = driver.find_element(By.NAME, "a8154")
        input_2.clear()
        input_2.send_keys(str(aa8154))
        input_2 = driver.find_element(By.NAME, "a8155")
        input_2.clear()
        input_2.send_keys(str(aa8155))
        input_2 = driver.find_element(By.NAME, "a8156")
        input_2.clear()
        input_2.send_keys(str(aa8156))
        input_2 = driver.find_element(By.NAME, "a8157")
        input_2.clear()
        input_2.send_keys(str(aa8157))
        input_2 = driver.find_element(By.NAME, "a8158")
        input_2.clear()
        input_2.send_keys(str(aa8158))
        input_2 = driver.find_element(By.NAME, "a8159")
        input_2.clear()
        input_2.send_keys(str(aa8159))
        input_2 = driver.find_element(By.NAME, "a8160")
        input_2.clear()
        input_2.send_keys(str(aa8160))
        input_2 = driver.find_element(By.NAME, "a8161")
        input_2.clear()
        input_2.send_keys(str(aa8161))
        #ครอบครอง
        input_2 = driver.find_element(By.NAME, "a8163")
        input_2.clear()
        input_2.send_keys(str(aa8163))
        input_2 = driver.find_element(By.NAME, "a8164")
        input_2.clear()
        input_2.send_keys(str(aa8164))
        input_2 = driver.find_element(By.NAME, "a8165")
        input_2.clear()
        input_2.send_keys(str(aa8165))
        input_2 = driver.find_element(By.NAME, "a8166")
        input_2.clear()
        input_2.send_keys(str(aa8166))
        input_2 = driver.find_element(By.NAME, "a8167")
        input_2.clear()
        input_2.send_keys(str(aa8167))
        input_2 = driver.find_element(By.NAME, "a8168")
        input_2.clear()
        input_2.send_keys(str(aa8168))
        input_2 = driver.find_element(By.NAME, "a8169")
        input_2.clear()
        input_2.send_keys(str(aa8169))
        input_2 = driver.find_element(By.NAME, "a8170")
        input_2.clear()
        input_2.send_keys(str(aa8170))
        input_2 = driver.find_element(By.NAME, "a8171")
        input_2.clear()
        input_2.send_keys(str(aa8171))
        #เสพ
        input_2 = driver.find_element(By.NAME, "a8173")
        input_2.clear()
        input_2.send_keys(str(aa8173))
        input_2 = driver.find_element(By.NAME, "a8174")
        input_2.clear()
        input_2.send_keys(str(aa8174))
        input_2 = driver.find_element(By.NAME, "a8175")
        input_2.clear()
        input_2.send_keys(str(aa8175))
        input_2 = driver.find_element(By.NAME, "a8176")
        input_2.clear()
        input_2.send_keys(str(aa8176))
        input_2 = driver.find_element(By.NAME, "a8177")
        input_2.clear()
        input_2.send_keys(str(aa8177))
        input_2 = driver.find_element(By.NAME, "a8178")
        input_2.clear()
        input_2.send_keys(str(aa8178))
        input_2 = driver.find_element(By.NAME, "a8179")
        input_2.clear()
        input_2.send_keys(str(aa8179))
        input_2 = driver.find_element(By.NAME, "a8180")
        input_2.clear()
        input_2.send_keys(str(aa8180))
        input_2 = driver.find_element(By.NAME, "a8181")
        input_2.clear()
        input_2.send_keys(str(aa8181))
        #ครอบครองเสพ
        input_2 = driver.find_element(By.NAME, "a8183")
        input_2.clear()
        input_2.send_keys(str(aa8183))
        input_2 = driver.find_element(By.NAME, "a8184")
        input_2.clear()
        input_2.send_keys(str(aa8184))
        input_2 = driver.find_element(By.NAME, "a8185")
        input_2.clear()
        input_2.send_keys(str(aa8185))
        input_2 = driver.find_element(By.NAME, "a8186")
        input_2.clear()
        input_2.send_keys(str(aa8186))
        input_2 = driver.find_element(By.NAME, "a8187")
        input_2.clear()
        input_2.send_keys(str(aa8187))
        input_2 = driver.find_element(By.NAME, "a8188")
        input_2.clear()
        input_2.send_keys(str(aa8188))
        input_2 = driver.find_element(By.NAME, "a8189")
        input_2.clear()
        input_2.send_keys(str(aa8189))
        input_2 = driver.find_element(By.NAME, "a8190")
        input_2.clear()
        input_2.send_keys(str(aa8190))
        input_2 = driver.find_element(By.NAME, "a8191")
        input_2.clear()
        input_2.send_keys(str(aa8191))
        #ครอบครองจำหน่าย
        input_2 = driver.find_element(By.NAME, "a8193")
        input_2.clear()
        input_2.send_keys(str(aa8193))
        input_2 = driver.find_element(By.NAME, "a8194")
        input_2.clear()
        input_2.send_keys(str(aa8194))
        input_2 = driver.find_element(By.NAME, "a8195")
        input_2.clear()
        input_2.send_keys(str(aa8195))
        input_2 = driver.find_element(By.NAME, "a8196")
        input_2.clear()
        input_2.send_keys(str(aa8196))
        input_2 = driver.find_element(By.NAME, "a8197")
        input_2.clear()
        input_2.send_keys(str(aa8197))
        input_2 = driver.find_element(By.NAME, "a8198")
        input_2.clear()
        input_2.send_keys(str(aa8198))
        input_2 = driver.find_element(By.NAME, "a8199")
        input_2.clear()
        input_2.send_keys(str(aa8199))
        input_2 = driver.find_element(By.NAME, "a8200")
        input_2.clear()
        input_2.send_keys(str(aa8200))
        input_2 = driver.find_element(By.NAME, "a8201")
        input_2.clear()
        input_2.send_keys(str(aa8201))
        #อื่นๆ
        input_2 = driver.find_element(By.NAME, "a8203")
        input_2.clear()
        input_2.send_keys(str(aa8203))
        input_2 = driver.find_element(By.NAME, "a8204")
        input_2.clear()
        input_2.send_keys(str(aa8204))
        input_2 = driver.find_element(By.NAME, "a8205")
        input_2.clear()
        input_2.send_keys(str(aa8205))
        input_2 = driver.find_element(By.NAME, "a8206")
        input_2.clear()
        input_2.send_keys(str(aa8206))
        input_2 = driver.find_element(By.NAME, "a8207")
        input_2.clear()
        input_2.send_keys(str(aa8207))
        input_2 = driver.find_element(By.NAME, "a8208")
        input_2.clear()
        input_2.send_keys(str(aa8208))
        input_2 = driver.find_element(By.NAME, "a8209")
        input_2.clear()
        input_2.send_keys(str(aa8209))
        input_2 = driver.find_element(By.NAME, "a8210")
        input_2.clear()
        input_2.send_keys(str(aa8210))
        input_2 = driver.find_element(By.NAME, "a8211")
        input_2.clear()
        input_2.send_keys(str(aa8211))
        #คำร้องฟื้นฟู
        input_2 = driver.find_element(By.NAME, "a17749")
        input_2.clear()
        input_2.send_keys(str(aa17749))            
        #จำหน่ายคดี
        input_2 = driver.find_element(By.NAME, "a17751")
        input_2.clear()
        input_2.send_keys(str(aa17751)) 

        end_time = time.time()
        elapsed_time = round(end_time - start_time, 2)
        msg = (
            f"✅ กรอกข้อมูลจำนวนคดีอาญาเสร็จเรียบร้อย!\n"
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