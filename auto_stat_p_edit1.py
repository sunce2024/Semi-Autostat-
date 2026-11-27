
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
        df_stat = all_sheets["บัญชีคดีแพ่งรายเดือน"].fillna(0)

        iid1x3_2 = get_safe_int(df_stat, 3, 8)  
        iid1x3_3 = get_safe_int(df_stat, 4, 8)
        iid1x3_4 = get_safe_int(df_stat, 5, 8)  
        iid1x3_6 = get_safe_int(df_stat, 7, 8)
        iid1x3_7 = get_safe_int(df_stat, 8, 8)
        iid1x3_8 = get_safe_int(df_stat, 9, 8)    
        iid1x3_9 = get_safe_int(df_stat, 10, 8)
        iid1x3_10 = get_safe_int(df_stat, 11, 8)
        iid1x3_11 = get_safe_int(df_stat, 12, 8)
        iid1x3_12 = get_safe_int(df_stat, 13, 8)
        iid1x3_13 = get_safe_int(df_stat, 14, 8)
        iid1x3_14 = get_safe_int(df_stat, 15, 8)
        iid1x3_15 = get_safe_int(df_stat, 16, 8)
        iid1x3_16 = get_safe_int(df_stat, 17, 8)
        iid1x3_17 = get_safe_int(df_stat, 18, 8)
        iid1x3_18 = get_safe_int(df_stat, 19, 8)
        iid1x3_19 = get_safe_int(df_stat, 20, 8)
        iid1x3_20 = get_safe_int(df_stat, 21, 8)
        iid1x3_21 = get_safe_int(df_stat, 22, 8)
        iid1x3_22 = get_safe_int(df_stat, 23, 8)
        iid1x3_23 = get_safe_int(df_stat, 24, 8)
        iid1x3_24 = get_safe_int(df_stat, 25, 8)
        iid1x3_25 = get_safe_int(df_stat, 26, 8)

        iid1x3_27= get_safe_int(df_stat, 28, 8)
        iid1x3_28 = get_safe_int(df_stat, 29, 8)

        iid1x3_29 = get_safe_int(df_stat, 3, 24)
        iid1x3_31 = get_safe_int(df_stat, 5, 24)
        iid1x3_32 = get_safe_int(df_stat, 6, 24)
        iid1x3_33= get_safe_int(df_stat, 7, 24)
        iid1x3_34 = get_safe_int(df_stat, 8, 24)

        iid1x3_36 = get_safe_int(df_stat, 10, 24)
        iid1x3_37 = get_safe_int(df_stat, 11, 24)
        iid1x3_38 = get_safe_int(df_stat, 12, 24)
        iid1x3_39 = get_safe_int(df_stat, 13, 24)
        iid1x3_40 = get_safe_int(df_stat, 14, 24)
        iid1x3_41 = get_safe_int(df_stat, 15, 24)
        iid1x3_42 = get_safe_int(df_stat, 16, 24)
        iid1x3_43 = get_safe_int(df_stat, 17, 24)
        iid1x3_44 = get_safe_int(df_stat, 18, 24)
        iid1x3_45 = get_safe_int(df_stat, 19, 24)
        iid1x3_46 = get_safe_int(df_stat, 20, 24)
        iid1x3_47 = get_safe_int(df_stat, 21, 24)

        iid1x3_49 = get_safe_int(df_stat, 23, 24)
        iid1x3_50 = get_safe_int(df_stat, 24, 24)
        iid1x3_51 = get_safe_int(df_stat, 25, 24)

        # แท๊ปใหม่ id2x3
        iid2x3_57 = get_safe_int(df_stat, 33, 25)
        iid2x3_58 = get_safe_int(df_stat, 34, 25)
        iid2x3_59 = get_safe_int(df_stat, 35, 25)
        iid2x3_60 = get_safe_int(df_stat, 36, 25)
        # แท๊ปใหม่ id3x3 คดีไม่มีข้อพิพาท
        iid3x3_53 = get_safe_int(df_stat, 33, 9)
        iid3x3_54 = get_safe_int(df_stat, 34, 9)
        iid3x3_55 = get_safe_int(df_stat, 35, 9)
        # แท๊ปใหม่ id5x3 ข้อหาอื่นๆนอกจากบัญชี
        iid5x3_63 = get_safe_int(df_stat, 41, 8)
        iid5x3_64 = get_safe_int(df_stat, 42, 8)

        iid5x3_66 = get_safe_int(df_stat, 44, 8)
        iid5x3_67 = get_safe_int(df_stat, 45, 8)
        iid5x3_68 = get_safe_int(df_stat, 46, 8)
        iid5x3_69 = get_safe_int(df_stat, 47, 8)
        iid5x3_70 = get_safe_int(df_stat, 46, 8)
        iid5x3_71 = get_safe_int(df_stat, 49, 8)
        iid5x3_72 = get_safe_int(df_stat, 50, 8)
        iid5x3_73 = get_safe_int(df_stat, 51, 8) 

        iid5x3_75 = get_safe_int(df_stat, 53, 8) 
        iid5x3_76 = get_safe_int(df_stat, 54, 8) 
        iid5x3_77 = get_safe_int(df_stat, 55, 8) 
        iid5x3_78 = get_safe_int(df_stat, 56, 8)            





        print(f"ดึงข้อมูลสำเร็จ -> คดีแพ่ง: {iid1x3_27}, คดีผู้บริโภค: {iid1x3_29}, ค้างแพ่งไม่เกิน1เดือน: ")
        # --- 2. เปลี่ยนไปยังหน้าฟอร์ม fid=27 ---
        print("กำลังเปลี่ยนไปหน้าฟอร์ม fid=28...")
        driver.get(
            "http://10.1.2.9/court_stat_program/form_case.php?fid=28"
        )

        # ตั้งเวลารอ Element โหลดสูงสุด 30 วินาที
        wait = WebDriverWait(driver, 30)

        # --- 3. เริ่มกรอกข้อมูลอัตโนมัติ ---
        print("กำลังกรอกข้อมูลอัตโนมัติ...")

    # ข้อหาคดีแพ่ง
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "id1x3_2")))
        input_1.clear()
        input_1.send_keys(str(iid1x3_2 ))

        input_2 = driver.find_element(By.NAME, "id1x3_3")
        input_2.clear()
        input_2.send_keys(str(iid1x3_3))
        input_2 = driver.find_element(By.NAME, "id1x3_4")
        input_2.clear()
        input_2.send_keys(str(iid1x3_4))

        input_2 = driver.find_element(By.NAME, "id1x3_6")
        input_2.clear()
        input_2.send_keys(str(iid1x3_6))
        input_2 = driver.find_element(By.NAME, "id1x3_7")
        input_2.clear()
        input_2.send_keys(str(iid1x3_7))
        input_2 = driver.find_element(By.NAME, "id1x3_8")
        input_2.clear()
        input_2.send_keys(str(iid1x3_8))
        input_2 = driver.find_element(By.NAME, "id1x3_9")
        input_2.clear()
        input_2.send_keys(str(iid1x3_9))
        input_2 = driver.find_element(By.NAME, "id1x3_10")
        input_2.clear()
        input_2.send_keys(str(iid1x3_10))
        input_2 = driver.find_element(By.NAME, "id1x3_11")
        input_2.clear()
        input_2.send_keys(str(iid1x3_11))
        input_2 = driver.find_element(By.NAME, "id1x3_12")
        input_2.clear()
        input_2.send_keys(str(iid1x3_12))
        input_2 = driver.find_element(By.NAME, "id1x3_13")
        input_2.clear()
        input_2.send_keys(str(iid1x3_13))
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
        input_2 = driver.find_element(By.NAME, "id1x3_18")
        input_2.clear()
        input_2.send_keys(str(iid1x3_18))
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
        input_2 = driver.find_element(By.NAME, "id1x3_24")
        input_2.clear()
        input_2.send_keys(str(iid1x3_24))
        input_2 = driver.find_element(By.NAME, "id1x3_25")
        input_2.clear()
        input_2.send_keys(str(iid1x3_25))

        input_2 = driver.find_element(By.NAME, "id1x3_27")
        input_2.clear()
        input_2.send_keys(str(iid1x3_27))
        input_2 = driver.find_element(By.NAME, "id1x3_28")
        input_2.clear()
        input_2.send_keys(str(iid1x3_28))
        input_2 = driver.find_element(By.NAME, "id1x3_29")
        input_2.clear()
        input_2.send_keys(str(iid1x3_29))
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

        input_2 = driver.find_element(By.NAME, "id1x3_36")
        input_2.clear()
        input_2.send_keys(str(iid1x3_36))
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
        input_2 = driver.find_element(By.NAME, "id1x3_43")
        input_2.clear()
        input_2.send_keys(str(iid1x3_43))
        input_2 = driver.find_element(By.NAME, "id1x3_44")
        input_2.clear()
        input_2.send_keys(str(iid1x3_44))
        input_2 = driver.find_element(By.NAME, "id1x3_45")
        input_2.clear()
        input_2.send_keys(str(iid1x3_45))
        input_2 = driver.find_element(By.NAME, "id1x3_46")
        input_2.clear()
        input_2.send_keys(str(iid1x3_46))
        input_2 = driver.find_element(By.NAME, "id1x3_47")
        input_2.clear()
        input_2.send_keys(str(iid1x3_47))

        input_2 = driver.find_element(By.NAME, "id1x3_49")
        input_2.clear()
        input_2.send_keys(str(iid1x3_49))
        input_2 = driver.find_element(By.NAME, "id1x3_50")
        input_2.clear()
        input_2.send_keys(str(iid1x3_50))
        input_2 = driver.find_element(By.NAME, "id1x3_51")
        input_2.clear()
        input_2.send_keys(str(iid1x3_51))

        # tab ใหม่ http  ใส่link driver.get("http://10.1.2.9/court_stat_program/login.php")
        #driver.get("http://10.1.2.9/court_stat_program/form_case.php?fid=28#view1")

       # driver.get("http://10.1.2.9/court_stat_program/form_case.php?fid=28")
        #time.sleep(1)

        # ----------------------------------------------------
        #  สลับไป Tab View 1
        # ----------------------------------------------------
        try:
            tab1_button = driver.find_element(By.XPATH, "//a[@href='#view1']")
            tab1_button.click()
            time.sleep(0.5)
        except Exception:
            pass  # ข้ามหากคลิกไม่ได้
        input_2 = driver.find_element(By.NAME, "id2x3_57")
        input_2.clear()
        input_2.send_keys(str(iid2x3_57))
        input_2 = driver.find_element(By.NAME, "id2x3_58")
        input_2.clear()
        input_2.send_keys(str(iid2x3_58))
        input_2 = driver.find_element(By.NAME, "id2x3_59")
        input_2.clear()
        input_2.send_keys(str(iid2x3_59))
        input_2 = driver.find_element(By.NAME, "id2x3_60")
        input_2.clear()
        input_2.send_keys(str(iid2x3_60))

    # tab ใหม่ http  ใส่link driver.get("http://10.1.2.9/court_stat_program/login.php")
        #driver.get("http://10.1.2.9/court_stat_program/form_case.php?fid=28#view2")
       # driver.get("http://10.1.2.9/court_stat_program/form_case.php?fid=28")
       # time.sleep(1)
        
                # ----------------------------------------------------
                #  สลับไป Tab View 2
                # ----------------------------------------------------
        try:
            tab2_button = driver.find_element(By.XPATH, "//a[@href='#view2']")
            tab2_button.click()
            time.sleep(0.5)
        except Exception:
            pass  # ข้ามหากคลิกไม่ได้
        input_2 = driver.find_element(By.NAME, "id3x3_53")
        input_2.clear()
        input_2.send_keys(str(iid3x3_53))
        input_2 = driver.find_element(By.NAME, "id3x3_54")
        input_2.clear()
        input_2.send_keys(str(iid3x3_54))
        input_2 = driver.find_element(By.NAME, "id3x3_55")
        input_2.clear()
        input_2.send_keys(str(iid3x3_55))
    
    # tab ใหม่ http  ใส่link driver.get("http://10.1.2.9/court_stat_program/login.php")
    #driver.get("http://10.1.2.9/court_stat_program/form_case.php?fid=28#view3")
       # driver.get("http://10.1.2.9/court_stat_program/form_case.php?fid=28")
       # time.sleep(1)
    
            # ----------------------------------------------------
            #  สลับไป Tab View 3
            # ----------------------------------------------------
        try:
            tab3_button = driver.find_element(By.XPATH, "//a[@href='#view3']")
            tab3_button.click()
            time.sleep(0.5)
        except Exception:
            pass  # ข้ามหากคลิกไม่ได้
        input_2 = driver.find_element(By.NAME, "id5x3_63")
        input_2.clear()
        input_2.send_keys(str(iid5x3_63))
        input_2 = driver.find_element(By.NAME, "id5x3_64")
        input_2.clear()
        input_2.send_keys(str(iid5x3_64))

        input_2 = driver.find_element(By.NAME, "id5x3_66")
        input_2.clear()
        input_2.send_keys(str(iid5x3_66))
        input_2 = driver.find_element(By.NAME, "id5x3_67")
        input_2.clear()
        input_2.send_keys(str(iid5x3_67))
        input_2 = driver.find_element(By.NAME, "id5x3_68")
        input_2.clear()
        input_2.send_keys(str(iid5x3_68))
        input_2 = driver.find_element(By.NAME, "id5x3_69")
        input_2.clear()
        input_2.send_keys(str(iid5x3_69))
        input_2 = driver.find_element(By.NAME, "id5x3_70")
        input_2.clear()
        input_2.send_keys(str(iid5x3_70))
        input_2 = driver.find_element(By.NAME, "id5x3_71")
        input_2.clear()
        input_2.send_keys(str(iid5x3_71))
        input_2 = driver.find_element(By.NAME, "id5x3_72")
        input_2.clear()
        input_2.send_keys(str(iid5x3_72))
        input_2 = driver.find_element(By.NAME, "id5x3_73")
        input_2.clear()
        input_2.send_keys(str(iid5x3_73))

        input_2 = driver.find_element(By.NAME, "id5x3_75")
        input_2.clear()
        input_2.send_keys(str(iid5x3_75))
        input_2 = driver.find_element(By.NAME, "id5x3_76")
        input_2.clear()
        input_2.send_keys(str(iid5x3_76))
        input_2 = driver.find_element(By.NAME, "id5x3_77")
        input_2.clear()
        input_2.send_keys(str(iid5x3_77))
        input_2 = driver.find_element(By.NAME, "id5x3_78")
        input_2.clear()
        input_2.send_keys(str(iid5x3_78))

        end_time = time.time()
        elapsed_time = round(end_time - start_time, 2)
        msg = (
            f"✅ กรอกข้อมูลคดีแพ่งเสร็จเรียบร้อย!\n"
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