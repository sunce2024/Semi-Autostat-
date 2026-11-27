
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
# ขั้นตอนที่ 1: เปิด Browser และไปหน้า Login (ปรับปรุงใหม่)
# ==========================================
def step1_open_login():
    """ขั้นตอนที่ 1: เปิดหน้าเว็บ Login (ข้ามหากเปิดค้างและเข้าสู่ระบบอยู่แล้ว)"""
    global driver
    try:
        # 1. ตรวจสอบว่ามี Driver อยู่แล้วหรือไม่ และเบราว์เซอร์ยังไม่ถูกผู้ใช้กดปิดไป
        if driver is not None:
            try:
                current_url = driver.current_url

                # เช็กว่าเข้าสู่ระบบอยู่แล้วหรือไม่ (ไม่อยู่หน้า login.php และอยู่ในระบบ court_stat_program)
                if (
                    "court_stat_program" in current_url
                    and "login.php" not in current_url
                ):
                    return "✅ เบราว์เซอร์เปิดอยู่และอยู่ในระบบแล้ว! ระบบจะดำเนินการกรอกข้อมูล"
            except Exception:
                # ถ้าเช็กแล้วเจอว่าผู้ใช้แอบกดปิดหน้าต่าง Chrome ไปก่อนหน้า ให้รีเซ็ต driver
                driver = None

        # 2. ถ้ายังไม่มี driver หรือเพิ่งถูกเปิดใหม่ ให้เปิด Chrome และไปหน้า Login
        if driver is None:
            options = webdriver.ChromeOptions()
            options.add_experimental_option("detach", True)
            driver = webdriver.Chrome(options=options)

        print("\n" + "=" * 50)
        print("1. กำลังเปิดหน้า Login...")
        driver.get("http://10.1.2.9/court_stat_program/login.php")

        print(
            "2. >>> กรุณาพิมพ์ Username/Password และ Login ด้วยตัวคุณเองบน Chrome <<<"
        )
        print("=" * 50)

        return "🌐 เปิดหน้าเว็บเรียบร้อยแล้ว! กรุณา Login บนหน้าต่าง Chrome ให้เสร็จสิ้น แล้วค่อยกดปุ่มขั้นตอนที่ 2"

    except Exception as e:
        return f"❌ เกิดข้อผิดพลาดในการเปิดหน้าเว็บ: {e}"


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
        df_stat = all_sheets["สถิติรายเดือน"].fillna(0)

        # ดึงค่าตัวแปรจากเซลล์ต่างๆ
        aa6947 = get_safe_int(df_stat, 9, 16)
        aa6948 = get_safe_int(df_stat, 10, 16)
        aa6959 = get_safe_int(df_stat, 15, 8)
        aa6960 = get_safe_int(df_stat, 16, 8)
        aa6961 = get_safe_int(df_stat, 17, 8)
        aa6963 = get_safe_int(df_stat, 15, 11)
        aa6964 = get_safe_int(df_stat, 16, 11)
        aa6965 = get_safe_int(df_stat, 17, 11)
        aa6967 = get_safe_int(df_stat, 15, 15)
        aa6968 = get_safe_int(df_stat, 16, 15)
        aa6969 = get_safe_int(df_stat, 17, 15)
        aa6971 = get_safe_int(df_stat, 15, 20)
        aa6972 = get_safe_int(df_stat, 16, 20)
        aa6973 = get_safe_int(df_stat, 17, 20) 
        aa6975 = get_safe_int(df_stat, 15, 25)
        aa6976 = get_safe_int(df_stat, 16, 25)
        aa6977 = get_safe_int(df_stat, 17, 25)
        aa6979 = get_safe_int(df_stat, 15, 33)
        aa6980 = get_safe_int(df_stat, 16, 33)
        aa6981 = get_safe_int(df_stat, 17, 33)
        aa6983 = get_safe_int(df_stat, 15, 38)
        aa6984 = get_safe_int(df_stat, 16, 38)
        aa6985 = get_safe_int(df_stat, 17, 38) 
        aa6987 = get_safe_int(df_stat, 15, 43)
        aa6988 = get_safe_int(df_stat, 16, 43)
        aa6989 = get_safe_int(df_stat, 17, 43) 
        aa6991 = get_safe_int(df_stat, 15, 51)
        aa6992 = get_safe_int(df_stat, 16, 51)
        aa6993 = get_safe_int(df_stat, 17, 51)
        aa7003 = get_safe_int(df_stat, 21, 14)
        aa7004 = get_safe_int(df_stat, 22, 14)
        aa7005 = get_safe_int(df_stat, 23, 14)
        aa7007 = get_safe_int(df_stat, 21, 24)
        aa7008 = get_safe_int(df_stat, 22, 24)
        aa7009 = get_safe_int(df_stat, 23, 24)  
        aa7011 = get_safe_int(df_stat, 21, 37)
        aa7012 = get_safe_int(df_stat, 22, 37)
        aa7013 = get_safe_int(df_stat, 23, 37) 
        aa7023 = get_safe_int(df_stat, 28, 14)
        aa7024 = get_safe_int(df_stat, 29, 14)
        aa7025 = get_safe_int(df_stat, 30, 14) 
        aa7027 = get_safe_int(df_stat, 28, 23)
        aa7028 = get_safe_int(df_stat, 29, 23)
        aa7029 = get_safe_int(df_stat, 30, 23)  
        aa7031 = get_safe_int(df_stat, 28, 36)
        aa7032 = get_safe_int(df_stat, 29, 36)
        aa7033 = get_safe_int(df_stat, 30, 36)  
        aa7120 = get_safe_int(df_stat, 59, 6) 
        aa7121 = get_safe_int(df_stat, 60, 6)
        aa7122 = get_safe_int(df_stat, 61, 6)
        aa7124 = get_safe_int(df_stat, 59, 9) 
        aa7125 = get_safe_int(df_stat, 60, 9)
        aa7126 = get_safe_int(df_stat, 61, 9)
        aa7128 = get_safe_int(df_stat, 59, 13) 
        aa7129 = get_safe_int(df_stat, 60, 13)
        aa7130 = get_safe_int(df_stat, 61, 13)
        aa7132 = get_safe_int(df_stat, 59, 19) 
        aa7133 = get_safe_int(df_stat, 60, 19)
        aa7134 = get_safe_int(df_stat, 61, 19)
        aa7136 = get_safe_int(df_stat, 59, 22) 
        aa7137 = get_safe_int(df_stat, 60, 22)
        aa7138 = get_safe_int(df_stat, 61, 22)
        aa7140 = get_safe_int(df_stat, 59, 29) 
        aa7141 = get_safe_int(df_stat, 60, 29)
        aa7142 = get_safe_int(df_stat, 61, 29)
        aa7144 = get_safe_int(df_stat, 59, 35) 
        aa7145 = get_safe_int(df_stat, 60, 35)
        aa7146 = get_safe_int(df_stat, 61, 35)
        aa7148 = get_safe_int(df_stat, 59, 39) 
        aa7149 = get_safe_int(df_stat, 60, 39)
        aa7150 = get_safe_int(df_stat, 61, 39)
        aa7152 = get_safe_int(df_stat, 59, 47) 
        aa7153 = get_safe_int(df_stat, 60, 47)
        aa7154 = get_safe_int(df_stat, 61, 47)
        #ขั้นตอนคดีค้าง ค้างจ่าย
        aa18225 = get_safe_int(df_stat, 65, 23)
        #รอพิจารณานัดแรก
        aa18226 = get_safe_int(df_stat, 66, 23)
        #สืบพยาน
        aa18227 = get_safe_int(df_stat, 67, 23)
        #ไม่ให้กรอกข้อมูล 18228
        aa18229 = get_safe_int(df_stat, 68, 23)
        aa18230 = get_safe_int(df_stat, 69, 23) 
        aa18231 = get_safe_int(df_stat, 70, 23)
        aa18232 = get_safe_int(df_stat, 71, 23)
        aa18233 = get_safe_int(df_stat, 72, 23)
        aa18234 = get_safe_int(df_stat, 73, 23)
        aa18235 = get_safe_int(df_stat, 74, 23)
        aa18236 = get_safe_int(df_stat, 75, 23)
        aa18237 = get_safe_int(df_stat, 76, 23)
        aa18238 = get_safe_int(df_stat, 77, 23)
        aa18239 = get_safe_int(df_stat, 78, 23)
       # ในexcel ไม่มี บรรทัดนี้  aa18240 = get_safe_int(df_stat, 77, 23)
    
        aa18241 = get_safe_int(df_stat, 65, 36)
        aa18242 = get_safe_int(df_stat, 66, 36)
        aa18243 = get_safe_int(df_stat, 67, 36)
    
        aa18245 = get_safe_int(df_stat, 68, 36)
        aa18246 = get_safe_int(df_stat, 69, 36)
        aa18247 = get_safe_int(df_stat, 70, 36)
        aa18248 = get_safe_int(df_stat, 71, 36)
        aa18249 = get_safe_int(df_stat, 72, 36)
        aa18250 = get_safe_int(df_stat, 73, 36)
        aa18251 = get_safe_int(df_stat, 74, 36)
        aa18252 = get_safe_int(df_stat, 75, 36)
        aa18253 = get_safe_int(df_stat, 76, 36)
        aa18254 = get_safe_int(df_stat, 77, 36)
        aa18255 = get_safe_int(df_stat, 78, 36)
        aa18256 = get_safe_int(df_stat, 79, 36)
    # จำหน่ายคดีชั่วคราว
        aa10497 = get_safe_int(df_stat, 80, 5)
        aa10498 = get_safe_int(df_stat, 81, 5)
        aa10499 = get_safe_int(df_stat, 82, 5)    

        print(
            f"ดึงข้อมูลสำเร็จ -> aa6947: {aa6947}, aa6948: {aa6948}, aa6959: {aa6959}, aa6960: {aa6960}"
        )

        # --- 2. เปลี่ยนไปยังหน้าฟอร์ม fid=27 ---
        print("กำลังเปลี่ยนไปหน้าฟอร์ม fid=27...")
        driver.get(
            "http://10.1.2.9/court_stat_program/form_statistic.php?fid=27"
        )

        # ตั้งเวลารอ Element โหลดสูงสุด 30 วินาที
        wait = WebDriverWait(driver, 30)

        # --- 3. เริ่มกรอกข้อมูลอัตโนมัติ ---
        print("กำลังกรอกข้อมูลอัตโนมัติ...")

        # ช่อง a6947
        input_1 = wait.until(
            EC.presence_of_element_located((By.NAME, "a6947"))
        )
        input_1.clear()
        input_1.send_keys(str(aa6947))

        # ช่อง a6948
        input_2 = wait.until(
            EC.presence_of_element_located((By.NAME, "a6948"))
        )
        input_2.clear()
        input_2.send_keys(str(aa6948))

        # ช่อง a6959
        input_3 = wait.until(
            EC.presence_of_element_located((By.NAME, "a6959"))
        )
        input_3.clear()
        input_3.send_keys(str(aa6959))

        # ช่อง a6960
        input_4 = wait.until(
            EC.presence_of_element_located((By.NAME, "a6960"))
        )
        input_4.clear()
        input_4.send_keys(str(aa6960))
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "a6961")))
        input_1.clear()
        input_1.send_keys(str(aa6961))
        #
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "a6963")))
        input_1.clear()
        input_1.send_keys(str(aa6963))
        #
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "a6964")))
        input_1.clear()
        input_1.send_keys(str(aa6964))
        #
        input_1 = wait.until(EC.presence_of_element_located((By.NAME, "a6965")))
        input_1.clear()
        input_1.send_keys(str(aa6965))
        #
        input_2 = driver.find_element(By.NAME, "a6967")
        input_2.clear()
        input_2.send_keys(str(aa6967))
        #
        input_2 = driver.find_element(By.NAME, "a6968")
        input_2.clear()
        input_2.send_keys(str(aa6968))
        #
        input_2 = driver.find_element(By.NAME, "a6969")
        input_2.clear()
        input_2.send_keys(str(aa6969))
        #
        input_2 = driver.find_element(By.NAME, "a6971")
        input_2.clear()
        input_2.send_keys(str(aa6971))
        #
        input_2 = driver.find_element(By.NAME, "a6972")
        input_2.clear()
        input_2.send_keys(str(aa6972))
        #
        input_2 = driver.find_element(By.NAME, "a6973")
        input_2.clear()
        input_2.send_keys(str(aa6973))
        #
        input_2 = driver.find_element(By.NAME, "a6975")
        input_2.clear()
        input_2.send_keys(str(aa6975))   
        #
        input_2 = driver.find_element(By.NAME, "a6976")
        input_2.clear()
        input_2.send_keys(str(aa6976))   
        #
        input_2 = driver.find_element(By.NAME, "a6977")
        input_2.clear()
        input_2.send_keys(str(aa6977))      
        #
        input_2 = driver.find_element(By.NAME, "a6979")
        input_2.clear()
        input_2.send_keys(str(aa6979))  
        #
        input_2 = driver.find_element(By.NAME, "a6980")
        input_2.clear()
        input_2.send_keys(str(aa6980))  
        #
        input_2 = driver.find_element(By.NAME, "a6981")
        input_2.clear()
        input_2.send_keys(str(aa6981))
        #
        input_2 = driver.find_element(By.NAME, "a6983")
        input_2.clear()
        input_2.send_keys(str(aa6983)) 
        #
        input_2 = driver.find_element(By.NAME, "a6984")
        input_2.clear()
        input_2.send_keys(str(aa6984))  
        #
        input_2 = driver.find_element(By.NAME, "a6985")
        input_2.clear()
        input_2.send_keys(str(aa6985))
        #
        input_2 = driver.find_element(By.NAME, "a6987")
        input_2.clear()
        input_2.send_keys(str(aa6987)) 
        #
        input_2 = driver.find_element(By.NAME, "a6988")
        input_2.clear()
        input_2.send_keys(str(aa6988))  
        #
        input_2 = driver.find_element(By.NAME, "a6989")
        input_2.clear()
        input_2.send_keys(str(aa6989))
        #
        input_2 = driver.find_element(By.NAME, "a6991")
        input_2.clear()
        input_2.send_keys(str(aa6991)) 
        #
        input_2 = driver.find_element(By.NAME, "a6992")
        input_2.clear()
        input_2.send_keys(str(aa6992))  
        #
        input_2 = driver.find_element(By.NAME, "a6993")
        input_2.clear()
        input_2.send_keys(str(aa6993))
        #
        input_2 = driver.find_element(By.NAME, "a7003")
        input_2.clear()
        input_2.send_keys(str(aa7003))
        #
        input_2 = driver.find_element(By.NAME, "a7004")
        input_2.clear()
        input_2.send_keys(str(aa7004))
        #
        input_2 = driver.find_element(By.NAME, "a7005")
        input_2.clear()
        input_2.send_keys(str(aa7005))
        #
        input_2 = driver.find_element(By.NAME, "a7007")
        input_2.clear()
        input_2.send_keys(str(aa7007))
        #
        input_2 = driver.find_element(By.NAME, "a7008")
        input_2.clear()
        input_2.send_keys(str(aa7008))
        #
        input_2 = driver.find_element(By.NAME, "a7009")
        input_2.clear()
        input_2.send_keys(str(aa7009))
        #
        input_2 = driver.find_element(By.NAME, "a7011")
        input_2.clear()
        input_2.send_keys(str(aa7011))
        #
        input_2 = driver.find_element(By.NAME, "a7012")
        input_2.clear()
        input_2.send_keys(str(aa7012))
        #
        input_2 = driver.find_element(By.NAME, "a7013")
        input_2.clear()
        input_2.send_keys(str(aa7013))
        #
        input_2 = driver.find_element(By.NAME, "a7023")
        input_2.clear()
        input_2.send_keys(str(aa7023))
        #
        input_2 = driver.find_element(By.NAME, "a7024")
        input_2.clear()
        input_2.send_keys(str(aa7024))
        #
        input_2 = driver.find_element(By.NAME, "a7025")
        input_2.clear()
        input_2.send_keys(str(aa7025))
        #
        input_2 = driver.find_element(By.NAME, "a7027")
        input_2.clear()
        input_2.send_keys(str(aa7027))
        #
        input_2 = driver.find_element(By.NAME, "a7028")
        input_2.clear()
        input_2.send_keys(str(aa7028))
        #
        input_2 = driver.find_element(By.NAME, "a7029")
        input_2.clear()
        input_2.send_keys(str(aa7029))
        #
        input_2 = driver.find_element(By.NAME, "a7031")
        input_2.clear()
        input_2.send_keys(str(aa7031))
        #
        input_2 = driver.find_element(By.NAME, "a7032")
        input_2.clear()
        input_2.send_keys(str(aa7032))
        #
        input_2 = driver.find_element(By.NAME, "a7033")
        input_2.clear()
        input_2.send_keys(str(aa7033))
        #
        input_2 = driver.find_element(By.NAME, "a7120")
        input_2.clear()
        input_2.send_keys(str(aa7120))
        #
        input_2 = driver.find_element(By.NAME, "a7121")
        input_2.clear()
        input_2.send_keys(str(aa7121))
        #
        input_2 = driver.find_element(By.NAME, "a7122")
        input_2.clear()
        input_2.send_keys(str(aa7122))
        #
        input_2 = driver.find_element(By.NAME, "a7124")
        input_2.clear()
        input_2.send_keys(str(aa7124))
        #
        input_2 = driver.find_element(By.NAME, "a7125")
        input_2.clear()
        input_2.send_keys(str(aa7125))
        #
        input_2 = driver.find_element(By.NAME, "a7126")
        input_2.clear()
        input_2.send_keys(str(aa7126))
        #
        input_2 = driver.find_element(By.NAME, "a7128")
        input_2.clear()
        input_2.send_keys(str(aa7128))
        #
        input_2 = driver.find_element(By.NAME, "a7129")
        input_2.clear()
        input_2.send_keys(str(aa7129))
        #
        input_2 = driver.find_element(By.NAME, "a7130")
        input_2.clear()
        input_2.send_keys(str(aa7130))
        #
        input_2 = driver.find_element(By.NAME, "a7132")
        input_2.clear()
        input_2.send_keys(str(aa7132))
        #
        input_2 = driver.find_element(By.NAME, "a7133")
        input_2.clear()
        input_2.send_keys(str(aa7133))
        #
        input_2 = driver.find_element(By.NAME, "a7134")
        input_2.clear()
        input_2.send_keys(str(aa7134))
        #
        input_2 = driver.find_element(By.NAME, "a7136")
        input_2.clear()
        input_2.send_keys(str(aa7136))
        #
        input_2 = driver.find_element(By.NAME, "a7137")
        input_2.clear()
        input_2.send_keys(str(aa7137))
        #
        input_2 = driver.find_element(By.NAME, "a7138")
        input_2.clear()
        input_2.send_keys(str(aa7138))
        #
        input_2 = driver.find_element(By.NAME, "a7140")
        input_2.clear()
        input_2.send_keys(str(aa7140))
        #
        input_2 = driver.find_element(By.NAME, "a7141")
        input_2.clear()
        input_2.send_keys(str(aa7141))
        #
        input_2 = driver.find_element(By.NAME, "a7142")
        input_2.clear()
        input_2.send_keys(str(aa7142))
        #
        input_2 = driver.find_element(By.NAME, "a7144")
        input_2.clear()
        input_2.send_keys(str(aa7144))
        #
        input_2 = driver.find_element(By.NAME, "a7145")
        input_2.clear()
        input_2.send_keys(str(aa7145))
        #
        input_2 = driver.find_element(By.NAME, "a7146")
        input_2.clear()
        input_2.send_keys(str(aa7146))
        #
        input_2 = driver.find_element(By.NAME, "a7148")
        input_2.clear()
        input_2.send_keys(str(aa7148))
        #
        input_2 = driver.find_element(By.NAME, "a7149")
        input_2.clear()
        input_2.send_keys(str(aa7149))
        #
        input_2 = driver.find_element(By.NAME, "a7150")
        input_2.clear()
        input_2.send_keys(str(aa7150))
        #
        input_2 = driver.find_element(By.NAME, "a7152")
        input_2.clear()
        input_2.send_keys(str(aa7152))
        #
        input_2 = driver.find_element(By.NAME, "a7153")
        input_2.clear()
        input_2.send_keys(str(aa7153))
        #
        input_2 = driver.find_element(By.NAME, "a7154")
        input_2.clear()
        input_2.send_keys(str(aa7154))
        # เริ่มกรอกขั้นตอนคดีค้าง
        input_2 = driver.find_element(By.NAME, "a18225")
        input_2.clear()
        input_2.send_keys(str(aa18225))
        # สลับกับ 18226 เพราะใน excel กับเว็บสลับหัวข้อกัน
        input_2 = driver.find_element(By.NAME, "a18226")
        input_2.clear()
        input_2.send_keys(str(aa18227))
        # # สลับกับ 18227 เพราะใน excel กับเว็บสลับหัวข้อกัน
        input_2 = driver.find_element(By.NAME, "a18227")
        input_2.clear()
        input_2.send_keys(str(aa18226))
        # ข้าม18228 กรอกไม่ได้
        input_2 = driver.find_element(By.NAME, "a18229")
        input_2.clear()
        input_2.send_keys(str(aa18229))
        # 
        input_2 = driver.find_element(By.NAME, "a18230")
        input_2.clear()
        input_2.send_keys(str(aa18230))
        # 
        input_2 = driver.find_element(By.NAME, "a18231")
        input_2.clear()
        input_2.send_keys(str(aa18231))
        # 
        input_2 = driver.find_element(By.NAME, "a18232")
        input_2.clear()
        input_2.send_keys(str(aa18232))
        # 
        input_2 = driver.find_element(By.NAME, "a18233")
        input_2.clear()
        input_2.send_keys(str(aa18233))
        # 
        input_2 = driver.find_element(By.NAME, "a18234")
        input_2.clear()
        input_2.send_keys(str(aa18234))
        # 
        input_2 = driver.find_element(By.NAME, "a18235")
        input_2.clear()
        input_2.send_keys(str(aa18235))
        # 
        input_2 = driver.find_element(By.NAME, "a18236")
        input_2.clear()
        input_2.send_keys(str(aa18236))
        # 
        input_2 = driver.find_element(By.NAME, "a18237")
        input_2.clear()
        input_2.send_keys(str(aa18237))
        # 
        input_2 = driver.find_element(By.NAME, "a18238")
        input_2.clear()
        input_2.send_keys(str(aa18238))
        # 
        input_2 = driver.find_element(By.NAME, "a18239")
        input_2.clear()
        input_2.send_keys(str(aa18239))
        # 
       # input_2 = driver.find_element(By.NAME, "a18240")
     #   input_2.clear()
      #  input_2.send_keys(str(aa18240))
        # 
        input_2 = driver.find_element(By.NAME, "a18241")
        input_2.clear()
        input_2.send_keys(str(aa18241))
        # หัวข้อสลับ
        input_2 = driver.find_element(By.NAME, "a18242")
        input_2.clear()
        input_2.send_keys(str(aa18249))
        # # สลับกับ 18243 เพราะใน excel กับเว็บสลับหัวข้อกัน
        input_2 = driver.find_element(By.NAME, "a18243")
        input_2.clear()
        input_2.send_keys(str(aa18242))
        # ข้อมูลเรียงสลับ
        input_2 = driver.find_element(By.NAME, "a18245")
        input_2.clear()
        input_2.send_keys(str(aa18243))
        # 
        input_2 = driver.find_element(By.NAME, "a18246")
        input_2.clear()
        input_2.send_keys(str(aa18246))
        # 
        input_2 = driver.find_element(By.NAME, "a18247")
        input_2.clear()
        input_2.send_keys(str(aa18247))
        # 
        input_2 = driver.find_element(By.NAME, "a18248")
        input_2.clear()
        input_2.send_keys(str(aa18248))
        # 
        input_2 = driver.find_element(By.NAME, "a18249")
        input_2.clear()
        input_2.send_keys(str(aa18249))
        # 
        input_2 = driver.find_element(By.NAME, "a18250")
        input_2.clear()
        input_2.send_keys(str(aa18250))
        # 
        input_2 = driver.find_element(By.NAME, "a18251")
        input_2.clear()
        input_2.send_keys(str(aa18251))
        # 
        input_2 = driver.find_element(By.NAME, "a18252")
        input_2.clear()
        input_2.send_keys(str(aa18252))
        # 
        input_2 = driver.find_element(By.NAME, "a18253")
        input_2.clear()
        input_2.send_keys(str(aa18253))
        # 
        input_2 = driver.find_element(By.NAME, "a18254")
        input_2.clear()
        input_2.send_keys(str(aa18254))
        # 
        input_2 = driver.find_element(By.NAME, "a18255")
        input_2.clear()
        input_2.send_keys(str(aa18255))
        # ตัวแปรสุดท้าย 18256 ไม่มีซึ่งเป็นยอดผลรวม ยอดสุดท้าย 18255
        input_2 = driver.find_element(By.NAME, "a18256")
        input_2.clear()
        input_2.send_keys(str(aa18255))
        # จำหน่ายคดี แพ่ง
        input_2 = driver.find_element(By.NAME, "a10497")
        input_2.clear()
        input_2.send_keys(str(aa10497))
        # จำหน่ายคดี ผบ
        input_2 = driver.find_element(By.NAME, "a10498")
        input_2.clear()
        input_2.send_keys(str(aa10498))
        # จำหน่ายคดี สิง่แวดล้อม
        input_2 = driver.find_element(By.NAME, "a10499")
        input_2.clear()
        input_2.send_keys(str(aa10499))    

        end_time = time.time()

        elapsed_time = round(end_time - start_time, 2)
        msg = (
            f"✅ กรอกข้อมูลเสร็จเรียบร้อย!\n"
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