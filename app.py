import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
from thefuzz import process, fuzz
import io
import os

# --- 1. ระบบความปลอดภัย (รหัสผ่านหน้าแรก) ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False

    if not st.session_state.password_correct:
        st.title("🔐 กรุณาใส่รหัสผ่านเพื่อใช้งาน")
        password = st.text_input("Password", type="password")
        if st.button("Login"):
            if password == "1234":  # <--- เปลี่ยนรหัสผ่านเข้าเว็บที่นี่
                st.session_state.password_correct = True
                st.rerun()
            else:
                st.error("รหัสผ่านไม่ถูกต้อง")
        return False
    return True

if not check_password():
    st.stop()

# --- 2. ฟังก์ชันดึงข้อมูลจาก HTML ---
def extract_html_data(uploaded_file):
    content = uploaded_file.read().decode("utf-8")
    soup = BeautifulSoup(content, 'html.parser')
    rows_data = []
    
    # ตัดนามสกุลไฟล์ออกเพื่อใช้เป็นชื่อคอลัมน์
    file_display_name = os.path.splitext(uploaded_file.name)[0]
    
    for row in soup.find_all('tr'):
        cols = [ele.text.strip() for ele in row.find_all(['td', 'th'])]
        if len(cols) >= 5:
            menu_name = cols[2].replace('\n', ' ').strip()
            qty_raw = cols[3].strip()
            if qty_raw.isdigit():
                rows_data.append({
                    'Menu Name': menu_name,
                    'Quantity': int(qty_raw),
                    'Source File': file_display_name
                })
    return pd.DataFrame(rows_data)

# --- 3. ส่วนการจัดการ Dataset (Admin Zone) ---
st.sidebar.header("⚙️ การจัดการข้อมูลหลัก")
DEFAULT_DATASET = "mapping_rules.xlsx"

# โหลด Dataset เริ่มต้น
if 'current_dataset' not in st.session_state:
    if os.path.exists(DEFAULT_DATASET):
        st.session_state.current_dataset = pd.read_excel(DEFAULT_DATASET)
    else:
        st.session_state.current_dataset = None

# ระบบ Admin สำหรับอัปเดตไฟล์ (รหัสผ่านชั้นที่ 2)
with st.sidebar.expander("🛡️ Admin Only (อัปเดต Dataset)"):
    admin_pw = st.text_input("Admin Password", type="password")
    if admin_pw == "admin99": # <--- รหัสสำหรับอัปเดตไฟล์ Dataset
        uploaded_new_ds = st.file_uploader("อัปโหลดไฟล์ mapping_rules ใหม่", type=['xlsx'])
        if uploaded_new_ds:
            st.session_state.current_dataset = pd.read_excel(uploaded_new_ds)
            st.success("อัปเดต Dataset ชั่วคราวเรียบร้อย!")

# ตรวจสอบว่ามี Dataset ให้ทำงานไหม
if st.session_state.current_dataset is None:
    st.error("ไม่พบไฟล์ Dataset (mapping_rules.xlsx) กรุณาติดต่อ Admin")
    st.stop()

df_mapping = st.session_state.current_dataset
keywords = df_mapping.iloc[:, 0].dropna().tolist()

# --- 4. ส่วนการทำงานหลัก (Main UI) ---
st.title("📊 ระบบ AI สรุปยอดขาย (Multi-File)")
uploaded_files = st.sidebar.file_uploader("📥 อัปโหลดไฟล์ HTML", type=['html'], accept_multiple_files=True)

if uploaded_files:
    all_processed_data = []
    all_unmatched = []

    for uploaded_file in uploaded_files:
        df_sales = extract_html_data(uploaded_file)
        if not df_sales.empty:
            matched_categories = []
            for menu in df_sales['Menu Name']:
                match, score = process.extractOne(menu, keywords, scorer=fuzz.token_sort_ratio)
                if score >= 60:
                    cat = df_mapping[df_mapping.iloc[:, 0] == match].iloc[0, 1]
                    matched_categories.append(cat)
                else:
                    matched_categories.append("Unknown")
                    if menu not in all_unmatched: all_unmatched.append(menu)
            
            df_sales['Category'] = matched_categories
            all_processed_data.append(df_sales)

            # แสดงสรุปรายไฟล์
            with st.expander(f"📄 ดูสรุปไฟล์: {os.path.splitext(uploaded_file.name)[0]}"):
                file_sum = df_sales.groupby('Category')['Quantity'].sum().reset_index()
                st.bar_chart(file_sum.set_index('Category'))
                st.table(file_sum)

    # --- 5. การสร้างตารางเปรียบเทียบและการดาวน์โหลด ---
    if all_processed_data:
        final_df = pd.concat(all_processed_data, ignore_index=True)
        
        # สร้าง Pivot Table (ชื่อไฟล์เป็นหัว Column)
        pivot_table = final_df.pivot_table(
            index='Category',
            columns='Source File',
            values='Quantity',
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        # เพิ่มยอดรวมท้ายสุด
        pivot_table['Total All Files'] = pivot_table.iloc[:, 1:].sum(axis=1)

        st.divider()
        st.header("📦 รายงานสรุปเปรียบเทียบรายไฟล์")
        st.dataframe(pivot_table, use_container_width=True)

        # ปุ่มดาวน์โหลด
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            pivot_table.to_excel(writer, index=False, sheet_name='Summary_By_Files')
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์สรุป (Excel)",
            data=buffer.getvalue(),
            file_name="Master_Summary_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if all_unmatched:
            st.warning(f"⚠️ ตรวจพบ {len(all_unmatched)} เมนูที่ไม่พบใน Dataset กรุณาแจ้ง Admin")
            with st.expander("ดูรายชื่อเมนูที่ไม่พบ"):
                for item in all_unmatched: st.write(f"- {item}")
else:
    st.info("เริ่มต้นโดยการอัปโหลดไฟล์ HTML ที่แถบด้านซ้าย")
