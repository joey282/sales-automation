import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
from thefuzz import process, fuzz
from streamlit_gsheets import GSheetsConnection # เพิ่มส่วนนี้
import io
import os

# --- ตั้งค่าหน้าจอและรหัสผ่าน (เหมือนเดิม) ---
st.set_page_config(page_title="AI Sales Summary Pro", layout="wide")

def check_password():
    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False
    if not st.session_state.password_correct:
        st.title("🔐 กรุณาใส่รหัสผ่านเพื่อใช้งาน")
        password = st.text_input("Password", type="password")
        if st.button("Login"):
            if password == "1234":
                st.session_state.password_correct = True
                st.rerun()
            else:
                st.error("รหัสผ่านไม่ถูกต้อง")
        return False
    return True

if not check_password():
    st.stop()

# --- เชื่อมต่อ Google Sheets ---
# วาง URL ของ Google Sheets ของคุณที่นี่
SHEET_URL = "https://docs.google.com/spreadsheets/d/1cM7lKy8jq3wcjvBz3fH1tAmR9jjKZ2QDwETPDeDdlME/edit?usp=sharing"

try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    df_mapping = conn.read(spreadsheet=SHEET_URL)
    keywords = df_mapping.iloc[:, 0].dropna().tolist()
    st.sidebar.success("✅ เชื่อมต่อฐานข้อมูลเมนูแล้ว")
except Exception as e:
    st.error("❌ ไม่สามารถเชื่อมต่อ Google Sheets ได้ กรุณาตรวจสอบ URL หรือการแชร์")
    st.stop()

# --- ส่วนประมวลผล HTML (เหมือนเดิม) ---
def extract_html_data(uploaded_file):
    content = uploaded_file.read().decode("utf-8")
    soup = BeautifulSoup(content, 'html.parser')
    rows_data = []
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

# --- หน้าจอหลัก ---
st.title("📊 ระบบสรุปยอดขาย (Sync with Google Sheets)")

# ปุ่มกด Refresh ข้อมูลจาก Google Sheets
if st.sidebar.button("🔄 อัปเดตข้อมูลเมนูใหม่"):
    st.cache_data.clear()
    st.rerun()

uploaded_files = st.sidebar.file_uploader("📥 อัปโหลดไฟล์ HTML", type=['html'], accept_multiple_files=True)

if uploaded_files:
    all_processed_data = []
    all_unmatched = []
    # วนลูปประมวลผลทีละไฟล์
    for uploaded_file in uploaded_files:
        df_sales = extract_html_data(uploaded_file)
        
        if not df_sales.empty:
            matched_categories = []
            for menu in df_sales['Menu Name']:
                # Fuzzy Matching ค้นหาตัวที่ใกล้เคียงที่สุด
                match, score = process.extractOne(menu, keywords, scorer=fuzz.token_sort_ratio)
                
                if score >= 60:
                    cat = df_mapping[df_mapping.iloc[:, 0] == match].iloc[0, 1]
                    matched_categories.append(cat)
                else:
                    matched_categories.append("Unknown")
                    if menu not in all_unmatched:
                        all_unmatched.append(menu)
            
            df_sales['Category'] = matched_categories
            all_processed_data.append(df_sales)
    if all_processed_data:
        final_df = pd.concat(all_processed_data, ignore_index=True)
        pivot_table = final_df.pivot_table(
            index='Category',
            columns='Source File',
            values='Quantity',
            aggfunc='sum',
            fill_value=0
        ).reset_index()
        pivot_table['Total'] = pivot_table.iloc[:, 1:].sum(axis=1)

        st.header("📦 รายงานสรุปเปรียบเทียบ")
        st.dataframe(pivot_table, use_container_width=True)

        # ปุ่มดาวน์โหลด Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            pivot_table.to_excel(writer, index=False)
        st.download_button("📥 ดาวน์โหลดไฟล์สรุป (Excel)", data=buffer.getvalue(), file_name="Summary.xlsx")

        if all_unmatched:
            with st.expander("⚠️ รายการเมนูที่ไม่พบในฐานข้อมูล"):
                st.write("กรุณาเพิ่มรายการเหล่านี้ใน Google Sheets:")
                for item in all_unmatched: st.write(f"- {item}")
