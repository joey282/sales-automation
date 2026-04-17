import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
from thefuzz import process, fuzz
import io
import os

# --- การตั้งค่าเบื้องต้น ---
st.set_page_config(page_title="AI Data Automation System", layout="wide")

def extract_html_data(uploaded_file):
    """ฟังก์ชันดึงข้อมูลจากไฟล์ HTML"""
    # อ่านไฟล์และจัดการเรื่อง Encoding
    content = uploaded_file.read().decode("utf-8")
    soup = BeautifulSoup(content, 'html.parser')
    rows_data = []
    
    # ตัดนามสกุลไฟล์ออกเพื่อใช้เป็นชื่อคอลัมน์
    file_display_name = os.path.splitext(uploaded_file.name)[0]
    
    for row in soup.find_all('tr'):
        cols = [ele.text.strip() for ele in row.find_all(['td', 'th'])]
        # โครงสร้าง: คอลัมน์ 3=เมนู, 4=จำนวน
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

# --- ส่วน UI ---
st.title("📊 ระบบสรุปยอดขายแยกตามประเภท (Multi-File)")
st.markdown("โปรแกรมจะดึงข้อมูลจากไฟล์ HTML และเปรียบเทียบยอดขายตามประเภทสินค้าในแต่ละไฟล์")

# 1. โหลด Dataset (mapping_rules.xlsx)
dataset_path = "mapping_rules.xlsx"
if not os.path.exists(dataset_path):
    st.error(f"❌ ไม่พบไฟล์ '{dataset_path}' ในเครื่อง")
    st.stop()

df_mapping = pd.read_excel(dataset_path)
# สมมติ Column A คือ Keyword, Column B คือ Category
keywords = df_mapping.iloc[:, 0].dropna().tolist()

# 2. ส่วนอัปโหลดไฟล์
st.sidebar.header("📁 จัดการไฟล์")
uploaded_files = st.sidebar.file_uploader("อัปโหลดไฟล์ HTML", type=['html'], accept_multiple_files=True)

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

            # แสดงผลรายงานรายไฟล์บนหน้าจอ
            with st.expander(f"📄 รายงานสรุป: {os.path.splitext(uploaded_file.name)[0]}"):
                file_sum = df_sales.groupby('Category')['Quantity'].sum().reset_index()
                st.bar_chart(file_sum.set_index('Category'))
                st.dataframe(file_sum, use_container_width=True)

    # 3. รวมข้อมูลและทำตารางเปรีย contrast (Pivot Table)
    if all_processed_data:
        final_df = pd.concat(all_processed_data, ignore_index=True)
        
        # สร้างตารางเปรียบเทียบ (ชื่อไฟล์เป็นหัว Column)
        pivot_table = final_df.pivot_table(
            index='Category',
            columns='Source File',
            values='Quantity',
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        # เพิ่มคอลัมน์ยอดรวมท้ายสุด
        pivot_table['Total All Files'] = pivot_table.iloc[:, 1:].sum(axis=1)

        st.divider()
        st.header("📦 สรุปยอดรวมทุกไฟล์ (แยกตามประเภท)")
        st.dataframe(pivot_table, use_container_width=True)

        # 4. ส่วนการดาวน์โหลด
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            pivot_table.to_excel(writer, index=False, sheet_name='Summary_By_Category')
            
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์สรุปจำนวนแยกตามประเภท (Excel)",
            data=buffer.getvalue(),
            file_name="Sales_Summary_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # แจ้งเตือนรายการที่ไม่รู้จัก
        if all_unmatched:
            st.warning("⚠️ รายการเมนูที่ไม่พบใน Dataset")
            st.write(", ".join(all_unmatched))
else:
    st.info("กรุณาอัปโหลดไฟล์ HTML เพื่อเริ่มต้นการประมวลผล")