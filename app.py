import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="BOM Checker Pro", layout="wide")
st.title("🚀 Đối chiếu BOM: Tối ưu Sai số, Công suất & Giá")

# --- 1. CÁC HÀM XỬ LÝ DỮ LIỆU NÂNG CAO ---

def is_valid_tech_type(type_val):
    t = str(type_val).upper()
    return "CAP" in t or "RES" in t

def clean_sync_value(val):
    if pd.isna(val): return None
    s = str(val).strip().upper()
    invalid_marks = ["#VALUE!", "#N/A", "NAN", "---", "", "NONE", "0", "#REF!"]
    if s in invalid_marks: return None
    return s

def clean_price(value):
    if pd.isna(value) or str(value).strip() in ["---", "", "NAN", "N/A"]: 
        return float('inf')
    text = str(value).replace('$', '').replace(',', '').strip()
    try: return float(text)
    except: return float('inf')

def parse_watt(watt_str):
    """Chuyển đổi 1/10W, 1/16W thành số thực 0.1, 0.0625"""
    if not watt_str: return 0.0
    try:
        if '/' in watt_str:
            num, den = watt_str.split('/')
            return float(num) / float(den)
        return float(watt_str)
    except: return 0.0

def extract_specs(desc):
    desc = str(desc).upper()
    specs = {'volt': 0.0, 'size': '', 'tol': 100.0, 'watt': 0.0}
    
    # 1. Bóc Size
    size_match = re.search(r'(0201|0402|0603|0805|1206|1210|2010|2512)', desc)
    if size_match: specs['size'] = size_match.group(1)
    
    # 2. Bóc Volt (Tụ)
    v_match = re.search(r'(\d+\.?\d*)\s*V', desc)
    if v_match: specs['volt'] = float(v_match.group(1))
    
    # 3. Bóc Sai số (%) - Lấy số trước dấu %
    tol_match = re.search(r'(\d+\.?\d*)\s*%', desc)
    if tol_match: specs['tol'] = float(tol_match.group(1))
    
    # 4. Bóc Công suất (W) - Lấy cụm trước W (xử lý cả phân số)
    w_match = re.search(r'(\d+/\d+|\d+\.?\d*)\s*W', desc)
    if w_match: specs['watt'] = parse_watt(w_match.group(1))
    
    return specs

# --- 2. GIAO DIỆN ---

master_file = st.sidebar.file_uploader("1. Nạp Master Data", type=['xlsx', 'xlsm'])
bom_file = st.sidebar.file_uploader("2. Nạp BOM List", type=['xlsx', 'xls'])

if master_file and bom_file:
    # Load Master Data
    dict_master = pd.read_excel(master_file, sheet_name=None)
    for s_name in dict_master:
        dict_master[s_name].columns = [str(c).strip() for c in dict_master[s_name].columns]

    # Load BOM
    df_bom = pd.read_excel(bom_file)
    df_bom.columns = [str(c).strip() for c in df_bom.columns]

    st.subheader("⚙️ Cấu hình cột")
    sample_cols = list(dict_master.values())[0].columns.tolist()
    
    col1, col2 = st.columns(2)
    with col1:
        st.info("Cấu hình Master")
        m_pn = st.selectbox("Cột P/N (Master):", sample_cols)
        m_val = st.selectbox("Cột Giá trị đồng bộ (Master):", sample_cols)
        m_desc = st.selectbox("Cột Mô tả (Master):", sample_cols)
        m_price = st.selectbox("Cột Giá 1000pcs (Master):", sample_cols)
        m_type = st.selectbox("Cột Loại hàng hóa (Master):", sample_cols)
        m_note = st.selectbox("Cột Ghi chú (Master):", sample_cols)
    with col2:
        st.info("Cấu hình BOM")
        b_cols = df_bom.columns.tolist()
        b_pn = st.selectbox("Cột P/N (BOM):", b_cols)
        b_qty = st.selectbox("Cột SL cần mua (BOM):", b_cols)
        b_val = st.selectbox("Cột Giá trị đồng bộ (BOM):", b_cols)
        b_desc = st.selectbox("Cột Mô tả (BOM):", b_cols)
        b_type = st.selectbox("Cột Loại hàng hóa (BOM):", b_cols)

    # Khởi tạo session state để lưu kết quả tạm
    if 'final_results' not in st.session_state:
        st.session_state.final_results = None

    if st.button("🚀 PHÂN TÍCH VÀ TÌM KIẾM", type="primary"):
        results = []
        for _, row in df_bom.iterrows():
            raw_type_bom = row.get(b_type, "")
            raw_val_bom = row.get(b_val, "")
            qty_bom = row.get(b_qty, 0)
            clean_v_bom = clean_sync_value(raw_val_bom)
            
            if not is_valid_tech_type(raw_type_bom) or clean_v_bom is None:
                continue
            
            pn_bom = str(row[b_pn]).strip().upper() if not pd.isna(row[b_pn]) else "N/A"
            specs_bom = extract_specs(row[b_desc])
            bom_size = specs_bom['size']

            # 1. Chỉ check 0402 và 0603
            if bom_size not in ["0402", "0603"]:
                results.append({"P/N BOM": pn_bom, "SL": qty_bom, "Size": bom_size, "Trạng thái": "⏩ GIỮ NGUYÊN", "Đề xuất": pn_bom, "Giá Đề Xuất": 0.0, "Lý do": "Size khác"})
                continue

            target_df = dict_master.get(bom_size, pd.DataFrame())
            if target_df.empty:
                results.append({"P/N BOM": pn_bom, "SL": qty_bom, "Size": bom_size, "Trạng thái": "❌ THIẾU SHEET", "Đề xuất": pn_bom, "Giá Đề Xuất": 0.0, "Lý do": f"Không có sheet {bom_size}"})
                continue

            # --- TÌM GIÁ GỐC ---
            original_item = target_df[target_df[m_pn].astype(str).str.upper().str.strip() == pn_bom]
            original_price = clean_price(original_item.iloc[0][m_price]) if not original_item.empty else float('inf')

            # --- LOGIC 1: ƯU TIÊN CHỮ "CHỌN" ---
            mask_select = (target_df[m_pn].astype(str).str.upper().str.strip() == pn_bom) & \
                          (target_df[m_note].astype(str).str.lower().str.contains("chọn", na=False))
            selected_item = target_df[mask_select]

            if not selected_item.empty:
                results.append({"P/N BOM": pn_bom, "SL": qty_bom, "Size": bom_size, "Trạng thái": "✅ ƯU TIÊN", "Đề xuất": pn_bom, "Giá Đề Xuất": original_price, "Lý do": "Đã duyệt 'Chọn'"})
            else:
                # --- LOGIC 2 & 3: TÌM ỨNG VIÊN THAY THẾ ---
                potential = target_df[
                    (target_df[m_type].apply(is_valid_tech_type)) & 
                    (target_df[m_val].apply(clean_sync_value) == clean_v_bom)
                ]
                
                valid_list = []
                for _, m_row in potential.iterrows():
                    m_specs = extract_specs(m_row[m_desc])
                    match = False
                    if "CAP" in str(raw_type_bom).upper():
                        # Tụ: Áp >= BOM
                        match = (m_specs['volt'] >= specs_bom['volt'])
                    else:
                        # Trở: Sai số <= BOM và Công suất >= BOM
                        match = (m_specs['tol'] <= specs_bom['tol']) and (m_specs['watt'] >= specs_bom['watt'])
                    
                    if match:
                        valid_list.append({
                            'pn': m_row[m_pn],
                            'price': clean_price(m_row[m_price]),
                            'is_bom_pn': (str(m_row[m_pn]).strip().upper() == pn_bom)
                        })
                
                if valid_list:
                    # Sắp xếp tìm thằng rẻ nhất
                    best = sorted(valid_list, key=lambda x: x['price'])[0]
                    res_status = "⚠️ CÓ MÃ THAY THẾ" if not best['is_bom_pn'] else "✅ GIỮ NGUYÊN (Rẻ nhất)"
                    results.append({"P/N BOM": pn_bom, "SL": qty_bom, "Size": bom_size, "Trạng thái": res_status, "Đề xuất": best['pn'], "Giá Đề Xuất": best['price'], "Lý do": "Tối ưu kỹ thuật & Giá"})
                else:
                    # TRƯỜNG HỢP 3: MÃ MỚI HOÀN TOÀN
                    results.append({"P/N BOM": pn_bom, "SL": qty_bom, "Size": bom_size, "Trạng thái": "✨ Mã MỚI", "Đề xuất": pn_bom, "Giá Đề Xuất": 0.0, "Lý do": "Chưa có mã đạt chuẩn trong Master"})

        st.session_state.final_results = pd.DataFrame(results)

    # --- 3. KHU VỰC NHẬP GIÁ VÀ XUẤT FILE ---
    if st.session_state.final_results is not None:
        st.divider()
        st.subheader("📝 Bước 2: Cập nhật giá cho Mã mới / Kiểm tra lại")
        st.write("Bạn có thể điền giá vào cột **Giá Đề Xuất** cho các dòng 'Mã MỚI' (Giá = 0) trước khi tải về.")
        
        # Cho phép sửa trực tiếp trên bảng
        edited_df = st.data_editor(
            st.session_state.final_results,
            column_config={
                "Giá Đề Xuất": st.column_config.NumberColumn("Giá 1000pcs", format="%.2f $"),
                "Trạng thái": st.column_config.TextColumn(disabled=True),
                "Size": st.column_config.TextColumn(disabled=True),
            },
            hide_index=True,
            use_container_width=True
        )

        # Xuất file
        col_dl1, col_dl2 = st.columns([1, 5])
        with col_dl1:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                edited_df.to_excel(writer, index=False)
            st.download_button("📥 TẢI FILE KẾT QUẢ", output.getvalue(), "BOM_Comparison_Final.xlsx", type="primary")
