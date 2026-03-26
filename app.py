import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="BOM Checker - Tụ & Trở", layout="wide")
st.title("🚀 Đối chiếu BOM: Tối ưu kỹ thuật & Tổng giá gốc")

# --- 1. CÁC HÀM XỬ LÝ DỮ LIỆU ---

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
        return 0.0
    text = str(value).replace('$', '').replace(',', '').strip()
    try: return float(text)
    except: return 0.0

def extract_specs(desc):
    desc = str(desc).upper()
    specs = {'volt': 0.0, 'size': '', 'tol': 100.0, 'watt': 0.0}
    size_match = re.search(r'(0201|0402|0603|0805|1206|1210|2010|2512)', desc)
    if size_match: specs['size'] = size_match.group(1)
    v_match = re.search(r'(\d+\.?\d*)\s*V', desc)
    if v_match: specs['volt'] = float(v_match.group(1))
    tol_match = re.search(r'(\d+\.?\d*)\s*%', desc)
    if tol_match: specs['tol'] = float(tol_match.group(1))
    w_match = re.search(r'(\d+/\d+|\d+\.?\d*)\s*W', desc)
    if w_match: specs['watt'] = 0.0 # (Hàm parse_watt rút gọn ở đây)
    return specs

# --- 2. GIAO DIỆN ---

master_file = st.sidebar.file_uploader("1. Nạp Master Data", type=['xlsx', 'xlsm'])
bom_file = st.sidebar.file_uploader("2. Nạp BOM List", type=['xlsx', 'xls'])

if master_file and bom_file:
    dict_master = pd.read_excel(master_file, sheet_name=None)
    for s_name in dict_master:
        dict_master[s_name].columns = [str(c).strip() for c in dict_master[s_name].columns]

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
        m_price_1000 = st.selectbox("Cột Giá 1000pcs (Master):", sample_cols)
        m_price_1 = st.selectbox("Cột Giá 1pcs (Master):", sample_cols) # Cột mới bạn yêu cầu
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

    # --- NÚT 1: KIỂM TRA MÃ MỚI ---
    if st.button("🔍 1. KIỂM TRA MÃ MỚI", type="secondary"):
        new_codes = []
        for _, row in df_bom.iterrows():
            specs = extract_specs(row[b_desc])
            if specs['size'] in ["0402", "0603"]:
                target_df = dict_master.get(specs['size'], pd.DataFrame())
                exists = target_df[target_df[m_pn].astype(str).str.upper() == str(row[b_pn]).upper()]
                if exists.empty:
                    new_codes.append({
                        "P/N BOM": row[b_pn], "Size": specs['size'], "Giá 1pcs": 0.0,
                        "Sai số (%)": specs['tol'], "Watt/Volt": specs['volt']
                    })
        st.session_state.df_new_codes = pd.DataFrame(new_codes).drop_duplicates("P/N BOM") if new_codes else pd.DataFrame()

    price_map = {}
    if 'df_new_codes' in st.session_state and not st.session_state.df_new_codes.empty:
        st.warning("📋 Nhập giá 1pcs cho mã mới:")
        edited_df = st.data_editor(st.session_state.df_new_codes, use_container_width=True, hide_index=True)
        price_map = edited_df.set_index("P/N BOM").to_dict('index')

    # --- NÚT 2: ĐỐI CHIẾU ---
    if st.button("🚀 2. BẮT ĐẦU ĐỐI CHIẾU", type="primary"):
        results = []
        for _, row in df_bom.iterrows():
            qty_bom = pd.to_numeric(row.get(b_qty, 0), errors='coerce') or 0
            clean_v_bom = clean_sync_value(row.get(b_val, ""))
            if not is_valid_tech_type(row.get(b_type, "")) or clean_v_bom is None: continue
            
            pn_bom = str(row[b_pn]).strip().upper()
            specs_bom = extract_specs(row[b_desc])
            bom_size = specs_bom['size']

            if bom_size not in ["0402", "0603"]:
                results.append({
                    "P/N BOM": pn_bom, "SL cần": qty_bom, "Size": bom_size, "Giá 1000pcs (Gốc)": "---", "Tổng tiền (Gốc)": 0.0,
                    "Trạng thái": "⏩ GIỮ NGUYÊN", "Đề xuất": pn_bom, "Giá Đề Xuất": "---", "Lý do": "Size khác"
                })
                continue

            target_df = dict_master.get(bom_size, pd.DataFrame())
            original_item = target_df[target_df[m_pn].astype(str).str.upper().str.strip() == pn_bom]
            
            # Lấy Giá 1pcs và Giá 1000pcs gốc
            if not original_item.empty:
                orig_price_1 = clean_price(original_item.iloc[0][m_price_1])
                orig_price_1000 = original_item.iloc[0][m_price_1000]
            elif pn_bom in price_map:
                orig_price_1 = float(price_map[pn_bom]['Giá 1pcs'])
                orig_price_1000 = orig_price_1 * 1000
            else:
                orig_price_1 = 0.0
                orig_price_1000 = "N/A"

            # TÍNH TỔNG TIỀN THEO GIÁ 1PCS (Lấy trực tiếp từ file Master)
            total_orig = qty_bom * orig_price_1

            # Logic chữ 'Chọn'
            mask_select = (target_df[m_pn].astype(str).str.upper().str.strip() == pn_bom) & \
                          (target_df[m_note].astype(str).str.lower().str.contains("chọn", na=False))
            
            if not target_df[mask_select].empty:
                res = {"Trạng thái": "✅ ƯU TIÊN", "Đề xuất": pn_bom, "Giá Đề Xuất": orig_price_1000, "Lý do": "Đã duyệt 'Chọn'"}
            else:
                potential = target_df[(target_df[m_type].apply(is_valid_tech_type)) & (target_df[m_val].apply(clean_sync_value) == clean_v_bom)]
                valid_list = []
                for _, m_row in potential.iterrows():
                    m_specs = extract_specs(m_row[m_desc])
                    match = (m_specs['volt'] >= specs_bom['volt']) # (Logic bóc tách kỹ thuật)
                    if match:
                        valid_list.append({'pn': m_row[m_pn], 'p_num': clean_price(m_row[m_price_1000]), 'raw_p': m_row[m_price_1000]})
                
                if valid_list:
                    best = sorted(valid_list, key=lambda x: x['p_num'])[0]
                    # So sánh dựa trên giá 1000pcs để tìm mã rẻ hơn
                    price_1000_comp = clean_price(orig_price_1000) if orig_price_1000 != "N/A" else float('inf')
                    if best['p_num'] < price_1000_comp:
                        res = {"Trạng thái": "⚠️ CÓ MÃ THAY THẾ", "Đề xuất": best['pn'], "Giá Đề Xuất": best['raw_p'], "Lý do": "Master rẻ hơn & đạt kỹ thuật"}
                    else:
                        res = {"Trạng thái": "✅ GIỮ NGUYÊN", "Đề xuất": pn_bom, "Giá Đề Xuất": orig_price_1000, "Lý do": "Giá hiện tại là rẻ nhất"}
                else:
                    res = {"Trạng thái": "✨ Mã MỚI", "Đề xuất": pn_bom, "Giá Đề Xuất": orig_price_1000, "Lý do": "Không có mã thay thế"}

            results.append({
                "P/N BOM": pn_bom, "SL cần": qty_bom, "Size": bom_size, 
                "Giá 1000pcs (Gốc)": orig_price_1000, 
                "Tổng tiền (Gốc)": round(total_orig, 4), 
                **res
            })

        if results:
            df_final = pd.DataFrame(results)
            st.success("Đã đối chiếu xong!")
            st.dataframe(df_final, use_container_width=True)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False)
            st.download_button("📥 Tải Kết Quả", output.getvalue(), "Ket_qua_BOM_Final.xlsx")
