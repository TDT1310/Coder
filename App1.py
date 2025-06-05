# Thay đổi từ def benford trở xuống
# ĐÂY LÀ PHẦN KHAI BÁO THƯ VIỆN

import pandas as pd
import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
import re

# DEF MỘT SỐ HÀM
# Data extract
def Data_extract (uploaded_file, name):
    # Đọc dữ liệu
    data = pd.read_excel(uploaded_file, sheet_name= name)

    # Tìm cột
    def ratio_string(col):
    # Đếm số giá trị kiểu chuỗi trong cột / tổng số giá trị không null
        non_null = col.dropna()
        if len(non_null) == 0:
            return 0
        num_str = non_null.apply(lambda x: isinstance(x, str)).sum()
        return num_str / len(non_null)
    
    # Tính tỷ lệ chuỗi cho từng cột
    ratios = {col: ratio_string(data[col]) for col in data.columns}
    # Sắp xếp theo tỷ lệ chuỗi giảm dần
    ratios_sorted = sorted(ratios.items(), key=lambda x: x[1], reverse=True)
    # Chọn cột có tỷ lệ chuỗi lớn nhất làm cột tài khoản
    account_col = ratios_sorted[0][0]
    account_col = data.columns.get_loc(account_col)

    #Tìm hàng
    def is_year (s):
        if isinstance(s, (np.int64, np.float64, int, float)):
            s = str(int(s))
        if isinstance(s, str):
            return bool(re.fullmatch(r'20\d{2}|19\d{2}', s.strip()))
        return False
    
    data = pd.read_excel(uploaded_file, sheet_name=name, header = None)
    data.fillna(0, inplace=True)
    max_year_count = 0
    header_row_index = 0

    for i in range(min(20, len(data))):
        year_count = 0
        row = data.iloc[i]
        for cell in row:
            if(is_year(cell) == True): year_count = year_count +1
        if year_count > max_year_count: 
            max_year_count = year_count
            header_row_index = i
        
    #chuyển đổi data và format sheet
    data = pd.read_excel(uploaded_file, index_col= int(account_col), sheet_name=name, header= header_row_index)
    data.index = data.index.str.lower()
    data.index = data.index.str.strip()
    data.index = data.index.str.replace(" ","_")
    data = data.dropna(axis = 0, how = 'all')
    data = data.dropna(axis = 1, how = 'all')
    data.fillna(0, inplace=True)

    #tiến hành thay thế index bằng tên cụ thể hơn
    return data

# Chuẩn hóa data
def transformer (uploaded_file):
     # Tiến hành imput file để phân tích 
    excel_file = pd.ExcelFile(uploaded_file)
    sheets = excel_file.sheet_names
    combined_data = []
    for names in sheets:
        data_test = Data_extract(uploaded_file, names)
        combined_data.append(data_test)
    #Kết hợp và ghi vào file excel mới
    final_data = pd.concat(combined_data, ignore_index=False)
    #final_data.to_excel("final_data.xlsx", index=True)
    return final_data

def compute_benford_all_periods(df):
    year_columns = [col for col in df.columns]
    benford_results = {}

    def leading_digit(v):
        while v < 1:
            v *= 10
        return int(str(v)[0])

    for i in range(len(year_columns) - 1):
        y1 = str(year_columns[i])
        y2 = str(year_columns[i + 1])
        period = f"{y1}➞{y2}"

        values = pd.concat([df[y1], df[y2]]).values.flatten()
        values = [v for v in values if isinstance(v, (float, int)) and v > 0]
        leading_digits = [leading_digit(v) for v in values]

        actual_counts = pd.Series(leading_digits).value_counts().sort_index()
        actual_percentages = actual_counts / actual_counts.sum() * 100

        benford_dist = {d: np.log10(1 + 1/d) * 100 for d in range(1, 10)}
        benford_df = pd.Series(benford_dist)

        comparison_df = pd.DataFrame({
            'Benford (%)': benford_df,
            'Actual (%)': actual_percentages
        }).fillna(0)

        mad = np.mean(np.abs(comparison_df['Actual (%)'] - comparison_df['Benford (%)']))


        benford_results[period] = {
            "comparison_df": comparison_df,
            "mad": round(mad, 4)
        }

    return benford_results

def compute_m_score_components(uploaded_data):
    df = uploaded_data
    data_dict = df.T.to_dict()
    year_columns = [col for col in df.columns]
    results = []

    def get(item, year):
        return data_dict[item][str(year)]

    for i in range(len(year_columns) - 1):
        y1 = year_columns[i]
        y2 = year_columns[i + 1]

        dsri = (get("các_khoản_phải_thu", y2) / get("doanh_thu_thuần", y2)) / (get("các_khoản_phải_thu", y1) / get("doanh_thu_thuần", y1))
        gmi = ((get("doanh_thu_thuần", y1) - get("giá_vốn_hàng_bán", y1)) / get("doanh_thu_thuần", y1)) / \
                  ((get("doanh_thu_thuần", y2) - get("giá_vốn_hàng_bán", y2)) / get("doanh_thu_thuần", y2))
        aqi = (1 - (get("tài_sản_ngắn_hạn", y2) + get("tài_sản_cố_định", y2)) / get("tổng_cộng_tài_sản", y2)) / \
                  (1 - (get("tài_sản_ngắn_hạn", y1) + get("tài_sản_cố_định", y1)) / get("tổng_cộng_tài_sản", y1))
        sgi = get("doanh_thu_thuần", y2) / get("doanh_thu_thuần", y1)
        depi = (get("chi_phí_khấu_hao_tài_sản_cố_định", y1) / (get("chi_phí_khấu_hao_tài_sản_cố_định", y1) + get("tài_sản_cố_định", y1))) / \
                   (get("chi_phí_khấu_hao_tài_sản_cố_định", y2) / (get("chi_phí_khấu_hao_tài_sản_cố_định", y2) + get("tài_sản_cố_định", y2)))
        sgai = (get("chi_phí_quản_lý_doanh__nghiệp", y2) / get("doanh_thu_thuần", y2)) / \
                   (get("chi_phí_quản_lý_doanh__nghiệp", y1) / get("doanh_thu_thuần", y1))
        lvgi = (get("nợ_phải_trả", y2) / get("tổng_cộng_tài_sản", y2)) / (get("nợ_phải_trả", y1) / get("tổng_cộng_tài_sản", y1))
        tata = (get("lãi/(lỗ)_thuần_sau_thuế", y2) - get("lưu_chuyển_tiền_tệ_ròng_từ_các_hoạt_động_sản_xuất_kinh_doanh", y2)) / get("tổng_cộng_tài_sản", y2)

        m_score = -4.84 + 0.92*dsri + 0.528*gmi + 0.404*aqi + 0.892*sgi + 0.115*depi - 0.172*sgai + 4.679*tata - 0.327*lvgi

        results.append({
            "Period": f"{y1}➞{y2}",
            "DSRI": round(dsri, 4),
            "GMI": round(gmi, 4),
            "AQI": round(aqi, 4),
            "SGI": round(sgi, 4),
            "DEPI": round(depi, 4),
            "SGAI": round(sgai, 4),
            "LVGI": round(lvgi, 4),
            "TATA": round(tata, 4),
            "M-Score": round(m_score, 4)
        })
    return results


#TEST TÝ

# ĐÂY LÀ PHẦN GIAO DIỆN

# Tên và hướng dẫn người dùng (Chắc là sau sẽ phải làm tý des đủng)
st.title ("Ứng dụng phát hiện gian lận báo cáo tài chính")
st.markdown ("""Vui lòng tải lên báo cáo tài chính dưới định dạng CSV hoặc Excel để phân tích.""")

# Upload tài liệu và phân tích

uploaded_file = st.file_uploader("Tải lên báo cáo tài chính",type=["xlsx", "csv"])
if uploaded_file is not None:
    st.write (transformer(uploaded_file))
    final_data = transformer(uploaded_file)
    benford_results = compute_benford_all_periods(final_data)


    st.subheader("M-Score Chi tiết theo từng giai đoạn")
    m_score_table = compute_m_score_components(final_data)
    m_score_df = pd.DataFrame(m_score_table)
    st.line_chart(m_score_df.set_index("Period")[["M-Score"]])

    # Detect year columns
    year_options = [col for col in final_data.columns]
    year_pairs = [(year_options[i], year_options[i+1]) for i in range(len(year_options)-1)]

    # Let user choose a year pair
    selected_pair = st.selectbox("Chọn giai đoạn phân tích:", year_pairs, format_func=lambda x: f"{x[0]} ➞ {x[1]}")
    y1, y2 = selected_pair

    # Collapsed sections for both analyses
    with st.expander("Kết quả Beneish M-Score"):
        selected_period = f"{y1}➞{y2}"
        row_index = m_score_df.index[m_score_df["Period"] == selected_period].tolist()

        if row_index:
            idx = row_index[0]
            current = m_score_df.loc[idx]
            previous = m_score_df.loc[idx - 1] if idx > 0 else None

        def format_var(var):
            val = current[var]
            delta_str = ""
            if previous is not None:
                delta = val - previous[var]
                color = "red" if delta > 0 else "green" if delta < 0 else "gray"
                sign = "+" if delta > 0 else ""
                delta_str = f' <span style="color:{color}; font-size: 0.9em">({sign}{delta:.4f})</span>'
            else:
                delta_str = ' <span style="color:gray; font-size: 0.9em"></span>'
            return f"{val:.4f}{delta_str}"

        # Build HTML output
        html_output = f"""
        <h4>M-score Analysis ({selected_period})</h4>
        <ul>
        <li><b>DSRI</b>: {format_var('DSRI')}</li>
        <li><b>GMI</b>: {format_var('GMI')}</li>
        <li><b>AQI</b>: {format_var('AQI')}</li>
        <li><b>SGI</b>: {format_var('SGI')}</li>
        <li><b>DEPI</b>: {format_var('DEPI')}</li>
        <li><b>SGAI</b>: {format_var('SGAI')}</li>
        <li><b>LVGI</b>: {format_var('LVGI')}</li>
        <li><b>TATA</b>: {format_var('TATA')}</li>
        </ul>
        <h4><b>M-Score</b>: <code>{format_var('M-Score')}</code></h4>
        """

        st.markdown(html_output, unsafe_allow_html=True)

        # Interpretation
        score = current['M-Score']
        if score < -2.22:
            st.success("Unlikely Manipulation")
        elif score < -1.78:
            st.warning("Possible Manipulation")
        else:
            st.error("Likely Manipulation")




    with st.expander("Phân tích Benford's Law"):
        period_key = f"{y1}➞{y2}"
        bdata = benford_results[period_key]
        comparison_df = bdata["comparison_df"]
        mad = bdata["mad"]

        st.subheader(f"Benford Analysis ({period_key})")
        st.markdown("Comparison Table")
        st.dataframe(comparison_df.style.format("{:.2f}"))
        st.markdown("Bar Chart")
        fig, ax = plt.subplots(figsize=(10, 5))
        comparison_df.plot(kind='bar', ax=ax)
        ax.set_xlabel("Leading Digit")
        ax.set_ylabel("Percentage")
        ax.set_title(f"Benford's Law vs Actual ({period_key})")
        ax.grid(True)
        st.pyplot(fig)
        st.markdown("MAD (Mean Absolute Deviation) Test")
        st.markdown(f"**MAD:** `{mad:.4f}`")

        # Interpretation
        if mad <= 0.006:
            st.success("✅ Close conformity with Benford's Law")
        elif mad <= 0.012:
            st.info("🟡 Acceptable conformity")
        elif mad <= 0.015:
            st.warning("⚠️ Marginally acceptable conformity")
        else:
            st.error("❌ Nonconformity — possible anomaly")


else: st.info("Vui lòng tải lên các báo cáo tài chính cần phân tích")
