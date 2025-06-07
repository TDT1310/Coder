# ĐÂY LÀ PHẦN KHAI BÁO THƯ VIỆN
from account_mapping_utils import setup_account_mapping, robust_get
from flask import Flask, request, render_template, redirect, url_for, session
from flask_session import Session
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import re
import io
import openpyxl
from pathlib import Path

app = Flask(__name__)
app.secret_key = "your_secret_key" 
app.config["SESSION_TYPE"] = "filesystem"  
Session(app)

# Standardize the path for account mapping
MAPPING_PATH = Path(__file__).parent / "Data" / "account_mapping.xlsx"
standard_account_map, reverse_map, all_std_normalized, all_variant_normalized = setup_account_mapping(MAPPING_PATH, sheet_name=0)

# DEF MỘT SỐ HÀM
# Data extract

def dedup_names(names):
    counts = {}
    result = []
    for name in names:
        if name not in counts:
            counts[name] = 0
            result.append(name)
        else:
            counts[name] += 1
            result.append(f"{name}.{counts[name]}")
    return result

def Data_extract(uploaded_file, sheet_name):

    """
    Extracts the financial data table from an Excel sheet, detects header,
    aligns with bold formatting, and builds a unique Full Account label.
    """
    # Step 1: Detect header row (where the most year-like values are)
    data = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    data.fillna(0, inplace=True)

    def is_year(s):
        if isinstance(s, (np.int64, np.float64, int, float)):
            s = str(int(s))
        if isinstance(s, str):
            return bool(re.fullmatch(r'20\d{2}|19\d{2}', s.strip()))
        return False

    max_year_count = 0
    header_row_index = 0
    for i in range(min(20, len(data))):  # Scan first 20 rows
        year_count = sum(is_year(cell) for cell in data.iloc[i])
        if year_count > max_year_count:
            max_year_count = year_count
            header_row_index = i

    # Step 2: Read again with detected header row
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row_index)
    df = df.dropna(axis=1, how='all')
    df = df.dropna(axis=0, how='all')
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.reset_index(drop=True)
    df.columns = [str(col).strip() for col in df.columns]

    # Step 3: Find account column (auto-detect by string ratio)
    def ratio_string(col):
        non_null = col.dropna()
        if len(non_null) == 0:
            return 0
        num_str = non_null.apply(lambda x: isinstance(x, str)).sum()
        return num_str / len(non_null)

    ratios = {col: ratio_string(df[col]) for col in df.columns}
    ratios_sorted = sorted(ratios.items(), key=lambda x: x[1], reverse=True)
    account_col = ratios_sorted[0][0]
    account_col_idx = df.columns.get_loc(account_col)+1

    # Step 4: Get IsBold info for the account column (aligning to actual data table)
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb[sheet_name]
    # openpyxl is 1-based, pandas is 0-based (+2 = header + 1)
    first_data_row = header_row_index + 2
    is_bold = []
    for idx, row in enumerate(ws.iter_rows(min_row=first_data_row, max_row=ws.max_row, min_col=account_col_idx, max_col=account_col_idx)):
        cell = row[0]
        is_bold.append(cell.font.bold if cell.font else False)
    is_bold = is_bold[:len(df)]
    df['IsBold'] = is_bold

    # Step 5: Build Full Account with parent context and sheet name
    full_accounts = []
    parent_stack = []
    for idx, row in df.iterrows():
        label = str(row[account_col]).strip()
        bold = row['IsBold']
        if bold:
            parent_stack = [label]
            full_account = f"{sheet_name} - {label}"
        else:
            if parent_stack:
                full_account = f"{sheet_name} - {' - '.join(parent_stack)} - {label}"
            else:
                full_account = f"{sheet_name} - {label}"
        full_account = re.sub(r'\s+', ' ', full_account)
        full_accounts.append(full_account)
    df['Full Account'] = full_accounts
    df['Full Account'] = dedup_names(df['Full Account'])
    df = df.set_index('Full Account')
    return df

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
    final_data.index = final_data.index.str.lower()
    return final_data

# Function to calculate all M-score inputs and result
def compute_m_score(y1, y2, uploaded_file):
    # Load CSV
    df = uploaded_file
    y1 = str(y1)
    y2 = str(y2)
    year_cols = [col for col in df.columns if re.match(r'^20\d{2}|19\d{2}$', str(col))]
    data_dict = df[year_cols].to_dict(orient='index')
    # Get function
    def get(item, year):
        return robust_get (item, year, data_dict, reverse_map, all_std_normalized, all_variant_normalized)
    # Detect year columns
    year_columns = [col for col in df.columns if type(col) == int]
    year_columns.sort() 
    def dsri():
        return (get("bảng cân đối kế toán - các khoản phải thu", y2) / get("kết quả kinh doanh - doanh thu thuần", y2)) / (get("bảng cân đối kế toán - các khoản phải thu", y1) / get("kết quả kinh doanh - doanh thu thuần", y1))
    def gmi():
        gm_t = (get("kết quả kinh doanh - doanh thu thuần", y2) - get("kết quả kinh doanh - giá vốn hàng bán", y2)) / get("kết quả kinh doanh - doanh thu thuần", y2)
        gm_t1 = (get("kết quả kinh doanh - doanh thu thuần", y1) - get("kết quả kinh doanh - giá vốn hàng bán", y1)) / get("kết quả kinh doanh - doanh thu thuần", y1)
        return gm_t1 / gm_t
    def aqi():
        num_t = 1 - (get("bảng cân đối kế toán - tài sản ngắn hạn", y2) + get("bảng cân đối kế toán - tài sản cố định", y2)) / get("bảng cân đối kế toán - tổng cộng tài sản", y2)
        num_t1 = 1 - (get("bảng cân đối kế toán - tài sản ngắn hạn", y1) + get("bảng cân đối kế toán - tài sản cố định", y1)) / get("bảng cân đối kế toán - tổng cộng tài sản", y1)
        return num_t / num_t1
    def sgi():
        return get("kết quả kinh doanh - doanh thu thuần", y2) / get("kết quả kinh doanh - doanh thu thuần", y1)
    def depi():
        rate_t = get("thuyết minh - chi phí sản xuất theo yếu tố - chi phí khấu hao tài sản cố định", y2) / (get("thuyết minh - chi phí sản xuất theo yếu tố - chi phí khấu hao tài sản cố định", y2) + get("bảng cân đối kế toán - tài sản cố định", y2))
        rate_t1 = get("thuyết minh - chi phí sản xuất theo yếu tố - chi phí khấu hao tài sản cố định", y1) / (get("thuyết minh - chi phí sản xuất theo yếu tố - chi phí khấu hao tài sản cố định", y1) + get("bảng cân đối kế toán - tài sản cố định", y1))
        return rate_t1 / rate_t
    def sgai():
        return (get("kết quả kinh doanh - chi phí quản lý doanh nghiệp", y2) / get("kết quả kinh doanh - doanh thu thuần", y2)) / (get("kết quả kinh doanh - chi phí quản lý doanh nghiệp", y1) / get("kết quả kinh doanh - doanh thu thuần", y1))
    def lvgi():
        return (get("bảng cân đối kế toán - nợ phải trả", y2) / get("bảng cân đối kế toán - tổng cộng tài sản", y2)) / (get("bảng cân đối kế toán - nợ phải trả", y1) / get("bảng cân đối kế toán - tổng cộng tài sản", y1))
    def tata():
        return (get("kết quả kinh doanh - lãi/(lỗ) thuần sau thuế", y2) - get("lưu chuyển tiền tệ - lưu chuyển tiền tệ ròng từ các hoạt động sản xuất kinh doanh", y2)) / get("bảng cân đối kế toán - tổng cộng tài sản", y2)
    # Compute all inputs
    dsri_v = round(dsri(), 4)
    gmi_v = round(gmi(), 4)
    aqi_v = round(aqi(), 4)
    sgi_v = round(sgi(), 4)
    depi_v = round(depi(), 4)
    sgai_v = round(sgai(), 4)
    lvgi_v = round(lvgi(), 4)
    tata_v = round(tata(), 4)
    # Calculate M-score
    m_score = -4.84 + 0.92*dsri_v + 0.528*gmi_v + 0.404*aqi_v + 0.892*sgi_v + 0.115*depi_v - 0.172*sgai_v + 4.679*tata_v - 0.327*lvgi_v
    m_score = round(m_score, 4)
    if m_score < -2.22:
      interpretation = "Unlikely Manipulation"
    elif m_score < -1.78:
      interpretation = "Possible Manipulation"
    else: interpretation = "Likely Manipulation"
    return m_score, interpretation
# ĐÂY LÀ PHẦN GIAO DIỆN

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file-upload")
        print("filr uploased:", file)
        if file:
            try:
                print("file is not empty")
                # Đọc file về bộ nhớ tạm
                file_bytes = file.read()
                uploaded_file = io.BytesIO(file_bytes)
                # Chạy transformer để lấy data tổng hợp
                final_data = transformer(uploaded_file)
                # Reset lại io cho Data_extract vì pandas cần đọc lại file từ đầu
                uploaded_file.seek(0)
                extracted_data = Data_extract(uploaded_file, "Thuyết minh")
                # Store only the data as JSON or CSV, not HTML
                print("final_data index:", list(final_data.index))
                session["final_data"] = final_data.to_json()
                session["extracted_data"] = extracted_data.to_json()
                m_score, interpretation = compute_m_score(2021, 2022, final_data)
                session["m_score"] = m_score
                session["analysis_result"] = interpretation
                print("redirecting to dashboard")
                return redirect(url_for("dashboard"))
            
            except Exception as e:
                err_msg = f"Lỗi khi xử lý file: {e}"
                print("Error:", err_msg)
                return render_template("upload.html", error=err_msg)

    return render_template("upload.html")

@app.route("/dashboard")
def dashboard():
    if "analysis_result" not in session:
        return redirect(url_for("index"))
    table_html = pd.read_json(session["final_data"]).to_html(classes="table table-striped", border=0)
    extracted_html = pd.read_json(session["extracted_data"]).to_html(classes="table table-striped", border=0)
    return render_template(
        "dashboard.html",
        result=session.get("analysis_result"),
        m_score=session.get("m_score"),
        table_html=table_html,
        extracted_html=extracted_html
    )

if __name__ == "__main__":
    app.run(debug=True)