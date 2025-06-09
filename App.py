#!/usr/bin/env python
"""
App module for financial fraud detection using Beneish M-Score and Benford's Law.
"""

import re
import io
import base64
import json
import tempfile
from pathlib import Path
from io import BytesIO

import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import openpyxl
import markdown

from flask import Flask, request, render_template, redirect, url_for, session
from flask_session import Session
from account_mapping_utils import setup_account_mapping
from AI_model import prepare_excel, rag

matplotlib.use("Agg")  # Use non-GUI backend for server

app = Flask(__name__)
app.secret_key = "your_secret_key"
app.config["SESSION_TYPE"] = "filesystem"
Session(app)

MAPPING_PATH = Path(__file__).parent / "Data" / "account_mapping.xlsx"
standard_account_map, reverse_map, all_std_normalized, all_variant_normalized = (
    setup_account_mapping(MAPPING_PATH, sheet_name=0)
)
# Data extract

def dedup_names(names):
    """Ensure unique names by appending suffix to duplicates."""
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

def data_extract(uploaded_file, sheet_name):
    """
    Extract financial data from Excel, detect headers, bold formatting,
    and generate unique full account labels.
    """
    # Step 1: Detect header row (where the most year-like values are)
    data = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
    data.fillna(0, inplace=True)

    def is_year(s):
        if isinstance(s, (np.integer, int, float, np.floating)):
            s = str(int(s))
        if isinstance(s, str):
            return bool(re.fullmatch(r"20\d{2}|19\d{2}", s.strip()))
        return False

    max_year_count = 0
    header_row_index = 0
    for i in range(min(20, len(data))):
        year_count = sum(is_year(cell) for cell in data.iloc[i])
        if year_count > max_year_count:
            max_year_count = year_count
            header_row_index = i

    # Step 2: Read again with detected header row
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row_index)
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    df = df.loc[:, ~df.columns.duplicated()].reset_index(drop=True)
    df.columns = [str(col).strip() for col in df.columns]

    # Step 3: Find account column (auto-detect by string ratio)
    def ratio_string(col):
        non_null = col.dropna()
        return non_null.apply(lambda x: isinstance(x, str)).sum() / len(non_null) if len(non_null) else 0

    ratios = {col: ratio_string(df[col]) for col in df.columns}
    account_col = sorted(ratios.items(), key=lambda x: x[1], reverse=True)[0][0]
    account_col_idx = df.columns.get_loc(account_col) + 1

    # Step 4: Get IsBold info for the account column (aligning to actual data table)
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb[sheet_name]
    first_data_row = header_row_index + 2
    is_bold = [
        row[0].font.bold if row[0].font else False
        for row in ws.iter_rows(min_row=first_data_row, max_row=ws.max_row,
                                min_col=account_col_idx, max_col=account_col_idx)
    ]
    is_bold = is_bold[: len(df)]
    df["IsBold"] = is_bold

    # Step 5: Build Full Account with parent context and sheet name
    full_accounts = []
    parent_stack = []
    for _, row in df.iterrows():
        label = str(row[account_col]).strip()
        bold = row["IsBold"]
        if bold:
            parent_stack = [label]
            full_account = f"{sheet_name} - {label}"
        else:
            full_account = f"{sheet_name} - {' - '.join(parent_stack)} - {label}" if parent_stack else f"{sheet_name} - {label}"
        full_accounts.append(re.sub(r"\s+", " ", full_account))

    df["Full Account"] = dedup_names(full_accounts)
    return df.set_index("Full Account")

def transformer(uploaded_file):
    """Merge all sheets from the uploaded Excel file into a standardized DataFrame."""
    excel_file = pd.ExcelFile(uploaded_file)
    sheets = excel_file.sheet_names
    combined_data = [data_extract(uploaded_file, sheet) for sheet in sheets]
    final_data = pd.concat(combined_data, ignore_index=False)
    final_data.index = final_data.index.str.lower()
    return final_data

# Function to calculate all M-score inputs and result
def extract_year_columns(df):
    """Return columns that are likely year values."""
    return [
        col for col in df.columns
        if isinstance(col, int)
        or (isinstance(col, str) and col.strip().isdigit() and len(col.strip()) == 4)
    ]

# --- Beneish M-Score ---
def compute_m_score_components(df):
    """Compute Beneish M-Score components for each period."""
    data_dict = df.T.to_dict()
    year_columns = extract_year_columns(df)
    results = []

    def get(item, year):
        return data_dict[item][str(year)]

    for i in range(len(year_columns) - 1):
        y1, y2 = year_columns[i], year_columns[i + 1]

        try:
            dsri = (
                get("bảng cân đối kế toán - các khoản phải thu", y2)
                / get("kết quả kinh doanh - doanh thu thuần", y2)
            ) / (
                get("bảng cân đối kế toán - các khoản phải thu", y1)
                / get("kết quả kinh doanh - doanh thu thuần", y1)
            )

            gm_t = (
                get("kết quả kinh doanh - doanh thu thuần", y2)
                - get("kết quả kinh doanh - giá vốn hàng bán", y2)
            ) / get("kết quả kinh doanh - doanh thu thuần", y2)

            gm_t1 = (
                get("kết quả kinh doanh - doanh thu thuần", y1)
                - get("kết quả kinh doanh - giá vốn hàng bán", y1)
            ) / get("kết quả kinh doanh - doanh thu thuần", y1)

            gmi = gm_t1 / gm_t

            aqi = (
                1
                - (
                    get("bảng cân đối kế toán - tài sản ngắn hạn", y2)
                    + get("bảng cân đối kế toán - tài sản cố định", y2)
                )
                / get("bảng cân đối kế toán - tổng cộng tài sản", y2)
            ) / (
                1
                - (
                    get("bảng cân đối kế toán - tài sản ngắn hạn", y1)
                    + get("bảng cân đối kế toán - tài sản cố định", y1)
                )
                / get("bảng cân đối kế toán - tổng cộng tài sản", y1)
            )

            sgi = get("kết quả kinh doanh - doanh thu thuần", y2) / get(
                "kết quả kinh doanh - doanh thu thuần", y1
            )

            depi = (
                get("lưu chuyển tiền tệ - khấu hao tscđ và bđsđt", y1)
                / (
                    get("lưu chuyển tiền tệ - khấu hao tscđ và bđsđt", y1)
                    + get("bảng cân đối kế toán - tài sản cố định", y1)
                )
            ) / (
                get("lưu chuyển tiền tệ - khấu hao tscđ và bđsđt", y2)
                / (
                    get("lưu chuyển tiền tệ - khấu hao tscđ và bđsđt", y2)
                    + get("bảng cân đối kế toán - tài sản cố định", y2)
                )
            )

            sgai = (
                get("kết quả kinh doanh - chi phí quản lý doanh nghiệp", y2)
                / get("kết quả kinh doanh - doanh thu thuần", y2)
            ) / (
                get("kết quả kinh doanh - chi phí quản lý doanh nghiệp", y1)
                / get("kết quả kinh doanh - doanh thu thuần", y1)
            )

            lvgi = (
                get("bảng cân đối kế toán - nợ phải trả", y2)
                / get("bảng cân đối kế toán - tổng cộng tài sản", y2)
            ) / (
                get("bảng cân đối kế toán - nợ phải trả", y1)
                / get("bảng cân đối kế toán - tổng cộng tài sản", y1)
            )

            tata = (
                get("kết quả kinh doanh - lãi/(lỗ) thuần sau thuế", y2)
                - get("lưu chuyển tiền tệ - lưu chuyển tiền tệ ròng từ các hoạt động sản xuất kinh doanh", y2)
            ) / get("bảng cân đối kế toán - tổng cộng tài sản", y2)

            m_score = (-4.84 + 0.92 * dsri + 0.528 * gmi + 0.404 * aqi + 0.892 * sgi +
                       0.115 * depi - 0.172 * sgai + 4.679 * tata - 0.327 * lvgi)

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
        except (KeyError, ZeroDivisionError, TypeError):
            continue

    return results

# --- Benford's Distribution ---
def compute_benford_all_periods(df):
    # Filter by IsBold == True
    bold_df = df[df['IsBold'] == True]
    year_columns = extract_year_columns(df)
    benford_results = {}

    def leading_digit(value):
        while value < 1 and value != 0:
            value *= 10
        return int(str(abs(value))[0]) if value != 0 else 0

    for i in range(len(year_columns) - 1):
        y1, y2 = str(year_columns[i]), str(year_columns[i + 1])
        period = f"{y1}➞{y2}"
        values = pd.concat([bold_df[y1], bold_df[y2]]).values.flatten()
        values = [v for v in values if isinstance(v, (float, int)) and v > 0]
        digits = [leading_digit(v) for v in values]
        actual_counts = pd.Series(digits).value_counts().sort_index()
        actual_percentages = actual_counts / actual_counts.sum() * 100

        benford_dist = {d: np.log10(1 + 1 / d) * 100 for d in range(1, 10)}
        benford_df = pd.Series(benford_dist)

        comparison_df = pd.DataFrame({
            "Benford (%)": benford_df,
            "Actual (%)": actual_percentages
        }).fillna(0)

        mad = np.mean(np.abs(comparison_df["Actual (%)"] - comparison_df["Benford (%)"]))

        benford_results[period] = {
            "comparison_df": comparison_df,
            "mad": round(mad, 4)
        }

    return bold_df, benford_results

# ĐÂY LÀ PHẦN GIAO DIỆN
def fig_to_base64(fig):
    """Convert a Matplotlib figure to base64-encoded PNG image string."""
    buffer = BytesIO()
    fig.savefig(buffer, format="png", bbox_inches="tight")
    buffer.seek(0)
    img_base64 = base64.b64encode(buffer.read()).decode("utf-8")
    plt.close(fig)
    return img_base64

def mscore_line_chart(m_score_table):
    """Generate M-Score line chart and return base64 image string."""
    df = pd.DataFrame(m_score_table)
    fig, ax = plt.subplots(figsize=(6, 3))
    df.set_index("Period")["M-Score"].plot(ax=ax, marker="o", color="#1d4ed8", linewidth=2)
    ax.set_title("M-Score Over Time", color="#1d4ed8")
    ax.set_ylabel("M-Score", color="#1d4ed8")
    ax.set_xlabel("Period", color="#1d4ed8")
    ax.grid(True, color="#fef9c3", alpha=0.5)
    ax.tick_params(axis="x", colors="#4f5871")
    ax.tick_params(axis="y", colors="#4f5871")
    fig.patch.set_facecolor("#fff")
    return fig_to_base64(fig)

def benford_bar_chart(comparison_df, period):
    """Generate Benford comparison bar chart and return base64 image string."""
    fig, ax = plt.subplots(figsize=(6, 3))
    comparison_df.plot(
        kind="bar",
        ax=ax,
        color={"Benford (%)": "#ec4899", "Actual (%)": "#be1862"},
        edgecolor="#fff"
    )
    ax.set_xlabel("Leading Digit", color="#1d4ed8")
    ax.set_ylabel("Percentage", color="#1d4ed8")
    ax.set_title(f"Benford's Law vs Actual ({period})", color="#ec4899")
    ax.grid(True, color="#fef9c3", alpha=0.5)
    ax.tick_params(axis="x", colors="#4f5871")
    ax.tick_params(axis="y", colors="#4f5871")
    fig.patch.set_facecolor("#fff")
    return fig_to_base64(fig)

m_score_explanations = {
    "DSRI": (
        "A significant increase in receivables relative to sales may indicate premature revenue "
        "recognition to artificially boost earnings."
    ),
    "GMI": (
        "A declining gross margin may signal deteriorating business performance, prompting firms "
        "to manipulate profits."
    ),
    "AQI": (
        "An increase in long-term assets—excluding property, plant and equipment—relative to total "
        "assets suggests aggressive capitalization of costs, potentially inflating earnings."
    ),
    "SGI": (
        "While high growth does not imply manipulation, rapidly expanding firms may face greater "
        "pressure to meet market expectations, increasing the temptation to alter reported earnings."
    ),
    "DEPI": (
        "A decline in depreciation expense relative to net fixed assets may reflect changes in "
        "accounting estimates that increase reported income."
    ),
    "SGAI": (
        "A disproportionate rise in SG&A expenses compared to sales can be viewed negatively by "
        "analysts, incentivizing management to adjust earnings figures."
    ),
    "LVGI": (
        "An increase in leverage (total debt relative to total assets) can pressure firms to "
        "manipulate earnings to comply with debt covenants."
    ),
    "TATA": (
        "Higher accruals indicate greater use of discretionary accounting practices, which may be "
        "associated with earnings manipulation."
    )
}

def top_mscore_changes(m_score_table):
    """
    Returns a dictionary with period as key and list of top 3 changing variables as values.
    """
    variables = ["DSRI", "GMI", "AQI", "SGI", "DEPI", "SGAI", "LVGI", "TATA"]
    top_changes_by_period = {}

    for i in range(1, len(m_score_table)):
        curr = m_score_table[i]
        prev = m_score_table[i - 1]
        period = curr["Period"]

        deltas = []
        for var in variables:
            try:
                delta = curr[var] - prev[var]
                explanation = m_score_explanations.get(var, "")
                deltas.append((var, curr[var], delta, explanation))
            except Exception:
                continue

        top3 = sorted(deltas, key=lambda x: abs(x[2]), reverse=True)[:3]
        top_changes_by_period[period] = top3

    return top_changes_by_period

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file-upload")
        if file:
            try:
                # Read file into memory
                file_bytes = file.read()
                uploaded_file = io.BytesIO(file_bytes)
                # Process the file to get combined data
                final_data = transformer(uploaded_file)

                # Reset pointer for further reading
                uploaded_file.seek(0)
                # Save to a temp file for prepare_excel
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(file_bytes)
                    tmp_path = tmp.name
                session["excel_file_path"] = tmp_path  
                # Store data as JSON for dashboard use
                session["final_data"] = final_data.to_json()
                # Compute M-score and interpretation for dashboard
                m_score_table = compute_m_score_components(final_data)
                session["m_score_table"] = m_score_table  # Store the table for chart
                session["top_mscore_changes"] = top_mscore_changes(m_score_table)
                if m_score_table:
                    session["m_score"] = m_score_table[-1]["M-Score"]
                else:
                    session["m_score"] = None
                # Compute Benford results and store as JSON-serializable
                _, benford_results = compute_benford_all_periods(final_data)
                # Convert DataFrames to dict for JSON serialization
                benford_results_serializable = {}
                for period, res in benford_results.items():
                    benford_results_serializable[period] = {
                        "comparison_df": res["comparison_df"].to_dict(orient="index"),
                        "mad": res["mad"]
                    }
                session["benford_results"] = benford_results_serializable
                # Interpretation (optional)
                session["analysis_result"] = "See dashboard for details"
                # Redirect to dashboard to display results
                return redirect(url_for("dashboard"))
            except Exception as e:
                err_msg = f"Lỗi khi xử lý file: {e}"
                print("Error:", err_msg)
                return render_template("upload.html", error=err_msg)
    return render_template("upload.html")

@app.route("/dashboard", methods=["GET", "POST"])
def dashboard():
    if "analysis_result" not in session:
        return redirect(url_for("index"))
    final_data = pd.read_json(session["final_data"])
    m_score_table = session.get("m_score_table", [])
    benford_results = session.get("benford_results", {})

    # Prepare periods for dashboard
    year_columns = extract_year_columns(final_data)
    periods = []
    for i in range(len(year_columns) - 1):
        y1, y2 = year_columns[i], year_columns[i + 1]
        periods.append(f"{y1}-{y2}")

    # Get selected period from GET/POST param, default to most recent
    selected_period = request.values.get("selected_period")
    if not selected_period or selected_period not in periods:
        selected_period = periods[-1] if periods else None

    top_changes_by_period = session.get("top_mscore_changes", {})
    top_changes = top_changes_by_period.get(selected_period.replace("-", "➞"), [])

    # Find the corresponding M-Score value for the selected period
    m_score_value = None
    for row in m_score_table:
        if row.get("Period") == selected_period.replace("-", "➞"):
            m_score_value = row.get("M-Score")
            break

    # Prepare M-Score Plotly data (all periods for the line chart)
    mscore_plotly_data = {}
    if m_score_table:
        mscore_plotly_data = {
            "x": [row["Period"] for row in m_score_table],
            "y": [row["M-Score"] for row in m_score_table]
        }

    # Prepare Benford chart and MAD for the selected period
    mad = None
    plotly_benford_data = {}
    benford_key = selected_period.replace("-", "➞") if selected_period else None
    if benford_key and benford_key in benford_results:
        comparison_df = pd.DataFrame.from_dict(benford_results[benford_key]["comparison_df"], orient="index")
        mad = benford_results[benford_key]["mad"]
        plotly_benford_data = {
            "x": list(comparison_df.index),
            "benford": list(comparison_df["Benford (%)"]),
            "actual": list(comparison_df["Actual (%)"])
        }
    else:
        mad = None
        plotly_benford_data = {}

    # Prepare Beneish M-Score component bar chart data for selected period
    mscore_components_bar_data = {}
    variables = ["DSRI", "GMI", "AQI", "SGI", "DEPI", "SGAI", "LVGI", "TATA"]
    selected_period_label = selected_period.replace("-", "➞") if selected_period else None
    idx = None
    for i, row in enumerate(m_score_table):
        if row.get("Period") == selected_period_label:
            idx = i
            break
    if idx is not None:
        current = m_score_table[idx]
        previous = m_score_table[idx - 1] if idx > 0 else None
        mscore_components_bar_data = {
            "variables": variables,
            "current": [current[var] for var in variables],
            "current_label": selected_period_label.split("➞")[1] if selected_period_label and "➞" in selected_period_label else "T",
            "previous": [previous[var] for var in variables] if previous else [None]*len(variables),
            "previous_label": selected_period_label.split("➞")[0] if selected_period_label and "➞" in selected_period_label else "T-1"
        }
    # --- Q&A Block ---
    qa_history = session.get("qa_history", [])
    latest_answer = session.get("latest_answer", "")
    latest_qa = session.get("latest_qa", {})

    return render_template(
        "dashboard.html",
        result=session.get("analysis_result"),
        m_score=m_score_value,
        mad=mad,
        plotly_benford_data=json.dumps(plotly_benford_data),
        mscore_plotly_data=json.dumps(mscore_plotly_data),
        periods=periods,
        selected_period=selected_period,
        mscore_components_bar_data=json.dumps(mscore_components_bar_data),
        top_changes=top_changes,
        qa_history=qa_history,
        latest_answer=latest_answer,
        latest_qa=latest_qa,
    )

@app.route("/dashboard_analysis", methods=["GET"])
def dashboard_analysis():
    final_data = pd.read_json(session["final_data"])
    m_score_table = session.get("m_score_table", [])
    benford_results = session.get("benford_results", {})

    # Prepare periods for dashboard
    year_columns = extract_year_columns(final_data)
    periods = []
    for i in range(len(year_columns) - 1):
        y1, y2 = year_columns[i], year_columns[i + 1]
        periods.append(f"{y1}-{y2}")

    selected_period = request.args.get("selected_period")
    if not selected_period or selected_period not in periods:
        selected_period = periods[-1] if periods else None

    top_changes_by_period = session.get("top_mscore_changes", {})
    top_changes = top_changes_by_period.get(selected_period.replace("-", "➞"), [])

    # Find the corresponding M-Score value for the selected period
    m_score_value = None
    for row in m_score_table:
        if row.get("Period") == selected_period.replace("-", "➞"):
            m_score_value = row.get("M-Score")
            break

    # Prepare M-Score Plotly data (all periods for the line chart)
    mscore_plotly_data = {}
    if m_score_table:
        mscore_plotly_data = {
            "x": [row["Period"] for row in m_score_table],
            "y": [row["M-Score"] for row in m_score_table]
        }

    # Prepare Benford chart and MAD for the selected period
    mad = None
    plotly_benford_data = {}
    benford_key = selected_period.replace("-", "➞") if selected_period else None
    if benford_key and benford_key in benford_results:
        comparison_df = pd.DataFrame.from_dict(benford_results[benford_key]["comparison_df"], orient="index")
        mad = benford_results[benford_key]["mad"]
        plotly_benford_data = {
            "x": list(comparison_df.index),
            "benford": list(comparison_df["Benford (%)"]),
            "actual": list(comparison_df["Actual (%)"])
        }
    else:
        mad = None
        plotly_benford_data = {}

    # Prepare Beneish M-Score component bar chart data for selected period
    mscore_components_bar_data = {}
    variables = ["DSRI", "GMI", "AQI", "SGI", "DEPI", "SGAI", "LVGI", "TATA"]
    selected_period_label = selected_period.replace("-", "➞") if selected_period else None
    idx = None
    for i, row in enumerate(m_score_table):
        if row.get("Period") == selected_period_label:
            idx = i
            break
    if idx is not None:
        current = m_score_table[idx]
        previous = m_score_table[idx - 1] if idx > 0 else None
        mscore_components_bar_data = {
            "variables": variables,
            "current": [current[var] for var in variables],
            "current_label": selected_period_label.split("➞")[1] if selected_period_label and "➞" in selected_period_label else "T",
            "previous": [previous[var] for var in variables] if previous else [None]*len(variables),
            "previous_label": selected_period_label.split("➞")[0] if selected_period_label and "➞" in selected_period_label else "T-1"
        }

    return render_template(
        "dashboard_analysis_partial.html",
        m_score=m_score_value,
        mad=mad,
        plotly_benford_data=json.dumps(plotly_benford_data),
        mscore_plotly_data=json.dumps(mscore_plotly_data),
        mscore_components_bar_data=json.dumps(mscore_components_bar_data),
        top_changes=top_changes,
        selected_period=selected_period,
        periods=periods,
    )

@app.route("/ask", methods=["POST"])
def ask():
    question = request.form.get("user_question")
    excel_file_path = session.get("excel_file_path")
    excel_df = prepare_excel(excel_file_path) if excel_file_path else None
    answer = rag(excel_df, question) if excel_df else "Knowledge base not available."
    answer_html = markdown.markdown(answer, extensions=["fenced_code", "tables", "nl2br"])
    # Save latest Q&A to session
    session["latest_qa"] = {"question": question, "answer": answer_html}
    return answer_html  # Return only the answer HTML

if __name__ == "__main__":
    app.run(debug=True)
