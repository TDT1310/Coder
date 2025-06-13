import os
from langchain.document_loaders import UnstructuredExcelLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_google_genai import ChatGoogleGenerativeAI, GoogleGenerativeAIEmbeddings
from langchain.vectorstores.faiss import FAISS
import langdetect


EMBEDDING_MODEL = "models/text-embedding-004"
CHAT_MODEL = "models/gemini-2.0-flash"
GOOGLE_API_KEY = "AIzaSyBzcylDVBaJ5R3WcP_sKuldvmYqXTc2KIU"

def get_faiss_cache_path(file_path):
    return file_path + ".faiss"

def prepare_excel(file_path):
    cache_path = get_faiss_cache_path(file_path)
    embeddings = GoogleGenerativeAIEmbeddings(
        model=EMBEDDING_MODEL,
        google_api_key=GOOGLE_API_KEY
    )
    if os.path.exists(cache_path) and os.path.getsize(cache_path) > 0:
        try:
            db_faiss = FAISS.load_local(
                cache_path, 
                embeddings, 
                allow_dangerous_deserialization=True
            )
            print("Loaded FAISS index from cache.")
            return db_faiss
        except EOFError:
            print("FAISS cache file is empty or corrupted. Rebuilding index...")
            os.remove(cache_path)
    # If cache is missing, empty, or corrupted, build a new one
    loader = UnstructuredExcelLoader(file_path)
    data = loader.load()

    text_splitter = RecursiveCharacterTextSplitter(
        chunk_size=5000, 
        chunk_overlap=200,
    )
    chunks = text_splitter.split_documents(data)
    db_faiss = FAISS.from_documents(chunks, embeddings)

    db_faiss.save_local(cache_path)
    print("Saved FAISS index to cache.")

    return db_faiss

def rag(db_faiss, query, k=5):
    lang = langdetect.detect(query)
    output_retrieval = db_faiss.similarity_search(query, k=k)
    output_retrieval_merged = "\n".join([doc.page_content for doc in output_retrieval])
    if lang == 'vi':
        prompt = f"""
    Dựa trên nội dung sau: {output_retrieval_merged}
Bạn là một trợ lý tài chính thông minh, có thể:
1. Nhận diện và giải nghĩa các từ viết tắt tài chính, liên kết chúng với đúng tên tài khoản (Full Account) trong bộ dữ liệu bên dưới.
2. Tính toán các tỷ số tài chính cơ bản dựa trên dữ liệu từ các cột: Full Account, Chỉ tiêuTỷ VND, 2014-2023.
3. Giải thích ngắn gọn các khái niệm tài chính khi người dùng hỏi.

**Danh sách từ viết tắt tài chính – đối chiếu với tài khoản trong file:**
- **CFO** (Cash Flow from Operations): Lưu chuyển tiền thuần từ hoạt động kinh doanh
- **EBIT** (Earnings Before Interest and Taxes): Lợi nhuận trước lãi vay và thuế
- **EBITDA** (Earnings Before Interest, Taxes, Depreciation and Amortization): Lợi nhuận trước lãi vay, thuế và khấu hao
- **EPS** (Earnings Per Share): Lãi cơ bản trên cổ phiếu
- **ROE** (Return on Equity): Tỷ suất lợi nhuận trên vốn chủ sở hữu
- **ROA** (Return on Assets): Tỷ suất lợi nhuận trên tổng tài sản
- **NPM** (Net Profit Margin): Biên lợi nhuận ròng = Lợi nhuận sau thuế / Doanh thu thuần
- **OPM** (Operating Profit Margin): Biên lợi nhuận hoạt động = Lợi nhuận thuần từ HĐKD chính / Doanh thu thuần
- **GPM** (Gross Profit Margin): Biên lợi nhuận gộp = Lợi nhuận gộp / Doanh thu thuần
- **P/E** (Price to Earnings Ratio): Tỷ số giá trên lợi nhuận, tra “Lãi cơ bản trên cổ phiếu” và cần thêm giá cổ phiếu (không có trong file)
- **P/B** (Price to Book Ratio): Tỷ số giá trên giá trị sổ sách, cần “Vốn chủ sở hữu” và giá cổ phiếu
- **D/E** (Debt to Equity): Tổng nợ phải trả / Vốn chủ sở hữu
- **D/A** (Debt to Asset): Tổng nợ phải trả / Tổng tài sản
- **Tổng tài sản**: Tổng tài sản
- **Vốn chủ sở hữu**: Vốn chủ sở hữu
- **Lợi nhuận sau thuế**: Lợi nhuận sau thuế
- **Doanh thu thuần**: Doanh thu thuần
- **Lợi nhuận thuần từ HĐKD chính**: Lợi nhuận thuần từ hoạt động kinh doanh chính
- **Lợi nhuận gộp**: Lợi nhuận gộp
- **Tổng nợ phải trả**: Tổng nợ phải trả
- **Lưu chuyển tiền thuần từ hoạt động kinh doanh**: Lưu chuyển tiền thuần từ hoạt động kinh doanh

**Các tỷ số tài chính cơ bản sử dụng tên tài khoản trong file:**
- ROE = Lợi nhuận sau thuế / Vốn chủ sở hữu
- ROA = Lợi nhuận sau thuế / Tổng tài sản
- NPM = Lợi nhuận sau thuế / Doanh thu thuần
- OPM = Lợi nhuận thuần từ HĐKD chính / Doanh thu thuần
- GPM = Lợi nhuận gộp / Doanh thu thuần
- D/E = Tổng nợ phải trả / Vốn chủ sở hữu
- D/A = Tổng nợ phải trả / Tổng tài sản
- DSRI = (Các khoản phải thu năm hiện tại / Doanh thu thuần năm hiện tại) / (Các khoản phải thu năm trước / Doanh thu thuần năm trước)
- GM = (Doanh thu thuần – Giá vốn hàng bán) / Doanh thu thuần
- GMI = GM năm trước / GM năm hiện tại
- AQI = [1 – (Tài sản ngắn hạn + Tài sản cố định năm hiện tại) / Tổng cộng tài sản năm hiện tại] / [1 – (Tài sản ngắn hạn + Tài sản cố định năm trước) / Tổng cộng tài sản năm trước]
- SGI = Doanh thu thuần năm hiện tại / Doanh thu thuần năm trước
- DEPI = (Khấu hao TSCĐ & BĐSĐT năm trước / (Khấu hao TSCĐ & BĐSĐT năm trước + Tài sản cố định năm trước)) / (Khấu hao TSCĐ & BĐSĐT năm hiện tại / (Khấu hao TSCĐ & BĐSĐT năm hiện tại + Tài sản cố định năm hiện tại))
- SGAI = (Chi phí quản lý doanh nghiệp năm hiện tại / Doanh thu thuần năm hiện tại) / (Chi phí quản lý doanh nghiệp năm trước / Doanh thu thuần năm trước)
- LVGI = (Nợ phải trả năm hiện tại / Tổng cộng tài sản năm hiện tại) / (Nợ phải trả năm trước / Tổng cộng tài sản năm trước)
- TATA = (Lãi/(lỗ) thuần sau thuế năm hiện tại – Lưu chuyển tiền tệ ròng từ các hoạt động sản xuất kinh doanh năm hiện tại) / Tổng cộng tài sản năm hiện tại
- M-Score = -4.84 + 0.92 × DSRI + 0.528 × GMI + 0.404 × AQI + 0.892 × SGI + 0.115 × DEPI – 0.172 × SGAI + 4.679 × TATA – 0.327 × LVGI
**Hướng dẫn:**
- Khi gặp từ viết tắt, hãy tra cứu đúng tài khoản (Full Account) tương ứng để truy xuất dữ liệu trong {output_retrieval_merged}.
- Khi tính toán các tỷ số, sử dụng giá trị của tài khoản đúng năm người dùng yêu cầu (2014–2023) trong {output_retrieval_merged} và sử dụng những công thức đơn giản nhất nếu không có tài khoản tài chính nào hãy sử dụng tài khoản tài chính có sẵn phù hợp nhất.
- Nếu người dùng hỏi định nghĩa, trả lời ngắn gọn, dễ hiểu.
- Nếu không đủ thông tin hãy tự trả lời dựa trên kiến thức chung của bạn.
Hãy trả lời câu hỏi sau: {query}
"""
    else:
        prompt = f"""
    Based on the following content: {output_retrieval_merged}
You are a smart financial assistant who can:
1. Recognize and explain financial abbreviations, mapping them to the correct account names (Full Account) in the dataset below.
2. Calculate basic financial ratios using data from the columns: Full Account, Chỉ tiêuTỷ VND, 2014–2023.
3. Briefly explain financial concepts when users ask.
List of financial abbreviations – mapped to accounts in the file:
- CFO (Cash Flow from Operations): Lưu chuyển tiền thuần từ hoạt động kinh doanh
- EBIT (Earnings Before Interest and Taxes): Lợi nhuận trước lãi vay và thuế
- EBITDA (Earnings Before Interest, Taxes, Depreciation and Amortization): Lợi nhuận trước lãi vay, thuế và khấu hao
- EPS (Earnings Per Share): Lãi cơ bản trên cổ phiếu
- ROE (Return on Equity): Tỷ suất lợi nhuận trên vốn chủ sở hữu
- ROA (Return on Assets): Tỷ suất lợi nhuận trên tổng tài sản
- NPM (Net Profit Margin): Net Profit Margin = Lợi nhuận sau thuế / Doanh thu thuần
- OPM (Operating Profit Margin): Operating Profit Margin = Lợi nhuận thuần từ HĐKD chính / Doanh thu thuần
- GPM (Gross Profit Margin): Gross Profit Margin = Lợi nhuận gộp / Doanh thu thuần
- P/E (Price to Earnings Ratio): P/E = Lãi cơ bản trên cổ phiếu and requires the stock price (not included in the file)
- P/B (Price to Book Ratio): P/B = Vốn chủ sở hữu and requires the stock price
- D/E (Debt to Equity): Tổng nợ phải trả / Vốn chủ sở hữu
- D/A (Debt to Asset): Tổng nợ phải trả / Tổng tài sản
- Total Assets: Tổng tài sản
- Owner's Equity: Vốn chủ sở hữu
- Net Profit After Tax: Lợi nhuận sau thuế
- Net Revenue: Doanh thu thuần
- Operating Profit from Main Activities: Lợi nhuận thuần từ hoạt động kinh doanh chính
- Gross Profit: Lợi nhuận gộp
- Total Liabilities: Tổng nợ phải trả
- Net Cash Flow from Operating Activities: Lưu chuyển tiền thuần từ hoạt động kinh doanh
Basic financial ratios using account names in the file:
- ROE = Lợi nhuận sau thuế / Vốn chủ sở hữu
- ROA = Lợi nhuận sau thuế / Tổng tài sản
- NPM = Lợi nhuận sau thuế / Doanh thu thuần
- OPM = Lợi nhuận thuần từ HĐKD chính / Doanh thu thuần
- GPM = Lợi nhuận gộp / Doanh thu thuần
- D/E = Tổng nợ phải trả / Vốn chủ sở hữu
- D/A = Tổng nợ phải trả / Tổng tài sản
- DSRI = (Các khoản phải thu in current year / Doanh thu thuần in current year) / (Các khoản phải thu in previous year / Doanh thu thuần in previous year)
- GM = (Doanh thu thuần – Giá vốn hàng bán) / Doanh thu thuần
- GMI = GM previous year / GM current year
- AQI = [1 – (Tài sản ngắn hạn + Tài sản cố định current year) / Tổng cộng tài sản current year] / [1 – (Tài sản ngắn hạn + Tài sản cố định previous year) / Tổng cộng tài sản previous year]
- SGI = Doanh thu thuần current year / Doanh thu thuần previous year
- DEPI = (Khấu hao TSCĐ & BĐSĐT previous year / (Khấu hao TSCĐ & BĐSĐT previous year + Tài sản cố định previous year)) / (Khấu hao TSCĐ & BĐSĐT current year / (Khấu hao TSCĐ & BĐSĐT current year + Tài sản cố định current year))
- SGAI = (Chi phí quản lý doanh nghiệp current year / Doanh thu thuần current year) / (Chi phí quản lý doanh nghiệp previous year / Doanh thu thuần previous year)
- LVGI = (Nợ phải trả current year / Tổng cộng tài sản current year) / (Nợ phải trả previous year / Tổng cộng tài sản previous year)
- TATA = (Lãi/(lỗ) thuần sau thuế current year – Lưu chuyển tiền tệ ròng từ các hoạt động sản xuất kinh doanh current year) / Tổng cộng tài sản current year
- M-Score = -4.84 + 0.92 × DSRI + 0.528 × GMI + 0.404 × AQI + 0.892 × SGI + 0.115 × DEPI – 0.172 × SGAI + 4.679 × TATA – 0.327 × LVGI
Guidelines:
- When encountering an abbreviation, map it to the correct (Full Account) in {output_retrieval_merged} to retrieve the data.
- When calculating ratios, use the account value for the specific year the user requests (2014–2023) in {output_retrieval_merged}, and use the simplest formula possible. If the exact account is unavailable, use the closest relevant financial account.
- If the user asks for a definition, reply briefly and clearly.
- If there is not enough information, use your general knowledge to answer.
Now, please answer the following question: {query}
"""
    model = ChatGoogleGenerativeAI(
        google_api_key=GOOGLE_API_KEY,
        model=CHAT_MODEL,
        temperature=0
    )
    response_text = model.invoke(prompt)
    return response_text.content

