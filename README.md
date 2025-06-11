# FraudDetect Pro: Financial Statement Fraud Detection Dashboard

## Overview
FraudDetect Pro is a web-based dashboard for detecting potential financial statement fraud using advanced statistical and AI techniques. It combines the Beneish M-Score, Benford's Law, and Retrieval-Augmented Generation (RAG) AI Q&A to provide a comprehensive analysis of uploaded financial Excel reports.

## Features
- **Excel Upload & Parsing:** Upload annual financial statements in Excel format for automated analysis.
- **Beneish M-Score:** Detects potential earnings manipulation using the Beneish M-Score model.
- **Benford's Law Analysis:** Checks for anomalies in leading digit distributions, flagging possible data manipulation.
- **Interactive Dashboard:** Visualizes M-Score trends, Benford deviations, and key variable changes.
- **AI Q&A (RAG):** Ask questions about your data in Vietnamese or English, powered by Google Gemini AI and LangChain.
- **Recommendations & Summaries:** Automated insights and recommendations based on detected risks.

## Installation
1. **Clone the repository:**
   ```bash
   git clone <your-repo-url>
   cd <your-repo-directory>
   ```
2. **Install dependencies:**
   It is recommended to use a virtual environment.
   ```bash
   pip install -r requirements.txt
   ```

## Usage
1. **Start the Flask server:**
   ```bash
   python App.py
   ```
2. **Open your browser and go to:**
   [http://localhost:5000](http://localhost:5000)

3. **Upload your Excel file:**
   - The file should contain annual financial statements (multiple sheets supported).
   - The system will automatically detect headers, account names, and bold formatting for hierarchy.

4. **Explore the Dashboard:**
   - **Summary Cards:** See overall fraud risk, latest M-Score, and Benford deviation.
   - **M-Score Analysis:** View component breakdown and top variable changes.
   - **Benford Analysis:** Visualize leading digit distribution and deviation level.
   - **Trends:** Track M-Score over time.
   - **AI Q&A:** Ask questions about your data (supports Vietnamese and English).

## File Structure
- `App.py` — Main Flask app, routes, and dashboard logic.
- `AI_model.py` — Handles RAG Q&A, vector database, and Google Gemini AI integration.
- `account_mapping_utils.py` — Account name normalization and mapping utilities.
- `templates/` — HTML templates for upload and dashboard pages.
- `Data/` — Contains mapping files and (optionally) sample data.
- `requirements.txt` — Python dependencies.
- `test.py` — (Optional) Test and development scripts.

## Dependencies
- Flask, Flask-Session
- pandas, numpy, matplotlib, openpyxl
- plotly
- langchain-community, langchain-google-genai, langchain-openai
- faiss-cpu
- unstructured
- rapidfuzz
- markdown, networkx, langdetect

Install all dependencies with:
```bash
pip install -r requirements.txt
```

## AI Q&A (RAG)
- Powered by Google Gemini (via LangChain)
- Supports both Vietnamese and English queries
- Uses Retrieval-Augmented Generation: answers are based on your uploaded data

## Security Note
- **API Keys:** The Google API key is currently hardcoded for demo purposes. For production, use environment variables or a secure vault.
- **Data Privacy:** Uploaded files are processed in-memory and temporarily stored for analysis only.

## Contact
For questions, issues, or contributions, please open an issue or contact the maintainer. 