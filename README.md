# MCA Data Extraction System

A specialized Python tool designed to automate the extraction of Merchant Cash Advance (MCA) data from bank statements. This system identifies potential loan payments and aggregates client information into a standardized 50-column Excel report.

---

## 🛠️ Installation & Setup (Windows)

To run this tool on Windows, follow these simple steps:

1. **Install Python**:
   If you don't have Python installed, download and install the latest version from [python.org](https://www.python.org/downloads/windows/).  
   **⚠️ IMPORTANT**: During installation, make sure to check the box **"Add Python to PATH"**.

2. **Download the Project**:
   Download the source code and extract it to a folder of your choice (e.g., `C:\MCA-Extractor`).

3. **Open Terminal**:
   Open **PowerShell** or **Command Prompt** (Cmd) and navigate to your project folder:

   ```cmd
   cd C:\MCA-Extractor
   ```

4. **Create a Virtual Environment** (Recommended):
   This keeps the project dependencies isolated from your system:

   ```cmd
   python -m venv .venv
   ```

5. **Activate the Virtual Environment**:

   ```cmd
   .venv\Scripts\activate
   ```

   _(You should see `(.venv)` appear at the beginning of your command line)_.

6. **Install Required Libraries**:
   ```cmd
   pip install -r requirements.txt
   ```

---

## 📖 How to Run on Windows

1. Ensure your virtual environment is active.
2. Start the application:
   ```cmd
   python app.py
   ```
3. A menu will appear asking for the **Processing Mode**:
   - **Batch Mode**: Choose this if you have multiple PDFs to process. You will enter the number of clients and select a file for each.
   - **Individual Mode**: Choose this to process just one PDF.
4. Select your bank statement PDF using the file explorer window.

---

## 📄 Results & Output

- The report is saved as **`MCA_Final_Report.xlsx`** in the project folder.
- You can open this file directly in **Microsoft Excel**.
- **⚠️ NOTE**: Close the Excel file before running the script, otherwise, the program will crash when trying to save new data.
- The system also creates a **`MCA_processed_files.json`** file to keep track of which PDFs have already been analyzed to prevent duplicate entries.

---

## 🚀 Key Features

- **Standardized Output**: 50-column Excel report compatible with US financial standards.
- **Smart Detection**: Identifies potential MCA payments (LOAN, FUNDING, ADVANCE, etc.).
- **Automatic Deduplication**: Warns you if an attempt is made to process the same PDF twice.
- **Cumulative Mode**: New data is added to the same spreadsheet without overwriting old records.
- **English UI with Portuguese Comments**: Built for American clients, with code comments in Portuguese for developer study.
