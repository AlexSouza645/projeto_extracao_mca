# MCA Data Extraction System

A specialized Python tool designed to automate the extraction of Merchant Cash Advance (MCA) data from bank statements. This system identifies potential loan payments and aggregates client information into a standardized 50-column Excel report.

---

## 🛠️ Installation & Setup (Windows & macOS)

### 1. Install Python
- **Windows:** Download the latest installer from [python.org](https://www.python.org/downloads/windows/).  
  **⚠️ CRITICAL:** During installation, you MUST check the box **"Add Python to PATH"** at the bottom of the first screen.
- **macOS:** Download the macOS installer from [python.org](https://www.python.org/downloads/macos/) or use Homebrew (`brew install python`).

### 2. Download the Project
Download the source code and extract it to a folder of your choice (e.g., `Documents/MCA-Extractor`).

### 3. Open Terminal
Navigate to your project folder using your system's terminal:
- **Windows:** Open **PowerShell** and run `cd C:\Path\To\MCA-Extractor`
- **macOS / Linux:** Open **Terminal** and run `cd ~/Path/To/MCA-Extractor`

### 4. Install Dependencies
Run the following command to install the required libraries permanently to your system python:

```bash
python -m pip install -r requirements.txt
```
*(Note for macOS: You might need to type `python3` instead of `python` and `pip3` instead of `pip` in your terminal).*

---

## 📖 How to Run the Extractor

1. Open your terminal in the project folder.
2. Start the application by running:
   ```bash
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
- **⚠️ NOTE**: **Close the Excel file before running the script**, otherwise, the program will crash when trying to save new data.
- The system also creates a **`MCA_processed_files.json`** file to keep track of which PDFs have already been analyzed to prevent duplicate entries.

---

## 🔍 Troubleshooting PDFs (`diagnostico.py`)

Sometimes, banks provide PDFs that are actually just scanned images, containing no selectable text. The MCA extractor cannot read scanned images without OCR.

To test if a batch of PDFs are readable or if they are just images:
1. Place the PDFs you want to test inside the `arquivos/` folder inside the project.
2. Run the diagnostic tool:
   ```bash
   python diagnostico.py
   ```
3. The tool will check all PDFs inside the `arquivos/` folder and tell you if they contain **Native Text (✓)** or if they are **Scanned (✗)**.

---

## 🚀 Key Features

- **Cross-Platform**: Works flawlessly on Windows, macOS, and Linux.
- **Standardized Output**: 50-column Excel report compatible with US financial standards.
- **Smart Detection**: Identifies potential MCA payments (LOAN, FUNDING, ADVANCE, etc.).
- **Automatic Deduplication**: Warns you if an attempt is made to process the same PDF twice.
- **Cumulative Mode**: New data is added to the same spreadsheet without overwriting old records.
