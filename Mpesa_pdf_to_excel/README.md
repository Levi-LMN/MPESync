# M-Pesa PDF to Excel Converter

## Overview
This is a Flask-based web application that extracts transaction data from M-Pesa PDF statements and converts it into an Excel file. The generated Excel file includes:

- **Raw Transactions**: Extracted data in tabular format.
- **Daily Totals**: Summarized daily totals of deposits, withdrawals, and net amounts.

## Features
- Upload M-Pesa PDF statements.
- Extract financial transactions from the PDF.
- Convert data into structured Excel format.
- Generate daily total summaries for better financial analysis.
- Automatically format the Excel file for readability.

## Requirements
Ensure you have Python installed, then install the required dependencies:

```sh
pip install flask pandas pdfplumber openpyxl
```

## Installation
1. Clone this repository:
   ```sh
   git clone https://github.com/yourusername/mpesa-pdf-to-excel.git
   ```
2. Navigate to the project folder:
   ```sh
   cd mpesa-pdf-to-excel
   ```
3. Install dependencies:
   ```sh
   pip install -r requirements.txt
   ```
4. Run the Flask app:
   ```sh
   python app.py
   ```
5. Open a web browser and go to:
   ```
   http://127.0.0.1:5000/
   ```

## Usage
1. Click **Choose File** and upload an M-Pesa PDF statement.
2. Click **Convert to Excel**.
3. Download the generated Excel file.
4. View transactions and daily totals in separate sheets.

## Folder Structure
```
mpesa-pdf-to-excel/
├── app.py          # Main Flask application
├── uploads/        # Directory for storing uploaded and processed files
├── requirements.txt # Dependencies
├── README.md       # Project documentation
```

## Notes
- Ensure the M-Pesa statement is in a structured tabular format.
- The app extracts columns such as `Completion Time`, `Paid In`, `Withdrawn`, and `Balance`.
- Withdrawn amounts are converted to negative values for correct calculations.

## License
This project is open-source under the MIT License.

