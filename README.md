# MPESync

## Overview
MPESync is a powerful and lightweight tool that extracts transactions from M-PESA PDF statements and converts them into structured Excel reports. It helps businesses, freelancers, and financial analysts automate the process of tracking M-PESA transactions, eliminating the need for manual data entry.

## Features
- **Extract Transactions**: Automatically pulls transaction details (Date, Reference, Amount, Name, and Account) from M-PESA PDF statements.
- **Generate Excel Reports**: Converts raw transaction data into well-structured Excel spreadsheets with detailed and summary views.
- **Monthly Summaries**: Organizes transactions by month for easy financial analysis.
- **Professional Formatting**: Creates Excel files with clean formatting, auto-calculated totals, and structured columns.
- **Fast & Secure**: Processes statements quickly while ensuring data privacy.
- **Business Insights**: Provides data-driven insights to help businesses manage cash flow efficiently.
- **Easy Integration**: Can be adapted into existing financial software and workflows.

## How MPESync Helps Businesses

### **1. Automating Financial Record-Keeping**
Businesses dealing with frequent M-PESA transactions often struggle with manual data entry, which is both time-consuming and error-prone. MPESync eliminates the need for manual work by instantly extracting and structuring transaction data into an Excel format, allowing businesses to maintain accurate financial records effortlessly.

### **2. Enhancing Cash Flow Management**
With a clear view of transactions categorized by month, businesses can better monitor cash inflows and outflows. This enables better financial planning, timely bill payments, and improved budgeting.

### **3. Simplifying Tax Preparation**
Many businesses face challenges during tax season when they need to compile financial statements. MPESync helps streamline this process by generating structured reports that can be easily shared with accountants and tax professionals, ensuring compliance with financial regulations.

### **4. Improving Financial Analysis**
The tool provides businesses with actionable insights into their transaction patterns. By identifying spending trends, high-frequency customers, and payment cycles, businesses can make informed financial decisions.

### **5. Reducing Operational Costs**
Manual data entry often requires hiring additional staff or allocating significant time to administrative work. By automating M-PESA transaction extraction, businesses can save costs on labor and redirect resources towards revenue-generating activities.

### **6. Enhancing Customer & Vendor Management**
With clear transaction tracking, businesses can efficiently manage payments received from customers and payments made to vendors. This leads to better accountability and improved financial relationships.

### **7. Generating Reports for Business Loans & Investors**
Accurate financial records are crucial when applying for loans or seeking investment. MPESync provides businesses with professionally formatted transaction summaries that can be presented to banks, investors, or financial institutions to showcase cash flow and financial stability.

## Installation
Ensure you have Python installed, then install the required dependencies:

```bash
pip install flask pandas PyPDF2 xlsxwriter
```

## Usage
1. Run the Flask app:

```bash
python app.py
```

2. Open your browser and go to `http://127.0.0.1:5000/`.
3. Upload your M-PESA PDF statement and download the extracted Excel file.

## API Endpoints
- **`GET /`** – Renders the file upload interface.
- **`POST /process`** – Processes the uploaded M-PESA PDF and returns an Excel file.

## Example Output
The output Excel file contains:
- **Monthly Summary Sheet**: Aggregated transactions by name and account.
- **Detailed Transactions Sheet**: A complete list of transactions sorted by date.

## Contributing
Feel free to fork this repository and submit pull requests with improvements. For major changes, please open an issue first to discuss the updates.

## License
This project is licensed under the MIT License.

## Contact
For questions or suggestions, reach out via GitHub issues.

