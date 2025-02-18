from flask import Flask, request, send_file, render_template_string
import PyPDF2
import pandas as pd
import re
from datetime import datetime
import io
import logging
import traceback

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)


def extract_transaction_details(text_content):
    """Extract transaction details from M-PESA statement."""
    transactions = []
    current_transaction = {}

    # Split into lines
    lines = text_content.split('\n')
    logger.info(f"Number of lines: {len(lines)}")

    i = 0
    while i < len(lines):
        try:
            line = lines[i].strip()

            # Updated pattern to match transactions like "SAO4YDEXQY 2024-01-24"
            if re.match(r'^[A-Z0-9]{9,10}\s+\d{4}-\d{2}-\d{2}', line):
                # Save previous transaction if exists
                if current_transaction:
                    transactions.append(current_transaction)

                # Start new transaction
                parts = line.split()
                current_transaction = {
                    'Reference': parts[0],
                    'Date': parts[1],
                    'Time': None,  # Initialize Time as None
                    'Amount': None,  # Initialize Amount as None
                    'Name': None,  # Initialize Name as None
                    'Account': None  # Initialize Account as None
                }

                # Get time from next line if it exists
                time_line = line
                if re.search(r'\d{2}:\d{2}:\d{2}', time_line):
                    time_match = re.search(r'(\d{2}:\d{2}:\d{2})', time_line)
                    if time_match:
                        current_transaction['Time'] = time_match.group(1)

                # Collect transaction details from subsequent lines until next transaction
                details_lines = []
                j = i + 1
                while j < len(lines):
                    next_line = lines[j].strip()
                    if re.match(r'^[A-Z0-9]{9,10}\s+\d{4}-\d{2}-\d{2}', next_line):
                        break
                    if next_line:
                        details_lines.append(next_line)
                    j += 1

                # Join collected lines
                details_text = ' '.join(details_lines)

                # Extract amount
                amount_match = re.search(r'Completed\s+([\d,]+\.\d{2})', details_text)
                if amount_match:
                    amount_str = amount_match.group(1).replace(',', '')
                    current_transaction['Amount'] = float(amount_str)

                # Extract name and account
                # Try various patterns that appear in your PDF
                name_patterns = [
                    r'from\s+254\*+\d+\s*-\s*([^-]+?)\s+Acc\.',  # Matches "from 254***** - NAME Acc."
                    r'-\s*([^-]+?)\s+Acc\.',  # Matches "- NAME Acc."
                ]

                for pattern in name_patterns:
                    name_match = re.search(pattern, details_text)
                    if name_match:
                        current_transaction['Name'] = name_match.group(1).strip()
                        break

                # Extract account number
                acc_match = re.search(r'Acc\.\s*([^C]+?)(?=\s*Completed|$)', details_text)
                if acc_match:
                    current_transaction['Account'] = acc_match.group(1).strip()

            i += 1

        except Exception as e:
            logger.error(f"Error processing line {i}: {str(e)}")
            logger.error(traceback.format_exc())
            i += 1
            continue

    # Add the last transaction
    if current_transaction:
        transactions.append(current_transaction)

    logger.info(f"Extracted {len(transactions)} transactions")
    return transactions


def process_pdf(pdf_file):
    """Process PDF and extract transactions."""
    all_transactions = []

    try:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        logger.info(f"PDF loaded successfully. Number of pages: {len(pdf_reader.pages)}")

        for page_num in range(len(pdf_reader.pages)):
            try:
                page = pdf_reader.pages[page_num]
                text = page.extract_text()
                logger.info(f"\nProcessing page {page_num + 1}")
                logger.info(f"Page text sample: {text[:200]}")

                transactions = extract_transaction_details(text)
                all_transactions.extend(transactions)

                logger.info(f"Transactions found on page {page_num + 1}: {len(transactions)}")

            except Exception as e:
                logger.error(f"Error processing page {page_num + 1}: {str(e)}")
                logger.error(traceback.format_exc())
                continue

    except Exception as e:
        logger.error(f"Error reading PDF: {str(e)}")
        raise

    logger.info(f"Total transactions found across all pages: {len(all_transactions)}")
    return all_transactions


def create_monthly_summary(transactions):
    """Create monthly summary DataFrame from transactions."""
    if not transactions:
        return pd.DataFrame()

    # Convert transactions to DataFrame
    df = pd.DataFrame(transactions)

    # Ensure Date is datetime
    df['Date'] = pd.to_datetime(df['Date'])

    # Sort by date
    df = df.sort_values('Date')

    # Extract month and year
    df['Month'] = df['Date'].dt.strftime('%B %Y')

    # Create a temporary month key for sorting
    df['Month_Sort'] = df['Date'].dt.strftime('%Y-%m')

    # Create pivot table
    pivot_df = df.pivot_table(
        index=['Name', 'Account'],
        columns='Month',
        values='Amount',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    # Calculate total for sorting
    pivot_df['Total'] = pivot_df.select_dtypes(include=['float64']).sum(axis=1)

    # Sort first by Name alphabetically, then by Total amount descending
    pivot_df = pivot_df.sort_values(['Name', 'Total'], ascending=[True, False])

    # Sort the columns (keeping Name, Account, and Total in place)
    static_cols = ['Name', 'Account']
    month_cols = [col for col in pivot_df.columns if col not in static_cols + ['Total']]

    # Convert month names to datetime for sorting
    month_dates = pd.to_datetime([datetime.strptime(m, '%B %Y') for m in month_cols])
    sorted_month_cols = [d.strftime('%B %Y') for d in sorted(month_dates)]

    # Reorder columns with sorted months
    pivot_df = pivot_df[static_cols + sorted_month_cols + ['Total']]

    return pivot_df

@app.route('/')
def index():
    """Render the upload form."""
    return '''
    <html>
        <head>
            <title>M-PESA Statement Processor</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    max-width: 800px;
                    margin: 40px auto;
                    padding: 20px;
                    background-color: #f8f9fa;
                }
                .container {
                    background-color: white;
                    padding: 30px;
                    border-radius: 10px;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                }
                h1 {
                    color: #28a745;
                    margin-bottom: 20px;
                }
                .description {
                    color: #6c757d;
                    margin-bottom: 25px;
                    line-height: 1.5;
                }
                form {
                    margin: 20px 0;
                }
                .upload-btn {
                    background-color: #28a745;
                    color: white;
                    padding: 12px 24px;
                    border: none;
                    border-radius: 5px;
                    cursor: pointer;
                    font-size: 16px;
                    transition: background-color 0.3s ease;
                }
                .upload-btn:hover {
                    background-color: #218838;
                }
                .file-input {
                    margin-bottom: 20px;
                }
                .file-input input {
                    padding: 10px;
                    border: 1px solid #ced4da;
                    border-radius: 5px;
                    width: 100%;
                    max-width: 400px;
                }
                .features {
                    margin-top: 30px;
                    padding-top: 20px;
                    border-top: 1px solid #dee2e6;
                }
                .features h2 {
                    color: #495057;
                    font-size: 1.2em;
                    margin-bottom: 15px;
                }
                .features ul {
                    color: #6c757d;
                    padding-left: 20px;
                }
                .features li {
                    margin-bottom: 8px;
                }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>M-PESA Statement Processor</h1>
                <p class="description">
                    Upload your M-PESA statement PDF to automatically convert it into an organized Excel spreadsheet.
                    The processor will create both a monthly summary and detailed transaction list.
                </p>
                <form action="/process" method="post" enctype="multipart/form-data">
                    <div class="file-input">
                        <input type="file" name="pdf_file" accept=".pdf" required>
                    </div>
                    <button type="submit" class="upload-btn">Process Statement</button>
                </form>
                <div class="features">
                    <h2>Features:</h2>
                    <ul>
                        <li>Monthly transaction summaries by account</li>
                        <li>Detailed transaction list with dates and times</li>
                        <li>Automatic sorting and organization</li>
                        <li>Professional Excel formatting</li>
                        <li>Totals calculation and analysis</li>
                    </ul>
                </div>
            </div>
        </body>
    </html>
    '''


@app.route('/process', methods=['POST'])
def process():
    """Process the uploaded PDF file and return Excel file."""
    if 'pdf_file' not in request.files:
        return 'No file uploaded', 400

    pdf_file = request.files['pdf_file']
    if pdf_file.filename == '':
        return 'No file selected', 400

    try:
        transactions = process_pdf(pdf_file)

        if not transactions:
            return 'No transactions found. Please check the PDF format.', 400

        # Create monthly summary
        monthly_summary = create_monthly_summary(transactions)

        # Create detailed transactions DataFrame
        detailed_df = pd.DataFrame(transactions)

        # Ensure all required columns exist with default values
        for col in ['Date', 'Reference', 'Name', 'Account', 'Amount', 'Time']:
            if col not in detailed_df.columns:
                detailed_df[col] = None

        # Convert Date to datetime
        detailed_df['Date'] = pd.to_datetime(detailed_df['Date'])

        # Sort by date
        detailed_df = detailed_df.sort_values('Date')

        # Reorder columns
        detailed_df = detailed_df[['Date', 'Reference', 'Name', 'Account', 'Amount', 'Time']]

        # Create Excel file in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Write monthly summary
            monthly_summary.to_excel(writer, sheet_name='Monthly Summary', index=False)

            # Write detailed transactions
            detailed_df.to_excel(writer, sheet_name='Detailed Transactions', index=False)

            workbook = writer.book

            # Format Monthly Summary sheet
            summary_worksheet = writer.sheets['Monthly Summary']

            # Add formats
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#D3D3D3',
                'border': 1,
                'text_wrap': True,
                'align': 'center',
                'valign': 'vcenter'
            })
            money_format = workbook.add_format({
                'num_format': '#,##0.00',
                'border': 1,
                'align': 'right'
            })
            text_format = workbook.add_format({
                'border': 1,
                'text_wrap': True
            })
            total_format = workbook.add_format({
                'bold': True,
                'num_format': '#,##0.00',
                'bg_color': '#F0F0F0',
                'border': 1,
                'align': 'right'
            })

            # Apply formats to Monthly Summary
            for col_num, value in enumerate(monthly_summary.columns.values):
                summary_worksheet.write(0, col_num, value, header_format)

                # Set column formats and widths
                if col_num < 2:  # Name and Account columns
                    summary_worksheet.set_column(col_num, col_num, 30, text_format)
                else:  # Amount columns
                    summary_worksheet.set_column(col_num, col_num, 15, money_format)

            # Add totals row
            total_row = len(monthly_summary) + 1
            summary_worksheet.write(total_row, 0, 'TOTAL', total_format)
            summary_worksheet.write(total_row, 1, '', total_format)

            for col_num in range(2, len(monthly_summary.columns)):
                col_letter = chr(65 + col_num)
                formula = f'=SUM({col_letter}2:{col_letter}{total_row})'
                summary_worksheet.write_formula(total_row, col_num, formula, total_format)

            # Format Detailed Transactions sheet
            detailed_worksheet = writer.sheets['Detailed Transactions']

            # Set formats for detailed transactions
            date_format = workbook.add_format({
                'num_format': 'yyyy-mm-dd',
                'border': 1,
                'align': 'center'
            })
            time_format = workbook.add_format({
                'num_format': 'hh:mm:ss',
                'border': 1,
                'align': 'center'
            })

            # Set column widths and formats
            detailed_worksheet.set_column('A:A', 12, date_format)  # Date
            detailed_worksheet.set_column('B:B', 15, text_format)  # Reference
            detailed_worksheet.set_column('C:C', 30, text_format)  # Name
            detailed_worksheet.set_column('D:D', 15, text_format)  # Account
            detailed_worksheet.set_column('E:E', 15, money_format)  # Amount
            detailed_worksheet.set_column('F:F', 10, time_format)  # Time

            # Apply header format
            for col_num, value in enumerate(detailed_df.columns.values):
                detailed_worksheet.write(0, col_num, value, header_format)

        output.seek(0)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='mpesa_transactions.xlsx'
        )

    except Exception as e:
        logger.error(f"Error in process route: {str(e)}")
        logger.error(traceback.format_exc())
        return f'Error processing file: {str(e)}', 500


if __name__ == '__main__':
    app.run(debug=True)