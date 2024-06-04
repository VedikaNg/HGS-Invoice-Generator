from flask import Flask, request, send_file, render_template, url_for
import pandas as pd
import pdfkit
from io import BytesIO
from datetime import datetime, timedelta
import os
import xlsxwriter

app = Flask(__name__)

def generate_invoice(invoice_date, due_date, account_id, items, beginTime, endTime, billing_address):
    total_mou = round(sum(float(item['mou']) for item in items), 2)
    total_amount = round(sum(float(item['amount']) for item in items), 2)
    
    # Generate PDF
    rendered_html = render_template('invoice_template.html', 
                                    invoice_date=invoice_date, 
                                    due_date=due_date, 
                                    account_id=account_id, 
                                    beginTime=beginTime,
                                    endTime=endTime,
                                    items=items,
                                    total_mou=total_mou,
                                    total_amount=total_amount,
                                    billing_address=billing_address,
                                    logo_url=url_for('static', filename='logo.jpg', _external=True))
    pdf = pdfkit.from_string(rendered_html, False)
    pdf_buffer = BytesIO(pdf)
    
    # Generate Excel using xlsxwriter
    excel_buffer = BytesIO()
    workbook = xlsxwriter.Workbook(excel_buffer, {'in_memory': True})
    worksheet = workbook.add_worksheet()

    # Set up formats
    title_format = workbook.add_format({'bold': True, 'font_size': 14})
    header_format = workbook.add_format({'bold': True, 'bg_color': '#F9DA04', 'border': 1})
    cell_format = workbook.add_format({'border': 1})
    currency_format = workbook.add_format({'num_format': '$#,##0.00', 'border': 1})
    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1})

    # Write title
    worksheet.merge_range('A1:F1', 'Invoice', title_format)

    # Write billing details
    worksheet.write('A3', 'Invoice Date:', cell_format)
    worksheet.write('B3', invoice_date.strftime('%Y-%m-%d'), date_format)
    worksheet.write('A4', 'Due Date:', cell_format)
    worksheet.write('B4', due_date.strftime('%Y-%m-%d'), date_format)
    worksheet.write('A5', 'Account ID:', cell_format)
    worksheet.write('B5', account_id, cell_format)
    worksheet.write('A6', 'Billing Address:', cell_format)
    worksheet.write('B6', f"{billing_address['name']}, {billing_address['address']}", cell_format)

    # Write headers
    headers = ['Area Name', 'Minutes of Use (MOU)', 'Rate per MOU', 'Quality', 'Amount']
    for col_num, header in enumerate(headers):
        worksheet.write(8, col_num, header, header_format)

    # Write item rows
    row_num = 9
    for item in items:
        worksheet.write(row_num, 0, item['area_name'], cell_format)
        worksheet.write(row_num, 1, float(item['mou']), cell_format)
        worksheet.write(row_num, 2, float(item['rate']), cell_format)
        worksheet.write(row_num, 3, item['quality'], cell_format)
        worksheet.write(row_num, 4, float(item['amount']), currency_format)
        row_num += 1

    # Write totals
    worksheet.write(row_num, 3, 'Total MOU:', header_format)
    worksheet.write(row_num, 4, total_mou, cell_format)
    row_num += 1
    worksheet.write(row_num, 3, 'Total Amount:', header_format)
    worksheet.write(row_num, 4, total_amount, currency_format)

    workbook.close()
    excel_buffer.seek(0)
    
    return pdf_buffer, excel_buffer
@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        invoice_date_str = request.form['invoice_date']
        invoice_date = datetime.strptime(invoice_date_str, '%Y-%m-%d')
        due_date = invoice_date + timedelta(days=7)
        
        excel1 = request.files['excel1']
        excel2 = request.files['excel2']

        df1 = pd.read_excel(excel1)
        df2 = pd.read_excel(excel2)

        # Create a dictionary for billing addresses
        billing_addresses = {}
        for _, row in df1.iterrows():
            account_id = row['account_id']
            billing_addresses[account_id] = {
                'name': row['Company Name'],
                'address': row['Company Address'],
                'code': row['invoice account']
            }

        invoices = []
        for account_id, group in df2.groupby('Account id'):
            items = []
            billing_address = billing_addresses.get(account_id, {'name': {account_id}, 'address': ' ', 'code' : {account_id}})
            for _, row in group.iterrows():
                mou = row['Total duration'] / 60
                mou = format(mou, '.2f')
                rate = row['Total charges'] / float(mou)
                rate = format(rate, '.4f')
                beginTime = datetime.strptime(row['Begin time'], '%Y-%m-%d')
                endTime = datetime.strptime(row['End time'], '%Y-%m-%d')
                items.append({
                    'area_name': row['Area name'],
                    'mou': mou,
                    'rate': rate,
                    'quality': row.get('quality', " "),
                    'amount': format(row['Total charges'], '.2f')
                })
            pdf_buffer, excel_buffer = generate_invoice(invoice_date, due_date, account_id, items, beginTime, endTime, billing_address)
            pdf_filename = f"{account_id}-{beginTime.strftime('%d-%b-%y')}_to_{endTime.strftime('%d-%b-%y')}.pdf"
            excel_filename = f"{account_id}-{beginTime.strftime('%d-%b-%y')}_to_{endTime.strftime('%d-%b-%y')}.xlsx"
            invoices.append((pdf_filename, pdf_buffer, excel_filename, excel_buffer))
        
        response_html = "<h1>Invoices Generated:</h1><ul>"
        for pdf_filename, pdf_buffer, excel_filename, excel_buffer in invoices:
            pdf_path = os.path.join('invoices', pdf_filename)
            with open(pdf_path, 'wb') as f:
                f.write(pdf_buffer.getbuffer())
            excel_path = os.path.join('invoices', excel_filename)
            with open(excel_path, 'wb') as f:
                f.write(excel_buffer.getbuffer())
            response_html += f'<li>PDF: <a href="/download/{pdf_filename}" target="_blank">{pdf_filename}</a> | Excel: <a href="/download/{excel_filename}" target="_blank">{excel_filename}</a></li>'
        response_html += "</ul>"

        return response_html

    return '''
    <!doctype html>
    <title>Upload Files</title>
    <center><h1>HGS INVOICE GENERATOR</center> </h1><br><br>
    <form method=post enctype=multipart/form-data>
    <table align="center">
    <tr><td style="padding: 10px; text-align: left; font-size: 20px;"><label><strong>Enter the Invoice date: </strong></label></td>
      <td style="padding: 10px; text-align: left"><input type="date" name="invoice_date" required></td></tr>
      <tr><td style="padding: 10px; font-size: 20px;"><label><strong>Upload Customer Detail Excel: </strong></label></td>
      <td style="padding: 10px; text-align: left"><input type="file" name="excel1" required></td></tr>
      <tr><td style="padding: 10px; font-size: 20px;"><label><strong>Upload latest Traffic Excel: </strong></label></td>
      <td style="padding: 10px; "><input type="file" name="excel2" required></td></tr>
      <tr><td style="padding: 10px;"><input type="submit" value="Upload"></td></tr>
      </table>
    </form>
    '''

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    file_path = os.path.join('invoices', filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True, download_name=filename)
    else:
        return "File not found", 404

if __name__ == '__main__':
    if not os.path.exists('invoices'):
        os.makedirs('invoices')
    app.run(debug=True)
