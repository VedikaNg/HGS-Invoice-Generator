from flask import Flask, request, send_file, render_template, render_template_string, make_response, url_for
import pandas as pd
import pdfkit
from io import BytesIO
from datetime import datetime, timedelta
import os


app = Flask(__name__)

def generate_invoice(invoice_date, due_date, account_id, items, beginTime, endTime, billing_address):
    total_mou = round(sum(item['mou'] for item in items),2)
    total_amount = round(sum(item['amount'] for item in items),2)
    
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
    return BytesIO(pdf)

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
                mou = round(mou, 2)
                rate = row['Total charges'] / mou 
                rate = format(rate, '.4f')
                beginTime = datetime.strptime(row['Begin time'], '%Y-%m-%d')
                endTime = datetime.strptime(row['End time'], '%Y-%m-%d')
                items.append({
                    'area_name': row['Area name'],
                    'mou': mou,
                    'rate': rate,
                    'quality': row.get('quality', " "),
                    'amount': round(row['Total charges'],2)
                })
            invoice_buffer = generate_invoice(invoice_date, due_date, account_id, items, beginTime, endTime, billing_address)
            invoices.append((f"{account_id}-{beginTime.strftime('%d-%b-%y')}_to_{endTime.strftime('%d-%b-%y')}.pdf", invoice_buffer))
        
        response_html = "<h1>Invoices Generated:</h1><ul>"
        for filename, buffer in invoices:
            with open(os.path.join('invoices', filename), 'wb') as f:
                f.write(buffer.getbuffer())
            response_html += f'<li><a href="/download/{filename}" target="_blank">{filename}</a></li>'
        response_html += "</ul>"

        return response_html

    return '''
    <!doctype html>
    <title>Upload Files</title>
    <center><h1>HUBGLOBE BILLING</center> </h1><br><br>
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
