from flask import Flask, request, send_file, render_template, url_for
import pandas as pd
import pdfkit
from io import BytesIO
from datetime import datetime, timedelta
import os

app = Flask(__name__)

def generate_invoice(invoice_date, due_date, account_id, items, beginTime, endTime, billing_address):
    total_mou = round(sum(item['mou'] for item in items), 2)
    total_amount = round(sum(item['amount'] for item in items), 2)
    
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
        df2['Begin time'] = pd.to_datetime(df2['Begin time'])
        df2['End time'] = pd.to_datetime(df2['End time'])

        grouped = df2.groupby(['Account id', 'Begin time', 'End time'])
        for (account_id, beginTime, endTime), group in grouped:
            items = []
            billing_address = billing_addresses.get(account_id, {'name': {account_id}, 'address': ' ', 'code': {account_id}})
            for _, row in group.iterrows():
                mou = row['Total duration'] / 60
                mou = round(mou, 2)
                rate = row['Total charges'] / mou 
                rate = format(rate, '.4f')
                items.append({
                    'area_name': row['Area name'],
                    'mou': mou,
                    'rate': rate,
                    'quality': row.get('quality', " "),
                    'amount': round(row['Total charges'], 2)
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
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Billing HGS</title>
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap');

            body {
                font-family: 'Roboto', sans-serif;
                background-color: #f0f2f5;
                color: #333;
                margin: 0;
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
            }
            .container {
                background: white;
                padding: 30px;
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
                border-radius: 10px;
                width: 100%;
                max-width: 600px;
                text-align: center;
            }
            .header {
                display: flex;
                align-items: center;
                justify-content: center;
                margin-bottom: 20px;
            }
            .header img {
                width: 50px;
                margin-right: 15px;
            }
            .header h1 {
                font-size: 28px;
                margin: 0;
                color: #007BFF;
            }
            form {
                margin-top: 20px;
            }
            table {
                margin: 0 auto;
                font-size: 16px;
                width: 100%;
            }
            td {
                padding: 10px;
                text-align: left;
            }
            input[type="submit"] {
                background-color: #007BFF;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 16px;
                cursor: pointer;
                border-radius: 5px;
                transition: background-color 0.3s ease;
            }
            input[type="submit"]:hover {
                background-color: #0056b3;
            }
            input[type="date"], input[type="file"] {
                padding: 10px;
                font-size: 16px;
                border: 1px solid #ccc;
                border-radius: 5px;
                width: calc(100% - 22px);
                margin: 5px 0 20px 0;
                background-color: #e6f7ff;
            }
            .input-container {
                margin-bottom: 20px;
                text-align: left;
            }
            .input-container label {
                display: block;
                margin-bottom: 8px;
                font-weight: bold;
                font-size: 16px;
            }
            .custom-file-input {
                position: relative;
                display: inline-block;
                width: calc(100% - 22px);
                margin: 5px 0 20px 0;
            }
            .custom-file-input input[type="file"] {
                position: absolute;
                opacity: 0;
                width: 100%;
                height: 100%;
                cursor: pointer;
            }
            .custom-file-input label {
                background-color: #007BFF;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 16px;
                cursor: pointer;
                border-radius: 5px;
                transition: background-color 0.3s ease;
            }

            .custom-file-input label:hover {
                background-color: #0056b3;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <img src="static/logo.jpg" alt="Logo">
                <h1>BILLING HUBGLOBE</h1>
            </div>
            <form method="post" enctype="multipart/form-data">
                <div class="input-container">
                    <label for="invoice_date"><strong>Enter the Invoice date:</strong></label>
                    <input type="date" name="invoice_date" id="invoice_date" required>
                </div>
                <div class="input-container custom-file-input">
                    <label for="excel1" id="label-excel1"><strong>Upload Customer Detail Excel:</strong></label>
                    <input type="file" name="excel1" id="excel1" required onchange="updateFileName('excel1')">
                </div>
                <div class="input-container custom-file-input">
                    <label for="excel2" id="label-excel2"><strong>Upload latest Traffic Excel:</strong></label>
                    <input type="file" name="excel2" id="excel2" required onchange="updateFileName('excel2')">
                </div>
                <input type="submit" value="Upload">
            </form>
        </div>
        <script>
            function updateFileName(inputId) {
                const input = document.getElementById(inputId);
                const label = document.getElementById('label-' + inputId);
                const fileName = input.files[0].name;
                label.innerHTML = `<strong>${fileName}</strong>`;
            }
        </script>
    </body>
    </html>
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
