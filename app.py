from flask import Flask, request, send_file, render_template, url_for
import pandas as pd
import pdfkit
from io import BytesIO
from datetime import datetime, timedelta
import os

app = Flask(__name__)

LOG_FILENAME = 'log.xlsx'

def generate_invoice(invoice_date, due_date, account_id, items, beginTime, endTime, billing_address, address_type):
    total_amount = int(round(sum(item['amount'] for item in items)))

    if address_type == 'Malaysia':
        total_mou = int(round(sum(item['mou'] for item in items)))
        template = 'invoice_template_malaysia.html'
        context = {
            'invoice_date': invoice_date,
            'due_date': due_date,
            'account_id': account_id,
            'beginTime': beginTime,
            'endTime': endTime,
            'items': items,
            'total_mou': total_mou,
            'total_amount': total_amount,
            'billing_address': billing_address,
            'logo_url': url_for('static', filename='logo.jpg', _external=True)
        }
    else:
        template = 'invoice_template_usa.html'
        context = {
            'invoice_date': invoice_date,
            'due_date': due_date,
            'account_id': account_id,
            'beginTime': beginTime,
            'endTime': endTime,
            'items': items,
            'total_amount': total_amount,
            'billing_address': billing_address,
            'logo_url': url_for('static', filename='logo.jpg', _external=True)
        }

    rendered_html = render_template(template, **context)
    pdf = pdfkit.from_string(rendered_html, False)
    return BytesIO(pdf)

def load_existing_logs(log_filename):
    if os.path.exists(log_filename):
        return pd.read_excel(log_filename, sheet_name=None)
    return {'Malaysia': pd.DataFrame(columns=['Account ID', 'Begin Time', 'End Time', 'Date Produced', 'Total Charges', 'Rate']),
            'USA': pd.DataFrame(columns=['Account ID', 'Begin Time', 'End Time', 'Date Produced', 'Total Charges', 'Rate'])}

def save_log_data(log_filename, log_data):
    with pd.ExcelWriter(log_filename) as writer:
        for address_type, data in log_data.items():
            df_log = pd.DataFrame(data)
            df_log.to_excel(writer, sheet_name=address_type, index=False)

def is_duplicate_entry(log_data, account_id, begin_time, end_time, invoice_date, rate_string):
    for entry in log_data:
        if (entry['Account ID'] == account_id and
            entry['Begin Time'] == begin_time.strftime('%Y-%m-%d') and
            # entry['Total Charges'] == rate_string and
            entry['Invoice Date'] == invoice_date.strftime('%Y-%m-%d') and
            entry['End Time'] == end_time.strftime('%Y-%m-%d')):
            return True
    return False

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        invoice_date_str = request.form['invoice_date']
        invoice_date = datetime.strptime(invoice_date_str, '%Y-%m-%d')

        excel1 = request.files['excel1']
        excel2 = request.files['excel2']

        df1 = pd.read_excel(excel1)
        df2 = pd.read_excel(excel2)

        billing_addresses = {}
        missing_account_ids = [] 
        missing_company_names = [] 
        
        for _, row in df1.iterrows():
            account_id = row['account_id']
            if pd.isnull(row['Company Name']):
                missing_company_names.append(account_id)
                continue
            
            billing_addresses[account_id] = {
                'name': row['Company Name'],
                'address': row['Company Address'],
                'code': row['invoice account'],
                'BU': row['BU']
            }

        invoices = []
        existing_logs = load_existing_logs(LOG_FILENAME)
        log_data = {'Malaysia': existing_logs.get('Malaysia').to_dict('records'), 'USA': existing_logs.get('USA').to_dict('records')}
        
        df2['Begin time'] = pd.to_datetime(df2['Begin time'])
        df2['End time'] = pd.to_datetime(df2['End time'])

        grouped = df2.groupby(['Account id', 'Begin time', 'End time'])
        for (account_id, beginTime, endTime), group in grouped:
            if account_id not in billing_addresses:
                missing_account_ids.append(account_id)
                continue
            
            items = []
            billing_address = billing_addresses[account_id]
            address_type = billing_address['BU']  # Get address type from the customer file

            date_diff = (endTime - beginTime).days
            due_date = invoice_date + timedelta(days=15 if date_diff > 10 else 7)
            
            for _, row in group.iterrows():
                if address_type == 'Malaysia':
                    mou = row['Total duration'] / 60
                    mou = round(mou, 2)
                    rate_string = row['Call charges']
                    if isinstance(rate_string, str):
                        rate_string = rate_string.replace(",", "")
                    rate = float(rate_string) / mou
                    rate = format(rate, '.4f')
                
                    item = {
                        'area_name': row['Area name'],
                        'mou': mou,
                        'rate': rate,
                        'quality': row.get('quality', " "),
                        'amount': round(float(rate_string), 2)
                    }
                if address_type == 'USA':
                    rate_string = row['Call charges']
                    if isinstance(rate_string, str):
                        rate_string = rate_string.replace(",", "")
                    item = {
                        'area_name': row['Area name'],
                        'amount': round(float(rate_string), 2)
                    }
                items.append(item)
            invoice_buffer = generate_invoice(invoice_date, due_date, account_id, items, beginTime, endTime, billing_address, address_type)
            invoices.append((f"{account_id}-{beginTime.strftime('%d-%b-%y')}_to_{endTime.strftime('%d-%b-%y')}.pdf", invoice_buffer))

            # Log data for each group
            if not is_duplicate_entry(log_data[address_type], account_id, beginTime, endTime, invoice_date, rate_string):
                log_entry = {
                    'Account ID': account_id,
                    'Begin Time': beginTime.strftime('%Y-%m-%d'),
                    'End Time': endTime.strftime('%Y-%m-%d'),
                    'Invoice Date': invoice_date.strftime('%Y-%m-%d'),
                    'Date Produced': datetime.now().strftime('%Y-%m-%d'),
                    'Total Charges': rate_string,
                    'MOU': mou if address_type == 'Malaysia' else 'N/A'
                }
                log_data[address_type].append(log_entry)
        
        # Save log data to Excel file
        save_log_data(LOG_FILENAME, log_data)

        response_html = "<h1>Invoices Generated:</h1><ul>"
        for filename, buffer in invoices:
            with open(os.path.join('invoices', filename), 'wb') as f:
                f.write(buffer.getbuffer())
            response_html += f'<li><a href="/download/{filename}" target="_blank">{filename}</a></li>'
        response_html += "</ul>"

        if missing_account_ids:
            response_html += "<h2>Missing Account IDs:</h2><ul>"
            for account_id in missing_account_ids:
                response_html += f"<li>{account_id}</li>"
            response_html += "</ul>"

        response_html += f'<h2>Log File:</h2><ul><li><a href="/download/{LOG_FILENAME}" target="_blank">{LOG_FILENAME}</a></li></ul>'

        return response_html

    return '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Billing HGS</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f0f0f0;
            margin: 0;
            padding: 0;
        }
        .container {
            max-width: 600px;
            margin: 50px auto;
            background-color: #fff;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
        }
        h1 {
            text-align: center;
            margin-bottom: 30px;
            color: #007bff;
        }
        label {
            font-weight: bold;
            display: block;
            margin-bottom: 10px;
            color: #555;
        }
        input[type="date"],
        input[type="file"],
        select {
            width: 100%;
            padding: 10px;
            margin-bottom: 20px;
            border: 1px solid #ccc;
            border-radius: 5px;
            box-sizing: border-box;
        }
        input[type="file"] {
            background-color: #f8f9fa;
        }
        input[type="submit"] {
            background-color: #007bff;
            color: #fff;
            border: none;
            padding: 12px 20px;
            border-radius: 5px;
            cursor: pointer;
            width: 100%;
            font-size: 16px;
        }
        input[type="submit"]:hover {
            background-color: #0056b3;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1><img src="static/logo.jpg" alt="Logo" style="width: 50px; vertical-align: middle;"><i> BILLING HUBGLOBE</i></h1>
        <form method="post" enctype="multipart/form-data">
            <label for="invoice_date">Enter the Invoice date:</label>
            <input type="date" id="invoice_date" name="invoice_date" required>
            
            <label for="excel1">Upload Customer Detail Excel:</label>
            <input type="file" id="excel1" name="excel1" required>
            
            <label for="excel2">Upload latest Traffic Excel:</label>
            <input type="file" id="excel2" name="excel2" required>

            <input type="submit" value="Submit">
        </form>
    </div>
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
