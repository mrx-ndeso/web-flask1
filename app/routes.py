from flask import render_template, request, redirect, url_for, send_file
from app import app
from app.utils import upload_to_google_sheets, download_from_google_sheets

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        spreadsheet_url = request.form['spreadsheet_url']
        worksheet_name = request.form['worksheet_name']
        return redirect(url_for('download_files', spreadsheet_url=spreadsheet_url, worksheet_name=worksheet_name))
    return render_template('index.html')

@app.route('/upload-xlsx', methods=['POST'])
def upload_xlsx():
    # Implementasi untuk upload XLSX dan kirim ke Google Sheets
    pass

@app.route('/download', methods=['GET'])
def download_files():
    # Implementasi untuk download dari Google Sheets dan simpan ke XLSX
    pass
