from flask import Blueprint, render_template, request, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import io

main = Blueprint('main', __name__)

# Konfigurasi
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

@main.route('/')
def home():
    return render_template('index.html')

@main.route('/about')
def about():
    return render_template('about.html')

@main.route('/application')
def application():
    return render_template('application.html')

@main.route('/dgs', methods=['GET', 'POST'])
def dgs():
    if request.method == 'POST':
        sheet_url = request.form['sheet_url']
        sheet_name = request.form['sheet_name']
        
        try:
            # Ambil spreadsheet dan worksheet
            spreadsheet = client.open_by_url(sheet_url)
            worksheet = spreadsheet.worksheet(sheet_name)
            
            # Ambil data dari worksheet
            data = worksheet.get_all_records()
            df = pd.DataFrame(data)
            
            # Konversi semua kolom menjadi string
            df = df.astype(str)
            
            # Simpan DataFrame sebagai XLSX
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            output.seek(0)
            
            return send_file(output, as_attachment=True, download_name='download.xlsx')
        except Exception as e:
            return f"An error occurred: {e}"

    return render_template('dgs.html')

@main.route('/upload_and_send', methods=['GET', 'POST'])
def upload_and_send():
    if request.method == 'POST':
        sheet_url = request.form['sheet_url']
        sheet_name = request.form['sheet_name']
        file = request.files['file']

        try:
            # Baca file XLSX
            df = pd.read_excel(file)
            
            # Konversi semua kolom menjadi string
            df = df.astype(str)
            
            # Debug: Tampilkan data yang akan dikirim
            print("Data from XLSX file:")
            print(df.head())

            # Ambil spreadsheet dan worksheet
            spreadsheet = client.open_by_url(sheet_url)
            worksheet = spreadsheet.worksheet(sheet_name)
            
            # Kosongkan worksheet sebelum mengisi data baru
            worksheet.clear()
            
            # Konversi DataFrame ke format list of lists
            data = df.values.tolist()

            # Debug: Tampilkan data yang akan diupdate ke Google Sheets
            print("Data to be uploaded to Google Sheets:")
            print(data)

            # Update worksheet dengan data baru
            worksheet.update('A1', [df.columns.tolist()] + data)

            return "Data successfully uploaded to Google Sheets!"

        except Exception as e:
            return f"An error occurred: {e}"

    # Untuk metode GET, render form
    return render_template('upload_and_send.html')
