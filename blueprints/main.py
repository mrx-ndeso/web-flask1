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

@main.route('/convert_csv_to_xlsx', methods=['GET', 'POST'])
def convert_csv_to_xlsx():
    if request.method == 'POST':
        # Ambil file dan delimiter dari form
        file = request.files['file']
        delimiter = request.form['delimiter']

        try:
            # Baca CSV menggunakan delimiter yang diinputkan
            df = pd.read_csv(file, delimiter=delimiter)
            df = df.astype(str)
            # Simpan DataFrame sebagai file XLSX
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            output.seek(0)

            # Kirim file XLSX sebagai lampiran
            return send_file(output, as_attachment=True, download_name='converted.xlsx')
        except Exception as e:
            return f"An error occurred: {e}"

    return render_template('convert_csv_to_xlsx.html')


@main.route('/combine_xlsx', methods=['GET', 'POST'])
def combine_xlsx():
    if request.method == 'POST':
        files = request.files.getlist('files')

        try:
            # List untuk menyimpan DataFrame dari setiap file
            dfs = []

            for file in files:
                df = pd.read_excel(file)
                df = df.astype(str)
                dfs.append(df)

            # Gabungkan semua DataFrame menjadi satu
            combined_df = pd.concat(dfs, ignore_index=True)

            # Simpan DataFrame gabungan sebagai file XLSX
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                combined_df.to_excel(writer, index=False, sheet_name='Combined')
            output.seek(0)

            # Kirim file XLSX gabungan sebagai lampiran
            return send_file(output, as_attachment=True, download_name='combined.xlsx')
        except Exception as e:
            return f"An error occurred: {e}"

    return render_template('combine_xlsx.html')
