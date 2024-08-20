from flask import Blueprint, render_template, request, send_file
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import io
import requests
import os
import zipfile


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


@main.route('/download_files', methods=['GET', 'POST'])
def download_files():
    if request.method == 'POST':
        sheet_url = request.form['sheet_url']
        sheet_name = request.form['sheet_name']
        formats = request.form['formats'].split(',')  # Ambil format file dari form

        try:
            # Buka spreadsheet dan worksheet
            spreadsheet = client.open_by_url(sheet_url)
            worksheet = spreadsheet.worksheet(sheet_name)

            # Ambil data dari sheet
            rows = worksheet.get_all_values()
            headers = rows[0]
            data = rows[1:]

            # Inisialisasi file ZIP
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for index, row in enumerate(data):
                    if len(row) < 2:
                        continue  # Skip rows without enough data
                    
                    unique_name = row[0]  # Nama unik dari kolom pertama
                    links = row[1:3]  # Link dari kolom kedua dan ketiga

                    # Jika ada lebih banyak format daripada link, kita gunakan format default untuk sisa link
                    if len(formats) < len(links):
                        formats.extend(['.unknown'] * (len(links) - len(formats)))

                    for i, (link, file_format) in enumerate(zip(links, formats)):
                        if link and link.startswith("https://drive.google.com"):
                            try:
                                # Extract file ID from the link
                                file_id = link.split('/')[-2]
                                # Generate the download URL
                                download_url = f"https://drive.google.com/uc?id={file_id}&export=download"
                                
                                # Mendapatkan nama file dari header link dan format
                                header = headers[i + 1] if i + 1 < len(headers) else "unknown"
                                file_name = f"{index + 1}_{unique_name}_{header}{file_format}"  # Format file ditambahkan di sini

                                # Unduh file
                                response = requests.get(download_url, stream=True)
                                if response.status_code == 200:
                                    zip_file.writestr(file_name, response.content)
                                else:
                                    print(f"Failed to download file from {link}")

                            except Exception as e:
                                print(f"Error downloading file from link: {link}\n{e}")

            zip_buffer.seek(0)
            return send_file(zip_buffer, as_attachment=True, download_name='files.zip')

        except Exception as e:
            return f"An error occurred: {e}"

    return render_template('download_files.html')
