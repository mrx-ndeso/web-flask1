# Flask XLSX to Google Sheets

A Flask web application that allows users to upload XLSX files and send their contents to Google Sheets. This application provides a simple interface to specify the Google Sheets URL and target sheet name.

## Features

- Upload XLSX files
- Send data from XLSX files to a specified Google Sheets
- User-friendly interface with Bootstrap styling

## Requirements

- Python 3.x
- Flask
- Openpyxl (for handling XLSX files)
- Google API Client Library

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/mrx-ndeso/web-flask1.git
   cd web-flask1
2. Create and activate a virtual environment:

python -m venv venv
source venv/bin/activate  # On Windows use `venv\Scripts\activate`


3. Install the required packages:

pip install -r requirements.txt

4. Add credentials.json:

Since credentials.json is listed in .gitignore, you need to manually place it in the root directory of the project. This file contains your Google API credentials required for authentication. Follow these steps to obtain credentials.json:

Go to the Google Cloud Console.
Create a new project or select an existing one.
Enable the Google Sheets API and Google Drive API.
Create credentials for OAuth 2.0 Client IDs and download the credentials.json file.
Place this credentials.json file in the root directory of your project.


5. Run the application:

flask run
