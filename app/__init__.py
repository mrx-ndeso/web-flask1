from flask import Flask

app = Flask(__name__)

# Load configuration
app.config.from_pyfile('config.py')

from app import routes
