from flask import Flask
from blueprints.main import main

app = Flask(__name__)

# Register Blueprints
app.register_blueprint(main)

print("Debug mode is", app.debug)


if __name__ == '__main__':
    app.run(debug=True)
