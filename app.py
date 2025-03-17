from flask import Flask
from flask_cors import CORS
from controllers.file_controller import file_bp

app = Flask(__name__)
CORS(app)  # Enable CORS

# Register the Blueprint
app.register_blueprint(file_bp)

if __name__ == "__main__":
    app.run(debug=True, port=5000)
