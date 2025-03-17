from flask import Blueprint, request, jsonify
from models.file_processor import process_file

file_bp = Blueprint("file_bp", __name__)

@file_bp.route("/upload", methods=["POST"])
def upload_file():
    """Handles multiple file uploads from one 'file' variable and processes them without saving"""
    if "file" not in request.files:
        return jsonify({"error": "No files uploaded"}), 400

    files = request.files.getlist("file")  # Get multiple files from one 'file' field
    results = [process_file(file) for file in files]  # Process each file

    return jsonify(results)
