from flask import Blueprint, jsonify, send_file
from app.services.excel_service import read_excel, write_excel
from app.config import Config


api_blueprint = Blueprint("api", __name__, url_prefix="/api")

@api_blueprint.route("/read-excel", methods=["GET"])
def read_excel_route():
    """Reads an Excel file."""
    return jsonify(read_excel())

@api_blueprint.route("/write-excel", methods=["GET"])
def write_excel_route():
    """Writes dummy data to a new Excel file."""
    return jsonify(write_excel())

@api_blueprint.route("/download-excel", methods=["GET"])
def download_excel_route():
    """Allows users to download the modified Excel file."""
    return send_file(Config.MODIFIED_FILE, as_attachment=True)
