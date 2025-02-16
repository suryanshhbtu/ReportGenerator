
from flask import Flask, jsonify, request
from excel_utils.excel_handler import read_excel, write_excel
from flask_swagger_ui import get_swaggerui_blueprint

app = Flask(__name__)

# flask swagger configs
SWAGGER_URL = '/swagger'
API_URL = '/static/swagger.json'
SWAGGERUI_BLUEPRINT = get_swaggerui_blueprint(
    SWAGGER_URL,
    API_URL,
    config={
        'app_name': "Read/Write Excel API"
    }
)
app.register_blueprint(SWAGGERUI_BLUEPRINT, url_prefix=SWAGGER_URL)

@app.route("/read-excel", methods=["GET"])
def read_excel_route():
    """API route to read the Excel file."""
    data = read_excel()
    return jsonify(data)

@app.route("/write-excel", methods=["POST"])
def write_excel_route():
    """API route to write dummy data to the Excel file."""
    json_data = request.get_json()
    # Validate input
    # if not isinstance(json_data, list) or not all(isinstance(row, dict) for row in json_data):
    #     return jsonify({"error": "Invalid data format. Expected a list of dictionaries."}), 400
    result = write_excel(json_data)
    return jsonify(result)


@app.route("/")
def home():
    return jsonify({"message": "Welcome to Flask!"})


if __name__ == "__main__":
    app.run(debug=True)
