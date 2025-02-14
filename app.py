
from flask import Flask, jsonify
from excel_utils.excel_handler import read_excel, write_excel

app = Flask(__name__)

@app.route("/read-excel", methods=["GET"])
def read_excel_route():
    """API route to read the Excel file."""
    data = read_excel()
    return jsonify(data)

@app.route("/write-excel", methods=["GET"])
def write_excel_route():
    """API route to write dummy data to the Excel file."""
    result = write_excel()
    return jsonify(result)


@app.route("/")
def home():
    return jsonify({"message": "Welcome to Flask!"})


if __name__ == "__main__":
    app.run(debug=True)
