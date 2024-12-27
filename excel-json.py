from flask import Flask, request, Response, jsonify
import pandas as pd
import os
import json
from datetime import datetime

app = Flask(__name__)

# Directories for uploads, JSON outputs, and Excel outputs
UPLOAD_FOLDER = 'uploads'
JSON_FOLDER = 'json_outputs'
EXCEL_FOLDER = 'excel_outputs'

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['JSON_FOLDER'] = JSON_FOLDER
app.config['EXCEL_FOLDER'] = EXCEL_FOLDER

# Ensure the required folders exist
for folder in [UPLOAD_FOLDER, JSON_FOLDER, EXCEL_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)


# Route to display the form for file uploads
@app.route('/')
def upload_form():
    return '''
        <h2>Upload an Excel or JSON File</h2>
        <form method="POST" enctype="multipart/form-data" action="/upload_excel">
            <h3>Excel to JSON:</h3>
            <input type="file" name="file" accept=".xlsx, .xls" />
            <input type="submit" value="Upload Excel" />
        </form>
        <br>
        <form method="POST" enctype="multipart/form-data" action="/upload_json">
            <h3>JSON to Excel:</h3>
            <input type="file" name="file" accept=".json" />
            <input type="submit" value="Upload JSON" />
        </form>
    '''


# Route to handle Excel to JSON conversion
@app.route('/upload_excel', methods=['POST'])
def handle_excel():
    if 'file' not in request.files:
        return "No file part"
    
    file = request.files['file']
    if file.filename == '':
        return "No selected file"

    # Save the uploaded file
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    try:
        # Read all worksheets into a dictionary of DataFrames
        workbook = pd.read_excel(filepath, sheet_name=None, engine='openpyxl')

        result = {}
        for sheet_name, df in workbook.items():
            df.fillna(method='ffill', inplace=True)
            df.fillna(method='bfill', inplace=True)
            result[sheet_name] = json.loads(df.to_json(orient='records'))

        # Save JSON to a local file with a timestamp
        json_filename = f"{os.path.splitext(file.filename)[0]}_{datetime.now().strftime('%Y%m%d%H%M%S')}.json"
        json_filepath = os.path.join(app.config['JSON_FOLDER'], json_filename)

        with open(json_filepath, 'w', encoding='utf-8') as json_file:
            json.dump(result, json_file, indent=4)

        return jsonify({
            "message": "Excel converted to JSON and saved successfully",
            "json_file": json_filename,
            "json_path": json_filepath,
            "data_preview": result
        })

    except Exception as e:
        return f"Error processing Excel file: {e}"


# Route to handle JSON to Excel conversion
@app.route('/upload_json', methods=['POST'])
def handle_json():
    if 'file' not in request.files:
        return "No file part"
    
    file = request.files['file']
    if file.filename == '':
        return "No selected file"

    # Save the uploaded JSON file
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    try:
        # Read JSON data
        with open(filepath, 'r', encoding='utf-8') as json_file:
            json_data = json.load(json_file)

        # Check if JSON contains multiple sheets
        if not isinstance(json_data, dict):
            return "Invalid JSON format. JSON must contain multiple sheets as top-level keys."

        # Write data to Excel file
        excel_filename = f"{os.path.splitext(file.filename)[0]}_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
        excel_filepath = os.path.join(app.config['EXCEL_FOLDER'], excel_filename)

        with pd.ExcelWriter(excel_filepath, engine='openpyxl') as writer:
            for sheet_name, data in json_data.items():
                if isinstance(data, list):
                    df = pd.DataFrame(data)
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    return f"Invalid data format in sheet: {sheet_name}"

        return jsonify({
            "message": "JSON converted to Excel and saved successfully",
            "excel_file": excel_filename,
            "excel_path": excel_filepath
        })

    except Exception as e:
        return f"Error processing JSON file: {e}"


if __name__ == '__main__':
    app.run(debug=True)
