from flask import Flask, request, Response, jsonify
import pandas as pd
import os
import json
from datetime import datetime

app = Flask(__name__)

# Directories for uploads and JSON outputs
UPLOAD_FOLDER = 'uploads'
JSON_FOLDER = 'json_outputs'

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['JSON_FOLDER'] = JSON_FOLDER

# Ensure the required folders exist
for folder in [UPLOAD_FOLDER, JSON_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# Route to display the form and upload the file
@app.route('/')
def upload_form():
    return '''
        <form method="POST" enctype="multipart/form-data">
            <input type="file" name="file" />
            <input type="submit" />
        </form>
    '''

# Route to handle the uploaded file, save JSON, and return the result
@app.route('/', methods=['POST'])
def handle_file():
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
            # Handle merged cells by filling NaNs
            df.fillna(method='ffill', inplace=True)
            df.fillna(method='bfill', inplace=True)

            # Convert each sheet's DataFrame to JSON
            result[sheet_name] = json.loads(df.to_json(orient='records'))

        # Save JSON to a local file with a timestamp
        json_filename = f"{os.path.splitext(file.filename)[0]}_{datetime.now().strftime('%Y%m%d%H%M%S')}.json"
        json_filepath = os.path.join(app.config['JSON_FOLDER'], json_filename)

        with open(json_filepath, 'w', encoding='utf-8') as json_file:
            json.dump(result, json_file, indent=4)

        # Return a success message with JSON file details
        return jsonify({
            "message": "File processed and JSON saved successfully",
            "json_file": json_filename,
            "json_path": json_filepath,
            "data_preview": result  # Returning a preview of JSON data
        })
    
    except Exception as e:
        return f"Error processing the file: {e}"

if __name__ == '__main__':
    app.run(debug=True)
