from flask import Flask, request, Response
import pandas as pd
import os
import json

app = Flask(__name__)

# Create an upload folder for the files
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure the upload folder exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Route to display the form and upload the file
@app.route('/')
def upload_form():
    return '''
        <form method="POST" enctype="multipart/form-data">
            <input type="file" name="file" />
            <input type="submit" />
        </form>
    '''

# Route to handle the uploaded file and return the result in JSON format
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

    # Read the Excel file using pandas
    try:
        df = pd.read_excel(filepath)

        # Convert the DataFrame to JSON format
        json_data = df.to_json(orient='records')

        # Pretty-print the JSON data
        formatted_json = json.dumps(json.loads(json_data), indent=4)

        # Return the formatted JSON response
        return Response(formatted_json, mimetype='application/json')
    
    except Exception as e:
        return f"Error reading the file: {e}"

if __name__ == '__main__':
    app.run(debug=True)
