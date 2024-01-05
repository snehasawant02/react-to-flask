from flask import Flask, request, send_file
import pandas as pd
import tabula
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# Function to append tables from a PDF to an existing Excel file
def append_table_to_excel(pdf_path, excel_path):
    # Read the existing Excel file if it exists
    try:
        existing_data = pd.read_excel(excel_path)
    except FileNotFoundError:
        # If the file doesn't exist, create an empty DataFrame
        existing_data = pd.DataFrame()

    # Extract tables from the PDF
    tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)

    # Iterate through tables and append them to the existing data
    for i, table in enumerate(tables):
        if not existing_data.empty:
            # Append the table to the existing data
            existing_data = pd.concat([existing_data, table], ignore_index=True)
        else:
            # If there is no existing data, use the table directly
            existing_data = table

    # Write the combined data to the Excel file
    existing_data.to_excel(excel_path, index=False)

# Define a route to handle file uploads
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return {'error': 'No file part'}, 400

    file = request.files['file']

    if file.filename == '':
        return {'error': 'No selected file'}, 400

    if file and file.filename.endswith('.pdf'):
        # Save the uploaded file
        upload_path = 'uploaded_file.pdf'
        file.save(upload_path)

        # Process the PDF file
        output_excel_path = 'output_tables2.xlsx'
        append_table_to_excel(upload_path, output_excel_path)

        # Provide the processed file for download
        return send_file(output_excel_path, as_attachment=True)

    return {'error': 'Invalid file format'}, 400

if __name__ == '__main__':
    app.run(debug=True)
