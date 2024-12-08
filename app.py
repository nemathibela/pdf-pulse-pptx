from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from pdf2pptx import convert_pdf2pptx
from dotenv import load_dotenv
import tempfile
import os

app = Flask(__name__)
CORS(app)  # Adding CORS to all routes

load_dotenv()  # Load environment variables

@app.route('/convert-powerpoint', methods=['POST'])
def convert_pdf_to_pptx():
    # Check if a file is part of the request
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']

    # Check if the file is a PDF
    if not file.filename.endswith('.pdf'):
        return jsonify({'error': 'File is not a PDF'}), 400

    # Save the uploaded PDF file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
        temp_pdf_name = temp_pdf.name
        file.save(temp_pdf_name)

        # Define the output PPTX file path
        pptx_path = temp_pdf_name.replace('.pdf', '.pptx')

        # Convert PDF to PPTX
        # You may adjust the resolution, start_page, and page_count as needed
        resolution = 300  # Example resolution
        start_page = 0    # Start from the first page
        page_count = None  # Convert all pages; you can set a specific number if desired

        convert_pdf2pptx(temp_pdf_name, pptx_path, resolution, start_page, page_count)

    # Send the converted PPTX file back to the client
    return send_file(pptx_path, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
