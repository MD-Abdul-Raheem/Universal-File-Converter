from flask import Flask, request, jsonify, send_file, render_template, url_for
import os
import tempfile
from werkzeug.utils import secure_filename
import uuid

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

# Import converter after Flask app creation to avoid circular imports
try:
    from converters import FileConverter
    converter = FileConverter()
except ImportError as e:
    print(f"Warning: Could not import FileConverter: {e}")
    converter = None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/health')
def health():
    return "Universal File Converter is running! 🚀"

@app.route('/about')
def about():
    return render_template('about.html')

@app.route('/convert', methods=['POST'])
def convert_file():
    try:
        if converter is None:
            return jsonify({'error': 'File converter not available'}), 500
            
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        output_format = request.form.get('format', '').lower()
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not output_format:
            return jsonify({'error': 'No output format selected'}), 400
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        temp_id = str(uuid.uuid4())
        temp_path = os.path.join(tempfile.gettempdir(), f"{temp_id}_{filename}")
        file.save(temp_path)
        
        # Convert immediately
        output_path = converter.convert(temp_path, output_format)
        
        # Generate download filename
        base_name = os.path.splitext(filename)[0]
        download_name = f"{base_name}.{output_format}"
        
        # Return file directly and clean up after
        try:
            return send_file(
                output_path,
                as_attachment=True,
                download_name=download_name
            )
        finally:
            # Clean up both input and output files
            try:
                os.remove(temp_path)
                os.remove(output_path)
            except:
                pass
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Download route removed - files are served directly from /convert

if __name__ == '__main__':
    app.run(debug=True, port=5000)