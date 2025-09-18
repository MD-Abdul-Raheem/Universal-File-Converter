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
        
        # Convert immediately (optimized)
        output_path = converter.convert(temp_path, output_format)
        
        # Read file for text preview
        text_content = None
        if output_format in ['txt', 'html', 'xml', 'csv']:
            try:
                with open(output_path, 'r', encoding='utf-8') as f:
                    text_content = f.read()
            except:
                pass
        
        # Don't clean up files immediately - keep them for download
        # Files will be cleaned up by system temp cleanup
        
        output_filename = os.path.basename(output_path)
        
        return jsonify({
            'success': True,
            'text_content': text_content,
            'format': output_format,
            'download_url': f"/download/{output_filename}"
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        filepath = os.path.join(tempfile.gettempdir(), filename)
        if os.path.exists(filepath):
            return send_file(filepath, as_attachment=True)
        return jsonify({'error': f'File not found: {filename}'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)