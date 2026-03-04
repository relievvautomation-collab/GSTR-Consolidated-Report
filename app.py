import os
import uuid
from datetime import datetime, timedelta
from flask import Flask, render_template, request, jsonify, send_file, after_this_request
from werkzeug.utils import secure_filename

import gstr1_parser
import gstr3b_parser
import gstr2a_parser

app = Flask(__name__)
app.secret_key = 'gstr-consolidator-secret-key'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
ALLOWED_EXTENSIONS = {'json'}

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

def cleanup_old_files():
    now = datetime.now()
    for folder in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER']]:
        for fname in os.listdir(folder):
            fpath = os.path.join(folder, fname)
            if os.path.isfile(fpath):
                mtime = datetime.fromtimestamp(os.path.getmtime(fpath))
                if now - mtime > timedelta(hours=1):
                    os.remove(fpath)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    cleanup_old_files()
    if 'files[]' not in request.files:
        return jsonify({'error': 'No files uploaded'}), 400

    return_type = request.form.get('returnType', 'GSTR-3B')
    files = request.files.getlist('files[]')
    saved_paths = []
    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(path)
            saved_paths.append(path)

    if not saved_paths:
        return jsonify({'error': 'No valid JSON files'}), 400

    # ---------- GSTR-2A ----------
    if return_type == 'GSTR-2A':
        consolidated, err = gstr2a_parser.parse_gstr2a_files(saved_paths)
        if err:
            return jsonify({'error': err}), 400
        preview = {
            'type': 'GSTR-2A',
            'meta': consolidated['meta'],
            'sheets': {
                name: {
                    'columns': data['columns'],
                    'rows': data['rows'][:50]   # preview first 50 rows
                } for name, data in consolidated['sheets'].items()
            }
        }
        sheet_count = 4   # Meta_Data + B2B + CDN + TCS
        row_count = sum(len(data['rows']) for data in consolidated['sheets'].values())
        out_fname = gstr2a_parser.create_gstr2a_excel_file(consolidated)

    # ---------- GSTR-1 ----------
    elif return_type == 'GSTR-1':
        consolidated, err = gstr1_parser.parse_gstr1_files(saved_paths)
        if err:
            return jsonify({'error': err}), 400
        preview = {
            'type': 'GSTR-1',
            'meta': consolidated['meta'],
            'sheets': {
                name: {
                    'columns': data['columns'],
                    'rows': data['rows'][:50]
                } for name, data in consolidated['sheets'].items()
            }
        }
        sheet_count = 5
        row_count = sum(len(data['rows']) for data in consolidated['sheets'].values())
        out_fname = gstr1_parser.create_gstr1_excel_file(consolidated)

    # ---------- GSTR-3B ----------
    else:  # default GSTR-3B
        consolidated, err = gstr3b_parser.parse_gstr3b_files(saved_paths)
        if err:
            return jsonify({'error': err}), 400
        preview = {
            'type': 'GSTR-3B',
            'meta': consolidated['meta'],
            'rows': consolidated['rows']
        }
        sheet_count = 2
        row_count = len(consolidated['rows'])
        out_fname = gstr3b_parser.create_gstr3b_excel_file(consolidated)

    token = str(uuid.uuid4())
    if not hasattr(app, 'file_map'):
        app.file_map = {}
    app.file_map[token] = out_fname

    return jsonify({
        'success': True,
        'preview': preview,
        'token': token,
        'fileCount': len(saved_paths),
        'sheetCount': sheet_count,
        'rowCount': row_count,
        'monthCount': len(consolidated['meta']['months'])
    })

@app.route('/download')
def download():
    token = request.args.get('token')
    if not token or not hasattr(app, 'file_map') or token not in app.file_map:
        return jsonify({'error': 'Invalid token'}), 400
    fname = app.file_map[token]
    path = os.path.join(app.config['OUTPUT_FOLDER'], fname)
    if not os.path.exists(path):
        return jsonify({'error': 'File not found'}), 404

    @after_this_request
    def cleanup(response):
        try:
            os.remove(path)
            del app.file_map[token]
        except:
            pass
        return response

    return send_file(path, as_attachment=True, download_name=fname)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000, debug=True)
