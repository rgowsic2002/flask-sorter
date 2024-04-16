import os
import zipfile
from flask import Flask, flash, request, redirect, render_template, send_file
from werkzeug.utils import secure_filename
import pandas as pd

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def sort_and_save_excel(file_path):
    df = pd.read_excel(file_path)
    sorted_df = df.sort_values(by=['WCdesc'])
    sorted_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'sorted_' + os.path.basename(file_path))
    sorted_df.to_excel(sorted_file_path, index=False)
    return sorted_file_path

def zip_sorted_files(sorted_files):
    zip_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'sorted_files.zip')
    with zipfile.ZipFile(zip_file_path, 'w') as zipf:
        for file_path in sorted_files:
            file_name = os.path.basename(file_path)
            zipf.write(file_path, file_name)
    return zip_file_path

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file part')
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        flash('No selected file')
        return redirect(request.url)
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        sorted_file_path = sort_and_save_excel(file_path)
        return render_template('results.html', zip_file_path=zip_sorted_files([sorted_file_path]))
    else:
        flash('Invalid file type, please upload an Excel file')
        return redirect(request.url)

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)

if __name__ == "__main__":
    app.secret_key = 'super_secret_key'
    app.run(debug=True)
