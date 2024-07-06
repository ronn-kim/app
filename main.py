from flask import Flask, render_template, request, redirect, url_for, flash
import os
from werkzeug.utils import secure_filename
import subprocess

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'txt', 'pdf', 'xls', 'xlsx', 'xlsm'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = 'e6934c9510c6bcae4092d17b6d0aaeedb5745c2f'

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

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
        flash('File uploaded successfully')
        return redirect(url_for('execute_script', filename=filename))
    else:
        flash('Allowed file types are txt, pdf, xls, xlsx, xlsm')
        return redirect(request.url)

@app.route('/execute/<filename>')
def execute_script(filename):
    # Execute the uploaded script using subprocess
    script_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    try:
        result = subprocess.run(['python', script_path], capture_output=True, text=True)
        if result.returncode == 1:
            flash('Script executed successfully')
        else:
            flash(f'Error executing script: {result.stderr}')
    except subprocess.CalledProcessError as e:
        flash(f'Error executing script: {e}')
    
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)

