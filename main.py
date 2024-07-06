from flask import Flask, render_template,redirect,url_for, request, send_file
import os
import pythoncom
import win32com.client as win32

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure the upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    if file:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        # Redirect to the route for executing the macro with the uploaded file
        return redirect(url_for('macro_execute', filename=file.filename))

@app.route('/macro_execute/<filename>')
def macro_execute(filename):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    execute_excel_macro(filepath)
    return send_file(filepath, as_attachment=True)

def execute_excel_macro(filepath):
    pythoncom.CoInitialize()  
    xl = win32.Dispatch('Excel.Application')  
    wb = xl.Workbooks.Open(filepath) 
    xl.Application.Run('final_script_mpesa')  
    wb.Save()  
    wb.Close() 
    xl.Quit()  

if __name__ == '__main__':
    app.run(debug=True)

    

