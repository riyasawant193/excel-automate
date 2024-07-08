from flask import Flask, request, render_template, send_file
import pdfplumber
import openpyxl
import re
import os

app = Flask(__name__)

def extract_text_from_pdf(pdf_path):
    text = ''
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text()
    return text

def parse_brochure(text):
    data = {
        'Rank': None,
        'University': None,
        'State': None,
        'Majors offered': None,
        'Acceptance rate': None,
        'Application Date': None,
        'Application Deadline': None,
        'Application fee': None,
        'Tuition fees': None,
        'Estimate': None,
        'GRE': None,
        'IELTS/TOEFLS': None,
        'LORs': None,
        'SOPs': None,
        'Upperhand': None,
        'Financial aid': None,
        'Curicullum': None,
        'Requirements': None
    }

    university_match = re.search(r'University Name:\s*(.*)', text)
    if university_match:
        data['University'] = university_match.group(1)
    
    state_match = re.search(r'State:\s*(.*)', text)
    if state_match:
        data['State'] = state_match.group(1)

    # Add more parsing logic with checks as needed

    return data

def write_to_excel(data, excel_path, row):
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active

    columns = [
        'Rank', 'University', 'State', 'Majors offered', 'Acceptance rate', 
        'Application Date', 'Application Deadline', 'Application fee', 
        'Tuition fees', 'Estimate', 'GRE', 'IELTS/TOEFLS', 'LORs', 
        'SOPs', 'Upperhand', 'Financial aid', 'Curicullum', 'Requirements'
    ]

    for col_num, column_name in enumerate(columns, 1):
        sheet.cell(row=row, column=col_num, value=data.get(column_name))

    wb.save(excel_path)

@app.route('/')
def upload_file():
    return '''
    <!doctype html>
    <title>Upload PDF File</title>
    <h1>Upload PDF File</h1>
    <form action="/upload" method=post enctype=multipart/form-data>
      <input type=file name=pdf_file>
      <input type=submit value=Upload>
    </form>
    '''

@app.route('/upload', methods=['POST'])
def upload_file_post():
    if 'pdf_file' not in request.files:
        return 'No file part'
    
    pdf_file = request.files['pdf_file']
    
    if pdf_file.filename == '':
        return 'No selected file'
    
    if pdf_file and pdf_file.filename.endswith('.pdf'):
        pdf_path = os.path.join(os.getcwd(), pdf_file.filename)
        excel_path = 'path_to_your_existing_excel_file.xlsx'  # Set the path to your existing Excel file
        
        pdf_file.save(pdf_path)
        
        text = extract_text_from_pdf(pdf_path)
        parsed_data = parse_brochure(text)
        write_to_excel(parsed_data, excel_path, row=2)  # Adjust row number as needed

        return send_file(excel_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
