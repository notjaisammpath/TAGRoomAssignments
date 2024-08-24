from datetime import datetime
from flask import Flask, request, render_template, redirect, url_for, flash
import openpyxl
import os
import emails
import sensitive_info

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.secret_key = sensitive_info.secret_key
data_by_person = {}


def parse_excel(file_path):
    global data_by_person
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    data_by_person = {}
    for row in sheet.iter_rows(values_only=True):
        if row[0] is None:
            continue

        number = row[0]
        last_name = row[1].capitalize()
        first_name = row[2].capitalize()
        exam = row[3]
        date = row[4]
        time = row[5]
        room_number = row[6]
        proctor = row[7]

        full_name = f"{first_name} {last_name}"
        exam_details = {
            "exam": exam,
            "date": date,
            "number": number,
            "am_pm": time,
            "room_number": room_number,
            "proctor": proctor
        }

        if full_name not in data_by_person:
            data_by_person[full_name] = []
        data_by_person[full_name].append(exam_details)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    if file:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        parse_excel(file_path)
        return redirect(url_for('query'))
    return redirect(request.url)


@app.route('/query', methods=['GET', 'POST'])
def query():
    if request.method == 'POST':
        full_name = f"{request.form['full_name']}"    
        results = data_by_person.get(
            full_name, "No data found for this person")
        return render_template('query.html', results=results, name=full_name)
    return render_template('query.html')

@app.route('/send_email', methods=['POST'])
def send_email():
    # Implement your email sending logic here
    emails.send_emails(data_by_person)
    # Redirect back to the query page with a success message
    return redirect(url_for('query', message='Emails have been sent successfully!'))


if __name__ == '__main__':
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    app.run(debug=True, host="0.0.0.0", port=8080)
