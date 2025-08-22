import json
import os
import subprocess
import time
from flask import Flask, flash, jsonify, render_template, request, redirect, url_for, session
from flask_cors import CORS
import pandas as pd
from werkzeug.utils import secure_filename
from append_students_to_excel import add_to_database



app = Flask(__name__)
CORS(app)
app.secret_key = os.urandom(24) # Needed for session

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Dummy Admin Credentials
ADMIN_USERNAME = 'admin'
ADMIN_PASSWORD = 'admin123'
FILES_FOLDER ="D:/auto_attendance/output"
EXCEL_PATH = "D:/auto_attendance/data.xlsx"
IMAGE_FOLDER  ="D:/auto_attendance/Images"
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg'}


@app.route('/')
def home():
    return redirect('/login')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        if request.form['username'] == 'admin' and request.form['password'] == 'admin123':
            session['admin_logged_in'] = True
            return redirect('/home')
        else:
            return "Invalid credentials"
    return render_template('login.html')

@app.route('/home')
def home_page():
    if not session.get('admin_logged_in'):
        return redirect('/login')
    return render_template('home.html')

@app.route('/admin')
def admin_page():
    if not session.get('admin_logged_in'):
        return redirect('/login')
    return render_template('admin.html')

@app.route('/change-time-slot', methods=['GET', 'POST'])
def change_time_slot():
    if not session.get('admin_logged_in'):
        return redirect('/login')
    
    if request.method == 'POST':
        start = request.form.get('start_time')
        end = request.form.get('end_time')
        if start and end:
            # Load existing slots
            if os.path.exists('time_slots.json'):
                with open('time_slots.json', 'r') as f:
                    slots = json.load(f)
            else:
                slots = []

            # Append new slot
            slots.append([start, end])
            with open('time_slots.json', 'w') as f:
                json.dump(slots, f)

            return "New Time Slot Added Successfully!"
    
    # Show current slots
    current_slots = []
    if os.path.exists('time_slots.json'):
        with open('time_slots.json', 'r') as f:
            current_slots = json.load(f)
    
    return render_template('change_time_slot.html', slots=current_slots)



@app.route('/delete-time-slot', methods=['POST'])
def delete_time_slot():
    if not session.get('admin_logged_in'):
        return redirect('/login')

    index = int(request.form.get('index'))

    if os.path.exists('time_slots.json'):
        with open('time_slots.json', 'r') as f:
            slots = json.load(f)

        if 0 <= index < len(slots):
            slots.pop(index)
            with open('time_slots.json', 'w') as f:
                json.dump(slots, f)

    return redirect('/change-time-slot')



@app.route('/new-student', methods=['GET', 'POST'])
def new_student():
    if request.method == 'POST':
        roll_no = request.form['roll_no']
        student_name = request.form['student_name']
        parent_no = request.form['parent_no']

        if 'image' not in request.files:
            return "No image file part"

        file = request.files['image']
        if file.filename == '':
            return "No selected image"

        if file and allowed_file(file.filename):
            filename = secure_filename(f"{student_name}.jpg")
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

            # Add student to JSON file
            with open('students.json', 'r+') as f:
                students = json.load(f)
                new_student = {
                    "S/No": len(students) + 1,
                    "Roll No": roll_no,
                    "Student Name": student_name,
                    "Parent No": parent_no
                }
                students.append(new_student)
                f.seek(0)
                json.dump(students, f, indent=4)

            # Call Excel update function
            add_to_database()

            return redirect(url_for('new_student'))

    return render_template('new_student.html')


@app.route('/files')
def list_files():
    if not session.get('admin_logged_in'):
        return {'error': 'Unauthorized'}, 401

    files = os.listdir(FILES_FOLDER)
    files = [f for f in files if f.endswith('.xlsx')]
    return jsonify(files)


@app.route('/logout')
def logout():
    session.pop('admin_logged_in', None)
    return redirect('/login')

if __name__ == '__main__':
    app.run(debug=True)
