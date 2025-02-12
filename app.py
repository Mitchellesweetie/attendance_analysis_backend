# Importing flask module in the project is mandatory
# An object of Flask class is our WSGI application.
from flask import Flask,render_template,request,redirect,url_for,jsonify,send_file
import pandas as pd
import numpy as num
from analys import processdata
import os
import logging
from pathlib import Path
import openpyxl as xls
from tabulate import tabulate
from spire.xls import *
from spire.xls.common import*

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

Path(OUTPUT_FOLDER).mkdir(parents=True, exist_ok=True)
Path(UPLOAD_FOLDER).mkdir(parents=True, exist_ok=True)

app = Flask(__name__)


# @app.route('/')
# def hello_world():
#     return 'Hello World'
@app.route('/list')
def attendance_list():
    list=os.listdir(OUTPUT_FOLDER)
    return jsonify(list)



@app.route('/index', methods=['GET', 'POST'])
def registered():
    """Displays a dropdown of available Excel files and renders the selected file."""
    
    file_list = [f for f in os.listdir(UPLOAD_FOLDER) if f.endswith(".xlsx")]  
    selected_file = request.form.get("selected_file")  
    table_html = ""  

    if selected_file:
        file_path = os.path.join(UPLOAD_FOLDER, selected_file)
        try:
            df = pd.read_excel(file_path)
            table_html = df.to_html(classes="table table-striped-columns", index=False)  
        except Exception as e:
            table_html = f"<p>Error reading file: {str(e)}</p>"

    return render_template('hello.html', file_list=file_list, table_html=table_html,selected_file=selected_file) 

@app.route('/analyse')
def analyse():
     return render_template('analyse.html')

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    """Allows downloading of a selected file."""
    try:
        output_path = os.path.join(OUTPUT_FOLDER, filename)

        if os.path.exists(output_path):
            return send_file(output_path, as_attachment=True)

        return jsonify({"error": "File not found"}), 404

    except Exception as e:  # Fixed indentation and syntax
        print(f"Error: {e}")
        return jsonify({"error": "An error occurred while processing the download"}), 500
  

#need to use javascript to redirect hence we use jsonify
@app.route('/submit', methods=["POST"])
def submit():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file part'}), 400

        f = request.files['file']
        if f.filename == '':
            return jsonify({'error': 'No selected file'}), 400

        file_path = os.path.join(UPLOAD_FOLDER, f.filename)
        f.save(file_path)

        sheets_dict = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")

        if not sheets_dict:
            return jsonify({'error': 'Excel file is empty or unreadable'}), 400

        cleaned_sheets = []

        # Process each sheet
        for sheet_name, df in sheets_dict.items():
            df.columns = df.columns.str.lower()
            if 'email' in df.columns:
                df = df.drop_duplicates(subset=['email'], keep='first')  # Remove duplicate emails
                cleaned_sheets.append(df)

        if not cleaned_sheets:
            return jsonify({'error': 'No valid sheets with an email column'}), 400

        merged_df = pd.concat(cleaned_sheets, ignore_index=True)

        # Save the cleaned file
        output_filename = f"cleaned_{f.filename}"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            merged_df.to_excel(writer, sheet_name="Merged_Data", index=False)

        print('Processing successful:', output_path)
        return redirect(url_for('download_file'))
        # return jsonify({'message': 'File processed successfully', 'output_file': output_filename})

    except Exception as e:
        print(f"Error: {e}")  # Debugging
        return jsonify({'error': f'There was an error: {str(e)}'}), 50




# @app.route('/analyze_attendance',methods=["GET"])
# def analyze():
#     try:
#         data = request.json
#         if  not data:
#             return jsonify('data needed is in json')
#         filename = data.get('filename')

#         if not filename:
#             return jsonify({'error': 'No file selected'})

#         file_path = os.path.join(UPLOAD_FOLDER, filename)
        
#         if not os.path.exists(file_path):
#             return jsonify({'error': 'File does not exist'})

#         # Read all sheets dynamically
#         sheets_dict = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")

#         if not sheets_dict:
#             return jsonify({'error': 'Excel file is empty or unreadable'})

#         cleaned_sheets = []
        
#         # Process each sheet
#         for sheet_name, df in sheets_dict.items():
#             if 'email' in df.columns:
#                 df = df.drop_duplicates(subset=['email'], keep='first')  # Remove duplicates
#                 cleaned_sheets.append(df)

#         # Merge all sheets
#         if cleaned_sheets:
#             merged_df = pd.concat(cleaned_sheets, ignore_index=True)
#         else:
#             return jsonify({'error': 'No valid sheets with an email column'})

#         # Ensure output folder exists
#         os.makedirs(OUTPUT_FOLDER, exist_ok=True)

#         # Save the cleaned file
#         output_path = os.path.join(OUTPUT_FOLDER, f"cleaned_{filename}")
#         with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
#             merged_df.to_excel(writer, sheet_name="Merged_Data", index=False)

#         print('Processing successful:', output_path)

#         return jsonify({'message': 'File processed successfully', 'output_file': output_path})

#     except Exception as e:
#         print(f"Error: {e}")  # Print actual error for debugging
#         return jsonify({'error': f'There is an error: {str(e)}'})



#counting
def process_excel(file_path):
    """Processes an Excel file and calculates the adult percentage."""
    try:
        df = pd.read_excel(file_path)

        # Check if requeired columns exist
        if 'email' not in df.columns or 'age' not in df.columns:
            return None, "Missing required columns (email, age)"

        # Drop rows where email is missing
        df = df.dropna(subset=['email'])

        # Calculate adult percentage (age >= 18)
        total_count = len(df)
        adult_count = len(df[df['age'] >= 18])
        adult_percentage = round((adult_count / total_count) * 100, 2) if total_count > 0 else 0

        # Save processed file
        output_file = os.path.join(OUTPUT_FOLDER, "processed_data.xlsx")
        df.to_excel(output_file, index=False)

        return adult_percentage, None
    except Exception as e:
        return None, str(e)



if __name__ == '__main__':
    app.run(debug=True, port=5000)

