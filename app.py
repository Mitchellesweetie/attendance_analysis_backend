# Importing flask module in the project is mandatory
# An object of Flask class is our WSGI application.
from flask import Flask,render_template,request,redirect,url_for,jsonify,send_file
import pandas as pd
import numpy as num
import os
import shutil
from datetime import datetime
import logging
import io
import base64
from pathlib import Path
import openpyxl as xls
from tabulate import tabulate
from werkzeug.utils import secure_filename
from spire.xls import *
import matplotlib.pyplot as plt 
from spire.xls.common import*

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
COMBINE_FOLDER = "combine"
ACCESS_FOLDER="access"
MERGED_FOLDER="merged"

Path(OUTPUT_FOLDER).mkdir(parents=True, exist_ok=True)
Path(UPLOAD_FOLDER).mkdir(parents=True, exist_ok=True)
Path(COMBINE_FOLDER).mkdir(parents=True, exist_ok=True)
Path(ACCESS_FOLDER).mkdir(parents=True, exist_ok=True)
Path(MERGED_FOLDER).mkdir(parents=True, exist_ok=True)



app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def uploadataforanalysis():
    # List files from the directories
    file_list = [f for f in os.listdir(OUTPUT_FOLDER) if f.endswith(".xlsx")]
    file_lis = [f for f in os.listdir(COMBINE_FOLDER) if f.endswith(".xlsx")]

    # Get the filter value and action
    filter_value = request.form.get('filter', '').strip().lower()
    action = request.form.get('action', '')

    emails = ''
    table_html = ''
    file_path = ''

    # Handle file view action
    if action == 'view' and request.form.get("selected_file"):
        selected_file = request.form.get("selected_file")
        file_path = os.path.join(OUTPUT_FOLDER, selected_file)

        try:
            df = pd.read_excel(file_path)
            emails = df['email'].count()
            table_html = df.to_html(classes="table table-striped-columns", index=False)
        except Exception as e:
            table_html = f"<p>Error reading file: {str(e)}</p>"

    # Handle search action
    elif action == 'search' and filter_value:
        selected_file = request.form.get("selected_file")  # Re-get the selected file

        if selected_file:
            file_path = os.path.join(OUTPUT_FOLDER, selected_file)

            if os.path.exists(file_path):
                try:
                    merged_df = pd.read_excel(file_path)

                    # Check if filter_value looks like an email (contains '@')
                    if '@' in filter_value:
                        merged_df = merged_df[merged_df['email'].str.contains(filter_value, case=False, na=False)]
                        emails = merged_df['email'].count()

                    # If filter_value is a number (for days), filter on the 'days' column
                    elif filter_value.isdigit():
                        merged_df = merged_df[merged_df['days'].astype(str).str.contains(filter_value, case=False, na=False)]
                        emails = merged_df['email'].count()

                    else:
                        # If it's neither an email nor a number, you can optionally handle this case
                        merged_df = merged_df[
                            merged_df['email'].str.contains(filter_value, case=False, na=False) | 
                            merged_df['days'].astype(str).str.contains(filter_value, case=False, na=False)
                            # emails = merged_df['email'].count()

                        ]
                    # emails = df['email'].count()

                    table_html = merged_df.to_html(classes="table table-striped-columns", index=False)
                except Exception as e:
                    table_html = f"<p>Error reading file: {str(e)}</p>"
            else:
                table_html = f"<p>Error: File not found at {file_path}</p>"
        else:
            table_html = "<p>Error: No file selected for filtering.</p>"

    return render_template(
        'academy.html',
        emails=emails,
        file_list=file_list,
        table_html=table_html,
        selected_file=request.form.get("selected_file"),
        file_lis=file_lis
    )

#filter
@app.route('/filter', methods=['POST'])
def filter_data():
    """Filter the selected file based on user input (email OR day, not both)"""
    
    file_list = [f for f in os.listdir(OUTPUT_FOLDER) if f.endswith(".xlsx")]  
    selected_file = request.form.get('selected_file', '').strip()
    filter_value = request.form.get('filter', '').strip().lower()
    emails_count=''

    if not selected_file:
        return render_template('academy.html', error="No file selected", file_list=file_list)

    file_path = os.path.join(OUTPUT_FOLDER, selected_file)

    if not os.path.exists(file_path):
        return render_template('academy.html', error="File does not exist", file_list=file_list)

    try:
        df = pd.read_excel(file_path, engine="openpyxl")
    except Exception as e:
        return render_template('academy.html', error=f"Error reading Excel file: {str(e)}", file_list=file_list)

    # Normalize column names
    df.columns = df.columns.str.lower()

    has_email = 'email' in df.columns
    has_days = 'days' in df.columns

    if not (has_email or has_days):
        return render_template('academy.html', error="Excel file must contain either 'Email' or 'Days' column", file_list=file_list)

    # Determine whether filtering by email or day
    is_email = has_email and '@' in filter_value
    is_day = has_days and filter_value.isdigit()

    if not (is_email or is_day):
        return render_template('academy.html', error="Invalid input. Enter either an email or a valid day", file_list=file_list)

    # Apply filtering
    if is_email:
        df = df[df['email'].astype(str).str.lower() == filter_value]
    elif is_day:
        df['days'] = pd.to_numeric(df['days'], errors='coerce')  # Convert 'days' to numeric
        df = df[df['days'] == int(filter_value)]

    if df.empty:
        return render_template('academy.html', error="No matching data found", file_list=file_list)

    # Convert filtered results into an HTML table
    table_html = df.to_html(classes="table table-bordered", index=False)

    # Count number of filtered attendees
    emails_count = df['email'].count() if 'email' in df.columns else 0

    return render_template(
        'academy.html',
        table_html=table_html,
        file_list=file_list,
        emails=emails_count,  # Pass the new email count
        selected_file=selected_file
    )


    


@app.route('/upload_master')
def uploadmaster():
    file_list = [f for f in os.listdir(OUTPUT_FOLDER) if f.endswith(".xlsx")]  
    file_lis = [f for f in os.listdir(ACCESS_FOLDER) if f.endswith(".xlsx")]  

    selected_file = request.form.get("selected_file", "")
    selected_fil = request.form.get("selected_fil", "")

    table_html = ""

    if selected_file and selected_fil:
        file_path1 = os.path.join(OUTPUT_FOLDER, selected_file)
        file_path2 = os.path.join(ACCESS_FOLDER, selected_fil)

        try:
            df1 = pd.read_excel(file_path1)
            df2 = pd.read_excel(file_path2)

            df1.columns = df1.columns.str.strip().str.lower()
            df2.columns = df2.columns.str.strip().str.lower()

            # print(df1[df1.duplicated(subset='email')])
            # print(df2[df2.duplicated(subset='email')])

          

            # df1=df1.dropna[]

            if "email" in df1.columns and "email" in df2.columns:
                df1["email"] = df1["email"].str.strip().str.lower()
                df2["email"] = df2["email"].str.strip().str.lower()
                merged_df = pd.merge(df1, df2, on="email", how="inner") #merged
                table_html = merged_df.to_html(classes="table table-striped-columns", index=False)
                merged_filename = f"merged_{secure_filename(selected_file)}_{secure_filename(selected_fil)}"
                merged_path = os.path.join(MERGED_FOLDER, merged_filename)
                merged_df.to_excel(merged_path, index=False)

                # print(f"Merged DF Shape: {merged_df.shape}")
                # print(f"Merged DF Sample:\n{merged_df.head()}")
                print(merged_path)
            else:
                table_html = "<p>Error: 'email' column not found in one or both files.</p>"

        except Exception as e:
            table_html = f"<p>Error merging files: {str(e)}</p>"
       
    # return jsonify(file_lis=file_lis)

    return render_template('upload_master.html', file_list=file_list, file_lis=file_lis, table_html=table_html, 
                           selected_file=selected_file, selected_fil=selected_fil)

# filter and download
# visualization


#merging two combined excel
@app.route('/analysis')
def aanalysis():
    return render_template('merge.html')

@app.route('/merge', methods=['POST', 'GET'])
def merge():
    """Displays a dropdown of available Excel files and allows merging of two selected files."""
    
    # Get lists of available files
    file_list = [f for f in os.listdir(OUTPUT_FOLDER) if f.endswith(".xlsx")]  
    file_lis = [f for f in os.listdir(ACCESS_FOLDER) if f.endswith(".xlsx")]  

    selected_file = request.form.get("selected_file", "")
    selected_fil = request.form.get("selected_fil", "")

    table_html = ""
    male = ''
    female = ''
    students=''
    other=''
    imgbase64=None
    img_base64 = None  # Store the histogram image

    if selected_file and selected_fil:
        file_path1 = os.path.join(OUTPUT_FOLDER, selected_file)
        file_path2 = os.path.join(ACCESS_FOLDER, selected_fil)

        try:
            df1 = pd.read_excel(file_path1)
            df2 = pd.read_excel(file_path2)

            df1.columns = df1.columns.str.strip().str.lower()
            # df2.columns = df2.columns.str.strip().str.lower()

            df2.columns = df2.columns.astype(str).str.strip().str.lower()
            

            if "email" in df1.columns and "email" in df2.columns:
                df1["email"] = df1["email"].str.strip().str.lower()
                df2["email"] = df2["email"].str.strip().str.lower()
                merged_df = pd.merge(df1, df2, on="email", how="inner")  # Merged data
                merged_df.columns = merged_df.columns.str.replace(r'(_x|_y)$', '', regex=True)
                female = merged_df['gender'].value_counts().get('Female', 0)
                male = merged_df['gender'].value_counts().get('Male', 0)
                other = merged_df['select your job category'].value_counts().get('Public Servant', 0)
                students = merged_df['select your job category'].value_counts().get('Other', 0)

                

                # Convert merged data to HTML table
                table_html = merged_df.to_html(classes="table table-striped-columns", index=False)
                merged_filename = f"merged_{secure_filename(selected_file)}_{secure_filename(selected_fil)}.xlsx"
                merged_path = os.path.join(MERGED_FOLDER, merged_filename)
                
                merged_df.to_excel(merged_path, index=False)
                

                # Generate histogram
                plt.figure(figsize=(6, 4))
                merged_df['gender'].value_counts().plot(kind='bar', color=['blue', 'blue'])
                plt.xlabel('Gender')
                plt.ylabel('Count')
                plt.title('Gender Distribution')

                # Convert plot to base64
                img_io = io.BytesIO()
                plt.savefig(img_io, format='png')
                img_io.seek(0)
                plt.close()
                img_base64 = base64.b64encode(img_io.getvalue()).decode('utf-8')
                # print({img_base64})
               
                #students and other
                plt.figure(figsize=(6, 4))
                merged_df['select your job category'].value_counts().plot(kind='bar', color=['blue', 'blue'])
                plt.xlabel('Gender')
                plt.ylabel('Count')
                plt.title('Category interms of job Distribution')

                # Convert plot to base64
                img_io = io.BytesIO()
                plt.savefig(img_io, format='png')
                img_io.seek(0)
                plt.close()
                imgbase64 = base64.b64encode(img_io.getvalue()).decode('utf-8')

        except Exception as e:
            table_html = f"<p>Error merging files: {str(e)}</p>"


    return render_template(
        'upload_master.html',
        file_list=file_list,
        file_lis=file_lis,
        table_html=table_html,
        selected_file=selected_file,
        selected_fil=selected_fil,
        male=male,
        female=female,
        img_base64=img_base64,
        imgbase64=imgbase64,other=other,students=students
    )

#posting the excel to be merged and validate with
@app.route('/combine',methods=['POST','GET'])
def combine_excel():
    try:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        if 'file' not in request.files:
            return jsonify({'error': 'No file part'}), 400

        f = request.files['file']
        if f.filename == '':
            return jsonify({'error': 'No selected file'}), 400
        
        filename, file_extension = os.path.splitext(f.filename)
        newfile=f"{filename}_{timestamp}{file_extension}"

        file_path = os.path.join(COMBINE_FOLDER, newfile)
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

        # merged_df = pd.concat(cleaned_sheets, ignore_index=True)
        merged_df = cleaned_sheets[0]  

        for df in cleaned_sheets[1:]:
            merged_df = merged_df.merge(df, on="email", how="right")
            print(merged_df)

        # Save the cleaned file
        merged_df = merged_df.drop(columns=["name_y","organization"])
        merged_df.columns = merged_df.columns.str.replace(r'(_x|_y)$', '', regex=True)
        merged_df = merged_df.rename(columns={'name_x': 'Full Name'})


        output_filename = f"cleaned_{newfile}"
        output_path = os.path.join(ACCESS_FOLDER, output_filename)
        
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            merged_df.to_excel(writer, sheet_name="Merged_Data", index=False)

        print('Processing successful:', output_path)
        return redirect(url_for('uploadmaster'))
        # return jsonify({'message': 'File processed successfully', 'output_file': output_filename})

    except Exception as e:
        print(f"Error: {e}")  # Debugging
        return jsonify({'error': f'There was an error: {str(e)}'}), 500

    
@app.route('/list')
def attendance_list():
    list=os.listdir(OUTPUT_FOLDER)
    return jsonify(list)



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
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        if 'file' not in request.files:
            return jsonify({'error': 'No file part'}), 400

        f = request.files['file']
        if f.filename == '':
            return jsonify({'error': 'No selected file'}), 400
        
        filename, file_extension = os.path.splitext(f.filename)
        newfile = f"{filename}_{timestamp}{file_extension}"

        file_path = os.path.join(UPLOAD_FOLDER, newfile)
        f.save(file_path)

        sheets_dict = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")

        if not sheets_dict:
            return jsonify({'error': 'Excel file is empty or unreadable'}), 400

        all_emails = []

        # Process each sheet
        cleaned_sheets = []
        for sheet_name, df in sheets_dict.items():
            df.columns = df.columns.astype(str).str.lower() 
            if 'email' in df.columns:
                df = df.drop_duplicates(subset=['email'])
                all_emails.append(df) 
                cleaned_sheets.append(df)

        if not cleaned_sheets:
            return jsonify({'error': 'No valid sheets with an email column'}), 400

        # Combine all emails into a single DataFrame
        combined_df = pd.concat(all_emails, ignore_index=True)

        # Count occurrences of each email
        email_counts = combined_df['email'].value_counts().reset_index()
        email_counts.columns = ['email', 'days']

        # Merge email count back into the final merged sheet
        merged_df = pd.concat(cleaned_sheets).drop_duplicates().merge(email_counts, on="email", how="left")

        # emails = merged_df['email'].count()     
        # print('count',emails)

        # Save the cleaned file
        output_filename = f"cleaned_{newfile}"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            merged_df.to_excel(writer, sheet_name="Merged_Data", index=False)

        print('Processing successful:', output_path)
        return redirect(url_for('uploadataforanalysis'))

    except Exception as e:
        print(f"Error: {e}")  # Debugging
        return jsonify({'error': f'There was an error: {str(e)}'}), 500


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

