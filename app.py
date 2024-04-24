from flask import Flask, jsonify, send_from_directory
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from werkzeug.utils import secure_filename
import os
import json
from datetime import datetime
import pandas as pd
# referencing an external py file
from descriptive_statistics import analyze_file

# Initialize Flask application
app = Flask(__name__)
app.secret_key = 'supersecretkey'  # Required to use flash messages

# Set the path for the uploads folder
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'csv', 'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB upload limit

# Set the path for the json folder
JSON_FOLDER = 'json_objects'
JSON_EXTENSIONS = {'json'}
app.config['JSON_FOLDER'] = JSON_FOLDER
os.makedirs(JSON_FOLDER, exist_ok=True)

# Ensure the upload directory exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Function to convert statistical results to JSON format
def create_json_graph(file_path):
    try:
        stats = analyze_file(file_path)

        if stats is None:
            return {'error': 'Failed to generate statistics'}, None

        if 'error' in stats:
            return {'error': stats['error']}, None

        def convert(obj):
            if isinstance(obj, (int, float)):
                return obj
            elif isinstance(obj, list):
                return [convert(item) for item in obj]
            elif isinstance(obj, dict):
                return {key: convert(value) for key, value in obj.items()}
            return str(obj)

        stats = convert(stats)

        # Create JSON file with statistics
        json_file_name = os.path.basename(file_path).replace('.', '_') + '.json'
        json_file_path = os.path.join(app.config['JSON_FOLDER'], json_file_name)


        with open(json_file_path, 'w') as json_file:
            json.dump(stats, json_file, indent=4)

        return stats, os.path.basename(json_file_path)


    except Exception as e:
        return {'error': str(e)}, None
    
# Function to check if file has allowed extension
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Route for home page
@app.route('/')
def home():
    breadcrumbs = [("Home", "/")]
    return render_template('index.html', breadcrumbs=breadcrumbs)

# Function to drop 'Export Summary' sheet from Excel files
def drop_specific_sheet(file_path):
    # Check if the file is an Excel file
    if not file_path.endswith(('.xlsx', '.xls')):
        return None, 'Unsupported file format'

    # Load the Excel file
    xls = pd.ExcelFile(file_path)

    # Check if the sheet 'Export Summary' exists
    if 'Export Summary' in xls.sheet_names:
        # Read all sheets except 'Export Summary'
        sheets = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names if sheet_name != 'Export Summary'}

        # Save the modified Excel file
        modified_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"modified_{os.path.basename(file_path)}")
        with pd.ExcelWriter(modified_file_path) as writer:
            for sheet_name, df in sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        return modified_file_path, 'Sheet "Export Summary" dropped successfully'

    else:
        return None, 'Sheet "Export Summary" not found'

# Route for uploading data files
@app.route('/data-upload', methods=['GET', 'POST'])
def data_upload():
    breadcrumbs = [("Home", "/"), ("Data Upload", "/data-upload")]
    if request.method == 'POST':
        file = request.files['dataFile']
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            
            # Drop 'Export Summary' sheet
            modified_file_path, message = drop_specific_sheet(file_path)
            
            if modified_file_path:
                flash(f'😄 File {filename} uploaded and "Export Summary" sheet dropped successfully!', 'success')
            else:
                flash(f'😞 File {filename} uploaded but "Export Summary" sheet was not found.', 'warning')
                files = os.listdir(app.config['UPLOAD_FOLDER'])
                files_info = [{
                 'name': file,
                 'size': f"{os.stat(os.path.join(app.config['UPLOAD_FOLDER'], file)).st_size / 1024:.2f} KB",
                 'upload_time': datetime.fromtimestamp(os.stat(os.path.join(app.config['UPLOAD_FOLDER'], file)).st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            } for file in files]
    return render_template('data_upload.html', breadcrumbs=breadcrumbs, files=files_info)

# Route for displaying datasets
@app.route('/datasets')
def datasets():
    files = []
    for filename in os.listdir(app.config['UPLOAD_FOLDER']):
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        stats = os.stat(filepath)
        files.append({
            'name': filename,
            'size': f"{stats.st_size / 1024:.2f} KB",
            'upload_time': datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
        })
    return render_template('datasets.html', files=files)

# Route for exploring data
@app.route('/explore-data')
def explore_data():
    breadcrumbs = [("Home", "/"), ("Explore Data", "/explore-data")]
    files = [f for f in os.listdir(app.config['UPLOAD_FOLDER']) if allowed_file(f)]
    return render_template('explore_data.html', files=files, breadcrumbs=breadcrumbs)

# Route for analyzing data
@app.route('/analyze-data', methods=['POST'])
def analyze_data():
    selected_file = request.form.get('selectedFile')
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], selected_file)
    if not os.path.exists(file_path):
        flash("File does not exist.", 'danger')
        return redirect(url_for('explore_data'))

    breadcrumbs = [("Home", "/"), ("Explore Data", "/explore-data"), ("Analyze Data", "/analyze-data")]

    try:
        # Check the file extension and read accordingly
        if selected_file.endswith(('.xlsx', '.xls')):
            # Read all sheets; each sheet's DataFrame is stored in a dictionary
            xls = pd.ExcelFile(file_path)
            sheets = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names}
        elif selected_file.endswith('.csv'):
            df = pd.read_csv(file_path)
            sheets = {'CSV': df}
        else:
            flash("Unsupported file format.", 'danger')
            return redirect(url_for('explore_data'))

        # Convert each DataFrame to HTML table
        tables_html = {sheet_name: df.to_html(classes='table table-striped', index=False) for sheet_name, df in sheets.items()}

        # Call create_json_graph to print JSON data to console
        for sheet_name, df in sheets.items():
            json_data = create_json_graph(file_path)  # Assuming you want to print JSON for each sheet
            print(f"JSON data for {sheet_name}:")
            print(json_data)

    except Exception as e:
        flash(f"Failed to read the file: {str(e)}", 'danger')
        return redirect(url_for('explore_data'))

    return render_template('analysis_results.html', tables_html=tables_html, filename=selected_file, sheet_names=list(tables_html.keys()), json_data=json_data, breadcrumbs=breadcrumbs)

# Route for listing JSON files
@app.route('/list_files', methods=['GET'])
def list_files():
    files = [f for f in os.listdir(app.config['JSON_FOLDER']) if f.endswith('.json')]
    return jsonify(files)

# Route for getting JSON data
@app.route('/get_json_data/<filename>', methods=['GET'])
def get_json_data(filename):
    json_file_path = os.path.join(app.config['JSON_FOLDER'], filename)
    
    if not os.path.exists(json_file_path):
        return jsonify({'error': 'File not found'}), 404
    
    with open(json_file_path, 'r') as file:
        data = json.load(file)
        
    return jsonify(data)

# Route for descriptive statistics
@app.route('/descriptive-statistics')
def descriptive_statistics():
    breadcrumbs = [("Home", "/"), ("Descriptive Statistics", "/descriptive-statistics")]
    files = [f for f in os.listdir(app.config['UPLOAD_FOLDER']) if allowed_file(f)]
    stats = {}
    json_files = {}
    
    for file in files:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file)
        stats[file], json_files[file] = create_json_graph(file_path)

        # Call create_json_graph to print JSON data to console
        print(f"JSON data for {file}:")
        print(stats[file])
    
    return render_template('descriptive_statistics.html', files=files, stats=stats, json_files=json_files, breadcrumbs=breadcrumbs)

# Route for descriptive statistics viewer
@app.route('/descriptive-statistics-viewer')
def descriptive_statistics_viewer():
    breadcrumbs = [("Home", "/"), ("Descriptive Statistics Viewer", "/descriptive-statistics-viewer")]
    files = [f for f in os.listdir(app.config['UPLOAD_FOLDER']) if f.endswith(('.xlsx', '.xls', '.csv'))]  # Filter to include only relevant file types
    return render_template('descriptive_statistics_viewer.html', files=files, breadcrumbs=breadcrumbs)

# Route for analyzing a specific file
@app.route('/analyze/<filename>')
def analyze(filename):
    breadcrumbs = [("Home", "/"), ("Descriptive Statistics", "/descriptive-statistics"), ("View Statistics", "/")]
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(file_path):
        flash("File does not exist.", 'danger')
        return redirect(url_for('descriptive_statistics'))
    results = analyze_file(file_path)
    return render_template('view_statistics.html', filename=filename, results= results,breadcrumbs=breadcrumbs)

# Route for network visualizer
@app.route('/network-visualiser')
def network_visualiser():
    breadcrumbs = [("Home", "/"), ("Network Visualiser", "/network-visualiser")]
    return render_template('network_visualiser.html', breadcrumbs=breadcrumbs)

# Route for serving static JSON files
@app.route('/json_objects/<path:filename>')
def serve_static(filename):
    return send_from_directory(app.config['JSON_FOLDER'], filename)

# Route for report generator
@app.route('/report-generator')
def report_generator():
    breadcrumbs = [("Home", "/"), ("Report Generator", "/report-generator")]
    files = []
    for filename in os.listdir(app.config['UPLOAD_FOLDER']):
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        stats = os.stat(filepath)
        files.append({
            'name': filename,
        })
    return render_template('report_generator.html', breadcrumbs=breadcrumbs)

# Main function to run the application
if __name__ == '__main__':
    app.run(debug=True)
