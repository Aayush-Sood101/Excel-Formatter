"""
Flask web application to remove consecutive duplicate rows from Excel files.
Duplicates are identified based on Column B (Author), Column C (Title), and Column G.
"""

from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import pandas as pd
import os
from werkzeug.utils import secure_filename
import io

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this-in-production'  # Change this for security
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    """Check if the uploaded file has an allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def remove_consecutive_duplicates(file_stream):
    """
    Remove consecutive duplicate rows from an Excel file based on columns B, C, and G.
    
    Args:
        file_stream: File stream of the uploaded Excel file
        
    Returns:
        tuple: (cleaned DataFrame, stats dictionary)
    """
    # Load the Excel file from the stream
    df = pd.read_excel(file_stream)
    
    # Store initial row count
    initial_rows = len(df)
    
    # Check if file has at least 7 columns (to access column G)
    if len(df.columns) < 7:
        raise ValueError("The file must have at least 7 columns to access Column G.")
    
    # Get column names for the 2nd (index 1), 3rd (index 2), and 7th (index 6) columns
    col_b = df.columns[1]  # Column B (2nd column)
    col_c = df.columns[2]  # Column C (3rd column)
    col_g = df.columns[6]  # Column G (7th column)
    
    # Create a boolean mask to identify rows that are NOT consecutive duplicates
    # A row is kept if it's different from the previous row in any of Column B, C, or G
    mask = ((df[col_b] != df[col_b].shift(1)) | 
            (df[col_c] != df[col_c].shift(1)) | 
            (df[col_g] != df[col_g].shift(1)))
    
    # Apply the mask to keep only non-duplicate rows
    df_cleaned = df[mask].copy()
    
    # Calculate statistics
    final_rows = len(df_cleaned)
    removed_rows = initial_rows - final_rows
    
    stats = {
        'initial_rows': initial_rows,
        'final_rows': final_rows,
        'removed_rows': removed_rows,
        'col_b': col_b,
        'col_c': col_c,
        'col_g': col_g
    }
    
    return df_cleaned, stats

@app.route('/')
def index():
    """Render the main upload page."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and processing."""
    # Check if file was uploaded
    if 'file' not in request.files:
        flash('No file selected', 'error')
        return redirect(url_for('index'))
    
    file = request.files['file']
    
    # Check if filename is empty
    if file.filename == '':
        flash('No file selected', 'error')
        return redirect(url_for('index'))
    
    # Check if file type is allowed
    if not allowed_file(file.filename):
        flash('Invalid file type. Please upload an Excel file (.xlsx or .xls)', 'error')
        return redirect(url_for('index'))
    
    try:
        # Process the file
        df_cleaned, stats = remove_consecutive_duplicates(file)
        
        # Create output file in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_cleaned.to_excel(writer, index=False)
        output.seek(0)
        
        # Flash success message with statistics
        flash(f'Success! Removed {stats["removed_rows"]} consecutive duplicate(s).', 'success')
        flash(f'Original rows: {stats["initial_rows"]} → Final rows: {stats["final_rows"]}', 'info')
        flash(f'Compared columns: "{stats["col_b"]}", "{stats["col_c"]}", and "{stats["col_g"]}"', 'info')
        
        # Send the cleaned file for download
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='output_cleaned.xlsx'
        )
        
    except ValueError as ve:
        flash(str(ve), 'error')
        return redirect(url_for('index'))
    except Exception as e:
        flash(f'An error occurred while processing the file: {str(e)}', 'error')
        return redirect(url_for('index'))

if __name__ == '__main__':
    # Use environment variable for port (required for Render)
    port = int(os.environ.get('PORT', 5001))
    app.run(host='0.0.0.0', port=port, debug=False)