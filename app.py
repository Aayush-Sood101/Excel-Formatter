"""
Flask web application to remove consecutive duplicate rows from Excel files.
Duplicates are identified based on Column B (Author), Column C (Title), and Column G.
"""

from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import pandas as pd
import os
from werkzeug.utils import secure_filename
import io
import gc  # Garbage collection for memory management
import psutil  # For memory monitoring
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this-in-production'  # Change this for security
app.config['MAX_CONTENT_LENGTH'] = 8 * 1024 * 1024  # Reduced to 8MB max file size for better memory management

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    """Check if the uploaded file has an allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def check_memory_usage():
    """Check current memory usage and log it."""
    try:
        process = psutil.Process(os.getpid())
        memory_info = process.memory_info()
        memory_mb = memory_info.rss / 1024 / 1024
        logger.info(f"Current memory usage: {memory_mb:.2f} MB")
        return memory_mb
    except Exception as e:
        logger.error(f"Error checking memory: {e}")
        return 0

def remove_consecutive_duplicates(file_stream):
    """
    Remove consecutive duplicate rows from an Excel file based on columns B, C, and G.
    Memory-optimized version for deployment environments with chunked processing.
    
    Args:
        file_stream: File stream of the uploaded Excel file
        
    Returns:
        tuple: (cleaned DataFrame, stats dictionary)
    """
    try:
        logger.info("Starting file processing...")
        check_memory_usage()
        
        # Load the Excel file from the stream with memory optimization
        df = pd.read_excel(file_stream, engine='openpyxl')
        
        logger.info(f"Loaded DataFrame with {len(df)} rows and {len(df.columns)} columns")
        check_memory_usage()
        
        # Store initial row count
        initial_rows = len(df)
        
        # Check if the file is too large for the current memory constraints
        if initial_rows > 50000:  # Adjust this threshold based on your needs
            raise ValueError(f"File too large ({initial_rows} rows). Please use a file with fewer than 50,000 rows for the free tier.")
        
        # Check if file has at least 7 columns (to access column G)
        if len(df.columns) < 7:
            raise ValueError("The file must have at least 7 columns to access Column G.")
        
        # Get column names for the 2nd (index 1), 3rd (index 2), and 7th (index 6) columns
        col_b = df.columns[1]  # Column B (2nd column)
        col_c = df.columns[2]  # Column C (3rd column)
        col_g = df.columns[6]  # Column G (7th column)
        
        # Memory optimization: convert to string type if not already and handle NaN values
        for col in [col_b, col_c, col_g]:
            df[col] = df[col].fillna('').astype(str)
        
        logger.info("Data types optimized")
        check_memory_usage()
        
        # Create a boolean mask to identify rows that are NOT consecutive duplicates
        # A row is kept if it's different from the previous row in any of Column B, C, or G
        mask = ((df[col_b] != df[col_b].shift(1)) | 
                (df[col_c] != df[col_c].shift(1)) | 
                (df[col_g] != df[col_g].shift(1)))
        
        # Apply the mask to keep only non-duplicate rows
        df_cleaned = df[mask].copy()
        
        logger.info("Duplicate removal completed")
        check_memory_usage()
        
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
        
        # Force garbage collection to free memory
        del df, mask
        gc.collect()
        
        logger.info("Memory cleanup completed")
        check_memory_usage()
        
        return df_cleaned, stats
        
    except Exception as e:
        # Force garbage collection on error
        logger.error(f"Error processing file: {str(e)}")
        gc.collect()
        raise e

@app.route('/')
def index():
    """Render the main upload page."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and processing."""
    # Log initial memory usage
    initial_memory = check_memory_usage()
    
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
        
        # Create output file in memory with optimization
        output = io.BytesIO()
        try:
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_cleaned.to_excel(writer, index=False)
            output.seek(0)
            
            # Flash success message with statistics
            flash(f'Success! Removed {stats["removed_rows"]} consecutive duplicate(s).', 'success')
            flash(f'Original rows: {stats["initial_rows"]} → Final rows: {stats["final_rows"]}', 'info')
            flash(f'Compared columns: "{stats["col_b"]}", "{stats["col_c"]}", and "{stats["col_g"]}"', 'info')
            
            # Clean up memory before sending file
            del df_cleaned
            gc.collect()
            
            # Send the cleaned file for download
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='output_cleaned.xlsx'
            )
        except Exception as writer_error:
            # Clean up on writer error
            output.close()
            del df_cleaned
            gc.collect()
            raise writer_error
        
    except ValueError as ve:
        gc.collect()  # Clean up memory on error
        flash(str(ve), 'error')
        return redirect(url_for('index'))
    except Exception as e:
        gc.collect()  # Clean up memory on error
        flash(f'An error occurred while processing the file: {str(e)}', 'error')
        return redirect(url_for('index'))

if __name__ == '__main__':
    # Use environment variable for port (required for Render)
    port = int(os.environ.get('PORT', 5001))
    app.run(host='0.0.0.0', port=port, debug=False)