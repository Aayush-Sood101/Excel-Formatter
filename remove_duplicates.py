"""
Script to remove consecutive duplicate rows from an Excel file.
Duplicates are identified based on Column B (Author) and Column C (Title).
"""

import pandas as pd

def remove_consecutive_duplicates(input_file, output_file):
    """
    Remove consecutive duplicate rows from an Excel file based on columns B and C.
    
    Args:
        input_file (str): Path to the input Excel file
        output_file (str): Path to the output Excel file
    """
    
    # Load the Excel file
    print(f"Reading data from {input_file}...")
    df = pd.read_excel(input_file)
    
    # Display initial row count
    initial_rows = len(df)
    print(f"Initial number of rows (excluding header): {initial_rows}")
    
    # Get column names for the 2nd (index 1) and 3rd (index 2) columns
    # This ensures the script works regardless of the actual column names
    if len(df.columns) < 3:
        print("Warning: The file has fewer than 3 columns. Cannot proceed.")
        return
    
    col_b = df.columns[1]  # Column B (2nd column, index 1)
    col_c = df.columns[2]  # Column C (3rd column, index 2)
    
    print(f"Comparing consecutive rows based on columns: '{col_b}' and '{col_c}'")
    
    # Create a boolean mask to identify rows that are NOT consecutive duplicates
    # A row is kept if it's different from the previous row in either Column B or Column C
    # The first row is always kept (no previous row to compare)
    mask = (df[col_b] != df[col_b].shift(1)) | (df[col_c] != df[col_c].shift(1))
    
    # Apply the mask to keep only non-duplicate rows
    df_cleaned = df[mask].copy()
    
    # Display final row count
    final_rows = len(df_cleaned)
    removed_rows = initial_rows - final_rows
    print(f"Removed {removed_rows} consecutive duplicate row(s)")
    print(f"Final number of rows (excluding header): {final_rows}")
    
    # Save the cleaned data to a new Excel file
    print(f"Saving cleaned data to {output_file}...")
    df_cleaned.to_excel(output_file, index=False)
    
    print(f"Processing complete. Cleaned data saved to {output_file}")


if __name__ == "__main__":
    # Define input and output file names
    INPUT_FILE = "input.xlsx"
    OUTPUT_FILE = "output.xlsx"
    
    try:
        # Run the duplicate removal process
        remove_consecutive_duplicates(INPUT_FILE, OUTPUT_FILE)
    except FileNotFoundError:
        print(f"Error: {INPUT_FILE} not found in the current directory.")
        print("Please ensure the input file exists and try again.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
