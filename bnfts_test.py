import os
import re
import pandas as pd
from pathlib import Path

def parse_filename(filename):
    """Extract initial portion and version number from filename."""
    # Skip files with (1) in the name
    if "(1)" in filename:
        return None, None
        
    match = re.match(r'(\d{4}-\d+)\.(\d+)\.txt$', filename)
    if match:
        return match.group(1), float(match.group(2))
    return None, None

def get_latest_versions(folder_path):
    """Find the latest version of each initial portion."""
    files_dict = {}
    
    for filename in os.listdir(folder_path):
        initial_portion, version = parse_filename(filename)
        if initial_portion:
            if initial_portion not in files_dict or version > files_dict[initial_portion][1]:
                files_dict[initial_portion] = (filename, version)
    
    return [file_info[0] for file_info in files_dict.values()]

def extract_benefits_text(file_content):
    """Extract benefits text from file content."""
    # Define start patterns
    start_patterns = [
        r"ALL OCCUPATIONS LISTED ABOVE RECEIVE THE FOLLOWING BENEFITS:",
        r"ALL OCCUPATIONS LISTED ABOVE RECIEVE THE FOLLOWING BENEFITS:"
    ]
    
    # Define end patterns
    end_patterns = [
        r"THE OCCUPATIONS WHICH HAVE NUMBERED",
        r"\*\* HAZARDOUS PAY DIFFERENTIAL \*\*"
    ]
    
    # Find start of benefits text
    start_pos = -1
    for pattern in start_patterns:
        match = re.search(pattern, file_content)
        if match:
            start_pos = match.end()
            break
    
    if start_pos == -1:
        return None
    
    # Find end of benefits text
    end_pos = len(file_content)
    for pattern in end_patterns:
        match = re.search(pattern, file_content[start_pos:])
        if match:
            end_pos = start_pos + match.start()
            break
    
    benefits_text = file_content[start_pos:end_pos].strip()
    
    # Extract vacation section
    vacation_match = re.search(r'VACATION:(.*?)HOLIDAYS:', benefits_text, re.DOTALL)
    vacation_text = vacation_match.group(1).strip() if vacation_match else ""
    
    # Extract holidays section
    holidays_match = re.search(r'HOLIDAYS:(.*?)$', benefits_text, re.DOTALL)
    holidays_text = holidays_match.group(1).strip() if holidays_match else ""
    
    return benefits_text, vacation_text, holidays_text

def process_files(input_folder_path, output_folder_path, output_filename, max_cell_length=32000):
    """Process all files and create Excel output."""
    # Create output folder if it doesn't exist
    os.makedirs(output_folder_path, exist_ok=True)
    
    # Construct full output path
    output_file_path = os.path.join(output_folder_path, output_filename)
    
    latest_files = get_latest_versions(input_folder_path)
    results = []
    
    for filename in latest_files:
        file_path = os.path.join(input_folder_path, filename)
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
                extracted_texts = extract_benefits_text(content)
                
                if extracted_texts and extracted_texts[0]:  # If benefits text was found
                    benefits_text, vacation_text, holidays_text = extracted_texts
                    
                    # Split texts into chunks if needed
                    benefits_chunks = [benefits_text[i:i+max_cell_length] 
                                    for i in range(0, len(benefits_text), max_cell_length)]
                    vacation_chunks = [vacation_text[i:i+max_cell_length] 
                                    for i in range(0, len(vacation_text), max_cell_length)]
                    holidays_chunks = [holidays_text[i:i+max_cell_length] 
                                    for i in range(0, len(holidays_text), max_cell_length)]
                    
                    # Create row with filename and all text chunks
                    row = ([filename[:-4]] + 
                          benefits_chunks + 
                          vacation_chunks + 
                          holidays_chunks)
                    results.append(row)
        except Exception as e:
            print(f"Error processing {filename}: {str(e)}")
    
    if not results:
        print("No files were processed successfully.")
        return 0
        
    # Find maximum number of chunks for each section
    max_benefits = max(len([col for col in row[1:] if col]) for row in results)
    
    # Create column names
    columns = ['Filename']
    columns.extend([f'Benefits_Text_{i+1}' for i in range(max_benefits)])
    columns.extend(['Vacation_Section'])
    columns.extend(['Holidays_Section'])
    
    # Create DataFrame with all columns
    df = pd.DataFrame(results, columns=columns)
    
    # Save to Excel
    df.to_excel(output_file_path, index=False)
    
    return len(results)

def main():
    # Replace these paths with your actual paths
    input_folder_path = "path/to/input/folder"  # Where your text files are
    output_folder_path = "path/to/output/folder"  # Where you want to save the Excel file
    output_filename = "benefits_output.xlsx"  # Name of the output Excel file
    
    try:
        processed_files = process_files(input_folder_path, output_folder_path, output_filename)
        print(f"Successfully processed {processed_files} files.")
        print(f"Output saved to {os.path.join(output_folder_path, output_filename)}")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
