import os
import shutil
import pandas as pd
from fuzzywuzzy import process
import logging

def get_all_files(directory):
    """Recursively get all files in the directory and its subdirectories."""
    all_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            all_files.append(os.path.join(root, file))
    return all_files

def match_files(excel_file, source_dir, destination_dir, log_file):
    # Set up logging
    logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    # Load the Excel file
    df = pd.read_excel(excel_file)
    
    # Assuming the first column is "No." and the second is "Name"
    names_to_find = df['Name'].tolist()

    matched_count = 0
    total_files = len(names_to_find)
    unmatched_files = []
    score_threshold = 75

    logging.info("Starting file matching process...")
    
    # Get all files in the source directory and subdirectories
    all_files = get_all_files(source_dir)
    
    # Extract just the file names without extensions for matching
    all_file_names = [os.path.splitext(os.path.basename(f))[0] for f in all_files]
    
    # Table header
    print(f"{'No.':<5} {'File Name':<30} {'Matched File':<30} {'Status':<10} {'Score':<6}")
    print("-" * 80)
    
    # Iterate over the names in the Excel file with an ordered list format
    for index, name in enumerate(names_to_find, start=1):
        # Find the best match in the directory
        match, score = process.extractOne(name, all_file_names)
        
        # Determine status (PASSED or FAILED)
        status = 'PASSED' if score > score_threshold else 'FAILED'
        
        # Log and show the match in the terminal with justified formatting
        log_message = f"{str(index) + '.':<5} {name:<30} {match:<30} {status:<10} {str(score):<6}"
        print(log_message)
        logging.info(log_message)
        
        # If the score is above a certain threshold, consider it a match
        if score > score_threshold:
            # Find the actual file path that matches and copy it to the destination folder
            matched_file_path = next(f for f in all_files if os.path.splitext(os.path.basename(f))[0] == match)
            shutil.copy(matched_file_path, destination_dir)
            matched_count += 1
        else:
            unmatched_files.append(name)

    # Show and log the final score and result with indentation
    final_score_message = "\nFinal Score: {}/{}".format(matched_count, total_files)
    print(final_score_message)
    logging.info(final_score_message)
    
    if matched_count == total_files:
        result_message = "Overall Result: PASSED"
    else:
        result_message = "Overall Result: FAILED"
        unmatched_message = "\nUnmatched Files:"
        unmatched_list = "\n".join([f"    {index}. {file}" for index, file in enumerate(unmatched_files, start=1)])
        result_message += unmatched_message + "\n" + unmatched_list

    print(result_message)
    logging.info(result_message)

if __name__ == "__main__":
    excel_file = "FileList.xlsx"
    source_dir = "/Users/yaribhernandez/Documents/CodeExamples/FileSorting/Files"
    destination_dir = "/Users/yaribhernandez/Documents/CodeExamples/FileSorting/Results"
    log_file = "file_matching.log"
    
    match_files(excel_file, source_dir, destination_dir, log_file)