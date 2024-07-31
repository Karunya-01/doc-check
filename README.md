from docx import Document as DocxDocument
from spire.doc import Document as SpireDocument
from spire.doc.common import *
import re
import pandas as pd
import requests
from datetime import datetime, timedelta
import pytz
from tabulate import tabulate
from textwrap import fill
from nltk import word_tokenize, pos_tag

def read_tables_from_docx(file_path):
    doc = DocxDocument(file_path)
    all_table_data = []

    # Read all tables in the document
    for table in doc.tables:
        table_data = []
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            table_data.append(row_data)
        all_table_data.append(table_data)

    return all_table_data

def extract_tables_from_docx(docx_file):
    doc = docx.Document(docx_file)
    tables_data = []
    for table in doc.tables:
        table_data = []
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            table_data.append(row_data)
            # Create DataFrame from the entire table data
            df = pd.DataFrame(table_data)
            # Treat the first row as the header and reset the index
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)
            # Handle the first column (numbering) explicitly as string
            numbering_column = df.columns[0]
            df[numbering_column] = df[numbering_column].astype(str)
            # Append DataFrame to the list
            tables_data.append(df)

            return tables_data

def process_all_tables(all_table_data, value_list):
    total_rows = sum(len(table) - 1 for table in all_table_data)  # Exclude header rows from each table
    num_values = len(value_list)

    if total_rows != num_values:
        return "The number of values does not match the total number of rows across all tables"

    current_value_index = 0

    for table_data in all_table_data:
        header = table_data[0]
        if "Step #" in header:
            step_index = header.index("Step #")
        else:
            return "No 'Step #' column found in one of the tables"

        for i in range(1, len(table_data)):
            table_data[i][step_index] = value_list[current_value_index]
            current_value_index += 1

    return all_table_data

def extract_sorted_unique_versions(file_path):
    doc = SpireDocument()
    doc.LoadFromFile(file_path)
    text = doc.GetText()
    pattern = r"\b\d+\.\d+\.\d+\.\d+\b"
    matches = re.findall(pattern, text)
    unique_matches = sorted(set(matches), key=lambda s: list(map(int, s.split('.'))))
    return unique_matches

def check_hyperlink_accessibility(url):
    try:
        response = requests.head(url, allow_redirects=True, timeout=5)
        if response.status_code == 200:
            return True
        else:
            return False
    except requests.RequestException:
        return False



def check_screenshot_and_attachment1(df):
    results = []
    for test_procedure, actual_result in zip(df["Test Procedure"], df["Actual Result"]):
        screenshot_present = "CAPTURE SCREENSHOT(S)" in test_procedure
        attachment_present = "upload attachment" in actual_result.lower()
        hyperlink_match = re.search(r'(https?://\S+|www\.\S+)', actual_result)
        if attachment_present and hyperlink_match:
            result_message = "has attachment with a hyperlink."
        elif attachment_present and not hyperlink_match:
            result_message = "has attachment without a hyperlink."
        else:
            result_message = "has no attachment."

        results.append(result_message)
    return results


def check_screenshot_and_attachment(df):
    step_column = df["Requirement Reference"]
    test_procedure_column = df["Test Procedure"]
    actual_result_column = df["Actual Result"]
    results = []

    for step, test_procedure, actual_result in zip(step_column, test_procedure_column, actual_result_column):
        screenshot_present = "CAPTURE SCREENSHOT(S)" in test_procedure
        attachment_present = "upload attachment" in actual_result.lower()
        hyperlink_present = False
        clickable_hyperlink = False
        hyperlink_accessible = False
        
        if attachment_present:
            # Find hyperlink following "upload attachment"
            upload_attachment_index = actual_result.lower().find("upload attachment")
            hyperlink_match = re.search(r"(http[s]?://\S+)", actual_result[upload_attachment_index:])
            if hyperlink_match:
                hyperlink_present = True
                url = hyperlink_match.group(0)
                # Check if the hyperlink is clickable
                clickable_hyperlink = url.startswith(("http://", "https://"))
                if clickable_hyperlink:
                    hyperlink_accessible = check_hyperlink_accessibility(url)
        
        if attachment_present and hyperlink_present and clickable_hyperlink and hyperlink_accessible:
            results.append(f"{step}-- has attachment with accessible clickable hyperlink.")
        elif attachment_present and hyperlink_present and clickable_hyperlink and not hyperlink_accessible:
            results.append(f"{step}-- has attachment with hyperlink.")
        elif attachment_present and (not hyperlink_present or not clickable_hyperlink):
            results.append(f"{step}--  has 'upload attachment' but no clickable hyperlink.")
        else:
            results.append(f"{step}--  has no attachment.")

    return results
def extract_timestamp_and_step(df):
    step_column = df["Step #"]
    executed_column = df["Executed By & Date"]
    timestamp_pattern = r"\d{2}-[a-zA-Z]{3}-\d{4} \d{2}:\d{2}:\d{2} [AP]M \([A-Z]+\)"
    timestamps_and_steps = []
    prev_timestamp = None
    ist = pytz.timezone('Asia/Kolkata')
    for step, executed_str in zip(step_column, executed_column):
        timestamp_match = re.search(timestamp_pattern, executed_str)
        if timestamp_match:
            current_timestamp = timestamp_match.group(0)
            if prev_timestamp:
                try:
                    current_time = datetime.strptime(re.sub(r" \([A-Z]+\)", "", current_timestamp), "%d-%b-%Y %I:%M:%S %p")
                    current_time_ist = ist.localize(current_time)
                    
                    prev_time = datetime.strptime(re.sub(r" \([A-Z]+\)", "", prev_timestamp), "%d-%b-%Y %I:%M:%S %p")
                    prev_time_ist = ist.localize(prev_time)
                    
                    time_diff = current_time_ist - prev_time_ist
                    
                    if time_diff > timedelta(hours=1):
                        timestamps_and_steps.append(f"{step} - Time difference with previous step exceeds 1 hour")
                except ValueError as e:
                    timestamps_and_steps.append(f"Timestamp format error: {e}")
                   
            prev_timestamp = current_timestamp
    
    return timestamps_and_steps


def wrap_text(text, width):
    """Wrap text to a specified width."""
    return fill(text, width=width)

def check_step_results(df):
    step_column = df["Step #"]
    expected_column = df["Expected Result"]
    actual_column = df["Actual Result"]
    pass_fail_column = df["Pass/Fail"]

    results = []

    for step, expected, actual, result in zip(step_column, expected_column, actual_column, pass_fail_column):
        if result.lower() == "fail":
            wrapped_expected = wrap_text(expected, 20)
            wrapped_actual = wrap_text(actual, 20)
            results.append((step, "Fail", wrapped_expected, wrapped_actual))

    # Define table headers
    headers = ["Step #", "Status", "Expected Result", "Actual Result"]

    if results:
        table = tabulate(results, headers=headers, tablefmt="grid")
        print(table)
    else:
        print("No Fail on this Table")

def is_present_tense(sentence):
    tokens = word_tokenize(sentence)
    tags = pos_tag(tokens)
    for word, tag in tags:
        if tag in {"VB", "VBP", "VBZ"}:
            return True
    return False 

def check_actuals_present_tense(df):
    
    results = []
    for step,actual in zip(df["Step #"],df["Actual Result"]):
           tense_status = "Present Tense"
           if actual:  # Check if there is a sentence
            tense_status = "Present Tense" if is_present_tense(actual) else "Not Present Tense"
            results.append(f"{step}-{tense_status}")
           else:
               results.append(f"{step}-No sentence")
    return results



file_path = "Sample Test case.docx"

sorted_unique_versions = extract_sorted_unique_versions(file_path)

all_table_data = read_tables_from_docx(file_path)

processed_table_data = process_all_tables(all_table_data, sorted_unique_versions)
#for i, table_data in enumerate(processed_table_data):
    #print(f"Table {i+1}:")
 
    #df = pd.DataFrame(table_data[1:], columns=table_data[0])
    #timestamp_and_step_details = extract_timestamp_and_step(df)
    #print("----------------Timestamp Details:--------------")
   
    #for detail in timestamp_and_step_details:
        #print(detail)
        
    

    
    #print("------------------Pass/Fail Details:---------------------")
    #check_step_results(df)

    #Tense_details = check_actuals_present_tense(df)
    #print("-------------Tense Details:-------------")
    #for detail in Tense_details:
    #    print(detail)

tables = extract_tables_from_docx(file_path)

for i, table in enumerate(tables):
    print(f"Table {i+1}:")
    
    hyperlink_details = check_screenshot_and_attachment(table)
    print("----------------Hyperlink Details:-------------")
    for detail in hyperlink_details:
        print(detail)

    
   

