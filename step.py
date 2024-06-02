import docx
import pandas as pd
import re
from datetime import datetime, timedelta
import pytz
import requests

def extract_tables_from_docx(docx_file):
    doc = docx.Document(docx_file)
    tables_data = []
    for table in doc.tables:
        table_data = []
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            table_data.append(row_data)
        df = pd.DataFrame(table_data[1:], columns=table_data[0])
        tables_data.append(df)
    return tables_data

def extract_timestamp_and_step(df):
    step_column = df["Step #"]
    executed_column = df["Executed By & Date"]
    timestamps_and_steps = []
    timestamp_pattern = r"\d{2}-[a-zA-Z]{3}-\d{4} \d{2}:\d{2}:\d{2} [AP]M \([A-Z]+\)"
    prev_timestamp = None

    ist = pytz.timezone('Asia/Kolkata')

    for step, executed_str in zip(step_column, executed_column):
        new_value_match = re.search(r"New Value\s*(.*?)\s+", executed_str)
        old_value_match = re.search(r"Old Value\s*(.*?)\s+", executed_str)
        timestamp_match = re.search(timestamp_pattern, executed_str)

        new_value_timestamp = None
        old_value_timestamp = None

        if new_value_match and old_value_match:
            new_timestamp_match = re.search(timestamp_pattern, executed_str[new_value_match.end():])
            new_value_timestamp = new_timestamp_match.group(0) if new_timestamp_match else "Timestamp not found"

            old_timestamp_match = re.search(timestamp_pattern, executed_str[old_value_match.end():])
            old_value_timestamp = old_timestamp_match.group(0) if old_timestamp_match else "Timestamp not found"

            timestamps_and_steps.append(f"{step}--New Value: {new_value_timestamp}\n{step}--Old Value: {old_value_timestamp}")
            current_timestamp = old_value_timestamp
        else:
            timestamp = timestamp_match.group(0) if timestamp_match else "Timestamp not found"
            timestamps_and_steps.append(f"{step}--{timestamp}")
            current_timestamp = timestamp

        if current_timestamp != "Timestamp not found":
            if prev_timestamp:
                try:
                    current_time = datetime.strptime(re.sub(r" \([A-Z]+\)", "", current_timestamp), "%d-%b-%Y %I:%M:%S %p")
                    current_time_ist = ist.localize(current_time)
                    prev_time = datetime.strptime(re.sub(r" \([A-Z]+\)", "", prev_timestamp), "%d-%b-%Y %I:%M:%S %p")
                    prev_time_ist = ist.localize(prev_time)
                    time_diff = current_time_ist - prev_time_ist
                    if time_diff < timedelta(hours=1):
                        timestamps_and_steps.append(f"Time difference with previous step: {time_diff}")
                    else:
                        timestamps_and_steps.append("Time difference with previous step: More than 1 hour")
                except ValueError as e:
                    timestamps_and_steps.append(f"Timestamp format error: {e}")
            else:
                timestamps_and_steps.append("No previous timestamp to compare.")
            prev_timestamp = current_timestamp

    return timestamps_and_steps

def check_hyperlink_accessibility(url):
    try:
        response = requests.head(url, allow_redirects=True, timeout=5)
        if response.status_code == 200:
            return True
        else:
            return False
    except requests.RequestException:
        return False

def check_screenshot_and_attachment(df):
    step_column = df["Step #"]
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
            results.append(f"{step}-- has attachment with clickable hyperlink that is not accessible.")
        elif attachment_present and (not hyperlink_present or not clickable_hyperlink):
            results.append(f"{step}--  has 'upload attachment' but no clickable hyperlink.")
        else:
            results.append(f"{step}--  has no attachment.")

    return results

# Example usage
docx_file = "Step.docx"
tables = extract_tables_from_docx(docx_file)

for i, table in enumerate(tables):
    print(f"Table {i+1}:")
    if "Executed By & Date" in table.columns and "Step #" in table.columns:
        timestamps_and_steps = extract_timestamp_and_step(table)
        for timestamp_and_step in timestamps_and_steps:
            print(timestamp_and_step)
    else:
        print("No 'Executed By & Date' or 'Step #' column found.")
    
    if "Test Procedure" in table.columns and "Actual Result" in table.columns:
        screenshot_and_attachment_results = check_screenshot_and_attachment(table)
        for result in screenshot_and_attachment_results:
            print(result)
    
    print()
