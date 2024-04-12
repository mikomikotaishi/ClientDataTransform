import pandas as pd
from typing import Any, Dict, List, Union

COLUMN_NAMES_A: List[str] = [ "Status", "Payee Date of Birth", "Spouse Date of Birth", "Payee Gender", "Spouse Gender", "Province of Residence", # Columns A through F
                              "Postal Code", "Original Member's Date of Retirement", "Original Member's Date of Death", "Lifetime Monthly Pension", # Columns G through J
                              "Original Guarantee (years)", "Date Guarantee End", "Unlocated Member (Y/N)", "Surname", "Given Name", "Spouse Surname", # Columns K through P
                              "Spouse Given Name", "Marital Status", "Beneficiary Surname", "Beneficiary Given Name" ] # Columns Q through T
COLUMN_NAMES_B: List[str] = [ "Status", "Member DOB", "Spouse DOB", "Gender (M=1, F=2)", "Spouse Sex", "Postal Code", # Columns A through F
                              "DOR", "Date of Death", "Pension", "Member Name", "Spouse Name", "Beneficiary Name", # Columns G through L
                              "Marital Status", "Original Guarantee (years)", "Date Guarantee End", "Unlocated Member (Y/N)" ] # Columns M through P
RESULT_COLUMNS: List[str] = [ "ID", "Status", "Member DOB", "Missing DOB?", "Spouse Date of Birth", "Missing Spouse DOB?", # Columns A through F
                              "Payee Gender", "Member Gender Anomaly", "Spouse Gender", "Spouse Gender Anomaly", "Gender Mismatch", # Columns G through K
                              "Province of Residence", "Postal Code", "Postal Code Check", "Original Member's Date of Retirement", # Columns L through O
                              "Original Member's Date of Death", "Lifetime Monthly Pension", "Pension Amount Check", "Original Guarantee (years)", # Columns P through S
                              "Date Guarantee End", "Guarantee Check", "Unlocated Member (Y/N)", "Unlocated Check", "Member Name", "Member Name Check", # Columns T through Y
                              "Spouse Name", "Spouse Name Check", "Marital Status", "Beneficiary Name", "Beneficiary Name Check" ] # Columns Z through AD

def process_excel(input_file: str, file_type: str) -> pd.DataFrame:
    """
    Process the data from an Excel file based on the specified file type.

    Args:
        input_file (str): Name of the input Excel file
        file_type (str): Type of the Excel file (A or B)

    Returns:
        pd.DataFrame: Processed data as a DataFrame
    """
    # Validate the file type
    if file_type not in ['A', 'B']:
        raise ValueError("Invalid file type specified")

    identifier_counter: int = 1

    # Load the Excel file as a DataFrame
    df: pd.DataFrame = pd.read_excel(input_file, header = 0)

    # Create an empty DataFrame to store the processed data
    processed_data: pd.DataFrame = pd.DataFrame(columns = RESULT_COLUMNS)

    # Iterate through rows in the input DataFrame
    for _, row in df.iterrows():
        if identifier_counter > 9999999999:
            raise ValueError("Too many rows in the input file")
        # Process the row
        processed_row: Dict[str, Any] = process_data(row, file_type, identifier_counter)
        processed_data = pd.concat([processed_data, pd.DataFrame([processed_row])], ignore_index = True)
        identifier_counter += 1

    return processed_data

def validate_postal_code(postal_code: str) -> bool:
    """
    Validate the format of a postal code.
    Requires that the postal code is in the format "A1A 1A1".

    Args:
        postal_code (str): Postal code to validate

    Returns:
        bool: True if the postal code is valid, False otherwise
    """
    return (len(postal_code) == 7 and postal_code[0].isalpha() and postal_code[1].isdigit() and postal_code[2].isalpha() 
           and postal_code[3].isspace() and postal_code[4].isdigit() and postal_code[5].isalpha() and postal_code[6].isdigit())

def process_data(row: pd.Series, file_type: str, identifier_counter: int) -> Dict[Union[str, pd.Timestamp], Any]:
    """
    Process a row of data from an Excel file based on the specified file type.

    Args:
        row (pd.Series): Row of data from the Excel file
        file_type (str): Type of the Excel file (A or B)
        identifier_counter (int): Counter for generating unique identifiers

    Returns:
        Dict[str, Any]: Processed data as a dictionary
    """
    processed_row: Dict[str, Any] = {}
    column_names = COLUMN_NAMES_A if file_type == 'A' else COLUMN_NAMES_B
    
    # Process the row based on the file type   
    processed_row["ID"] = str(identifier_counter).zfill(10)
    processed_row["Status"] = "Unspecified" if pd.isnull(row[column_names.index("Status")]) else row[column_names.index("Status")]
    processed_row["Member DOB"] = row[column_names.index("Payee Date of Birth")] if file_type == 'A' else row[column_names.index("Member DOB")]
    processed_row["Missing DOB?"] = "Yes" if pd.isnull(processed_row["Member DOB"]) else "No"
    processed_row["Spouse Date of Birth"] = row[column_names.index("Spouse Date of Birth")] if file_type == 'A' else row[column_names.index("Spouse DOB")]
    processed_row["Missing Spouse DOB?"] = "Yes" if pd.isnull(processed_row["Spouse Date of Birth"]) else "No"
    # Process Payee Gender column
    if file_type == 'A':
        processed_row["Payee Gender"] = row[column_names.index("Payee Gender")]
    else:
        match row[column_names.index("Gender (M=1, F=2)")]:
            case 1:
                processed_row["Payee Gender"] = "M"
            case 2:
                processed_row["Payee Gender"] = "F"
            case _:
                # Error-checking case
                processed_row["Payee Gender"] = "X"
    # Process Member Gender Anomaly column
    if (pd.isnull(processed_row["Payee Gender"]) or len(processed_row["Payee Gender"]) == 0):
        processed_row["Member Gender Anomaly"] = "Missing gender"
    elif processed_row["Payee Gender"] not in ["M", "F"]:
        processed_row["Member Gender Anomaly"] = "Incorrect gender code"
    else:
        processed_row["Member Gender Anomaly"] = "Good"
    processed_row["Spouse Gender"] = row[column_names.index("Spouse Sex")] if file_type == 'B' else row[column_names.index("Spouse Gender")]
    # Process Spouse Gender Anomaly column
    if pd.isnull(processed_row["Spouse Gender"]) and row[column_names.index("Marital Status")] == "Yes":
        processed_row["Spouse Gender Anomaly"] = "Missing gender"
    elif processed_row["Spouse Gender"] not in ["M", "F"] and row[column_names.index("Marital Status")] == "Yes":
        processed_row["Spouse Gender Anomaly"] = "Incorrect gender code"
    else:
        processed_row["Spouse Gender Anomaly"] = "Good"
    processed_row["Gender Mismatch"] = "Yes" if (processed_row["Payee Gender"] == processed_row["Spouse Gender"]) else "No"
    processed_row["Province of Residence"] = row[column_names.index("Province of Residence")] if file_type == 'A' else "Unspecified" # Field not provided in type B
    processed_row["Postal Code"] = row[column_names.index("Postal Code")]
    processed_row["Postal Code Check"] = "Correct" if validate_postal_code(processed_row["Postal Code"]) else "Incorrect code"
    processed_row["Original Member's Date of Retirement"] = row[column_names.index("Original Member's Date of Retirement")] if file_type == 'A' else row[column_names.index("DOR")]
    processed_row["Original Member's Date of Death"] = row[column_names.index("Original Member's Date of Death")] if file_type == 'A' else row[column_names.index("Date of Death")]
    processed_row["Lifetime Monthly Pension"] = row[column_names.index("Lifetime Monthly Pension")] if file_type == 'A' else row[column_names.index("Pension")]
    processed_row["Pension Amount Check"] = "Correct" if processed_row["Lifetime Monthly Pension"] > 0 else "Incorrect amount"
    processed_row["Original Guarantee (years)"] = row[column_names.index("Original Guarantee (years)")]
    processed_row["Date Guarantee End"] = row[column_names.index("Date Guarantee End")]
    # Process Guarantee Check column
    if pd.isnull(processed_row["Date Guarantee End"]) and processed_row["Original Guarantee (years)"] != "":
        processed_row["Guarantee Check"] = "No"
    else:
        processed_row["Guarantee Check"] = "Yes"
    processed_row["Unlocated Member (Y/N)"] = row[column_names.index("Unlocated Member (Y/N)")]
    processed_row["Unlocated Check"] = "Correct" if pd.isnull(processed_row["Unlocated Member (Y/N)"]) else "Investigate" if processed_row["Unlocated Member (Y/N)"] == "Y" else "Correct"
    # Process Member Name column
    if file_type == 'A':
        if pd.isnull(row[column_names.index("Given Name")]) and pd.isnull(row[column_names.index("Surname")]):
            processed_row["Member Name"] = ""
        elif pd.isnull(row[column_names.index("Given Name")]):
            processed_row["Member Name"] = row[column_names.index("Surname")]
        elif pd.isnull(row[column_names.index("Surname")]):
            processed_row["Member Name"] = row[column_names.index("Given Name")]
        else:
            processed_row["Member Name"] = row[column_names.index("Given Name")] + " " + row[column_names.index("Surname")]
    else:
        processed_row["Member Name"] = row[column_names.index("Member Name")]
    processed_row["Member Name Check"] = "Yes" if len(processed_row["Member Name"]) > 0 else "No"
    # Process Spouse Name column
    if file_type == 'A':
        if pd.isnull(row[column_names.index("Spouse Given Name")]) and pd.isnull(row[column_names.index("Spouse Surname")]):
            processed_row["Spouse Name"] = ""
        elif pd.isnull(row[column_names.index("Spouse Given Name")]):
            processed_row["Spouse Name"] = row[column_names.index("Spouse Surname")]
        elif pd.isnull(row[column_names.index("Spouse Surname")]):
            processed_row["Spouse Name"] = row[column_names.index("Spouse Given Name")]
        else:
            processed_row["Spouse Name"] = row[column_names.index("Spouse Given Name")] + " " + row[column_names.index("Spouse Surname")]
    else:
        processed_row["Spouse Name"] = row[column_names.index("Spouse Name")]
    processed_row["Spouse Name Check"] = "No" if (pd.isnull(processed_row["Spouse Name"]) and row[column_names.index("Marital Status")] == "Yes") else "Yes"
    processed_row["Marital Status"] = row[column_names.index("Marital Status")]
    # Process Beneficiary Name column
    if file_type == 'A':
        if pd.isnull(row[column_names.index("Beneficiary Given Name")]) and pd.isnull(row[column_names.index("Beneficiary Surname")]):
            processed_row["Beneficiary Name"] = ""
        elif pd.isnull(row[column_names.index("Beneficiary Given Name")]):
            processed_row["Beneficiary Name"] = row[column_names.index("Beneficiary Surname")]
        elif pd.isnull(row[column_names.index("Beneficiary Surname")]):
            processed_row["Beneficiary Name"] = row[column_names.index("Beneficiary Given Name")]
        else:
            processed_row["Beneficiary Name"] = row[column_names.index("Beneficiary Given Name")] + " " + row[column_names.index("Beneficiary Surname")]
    else:
        processed_row["Beneficiary Name"] = row[column_names.index("Beneficiary Name")]
    processed_row["Beneficiary Name Check"] = "No" if ((pd.isnull(processed_row["Beneficiary Name"]) or len(processed_row["Beneficiary Name"]) == 0) and processed_row["Status"] == "Beneficiary") else "Yes"

    # Return the processed row
    return processed_row
