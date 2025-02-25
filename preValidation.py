# Rule 1 : Alpha Numeric Name is valid, Can't start with number, no special character is allowed except underscore
# Rule 2 : No NUll Entries in IB Variable Type
# Rule 3 : Duplicate entry 
import re
import logging
import pandas as pd  # Required if using Pandas for Excel reading

# for handler in logging.root.handlers[:]:
#     logging.root.removeHandler(handler)

# file_handler = logging.FileHandler("pre_validation_logs.log", mode = "a")
# file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))

logging.basicConfig(
    filename="pre_validation_logs1.log",
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s",
    filemode="a",
    force=True
)
# logger = logging.getLogger()

def pre_excel_rule1(file_path):
    """
    Checks if the project name in cell C4 of the 'project_info' sheet follows the naming convention.
    
    Rules:
    - Can be alphanumeric.
    - Cannot start with a number or special character.
    - Only underscore (_) is allowed as a special character.
    
    Returns:
    - (bool, str): True if valid, False otherwise with an error message.
    """
    try:
        # Read the full sheet
        sheet_name = "project_info"
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
 
        # Read value at C4 (Row 4, Column C -> index [3,2])
        project_name = str(df.iloc[3, 2]).strip()
 
        if project_name.lower() == "nan":  # Handle empty values
            logging.error("[ERROR] Project Name (C4) is empty or missing.")
            return False, "Project Name (C4) is empty or missing."
 
        # Validate naming convention
        if not re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", project_name):
            logging.error(f"[ERROR] Invalid project name '{project_name}' in C4.")
            return False, f"Invalid project name '{project_name}' in C4. Must be alphanumeric, start with a letter/underscore, and contain no special characters except '_'."
 
        logging.info(f"[SUCCESS] Project name '{project_name}' in C4 is valid.")
        return True, ""
    
    except Exception as ex:
        logging.exception(f"[ERROR] Rule_1 failed due to an exception: {ex}")
        return False, f"Critical Error: {ex}"
 
def is_empty(value):
    """
        Check if a value is empty (None, '', whitespace, or NaN).
    """
    return value is None or str(value).strip() == "" or (isinstance(value, float) and pd.isna(value))
 
def pre_ports_null_value_rule2(file_path, sheet_name):
    """
    Checks if any column in the given sheet contains an empty cell.
    
    Rules:
    - Every cell in every column must have a value.
    - If an empty cell is found, an error is raised.
 
    Returns:
    - (bool, str): True if valid, False otherwise with an error message.
    """
    try:
        # Read the sheet
        df = pd.read_excel(file_path, sheet_name=sheet_name)
 
        # Check for empty cells
        if df.isnull().values.any():
            empty_cells = df.isnull().sum().sum()
            logging.error(f"[ERROR] Found {empty_cells} empty cells in sheet '{sheet_name}'.")
            return False, f"Sheet '{sheet_name}' contains {empty_cells} empty cells. Fill all values."
 
        logging.info(f"[SUCCESS] No empty cells found in sheet '{sheet_name}'.")
        return True, ""
 
    except Exception as ex:
        logging.exception(f"[ERROR] Rule_2 failed due to an exception: {ex}")
        return False, f"Critical Error: {ex}"
    
def pre_duplicate_value_rule3(file_path):
    """
    Checks for duplicate values in predefined sheets and columns in an Excel file.
 
    Parameters:
        file_path (str): Path to the Excel file.
 
    Returns:
        (bool, list): Validation status and list of duplicate errors.
    """
    # Define sheets and columns to check
    sheets_columns_map = {
        "swc_info": ["C", "D", "E", "H", "I", "K"],
        "ib_data": ["B", "C"],
        # "ports": ["C"],
        "adt_primitive": ["B"]
    }
 
    duplicate_errors = []
    valid = True
 
    try:
        # Load the Excel file
        xls = pd.ExcelFile(file_path)
 
        for sheet, columns in sheets_columns_map.items():
            if sheet in xls.sheet_names:
                df = xls.parse(sheet)
                
                for col in columns:
                    # Ensure the column exists before checking
                    if col in df.columns:
                        duplicates = df[df.duplicated(subset=[col], keep=False)]
                        if not duplicates.empty:
                            valid = False
                            duplicate_errors.append(
                                f"Duplicates found in sheet '{sheet}', column '{col}': {duplicates[col].tolist()}"
                            )
 
    except Exception as e:
        return False, [f"Error processing file: {str(e)}"]
 
    return valid, duplicate_errors
    
def Rule_4(file_path):
    """
        Validates the consistency of interfaceType, interfaceName, dataElement, arguments, and applicationDataType.

            - If `interfaceName` is the same across columns:
        1. For interfaceType = SenderReceiver Interface / NvDataInterface / ParameterInterface:
            - interfaceType, dataElement, and applicationDataType should be the same.
        2. For interfaceType = Mode Switch, Client Server, Trigger:
            - interfaceType, interfaceName, dataElement, arguments, and applicationDataType must be the same.
            - If the conditions are not met, log errors and return False.

        Parameters:
            - file_path (str): Path to the Excel file.

        Returns:
            - (bool, list): Validation status and list of errors.
    """
 
    try:
        # Read required columns from 'ports' sheet using Pandas
        sheet_name = "ports"
        cols_to_read = ["D", "E", "F", "G", "H"]  # Assuming D=interfaceType, E=interfaceName, etc.
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=cols_to_read)
        df.columns = ["Interface Type", "Interface Name", "DataElement/OperatioAn/Parameter/Trigger/ModeGroup Name", "Argument/TriggerPeriod/Modes", "Application Data Type"]  # Rename columns
 
        # Remove leading/trailing spaces
        df = df.astype(str).applymap(lambda x: x.strip() if isinstance(x, str) else x)
 
        errors = []
 
        # Group data by `interfaceName`
        grouped = df.groupby("interfaceName")
 
        for name, group in grouped:
            interface_types = set(group["interfaceType"])
            data_elements = set(group["dataElement"])
            arguments_set = set(group["arguments"])
            app_data_types = set(group["applicationDataType"])
 
            if any(it in interface_types for it in ["SenderReceiver Interface", "NvDataInterface", "ParameterInterface"]):
                if len(interface_types) > 1 or len(data_elements) > 1 or len(app_data_types) > 1:
                    error_msg = f"[ERROR] Inconsistent values for Interface Name '{name}' (SenderReceiver/NvData/Parameter rule violated)."
                    logging.error(error_msg)
                    errors.append(error_msg)
 
            if any(it in interface_types for it in ["Mode Switch", "Client Server", "Trigger"]):
                if len(interface_types) > 1 or len(data_elements) > 1 or len(arguments_set) > 1 or len(app_data_types) > 1:
                    error_msg = f"[ERROR] Inconsistent values for Interface Name '{name}' (Mode Switch/Client Server/Trigger rule violated)."
                    logging.error(error_msg)
                    errors.append(error_msg)
 
        if errors:
            logging.error("Rule_4: Interface consistency check failed.")
            return False, errors
 
        return True, []
 
    except Exception as ex:
        logging.exception(f"Rule_4 encountered an error: {ex}")
        return False, [f"Error: {ex}"]
 
file_path = "C:\\Users\\hss930284\\Tata Technologies\\MBSE Team - SAARCONN - SAARCONN\\Eliminating_SystemDesk\\tests\\Harshit_validation_21_02\\Appl5_21_001.xlsx"
 
# Rule_1 Example
def pre_validation():
    is_valid, errors = pre_excel_rule1(file_path) 
    if not is_valid:
        raise ValueError("[ERROR] Name validation failed. Check logs.")
    
    # Rule_2 Example
    is_valid, errors = pre_ports_null_value_rule2(file_path, "ib_data")
    if not is_valid:
        raise ValueError("[ERROR] Required fields validation failed. Check logs.")
    
    # Rule_3 Example
    is_valid, errors = pre_duplicate_value_rule3(file_path)
    if not is_valid:
        raise ValueError("[ERROR] Duplicate values found. Check logs.")
    
    # Rule_4 Example
    # is_valid, errors = Rule_4(file_path)
    # if not is_valid:
    #     raise ValueError("[ERROR] Interface consistency check failed. Check logs.")