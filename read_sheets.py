# Filename: read_sheets.py
# Author: Tyler Stringer
# Date: 2024-09-03
# Details: Reads sheet names in an Excel workbook.

import pandas as pd
import logging

# Set up basic configuration for logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def read_sheet_names(path):
    try:
        logging.info(f"Loading Excel file from {path}")
        file = pd.ExcelFile(path)
        sheet_names = file.sheet_names
        logging.info(f"Found sheet names: {sheet_names}")
        return sheet_names
    except FileNotFoundError:
        logging.error(f"The file at '{path}' was not found.")
        return []
    except ValueError:
        logging.error(f"The file at '{path}' is not a valid Excel file.")
        return []
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        return []
