
import requests
import openpyxl
import os
from io import BytesIO


def load_excel(excile_path: str, data_only=False, read_only=True):
    """
    Load Excel from either local path or online URL.
    Returns workbook object.
    """
    try :
        if excile_path.startswith("http://") or excile_path.startswith("https://"):
            print("[INFO] Loading online Excel file:", excile_path)
            response = requests.get(excile_path)
            print("Content-Type:", response.headers.get("Content-Type"))
            
            response.raise_for_status()
            wb = openpyxl.load_workbook(filename=BytesIO(response.content), data_only=data_only, read_only=read_only)
            
        elif os.path.exists(excile_path):
            print("[INFO] Loading local Excel file:", excile_path)
            wb = openpyxl.load_workbook(filename=excile_path, data_only=data_only, read_only=read_only)

        else:
            raise FileNotFoundError(f"Excel source not found: {excile_path}")
        return wb
    except FileNotFoundError as e:
        raise FileNotFoundError(f'File not found in your pc: "{excile_path}"')
        