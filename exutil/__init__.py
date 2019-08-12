import os
from pathlib import Path

import pandas as pd


def extract_dir_name(input_file):
    """
    creates a directory path based on the specified file name
    :param input_file:
    :return:
    """
    s = input_file.split('.')
    name = '.'.join(s[:-1])
    return name


def open_dir(input_path):
    """
    Opens the specified input path and returns any located excel file
    :param input_path:
    :return:
    """
    for file in Path(input_path).glob('**/*.xls'):
        yield file

    for file in Path(input_path).glob('**/*.xlsx'):
        yield file


def shred_sheets(input_file, _format):
    """
    Opens an excel workbook, and converts all sheets to a new file of the specified format
    :param input_file:
    :param _format:
    :return:
    """
    name = extract_dir_name(input_file) + f"_{format}"
    try:
        os.makedirs(name)
    except:
        pass

    wb = pd.ExcelFile(input_file)
    for ws in wb.sheet_names:
        print(ws + '.' + _format, 'Done!')
        data = pd.read_excel(input_file, sheet_name=ws)
        new_file = os.path.join(name, ws + '.' + _format)
        if _format == 'json':
            data.to_json(new_file, orient="records")
        elif _format == 'csv':
            data.to_csv(new_file)
        else:
            raise AssertionError(f"Invalid format {_format}")
    print('Complete')
