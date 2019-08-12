import os
import pandas as pd
from pathlib import Path, PurePath


def extract_dir_name(input_file):
    """
    creates a directory path based on the specified file name
    :param input_file:
    :return:
    """
    fname = PurePath(input_file).__str__()
    s = fname.split('.')
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
    name = extract_dir_name(input_file)
    try:
        os.makedirs(name)
    except:
        pass

    wb = pd.ExcelFile(input_file)
    for ws in wb.sheet_names:
        print(f"\t Extracting WS: {ws}")
        data = pd.read_excel(input_file, sheet_name=ws)

        if _format == 'json' or _format == 'all':
            try:
                new_file = os.path.join(name, ws + '.json')
                data.to_json(new_file, orient="records")
            except Exception as e:
                print(e)
                continue

        if _format == 'csv' or _format == 'all':
            try:
                new_file = os.path.join(name, ws + '.csv')
                data.to_csv(new_file)
            except Exception as e:
                print(e)
                continue

        print("\tDone!")
