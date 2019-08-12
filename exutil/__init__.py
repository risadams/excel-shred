import os
import pandas as pd


def file_split( input_file):
    s = input_file.split('.')
    name = '.'.join(s[:-1]) # extract directory name
    return name


def openSheets( input_file):
    name = file_split(input_file)
    try:
        os.makedirs(name)
    except:
        pass

    wb = pd.ExcelFile(input_file)
    for ws in wb.sheet_names:
        print(ws+'.json', 'Done!')
        data = pd.read_excel(input_file, sheet_name=ws)
        new_file = os.path.join(name, ws + '.json')
        data.to_json(new_file, orient="records")
    print('Complete')
