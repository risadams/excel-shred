import os
import re
import click
import pandas as pd
from pymongo import MongoClient
from pathlib import Path, PurePath


def extract_dir_name(input_file):
    """
    creates a directory path based on the specified file name
    :param input_file: file bane
    :return: full path, minus extension
    """
    fname = PurePath(input_file).__str__()
    s = fname.split('.')
    name = '.'.join(s[:-1])
    return name


def prep_file_name(path, file):
    """
    append the original path and file name
     * strips special chars
     * remove spaces (replace with underscore)
     * convert to lowercase
    :param path: the path part of the new file name
    :param file: the original file name
    :return: sanitized name
    """
    name = path.__str__() + '~' + file.__str__()
    name = name.lower()
    name = name.replace(' ', '_')
    name = re.sub('[^a-z0-9\-_!.~]+', '', name)
    return name


def open_dir(input_path, patterns):
    """
    Opens the specified input path and returns any located excel file
    :param patterns: the file extensions to glob over (eg xls, csv)
    :param input_path: the starting path
    :return: generator of all found files
    """
    for ext in patterns:
        for file in Path(input_path).glob('**/*.' + ext):
            yield file


def shred_sheets(subdomain, audit_date, input_file, _format):
    """
    Opens an excel workbook, and converts all sheets to a new file of the specified format
    :param subdomain: appended to data frame
    :param audit_date: appended to data fram
    :param input_file: the path to the excel book
    :param _format: the format to convert all sheets
    :return:
    """
    name = extract_dir_name(input_file)
    fname = PurePath(input_file).name.__str__()
    try:
        os.makedirs(name)
    except:
        pass

    wb = pd.ExcelFile(input_file)
    for ws in wb.sheet_names:
        data = pd.read_excel(input_file, sheet_name=ws)
        # add constants
        data.index.names = ['ix']
        data['subdomin'] = subdomain
        data['audit_date'] = audit_date

        # strip chars we don't want in colum names
        cols = data.columns
        renamed = []
        for col in cols:
            col = re.sub('[^a-zA-Z0-9]', '', col)
            renamed.append(col)

        data.columns = renamed

        # build output formats
        if _format == 'mongo':
            client = MongoClient('mongodb://localhost:27017/')
            db = client.Sitebulb
            cl = db.August5

            try:
                cl.insert_many(data.to_dict('records'))
            except Exception as e:
                click.secho(f'\nERROR in [{input_file},{ws}] -- {e}', fg='red')
                continue

        if _format == 'json' or _format == 'all':
            try:
                new_file = os.path.join(name, fname + '~' + ws + '.json')
                data.to_json(new_file, orient="records")
            except Exception as e:
                click.secho(f'\nERROR in [{input_file},{ws}] -- {e}', fg='red')
                continue

        if _format == 'csv' or _format == 'all':
            try:
                new_file = os.path.join(name, fname + '~' + ws + '.csv')
                data.to_csv(new_file)
            except Exception as e:
                click.secho(f'\nERROR in [{input_file},{ws}] -- {e}', fg='red')
                continue
