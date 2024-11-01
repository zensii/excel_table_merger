import time

import pandas as pd
import xlwings as xw
from time import sleep
import numpy as np
import sys


def check_for_updated(filename):
    if '(UPDATED)' in filename:
        print('An updated file is already present in the current directory!')
        print('Please, delete the UPDATED file and try again.')
        print("Press 'Enter' key to quit.")
        input()
        sys.exit()


def check_for_df(dataframes):
    if dataframes:
        pass
    else:
        print('Files not found! Script terminating...')
        sleep(1)
        sys.exit()


def get_df(file_paths):
    """Function to scan the selected dir and extract the needed dataframes for each of the Excel files"""
    dataframes = {}
    df_main = None
    main_file = None
    main_file_open = False
    counter = 1

    for file_path in file_paths:
        filename = file_path.split('/')[-1]

        check_for_updated(filename)  # exit if updated file is present

        try:
            if not filename.endswith('.py') and not filename.startswith('~$'):
                df_main = pd.read_excel(f'{file_path}',
                                        sheet_name='raw date per vendor')  # locate the main data_frame file
                # if main file is located
                if len(df_main) != 0:
                    # checking if main file is already open
                    open_excels = xw.apps.active

                    if open_excels:
                        for book in open_excels.books:
                            if book.name == filename:
                                main_file_open = True
                                main_file = book
                                break

                    if not main_file_open:
                        # If the workbook is not open, open it
                        print('Opening main file in Excel. Please wait...')
                        main_file = xw.Book(file_path)  # start excel using xlwings

                        start_time = time.time()
                        timeout = 30
                        while time.time() - start_time < timeout:
                            try:
                                # Try to access a property that requires the workbook to be fully loaded
                                _ = main_file.sheets[0].used_range
                                print( f'File {filename} loaded successfully')
                                break
                            except Exception:
                                # If an exception occurs, the book isn't fully loaded yet
                                time.sleep(0.5)  # Short sleep to prevent excessive CPU usage

                    else:
                        # if main Excel file is open:
                        print('Main file already open - connecting...')
                        sleep(2)

        except ValueError:  # fill a dict with dataframes from the files in the folder
            print('Additional file located. Opening... ')
            key = f'df_to_add{counter}'

            try:  #
                # when reading, define the data type per column
                dataframes[key] = pd.read_excel(file_path, sheet_name='Sheet1', converters={'GL': str, 'Vendor': str})
                counter += 1
            except :
                print(f'File: {filename} is not valid!')
                continue

        except FileNotFoundError:
            print(f'{filename} not found!')

        except:
            print('There is a problem with the files.')
            sleep(2)
            sys.exit()

    return dataframes, df_main, main_file


def get_period(dataframes, file_year=None, file_period=None):
    """  get year and period from the file that has it (it is the one that has 'GL') """

    for key, df in dataframes.items():
        if 'GL' in df.columns:
            if file_year is None:
                file_year = int(df['Year'][0])
                file_period = int(df['Period'][0])
    return file_year, file_period


def clean_data(dataframes, columns):
    print('Cleaning data...')
    sleep(1)

    for key, df in dataframes.items():

        if 'GL' in df.columns:  # only the file we need to rename contains the 'GL' column name
            # dataframes[key] = this is also possible as rename() does not modify in place by default
            # (I used inplace=True)
            df.rename(columns={'GL': 'G/L Account', 'Vendor Name': 'Name 1', 'Amount_Act': 'Amount in Loc.Crcy 2'},
                      inplace=True)

            # replace empty values with '' - needed as otherwise in the pivot grand total is not correct
            # as the rows with no data are omitted?!
        for col in columns:
            df[col] = df[col].replace(np.nan, '')


def create_sub_pivots(dataframes):
    pv_tables = []
    print('Creating pivot tables...')
    sleep(1)
    # creating a list with the pivot tables of all files to be added in the folder
    for key, df in dataframes.items():
        pv_tables.append(pd.pivot_table(df, values=['Amount in Loc.Crcy 2', 'Amount in Local Currency'],
                                    index=['Company Code', 'G/L Account', 'Vendor', 'Name 1'],
                                    aggfunc='sum').reset_index())
    return pv_tables


def create_main_pivot(pv_tables, file_year, file_period):
    # Create the combined pivot table

    pivot = pd.concat(pv_tables, ignore_index=True)

    pivot['Year'] = file_year
    pivot['Period'] = file_period

    return pivot


def get_columns_to_update(pivot):
    print('Normalizing file formats...')
    sleep(1)
    # Defining the columns to update the main file with
    vendors = pivot['Vendor'].apply(
        lambda x: f"'{x}")  # adding ' in front of the numbers to be interpreted as text (needed for main file logic)
    g_l = pivot['G/L Account'].apply(
        lambda x: f"'{x}")  # adding ' in front of the numbers to be interpreted as text (needed for main file logic)
    name = pivot['Name 1']
    amount = pivot['Amount in Local Currency']
    year = pivot['Year']
    period = pivot['Period']

    return vendors, g_l, name, amount, year, period


def prep_update_package(pivot):
    data = []
    # save the data to be added as list in order to paste it in the Excel file in one shot
    vendors, g_l, name, amount, year, period = get_columns_to_update(pivot)
    print('Preparing data package...')
    sleep(1)
    for g, v, n, a, y, p in zip(g_l.values, vendors.values, name.values, amount.values, year.values, period.values):
        data.append([g, v, n, None, None, None, None, a, y, p])

    return data


def get_main_sheet(main_file):
    # which sheet we will update in the main file
    sheet = main_file.sheets['raw date per vendor']

    return sheet


def get_start_row(df_main):
    start_row = len(df_main.iloc[:, 0]) + 2  # find the first empty row

    return start_row


def update(data, sheet, start_row):
    print('Updating main file. Please wait...')
    sheet.range(f'B{start_row}').value = data  # write data to Excel file
    sleep(10)  # wait for the Excel to update formulas
