import xlwings as xw
import win32api
import string
from sqlalchemy import create_engine, inspect
import pandas as pd
import numpy as np
from sqlalchemy import MetaData, Table
from sqlalchemy.sql import select
import os
from dotenv import load_dotenv
script_dir = os.path.dirname(r"D:\Internship\mrpn\mrpn\input.txt")
dotenv_path = os.path.join(script_dir, '.env')
load_dotenv(dotenv_path)
user = os.getenv('mrpn_user')
password = os.getenv('HVzw60rpoc4A63IXi4LT2IwICfm0SZaP')
host = os.getenv('dpg-comd54v79t8c738rh8o0-a.singapore-postgres.render.com')
port = os.getenv('5432')
database = os.getenv('mrpn')
conn = f"postgresql://{mrpn_user}:{HVzw60rpoc4A63IXi4LT2IwICfm0SZaP}@{dpg-comd54v79t8c738rh8o0-a.singapore-postgres.render.com}:{5432}/{mrpn}"

def read_input_from_vba(file_path):
    with open(file_path, 'r') as file:
        input_data = file.read()
    return input_data


def fetch_data_from_database(engine, table_name):
    metadata = MetaData()
    metadata.reflect(bind=engine)
    table_df = Table(table_name, metadata, autoload=True, autoload_with=engine)

    with engine.connect() as conn:
        query = select('*').select_from(table_df)
        result = conn.execute(query)
        rows = result.fetchall()
        column_names = result.keys()

    return rows, column_names

# def update_excel_sheet(sheet, rows, column_names):
#     sheet.clear()


#     for i, col_name in enumerate(column_names):
#         sheet[f"{chr(65+i)}1"].value = col_name

#     for i, row in enumerate(rows):
#         for j, cell_value in enumerate(row):
#             sheet[f"{chr(65+j)}{i+2}"].value = cell_value


def update_excel_sheet(sheet, rows, column_names):
    sheet.clear()

    # Define a function to convert column index to Excel column name
    def get_column_name(index):
        if index < 26:
            return string.ascii_uppercase[index]
        else:
            first_letter = string.ascii_uppercase[index // 26 - 1]
            second_letter = string.ascii_uppercase[index % 26]
            return first_letter + second_letter

    # Writing column names
    for i, col_name in enumerate(column_names):
        sheet[f"{get_column_name(i)}1"].value = col_name

    # Writing data rows
    for i, row in enumerate(rows):
        for j, cell_value in enumerate(row):
            sheet[f"{get_column_name(j)}{i+2}"].value = cell_value


def main_g():

    script_dir = os.path.dirname(r"D:\Internship\mrpn\mrpn\~$mrpn.xlam")

    filename = "input.txt"
    file_path = os.path.join(script_dir, filename)

    input_from_vba = read_input_from_vba(file_path)

    input_from_vba = int(input_from_vba)

    table_names = ['az_nsr', 'country_plant',
                   'demand_incremental', 'hbl', 'oubp', 'sbl']
    selected_table = table_names[input_from_vba]

    engine = create_engine(conn)

    try:

        rows, column_names = fetch_data_from_database(engine, selected_table)

        wb = xw.Book.caller()
        sheet = wb.sheets[1]
        update_excel_sheet(sheet, rows, column_names)
        # wb.save()

    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:

        engine.dispose()


if __name__ == "__main__":
    xw.Book(r"mrpn.xlam").set_mock_caller()
    main_g()


def main_u():
    script_dir = os.path.dirname(r"D:\Internship\mrpn\mrpn\mrpn.xlam")

    filename = "input.txt"
    file_path = os.path.join(script_dir, filename)

    input_from_vba = read_input_from_vba(file_path)

    input_from_vba = int(input_from_vba)

    table_names = ['az_nsr', 'country_plant',
                   'demand_incremental', 'hbl', 'oubp', 'sbl']
    selected_table = table_names[input_from_vba]

    engine = create_engine(conn)

    try:

        wb = xw.Book.caller()
        sheet = wb.sheets[0]
        data = sheet.range('A1').expand().value
        column_names = data[0]
        data_values = data[1:]
        df = pd.DataFrame(data_values, columns=column_names)

        # Convert column names to lowercase
        df.columns = [str(col).lower() for col in df.columns]

        # Upload data to the database table
        df.to_sql(selected_table, engine, if_exists='append', index=False)
        message = f"Data Uploaded Succussfully from: {selected_table}"
        win32api.MessageBox(0, message, 'Success')

    except Exception as e:
        # print(f"An error occurred: {str(e)}")
        # xw.apps.active.msgbox(f"An error occurred: {str(e)}")
        error_message = f"An error occurred: {str(e)}"
        win32api.MessageBox(0, error_message, "Error")

    finally:
        # Dispose of the engine
        engine.dispose()


if __name__ == "__main__":
    xw.Book(r"mrpn.xlam").set_mock_caller()
    main_u()
