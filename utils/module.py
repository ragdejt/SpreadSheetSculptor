#--------------------------------------------------------------------------------------------------#
# Import necessary libraries.
#--------------------------------------------------------------------------------------------------#
import time
import pandas
from rich import print
from pathlib import Path
from utils.troubleshooting import *
from utils.ascii_text import *
#--------------------------------------------------------------------------------------------------#
# Path's.
#--------------------------------------------------------------------------------------------------#
user_path = Path.home()
script_path = user_path / ("Excel_Spreadsheet")
#--------------------------------------------------------------------------------------------------#
# Create directory.                                                                                #
#--------------------------------------------------------------------------------------------------#
def create_directory():
    script_path.mkdir(exist_ok=True)
#--------------------------------------------------------------------------------------------------#
# Date / Time.
#--------------------------------------------------------------------------------------------------#
def date_time():
    print(time.strftime(DATE_TIME).center(100, "-"))
    input("[Press ENTER to continue]".center(100, "-"))
#--------------------------------------------------------------------------------------------------#
# Read spreadsheet.
#--------------------------------------------------------------------------------------------------#
def read_spreadsheet():
    try:
        print(READ_SPREADSHEET)
        spreadsheet_name = str(input("[Enter the name of the spreadsheet]: "))
        spreadsheet_ext = str(input("[Enter the spreadsheet file format]: "))
        spreadsheet_path = script_path / (spreadsheet_name + spreadsheet_ext)
        spreadsheet = pandas.read_excel(spreadsheet_path)
        print("\n[SPREADSHEET]:\n")
        print(spreadsheet)
    except FileNotFoundError:
        print(FILE_NOT_FOUND_ERROR)
#--------------------------------------------------------------------------------------------------#
# Input data.
#--------------------------------------------------------------------------------------------------#
def input_data():
    try:
        print(INPUT_DATA)
        spreadsheet_name = str(input("[Enter the name of the spreadsheet]: "))
        spreadsheet_ext = str(input("[Enter the spreadsheet file format]: "))
        spreadsheet_path = script_path / (spreadsheet_name + spreadsheet_ext)
        spreadsheet = pandas.read_excel(spreadsheet_path)
        print(
            "\n[SPREADSHEET]: \n\n",
            spreadsheet
        )
    except FileNotFoundError:
        print(FILE_NOT_FOUND_ERROR)
    else:
        print("\n[Enter the requested data]:\n")
        product_name = input("[Product name]: ")
        product_description = input("[Product description]: ")
        product_category = input("[Product gategory]: ")
        product_code = input("[Product code]: ")
        product_weight = input("[Product weight]: ")
        product_dimension_height = input("[Product dimension (Height)]: ")
        product_dimension_width = input("[Product dimension (Width)]: ")
        product_dimension_lenght = input("[Product dimension (Lenght)]: ")
        product_price = input("[Product price]: ")
        product_stock = input("[Product quantity in stock]: ")
        
        new_data = {
            'Product name': [product_name],
            'Description': [product_description],
            'Category': [product_category],
            'Product Code': [product_code],
            'Weight (kg)': [product_weight],
            'Dimension (Height)': [product_dimension_height],
            'Dimension (Width)': [product_dimension_width],
            'Dimension (Lenght)': [product_dimension_lenght],
            'Price': [product_price],
            'Quantity in stock': [product_stock]
        }
        new_row = pandas.DataFrame(new_data)
        spreadsheet = pandas.concat([spreadsheet, new_row], ignore_index=True)
        print("\n[Data inserted into the spreadsheet]\n".center(100, "-"))
        spreadsheet.to_excel(spreadsheet_path, index=False)
#--------------------------------------------------------------------------------------------------#
# Remove data.
#--------------------------------------------------------------------------------------------------#
def remove_data():
        print(REMOVE_DATA)
        spreadsheet_name = str(input("[Enter the name of the spreadsheet]: "))
        spreadsheet_ext = str(input("[Enter the spreadsheet file format]: "))
        spreadsheet_path = script_path / (spreadsheet_name + spreadsheet_ext)
        spreadsheet = pandas.read_excel(spreadsheet_path)
        line_choice = int(input("[Enter the line you want to remove]: "))
        spreadsheet = spreadsheet.drop(line_choice)
        spreadsheet.to_excel(spreadsheet_path, index=False)
        print("[Line removed]".center(100, "-"))
#--------------------------------------------------------------------------------------------------#
# Create spreadsheet.
#--------------------------------------------------------------------------------------------------#
def create_spreadsheet():
        print(CREATE_SPREADSHEET)
        spreadsheet_name = str(input("[Enter the name of the spreadsheet]: "))
        spreadsheet_ext = str(input("[Enter the spreadsheet file format]: "))
        spreadsheet_path = script_path / (spreadsheet_name + spreadsheet_ext)
        spreadsheet = pandas.DataFrame()
        spreadsheet.to_excel(spreadsheet_path)
#--------------------------------------------------------------------------------------------------#
# Delete spreadsheet.
#--------------------------------------------------------------------------------------------------#
def delete_spreadsheet():
        print(DELETE_SPREADSHEET)
        spreadsheet_name = str(input("[Enter the name of the spreadsheet]: "))
        spreadsheet_ext = str(input("[Enter the spreadsheet file format]: "))
        spreadsheet_path = script_path / (spreadsheet_name + spreadsheet_ext)
        spreadsheet_path.unlink()
#--------------------------------------------------------------------------------------------------#
# Locate data.
#--------------------------------------------------------------------------------------------------#
def locate_data():
    try:
        print(LOCATE_DATA)
        spreadsheet_name = str(input("[Enter the name of the spreadsheet]: "))
        spreadsheet_ext = str(input("[Enter the spreadsheet file format]: "))
        spreadsheet_path = script_path / (spreadsheet_name + spreadsheet_ext)
        spreadsheet = pandas.read_excel(spreadsheet_path)
        column_spreadsheet = spreadsheet.columns
        print("\n[Spreadsheet columns]:\n")
        for i, column in enumerate(column_spreadsheet, start=1):
            print(f"[{i}]: [green]{column}[/]")
    except FileNotFoundError:
        print(FILE_EXISTS_ERROR)
    else:
        column_option = int(input("\n[Enter one of the valid options]: "))
        match column_option:
            case 0:
                print("\n[Product name]\n")
                column_value0 = spreadsheet['Product name'].values
                for i, item in enumerate(column_value0, start=1):
                    print(f"[{i}]: {item}")
                date_time()
            case 1:
                print("\n[Description]\n")
                column_value1 = spreadsheet['Description'].values
                for i, item in enumerate(column_value1, start=1):
                    print(f"[{i}]: {item}")
                date_time()
            case 2:
                print("\n[Category]\n")
                column_value2 = spreadsheet['Category'].values
                for i, item in enumerate(column_value2, start=1):
                    print(f"[{i}]: {item}")
                date_time()
            case 3:
                print("\n[Product Code]\n")
                column_value3 = spreadsheet['Product Code'].values
                for i, item in enumerate(column_value3, start=1):
                    print(f"[{i}]: {item}")
                date_time()
            case 4:
                print("\n[Weight (kg)]\n")
                column_value4 = spreadsheet['Weight (kg)'].values
                for i, item in enumerate(column_value4, start=1):
                    print(f"[{i}]: {item}")
                date_time()
            case 5:
                print("\n[Dimension (Height)]\n")
                column_value5 = spreadsheet['Dimension (Height)'].values
                for i, item in enumerate(column_value5, start=1):
                    print(f"[{i}]: {item}")
                date_time()
            case 6:
                print("\n[Dimension (Width)]\n")
                column_value6 = spreadsheet['Dimension (Width)'].values
                for i, item in enumerate(column_value6, start=1):
                    print(f"[{i}]: {item}")
                date_time()
            case 7:
                print("\n[Dimension (Lenght)]\n")
                column_value7 = spreadsheet['Dimension (Lenght)'].values
                for i, item in enumerate(column_value7, start=1):
                    print(f"[{i}]: {item}")
            case 8:
                print("\n[Price]\n")
                column_value8 = spreadsheet['Price'].values
                for i, item in enumerate(column_value8, start=1):
                    print(f"[{i}]: {item}")
                date_time()
            case 9:
                print("\n[Quantity in stock]\n")
                column_value9 = spreadsheet['Quantity in stock'].values
                for i, item in enumerate(column_value9, start=1):
                    print(f"[{i}]: {item}")
                date_time()
#--------------------------------------------------------------------------------------------------#
# Menu.
#--------------------------------------------------------------------------------------------------#
def menu():
    while True:
        try:
            print(MENU)
            menu_input = int(input("[Enter one of the valid options]: "))
            match menu_input:
                case 0:
                    exit()
                case 1:
                    input_data()
                case 2:
                    remove_data()
                case 3:
                    create_spreadsheet()
                case 4:
                    delete_spreadsheet()
                case 5:
                    read_spreadsheet()
                case 6:
                    locate_data()
                case _:
                    print(VALUE_ERROR)
        except ValueError:
            print(VALUE_ERROR)