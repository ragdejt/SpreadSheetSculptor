#--------------------------------------------------------------------------------------------------#
# Import necessary libraries.
#--------------------------------------------------------------------------------------------------#
import os
import time
import pandas
from rich import print
from pathlib import Path
from utils.ascii_text import *
from utils.troubleshooting import *
#--------------------------------------------------------------------------------------------------#
# Path's.
#--------------------------------------------------------------------------------------------------#
user_path = Path.home()
script_path = user_path / ("Excel_Spreadsheet")
#--------------------------------------------------------------------------------------------------#
# Date / Time.
#--------------------------------------------------------------------------------------------------#
def date_time():
    print(time.strftime(DATE_TIME).center(100, "-"))
    input("[Press ENTER to continue]".center(100, "-"))
    os.system("cls")
#--------------------------------------------------------------------------------------------------#
# List spreadsheet.
#--------------------------------------------------------------------------------------------------#
def list_spreadsheet():
    try:
        print("[Spreadsheet list]".center(100, "-"))
        for i, dir in enumerate(os.listdir(script_path), start=1):
            print(f"[{i}]: {dir}")
        date_time()
    except FileNotFoundError:
        print(FILE_NOT_FOUND_ERROR)
#--------------------------------------------------------------------------------------------------#
# Select the spreadsheet.
#--------------------------------------------------------------------------------------------------#
def select_spreadsheet():
    print(SELECT_SPREADSHEET)
    global spreadsheet_name
    spreadsheet_name = str(input("[Enter the name of the spreadsheet]: "))
    spreadsheet_ext = str(input("[Enter the spreadsheet file format]: "))
    global spreadsheet_path
    spreadsheet_path = script_path / (spreadsheet_name + spreadsheet_ext)
    date_time()
    menu()
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
        print("[Spreadsheet created]".center(100, "-"))
        date_time()
#--------------------------------------------------------------------------------------------------#
# Delete spreadsheet.
#--------------------------------------------------------------------------------------------------#
def delete_spreadsheet():
        print(DELETE_SPREADSHEET)
        spreadsheet_name = str(input("[Enter the name of the spreadsheet]: "))
        spreadsheet_ext = str(input("[Enter the spreadsheet file format]: "))
        spreadsheet_path = script_path / (spreadsheet_name + spreadsheet_ext)
        spreadsheet_path.unlink()
        print("[Spreadsheet deleted]".center(100, "-"))
        date_time()
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
        date_time()
    except FileNotFoundError:
        print(FILE_NOT_FOUND_ERROR)
#--------------------------------------------------------------------------------------------------#
# Menu spreadsheet.
#--------------------------------------------------------------------------------------------------#
def menu_spreadsheet():
    print(MENU_SPREADSHEET)
    menu_spreadsheet_input = int(input("[Enter one of the valid options]: "))
    match menu_spreadsheet_input:
        case 0:
            exit()
        case 1:
            list_spreadsheet()
            menu_spreadsheet()
        case 2:
            select_spreadsheet()
        case 3:
            create_spreadsheet()
            menu_spreadsheet()

        case 4:
            delete_spreadsheet()
            menu_spreadsheet()
        case 5:
            read_spreadsheet()
            menu_spreadsheet()

#--------------------------------------------------------------------------------------------------#
# Input data.
#--------------------------------------------------------------------------------------------------#
def insert_data():
    try:
        print(INSERT_DATA)
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
        spreadsheet.to_excel(spreadsheet_path, index=False)
        print("\n[Data inserted into the spreadsheet]\n".center(100, "-"))
        date_time()
#--------------------------------------------------------------------------------------------------#
# Remove data.
#--------------------------------------------------------------------------------------------------#
def delete_data():
        print(DELETE_DATA)
        spreadsheet = pandas.read_excel(spreadsheet_path)
        line_choice = int(input("[Enter the line you want to remove]: "))
        spreadsheet = spreadsheet.drop(line_choice)
        spreadsheet.to_excel(spreadsheet_path, index=False)
        print("[Data removed]".center(100, "-"))
        date_time()
#--------------------------------------------------------------------------------------------------#
# Add column
#--------------------------------------------------------------------------------------------------#
def add_column():
        try:
            print("")
            spreadsheet = pandas.read_excel(spreadsheet_path)
        except FileNotFoundError:
            print(FILE_EXISTS_ERROR)
        else:
            new_column_name = input("[Enter the column name]")
            new_column_value = input(f"[Enter value for {new_column_name}]: ")
            spreadsheet[new_column_name] = [new_column_value]
            spreadsheet.to_excel(spreadsheet_path, index=False)
            print("[Added column]".center(100, "-"))    
            date_time()
#--------------------------------------------------------------------------------------------------#
# Remove column
#--------------------------------------------------------------------------------------------------#
def remove_column():
        try:
            print("")
            spreadsheet = pandas.read_excel(spreadsheet_path)
        except FileNotFoundError:
            print(FILE_NOT_FOUND_ERROR)
        else:
            new_column_name = input("[Enter the column name]")
            spreadsheet = spreadsheet.drop(new_column_name, axis=1)
            spreadsheet.to_excel(spreadsheet_path, index=False)
            print("[Column removed]".center(100, "-"))
            date_time()
#--------------------------------------------------------------------------------------------------#
# Locate data.
#--------------------------------------------------------------------------------------------------#
def find_data():
    try:
        print(FIND_DATA)
        spreadsheet = pandas.read_excel(spreadsheet_path)
        column_spreadsheet = spreadsheet.columns
        print("\n[Spreadsheet columns]:\n")
        for i, column in enumerate(column_spreadsheet):
            print(f"[{i}]: [green]{column}[/]")
        selected_column_index = int(input("\n[Select column number]: "))
        selected_column = column_spreadsheet[selected_column_index]
        print(f"Selected Column: [green]{selected_column}[/]\n")
        print(spreadsheet[selected_column])
        date_time()
    except FileNotFoundError:
        print(FILE_EXISTS_ERROR)
#--------------------------------------------------------------------------------------------------#
# Main Menu.
#--------------------------------------------------------------------------------------------------#
def menu():
    try:
        print(MENU)
        menu_input = int(input("[Enter one of the valid options]: "))
        match menu_input:
            case 0:
                exit()
            case 1:
                insert_data()
                menu()
            case 2:
                delete_data()
                menu()
            case 3:
                add_column()
                menu()
            case 4:
                remove_column()
                menu()
            case 5:
                find_data()
                menu()
            case _:
                print(VALUE_ERROR)
    except ValueError:
        print(VALUE_ERROR)