# Script Name.

***`Author` Ragdejt***

***`Date Created` 23/01/2024***

#

***`Description:`***

This is a Python script for managing an Excel spreadsheet for product inventory. The script includes functions for creating a directory for the spreadsheet, displaying the current date and time, reading the contents of a spreadsheet, inputting new product data, removing product data, and locating product data in the spreadsheet. script uses the pandas library to read and write Excel files, and the rich library to print formatted text.

The script defines several constants at the top, such as:

- DATE_TIME
- MODULE_NOT_FOUND_ERROR
- FILE_NOT_FOUND_ERROR
- FILE_EXISTS_ERROR
- VALUE_ERROR
- PERMISSION_ERROR
- SYNTAX_ERROR
- INDENTATION_ERROR
- NAME_ERROR
- TYPE_ERROR
- INDEX_ERROR
- KEY_ERROR
- ATTRIBUTE_ERROR
- ZERO_DIVISION_ERROR
- IMPORT_ERROR

Which are used throughout the script to provide clear and consistent error messages and prompts to the user.
#
The `menu()` function displays a menu of options to the user and waits for them to enter a valid option. If the user enters an invalid option, the script will display an error message and prompt the user to try again. If the user enters the option to exit the program (0), the script will exit. Otherwise, the script will call the appropriate function based on the user's selection.
#
The `read_spreadsheet()` function prompts the user to enter the name and file format of the spreadsheet they want to read, then uses pandas to read the contents of the spreadsheet and display it to the user.
#
The `input_data()` function prompts the user to enter the name and file format of the spreadsheet they want to modify, then reads the contents of the spreadsheet using pandas. The function then prompts the user to enter the data for a new product, and adds a new row to the spreadsheet with this data. Finally, the function writes the updated spreadsheet back to disk.
#
The `remove_data()` function prompts the user to enter the name and file format of the spreadsheet they want to modify, then reads the contents of the spreadsheet using pandas. The function then prompts the user to enter the line number of the row they want to remove, and removes this row from the spreadsheet. Finally, the function writes the updated spreadsheet back to disk.
#
The `delete_spreadsheet()` function prompts the user to enter the name and file format of an existing Excel spreadsheet, then deletes this file from disk. It is used to permanently delete a spreadsheet.
#
The `create_spreadsheet()` function prompts the user to enter the name and file format of a new Excel spreadsheet, then creates a new, empty spreadsheet with this name and file format using the pandas library. It is used to create a new spreadsheet from scratch.
#
The `locate_data()` function prompts the user to enter the name and file format of the spreadsheet they want to search, then reads the contents of the spreadsheet using pandas. The function then displays a menu of columns in the spreadsheet and prompts the user to select one. Depending on the user's selection, the function displays the values in the corresponding column and waits for the user to press enter before continuing.
#
Overall, this script provides a convenient and user-friendly way to manage an Excel spreadsheet for product inventory. The script includes clear and concise error messages and prompts, and uses pandas to read and write Excel files, making it easy to use and modify.
#

# Security Notes

- Before running any script, carefully review and understand the code. This helps you to ensure that the script aligns with its intended purpose and does not have unintended consequences.

- This script is for educational purposes only.
    
- Do not use against any network that you don't own or have authorization to test.

- Make sure you run this script in a secure environment.This can help prevent unintended interactions with other applications and libraries.

#

# Troubleshooting

### ***`ModuleNotFoundError`*** 

***`Description:`***

- Indicates that the specified module is not installed.

***`Possible Solution:`*** 

- Run pip install example_module to install the missing module.

- Ensure that all dependencies are installed by running `pip install -r requirements.txt.`

#

### ***`ValueError`***

***`Description:`***

-  Indicates that the provided input format is not valid.

***`Possible Solution:`***

- Check if the input is in the correct format by referring to the documentation.

#

### ***`FileNotFoundError`***

***`Description:`***

- Indicates that the specified file was not found.

***`Possible Solution:`***

- Check if the file is in the correct location.

#

### ***`FileExistsError`***

***`Description:`***

- Indicates that the file you are trying to create already exists.

***`Possible Solution:`***

- Check if the file already exists before attempting to create it.

#

### ***`PermissionError`***

***`Description:`***

- Indicates that the program does not have permission to access the specified file.

***`Possible Solution:`***

- Check the file and directory permissions.

- Run the program with elevated permissions if necessary.

#

# Contributing

### Report Bugs

- *If you are reporting a bug, please include:*

        1. Your operating system name and version.

        2. Python Version.

        3. Any additional details about your local setup that might be helpful in troubleshooting.

        4. Detailed steps to reproduce the bug.

#

### Fix Bugs

- *If you are fixing a bug, please include:*

        1. Description of the Bug:

        Provide a clear and concise description of the bug.

        2. Steps to Reproduce:

        Outline the steps that lead to the occurrence of the bug.

        3. Expected Behavior:

        Describe what the expected behavior should have been without the bug.

        4. Environment Information:

        Mention your operating system, Python version, and any other relevant environment details.

#

### Implement Features

- *If you are proposing a feature:*

        1. Explain in detail how it would work.

        2. Keep the scope as narrow as possible, to make it easier to implement.

        3. Remember that this is a volunteer-driven project, and that contributions are welcome.

#

### Write Documentation

- If you are writing documentation, please include:

        1. Documented Feature/Function:
        
        Specify which feature or function the documentation covers

        2. Documentation Details:

        Provide clear and comprehensive details about the usage, parameters, and expected outcomes.

        3. Code Examples:

        Include code examples that demonstrate the usage of the documented feature.

        4. Formatting Guidelines:

        If there are specific formatting guidelines for documentation, mention them.

#

### Submit Feedback

- If you are submitting feedback, please include:

        1. Nature of Feedback:

        Clearly state whether the feedback is a suggestion, bug report, or general comment.

        2. Specific Section/Functionality:

        Identify the specific section or functionality your feedback is related to.

        3. Detailed Feedback:

        Provide detailed information to explain your feedback, including any suggestions for improvement.

        4. Reproducibility (if applicable):

        If it's a bug report, mention if you can consistently reproduce the issue and the steps to do so.

#

# Requirements

This script requires Python 3 and the following libraries:  

* pandas

* rich

* openpyxl

# Repository

- You can find the script and its source code in the following repository:   
https://www.github.com/ragdejt/
