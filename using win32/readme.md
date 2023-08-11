# Go Alerts Automation Project

## Setting up the Environment: Prerequisites
1. First of all, make sure you have python and IDE installed on your device. Here are the links:
   1. [Python](https://www.python.org/downloads/)
   2. [VSCode](https://code.visualstudio.com/Download) (or any IDE of your choice)
2. Let's also intall a python extension on VSCode for a better experience:
    1. Open VSCode and click on Extensions from the left pane.
    2. Search Python and download the one by Microsoft.
3. Verify python and install required libraries:
   1. Open terminal/command line using an icon in top-right of VSCode *OR* Open Command Prompt (as admin)
   2. Type `python --version`. If this displays the version, it means that we have python installed on our device. Now we can proceed to install required libraries.
   3. Type `pip install openpyxl pandas beautifulsoup4 pywin32 time schedule` to install all the required libraries for our python program.
    *OR*
    We can also install them step by step in this format: `pip install <library name>`
4. Create a new folder on your device. Paste the Template excel file and python program there. 
    1. Rename the excel file to your desired name and update the name in the code.    
    2. You can also rename the sheet1, sheet2, and sheet3. **BUT** make sure to update in the code
5. Setup the sender email by typing your email into the SENDER_EMAIL variable.
6. Copy the path of that excel file and paste into the MASTER_WORKBOOK_PATH variable.

## Running the Script
**VERY IMPORTANT: DO NOT KEEP EXCEL FILE OPEN WHILE RUNNING THE SCRIPT.**  
Run the program by either using the Run button on top-right or with the command `python main.py` in the terminal.

# Code Explanation: Email Table Data Extractor and Excel Updater
This Python script allows you to extract table data from emails, find "not working" entries, and update an Excel workbook with the extracted data. It opens outlook using pywin32 lib, searches for specific emails from a particular sender, extracts table data from the email content (assumed to be in HTML format), and updates the Excel workbook with the new data. The script also finds rows with "not working" entries in the "Message" column and copies them to a separate sheet in the workbook.

## Excel Sheets
- Sheet3: Local dynamic database of devices based on the most recent email received
- Sheet2: List of all the inactive devices based on the most recent email received
- Sheet1: List of all the alerts received

## Dependencies
 - `openpyxl`: This library is used to manipulate Excel files (xlsx format).
 - `pandas`: This library is used for data manipulation and analysis.
 - `beautifulsoup4` (bs4): This library is used for parsing HTML content.
 - `pywin32`: This library is used to handle all the outlook operations such as opening, selecting inbox, sender etc.

Please make sure you have these libraries installed before running the script.

## Configuration
1. Set the `SENDER_EMAIL` to the email address of the sender whose emails you want to process.
2. Set the `MASTER_WORKBOOK_PATH` to the path of the Excel workbook you want to update. If the workbook is in the same directory as the python file, simply paste the name of the workbook including the extenio (such as .xlsx)

## Class
### `EquipmentEntry`
A simple class representing an equipment entry extracted from the email table. It has the following attributes:

- `site`: The site of the equipment.
- `equipment`: The equipment's name.
- `message`: The message associated with the equipment.
- `last_state_change`: The last state change timestamp of the equipment.

## Functions
### `parse_html_email(html_content)`
This function takes the HTML content of an email and extracts tabular data from it and extract equipment entries. It returns the tabular data as a list of lists or `None` if no table is found. It returns a list of `EquipmentEntry` objects.

### `append_table_data_to_worksheet1(table_data, worksheet1, added_rows)`
This function appends the table data to the first sheet (`worksheet1`) of the master Excel workbook. It avoids adding duplicate rows by checking against the set of `added_rows`.

### `update_table_data(existing_entries, new_entries, worksheet3)`
This function updates the existing entries in the third sheet (`worksheet3`) of the master Excel workbook with the new entries. If an entry with the same equipment and site already exists, its message and last state change are updated with the new entry's data.

### `read_existing_data_from_excel(worksheet3)`
This function reads the existing data from the third sheet (`worksheet3`) of the master Excel workbook and returns a list of `EquipmentEntry` objects.

### `driver()`
The driver function is the main driver of this application and performs the following steps:
- Connects to the user's Outlook account
- Searches for emails from the specified sender and extracts the HTML table data.
- Opens the master workbook and accesses the first three sheets.
- Reads existing data from the third sheet into a list of `EquipmentEntry` objects.
- For each email, it fetches the raw email content, extracts table data from the HTML body, and appends it to the first sheet (`worksheet1`) of the master workbook.
- It also extracts equipment entries from the HTML content and updates the existing data in the third sheet (`worksheet3`) of the master workbook.
- Converts the combined data to a Pandas DataFrame and saves it directly to the third sheet of the master workbook.
- Finds all rows with "not working" in the "Message" column and copies them to the second sheet (`worksheet2`) of the master workbook.
- Saves the modified workbook and closes it.

### `main()`
This function handles the scheduling of the driver function. The program runs on a scheduled loop every day at 8:15am.