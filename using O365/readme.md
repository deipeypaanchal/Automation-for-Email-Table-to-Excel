# Go Alerts Automation Project

## Creating Azure Application to Handle Authentication
1. navigate to [the Azure portal](https://portal.azure.com) and sign in (you will need admin access to create the application).
2. Select 'App registrations' from the 'Azure services' section of the home page or search for it.
3. Click wither 'Register New Application' or 'New registration' to begin the process of creating the application.
4. Follow the following steps for the next page:
   1. Name: name the application whatever you would like.
   2. Supported account types: Accounts in any organizational directory.
   3. Redirect URI: Set platform to Web and paste the following URL into the text box: https://login.microsoftonline.com/common/oauth2/nativeclient.
   4. Click 'Register' to continue.
5. Next, Copy the 'Application (client) ID and use it to replace the existing Client ID variable in the code.
6. Click on 'Certificates & secrets' from the side menu.
   1. Click 'New client secret', and a description if desired, and click 'Add'.
   2. Copy the Value (NOT the 'Secret ID') and use it to replace the existing Client Secret variable in the code.
7. Next, select 'API permissions' from the side menu.
   1. Click 'Add a permission' and select 'Microsoft Graph' from the menu that pops up.
   2. Click 'Delegated permissions' and search for and check the boxes for 'email', 'offline_access', 'openid', and 'Mail.Read'.
   3. When finished, select 'Add permissions' and refresh the page to make sure they appear in the list.
8. The application has now been created! Now you can use the python program to authenticate with your Outlook account.

## Setting up the Environment: Prerequisites
1. First of all, make sure you have python and IDE installed on your device. Here are the links:
   1. [Python](https://www.python.org/downloads/)
   2. [VSCode](https://code.visualstudio.com/Download) (or any IDE of your choice)

2. Let's also intall a python extension on VSCode for a better experience:
    1. Open VSCode and click on Extensions from the left pane.
    2. Search Python and download the one by Microsoft.

3. Verify python and install required libraries:
   1. Open terminal/command line using an icon in top-right of VSCode *OR* Open Command Prompt (as admin).
   2. Type `python --version`. If this displays the version, it means that we have python installed on our device. Now we can proceed to install required libraries.
   3. Type `pip install openpyxl pandas O365 logging time schedule` to install all the required libraries for our python program.
    *OR*
    We can also install them step by step in this format: `pip install <library name>`.

4. Create a new folder on your device. Paste the Template excel file and python program there. 
    1. Rename the excel file to your desired name and update the name in the code (must include full path if it is not in the same folder as the python script).    
    2. You can also rename the sheet1, sheet2, and sheet3. **BUT** make sure to update in the code.

## Running the Script
**VERY IMPORTANT: DO NOT KEEP EXCEL FILE OPEN WHILE RUNNING THE SCRIPT.**  
1. Run the program by either using the Run button on top-right or with the command `python main-auth.py` in the terminal.
2. A link will be printed to the terminal. Navigate to this link and log into your Exelon Outlook account.
3. When the page refreshes, copy the URL from the search bar and paste it into the terminal.
4. The program should now run successfully and modify the Excel file.

# Code Explanation: Email Table Data Extractor and Excel Updater
This Python script allows you to extract table data from emails, find "not working" entries, and update an Excel workbook with the extracted data. It opens outlook using pywin32 lib, searches for specific emails from a particular sender, extracts table data from the email content (assumed to be in HTML format), and updates the Excel workbook with the new data. The script also finds rows with "not working" entries in the "Message" column and copies them to a separate sheet in the workbook.

## Excel Sheets
- Sheet3: Local dynamic database of devices based on the most recent email received.
- Sheet2: List of all the inactive devices based on the most recent email received.
- Sheet1: List of all the alerts received.

## Dependencies
 - `openpyxl`: This library is used to manipulate Excel files (xlsx format).
 - `pandas`: This library is used for data manipulation and analysis.
 - `O365`: This library allows authentication and access to the user's Outlook emails.
 - `logging`: This library helps with error checking and error logging.

Please make sure you have these libraries installed before running the script.

## Configuration
1. Set the `MASTER_WORKBOOK_PATH` to the path of the Excel workbook you want to update. If the workbook is in the same directory as the python file, simply paste the name of the workbook including the extenion (such as .xlsx).
2. Set the `EMAIL_FOLDER_NAME` to the name of the folder where your emails will be located.

## Class
### `EquipmentEntry`
A simple class representing an equipment entry extracted from the email table. It has the following attributes:

- `site`: The site of the equipment.
- `equipment`: The equipment's name.
- `message`: The message associated with the equipment.
- `last_state_change`: The last state change timestamp of the equipment.

## Functions
### `authenticate_account()`
This function uses the O365 variable to authenticate the user and give both basic and mailbox access to the program to be able to read the user's emails from the specified folder.

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
This function handles the scheduling of the driver function as well as user authentication to allow Outlook access. The program runs on a scheduled loop every day at 8:15am.