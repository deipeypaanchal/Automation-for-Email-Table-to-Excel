**Documentation and Purpose of the Code**

**Purpose:**
The purpose of this code is to automate the process of extracting, parsing, and updating equipment-related information received via emails, and then storing this information in a master Excel workbook. The code is designed to handle emails containing HTML tables that contain equipment data, and it performs the following main tasks:

1. **Email Retrieval and Processing:**
   - Connects to an IMAP email server using provided credentials.
   - Searches for emails from a specific sender's email address.
   - Downloads and processes emails to extract their content, specifically focusing on HTML content.
   - Extracts equipment-related information from the HTML tables within the email.

2. **Excel Workbook Handling:**
   - Loads a master Excel workbook that contains three sheets: `Sheet1`, `Sheet2`, and `Sheet3`.
   - Handles duplicate entries to avoid adding duplicate rows to `Sheet1`.

3. **Data Parsing and Updating:**
   - Parses the extracted HTML table data and creates a list of equipment entries.
   - Compares the extracted equipment entries with the existing entries in `Sheet3`.
   - Updates the existing equipment entries with new data if there are any changes.

4. **Data Storage:**
   - Converts the updated equipment data into a Pandas DataFrame.
   - Writes the updated equipment data to `Sheet3` of the master Excel workbook.

5. **Additional Data Processing:**
   - Identifies rows in `Sheet3` where the "Message" column contains "not working".
   - Appends these identified rows to `Sheet2`.

6. **Finalization:**
   - Saves the modified master Excel workbook.
   - Closes the workbook and disconnects from the email server.
   - Prints a message indicating that the table data has been added/updated in the master worksheet.

**Documentation:**

1. **Libraries Used:**
   - `imaplib`: Provides functionality to interact with IMAP email servers.
   - `email`: Helps in handling email messages.
   - `openpyxl`: Allows manipulation of Excel workbooks.
   - `pandas` (`pd` alias): Provides data manipulation capabilities using DataFrames.
   - `BeautifulSoup` (from `bs4`): A library for parsing HTML and XML content.

2. **Configuration:**
   - `IMAP_SERVER`: The IMAP server's hostname.
   - `EMAIL_ADDRESS`: The email address for authentication.
   - `PASSWORD`: The app password for authentication.
   - `SENDER_EMAIL`: The sender's email address to filter emails.
   - `MASTER_WORKBOOK_PATH`: The path to the master Excel workbook.

3. **Class:**
   - `EquipmentEntry`: Represents an equipment entry with attributes like `site`, `equipment`, `message`, and `last_state_change`.

4. **Functions:**
   - `parse_html_email`: Parses HTML content to extract equipment data, returns a list of `EquipmentEntry` instances.
   - `append_table_data_to_worksheet1`: Appends parsed table data to `Sheet1`, avoiding duplicates.
   - `update_table_data`: Updates existing entries in `Sheet3` with new data.
   - `read_existing_data_from_excel`: Reads existing equipment entries from `Sheet3` and returns a list of `EquipmentEntry` instances.
   - `main`: The main function that orchestrates the entire process, including email retrieval, data parsing, updating, storage, and finalization.

5. **Execution:**
   - The `main` function is executed when the script is run (`if __name__ == "__main__": main()`).
   - It connects to the email server, processes emails, updates the workbook, and then disconnects.

**Note:**
- The code assumes a specific structure of the HTML content in the emails, particularly the presence of tables containing equipment information.
- The code assumes that the equipment data table in the email contains columns for "Site," "Equipment," "Message," and "Last State Change."
- Make sure to replace placeholders (`IMAP_SERVER`, `EMAIL_ADDRESS`, etc.) with actual values before running the code.
