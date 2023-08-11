# Automation Project: Extract Email Table to Excel

Authors: Deipey Paanchal and Brandon Pyle

## Purpose
This program was written in order to automate the process of moving tables from Outlook emails to an excel file.

There are 3 different versions of this code using different libraries:
  1. using O365
  2. using imaplib
  3. using win32

# O365
This is the most functional and best version of the program as it allows the user to authenticate their account with Outlook, thereby ebnable the use of company accounts (with appropriate level of access). If using this program in a company with company-assigned emails, this is the way to go. This version, however, only works with outlook emails.

# imaplib
This was the first version of the program and uses imaplib to connect to email. The benefit of this is that it can work with any personal email, whether it's Gmail, Outlook, etc. However, this version will likely not work with company emails as they will generally block imap access.

# win32
This version uses the win32 library, which limits it to Windows users only. It requires your email to be set up within outlook, but can be used with any email provider. Similar to imaplib, company emails will not work with this version.
