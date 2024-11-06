#!/usr/bin/env python
# coding: utf-8

# In[57]:


from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import gspread
import os
import calendar
import datetime


# In[58]:


# Specify your desired download folder
current_directory = os.getcwd()
download_folder = os.path.join(current_directory, "Previous Month")

chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_folder,
    "download.prompt_for_download": False,  # To disable download prompt
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Initialize WebDriver with the configured options
driver = webdriver.Chrome(options=chrome_options)


# In[59]:


# Navigate to the webpage
driver.get('https://secure.datafinch.com/')

# Wait for the page and its elements to load
time.sleep(2)  # Consider using WebDriverWait for a more reliable wait

# Locate and fill in the username field
username_field = driver.find_element(By.ID, "Username")
username_field.send_keys("tpc6.ImranK")

# Locate and fill in the password field
password_field = driver.find_element(By.ID, "Password")
password_field.send_keys("Anwarkhan54#")

# Locate and click the login button
# If multiple elements have the same class, consider using a more specific selector or finding all and filtering
login_button = driver.find_element(By.CSS_SELECTOR, "button.btn.btn-primary.login-button")
login_button.click()

# Wait for the "Administration" link to be clickable after logging in
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[@href='/Administration']")))

# Locate and click the "Administration" link
administration_link = driver.find_element(By.XPATH, "//a[@href='/Administration']")
administration_link.click()

# Wait for the "Reports" link to be clickable
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "customReportsLink")))

# Locate and click the "Reports" link
reports_link = driver.find_element(By.ID, "customReportsLink")
reports_link.click()


# Wait for the "Date Range Reports" span to be present and clickable
date_range_reports_xpath = "//span[contains(text(), 'Date Range Reports')]"
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, date_range_reports_xpath)))

# Locate and click the "Date Range Reports" span
date_range_reports = driver.find_element(By.XPATH, date_range_reports_xpath)
date_range_reports.click()


# Wait for the link to be present and clickable based on its href attribute
bulk_timesheet_report_href = "/CustomReports/GetReport?reportId=9f46fca2-916f-45ec-820c-099783fd597f"
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, f"//a[@href='{bulk_timesheet_report_href}']")))

# Locate and click the link
driver.find_element(By.XPATH, f"//a[@href='{bulk_timesheet_report_href}']").click()

time.sleep(10)

# Wait for the first date input field to be present and visible
WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "parameters_0__Value")))

# Format the dates as required by your calendar widget
# Assuming your calendar needs "Select Friday, Mar 1, 2024" format

# Function to format dates
def format_date_for_calendar(date):
    # First, format the date parts individually
    day_of_week = date.strftime("%A")  # Day of the week, e.g., "Friday"
    month = date.strftime("%b")  # Abbreviated month name, e.g., "Mar"
    day = date.strftime("%d")  # Day of the month, e.g., "01"
    year = date.strftime("%Y")  # Year, e.g., "2024"
    
    # Remove leading zero from the day if present
    day = str(int(day))  # This converts the day part to integer, removing any leading zero, and back to string
    
    # Reassemble the formatted date string
    formatted_date = f"Select {day_of_week}, {month} {day}, {year}"
    return formatted_date

# Calculate the first and last day of the previous month
today = datetime.date.today()
first_day_of_this_month = datetime.date(today.year, today.month, 1)
last_day_of_previous_month = first_day_of_this_month - datetime.timedelta(days=1)
first_day_of_previous_month = datetime.date(last_day_of_previous_month.year, last_day_of_previous_month.month, 1)

# Use the function to format the start and end dates
start_date_str = format_date_for_calendar(first_day_of_previous_month)
end_date_str = format_date_for_calendar(last_day_of_previous_month)

print("Start Date:", start_date_str)
print("End Date:", end_date_str)

# Click the first date input field to trigger the calendar popup
first_date_input = driver.find_element(By.ID, "parameters_0__Value")
first_date_input.click()

# Click the button to go to the previous month
prev_month_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//a[@title='Show the previous month'][contains(@class,'datepick-cmd-prev')]"))
)
prev_month_button.click()

# Wait for and click the start date in the calendar
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"//a[@title='{start_date_str}']"))).click()

# Click the second date input field to trigger the calendar popup
second_date_input = driver.find_element(By.ID, "parameters_1__Value")
second_date_input.click()

# Click the button to go to the previous month for the second date input
prev_month_button_second = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//a[@title='Show the previous month'][contains(@class,'datepick-cmd-prev')]"))
)
prev_month_button_second.click()

# Wait for and click the end date in the calendar
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"//a[@title='{end_date_str}']"))).click()


# Click the "Run Report" button
run_report_button = driver.find_element(By.ID, "btnExecute")
run_report_button.click()


# Wait for the report to load (20 seconds in this example)
time.sleep(30)

# Click the "Excel" link to download the report
try:
    excel_link = driver.find_element(By.ID, "exportReport")
    excel_link.click()
except StaleElementReferenceException:
    # In case of StaleElementReferenceException, re-find and click the link
    excel_link = driver.find_element(By.ID, "exportReport")
    excel_link.click()

# Wait a bit for the download to finish
time.sleep(20)  # Adjust this wait time based on your download speed

# Close the WebDriver
driver.quit()


# In[60]:


# Assuming download_folder is defined and contains the path to your download directory
def get_latest_file(path):
    """Function to get the most recent file in a directory."""
    files = [os.path.join(path, f) for f in os.listdir(path) if f.endswith('.xlsx')]
    return max(files, key=os.path.getctime)

# Use the function to get the path of the most recently downloaded Excel file
excel_file_path = get_latest_file(download_folder)

# Load the Excel file into a DataFrame, skipping the first 3 rows and taking the header from the 4th row
df = pd.read_excel(excel_file_path, skiprows=3)

# Drop the 'Signature' column from the DataFrame
df = df.drop(columns=['Signature'])

from datetime import datetime  # This is correct but ensure it's not overwritten or blocked by another import

# Add two new columns to the DataFrame for today's date and today's time
today_date = datetime.now().strftime('%Y-%m-%d')  # Format the date as Year-Month-Day
today_time = datetime.now().strftime('%H:%M:%S')  # Format the time as Hour:Minute:Second

# Append new columns to the DataFrame
df['Date'] = today_date
df['Time'] = today_time


# In[61]:


import gspread
from datetime import datetime, timedelta  # Corrected import statement to include timedelta

# Authenticate with Google Sheets using service account
service_account_file = 'coral-gate-380914-289ef7fcdd78.json'  # Update with your file path
client = gspread.service_account(filename=service_account_file)

# Open the existing spreadsheet by its title
sheet = client.open("Catalyst Reports")

# Determine the previous month and year for the worksheet name
# First, get the first day of the current month
first_day_of_current_month = datetime.now().replace(day=1)
# Then, subtract one day to get a date in the previous month
previous_month_date = first_day_of_current_month - timedelta(days=1)
# Format the previous month date to get the name of the month and year
previous_month_year = previous_month_date.strftime('%B %Y')  # e.g., "Feb 2024"

try:
    # Try to open the worksheet by name (assuming it exists)
    worksheet = sheet.worksheet(previous_month_year)
    # Clear the entire worksheet if it exists
    worksheet.clear()
except gspread.exceptions.WorksheetNotFound:
    # If the worksheet does not exist, create it
    worksheet = sheet.add_worksheet(title=previous_month_year, rows="1000", cols="20")

# Assuming `df` is your DataFrame that you've prepared for updating the worksheet
# Replace `inf` and `-inf` with `NaN`
df.replace([float('inf'), float('-inf')], pd.NA, inplace=True)

# Replace NaN with an empty string "" for compatibility
df.fillna(value="", inplace=True)

# Update the worksheet with df, starting at cell A1
worksheet.update([df.columns.values.tolist()] + df.values.tolist())


# In[62]:


# Now let's append the data to the "Catalyst Logs" sheet
# Open the "Catalyst Logs" worksheet
logs_worksheet = sheet.worksheet("Catalyst Logs")

# Find the first empty row in the "Catalyst Logs" worksheet
logs_data = logs_worksheet.get_all_values()
first_empty_row = len(logs_data) + 1  # The first empty row is after the last filled row

# Get the current date and time (full date-time)
current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # e.g., "2024-09-30 14:05:45"

# Clean the 'Date' column by removing any leading apostrophes and converting to datetime
if df.shape[1] >= 3:
    df.iloc[:, 2] = df.iloc[:, 2].astype(str).str.replace("'", "", regex=False)
    df.iloc[:, 2] = pd.to_datetime(df.iloc[:, 2], errors='coerce')

# Calculate the maximum date from column C of the data
max_date = df.iloc[:, 2].max() if df.shape[1] >= 3 else ""
if pd.notna(max_date):
    max_date = max_date.strftime('%Y-%m-%d')  # Format the maximum date as "YYYY-MM-DD"

# Get the number of records in the data
num_records = len(df)

# Prepare the data to be added (current date, month for which data was uploaded, max date, and number of records)
log_entry = [[current_datetime, previous_month_year, max_date, num_records]]  # Use nested lists for proper range formatting

# Update the "Catalyst Logs" sheet in the first empty row (column 1 for date-time, column 2 for month-year, column 3 for max date, column 4 for number of records)
logs_worksheet.update(f'A{first_empty_row}:D{first_empty_row}', log_entry)

print("Data successfully uploaded and logs updated.")


# In[63]:


from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import gspread
import os
import calendar
import datetime


# In[64]:


# Specify your desired download folder
current_directory = os.getcwd()
download_folder = os.path.join(current_directory, "Current Month")

chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_folder,
    "download.prompt_for_download": False,  # To disable download prompt
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Initialize WebDriver with the configured options
driver = webdriver.Chrome(options=chrome_options)


# In[65]:


# Navigate to the webpage
driver.get('https://secure.datafinch.com/')

# Wait for the page and its elements to load
time.sleep(2)  # Consider using WebDriverWait for a more reliable wait

# Locate and fill in the username field
username_field = driver.find_element(By.ID, "Username")
username_field.send_keys("tpc6.ImranK")

# Locate and fill in the password field
password_field = driver.find_element(By.ID, "Password")
password_field.send_keys("Anwarkhan54#")

# Locate and click the login button
# If multiple elements have the same class, consider using a more specific selector or finding all and filtering
login_button = driver.find_element(By.CSS_SELECTOR, "button.btn.btn-primary.login-button")
login_button.click()

# Wait for the "Administration" link to be clickable after logging in
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[@href='/Administration']")))

# Locate and click the "Administration" link
administration_link = driver.find_element(By.XPATH, "//a[@href='/Administration']")
administration_link.click()

# Wait for the "Reports" link to be clickable
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "customReportsLink")))

# Locate and click the "Reports" link
reports_link = driver.find_element(By.ID, "customReportsLink")
reports_link.click()


# Wait for the "Date Range Reports" span to be present and clickable
date_range_reports_xpath = "//span[contains(text(), 'Date Range Reports')]"
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, date_range_reports_xpath)))

# Locate and click the "Date Range Reports" span
date_range_reports = driver.find_element(By.XPATH, date_range_reports_xpath)
date_range_reports.click()


# Wait for the link to be present and clickable based on its href attribute
bulk_timesheet_report_href = "/CustomReports/GetReport?reportId=9f46fca2-916f-45ec-820c-099783fd597f"
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, f"//a[@href='{bulk_timesheet_report_href}']")))

# Locate and click the link
driver.find_element(By.XPATH, f"//a[@href='{bulk_timesheet_report_href}']").click()

time.sleep(10)

# Wait for the first date input field to be present and visible
WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "parameters_0__Value")))

# Format the dates as required by your calendar widget
# Assuming your calendar needs "Select Friday, Mar 1, 2024" format

# Function to format dates
def format_date_for_calendar(date):
    # First, format the date parts individually
    day_of_week = date.strftime("%A")  # Day of the week, e.g., "Friday"
    month = date.strftime("%b")  # Abbreviated month name, e.g., "Mar"
    day = date.strftime("%d")  # Day of the month, e.g., "01"
    year = date.strftime("%Y")  # Year, e.g., "2024"
    
    # Remove leading zero from the day if present
    day = str(int(day))  # This converts the day part to integer, removing any leading zero, and back to string
    
    # Reassemble the formatted date string
    formatted_date = f"Select {day_of_week}, {month} {day}, {year}"
    return formatted_date

# Calculate the first and last day of the current month
today = datetime.date.today()
first_day_of_month = datetime.date(today.year, today.month, 1)
last_day_of_month = datetime.date(today.year, today.month, calendar.monthrange(today.year, today.month)[1])

# Use the adjusted function to format the start and end dates
start_date_str = format_date_for_calendar(first_day_of_month)
end_date_str = format_date_for_calendar(last_day_of_month)

print("Start Date:", start_date_str)
print("End Date:", end_date_str)







# Click the first date input field to trigger the calendar popup
first_date_input = driver.find_element(By.ID, "parameters_0__Value")
first_date_input.click()

# Wait for and click the start date in the calendar
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"//a[@title='{start_date_str}']"))).click()

# Click the second date input field to trigger the calendar popup
second_date_input = driver.find_element(By.ID, "parameters_1__Value")
second_date_input.click()

# Wait for and click the end date in the calendar
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"//a[@title='{end_date_str}']"))).click()

# Click the "Run Report" button
run_report_button = driver.find_element(By.ID, "btnExecute")
run_report_button.click()


# Wait for the report to load (20 seconds in this example)
time.sleep(30)

# Click the "Excel" link to download the report
try:
    excel_link = driver.find_element(By.ID, "exportReport")
    excel_link.click()
except StaleElementReferenceException:
    # In case of StaleElementReferenceException, re-find and click the link
    excel_link = driver.find_element(By.ID, "exportReport")
    excel_link.click()

# Wait a bit for the download to finish
time.sleep(20)  # Adjust this wait time based on your download speed

# Close the WebDriver
driver.quit()


# In[66]:


# Assuming download_folder is defined and contains the path to your download directory
def get_latest_file(path):
    """Function to get the most recent file in a directory."""
    files = [os.path.join(path, f) for f in os.listdir(path) if f.endswith('.xlsx')]
    return max(files, key=os.path.getctime)

# Use the function to get the path of the most recently downloaded Excel file
excel_file_path = get_latest_file(download_folder)

# Load the Excel file into a DataFrame, skipping the first 3 rows and taking the header from the 4th row
df = pd.read_excel(excel_file_path, skiprows=3)

# Drop the 'Signature' column from the DataFrame
df = df.drop(columns=['Signature'])

from datetime import datetime  # This is correct but ensure it's not overwritten or blocked by another import

# Add two new columns to the DataFrame for today's date and today's time
today_date = datetime.now().strftime('%Y-%m-%d')  # Format the date as Year-Month-Day
today_time = datetime.now().strftime('%H:%M:%S')  # Format the time as Hour:Minute:Second

# Append new columns to the DataFrame
df['Date'] = today_date
df['Time'] = today_time


# In[67]:


import gspread
from datetime import datetime  # Corrected import statement

# Authenticate with Google Sheets using service account
service_account_file = 'coral-gate-380914-289ef7fcdd78.json'  # Update with your file path
client = gspread.service_account(filename=service_account_file)

# Open the existing spreadsheet by its title
sheet = client.open("Catalyst Reports")

# Determine the current month and year for the worksheet name
current_month_year = datetime.now().strftime('%B %Y')  # e.g., "Mar 2024"

try:
    # Try to open the worksheet by name (assuming it exists)
    worksheet = sheet.worksheet(current_month_year)
    # Clear the entire worksheet if it exists
    worksheet.clear()
except gspread.exceptions.WorksheetNotFound:
    # If the worksheet does not exist, create it
    worksheet = sheet.add_worksheet(title=current_month_year, rows="1000", cols="20")

# Assuming `df` is your DataFrame that you've prepared for updating the worksheet
# Replace `inf` and `-inf` with `NaN`
df.replace([float('inf'), float('-inf')], pd.NA, inplace=True)

# Replace NaN with an empty string "" for compatibility
df.fillna(value="", inplace=True)

# Update the worksheet with df, starting at cell A1
worksheet.update([df.columns.values.tolist()] + df.values.tolist())


# In[68]:


# Now let's append the data to the "Catalyst Logs" sheet
# Open the "Catalyst Logs" worksheet
logs_worksheet = sheet.worksheet("Catalyst Logs")

# Find the first empty row in the "Catalyst Logs" worksheet
logs_data = logs_worksheet.get_all_values()
first_empty_row = len(logs_data) + 1  # The first empty row is after the last filled row

# Get the current date and time (full date-time)
current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # e.g., "2024-09-30 14:05:45"

# Clean the 'Date' column by removing any leading apostrophes and converting to datetime
if df.shape[1] >= 3:
    df.iloc[:, 2] = df.iloc[:, 2].astype(str).str.replace("'", "", regex=False)
    df.iloc[:, 2] = pd.to_datetime(df.iloc[:, 2], errors='coerce')

# Calculate the maximum date from column C of the data
max_date = df.iloc[:, 2].max() if df.shape[1] >= 3 else ""
if pd.notna(max_date):
    max_date = max_date.strftime('%Y-%m-%d')  # Format the maximum date as "YYYY-MM-DD"

# Get the number of records in the data
num_records = len(df)

# Prepare the data to be added (current date, month for which data was uploaded, max date, and number of records)
log_entry = [[current_datetime, current_month_year, max_date, num_records]]  # Use nested lists for proper range formatting

# Update the "Catalyst Logs" sheet in the first empty row (column 1 for date-time, column 2 for month-year, column 3 for max date, column 4 for number of records)
logs_worksheet.update(f'A{first_empty_row}:D{first_empty_row}', log_entry)

print("Data successfully uploaded and logs updated.")


# In[ ]:





# In[ ]:




