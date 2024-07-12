from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import openpyxl
import time

# Function to read Excel data
def read_excel_data(filepath):
    wb = openpyxl.load_workbook(filepath)
    sheet = wb.active
    data = []
    for row in sheet.iter_rows(min_row=2,values_only=True):
        data.append({
            'First name': row[0],
            'Last name': row[1],
            'Business email': row[2],
            'Company name': row[3],
            'Country': row[4],
            'Message': row[5],
            'Organization ID': row[6]  # Assuming Organization ID is in the 7th column (index 6)
        })
    return data

# Initialize WebDriver
driver = webdriver.Chrome()  # Adjust as per your WebDriver setup

# Open the web page
driver.get("https://openai.com/6a0bdd42abee7/")  # URL of the Microsoft for Startups form

# Read data from Excel file
excel_file = 'C:\\Users\n\\Downloads\\Batch 1.xlsx'  # Update with your Excel file path
users_data = read_excel_data(excel_file)

# Fill out the form for each user
for user in users_data:
    driver.find_element_by_name("First name").send_keys(user['First name'])
    driver.find_element_by_name("Last name").send_keys(user['Last name'])
    driver.find_element_by_name("Business email").send_keys(user['Business email'])
    driver.find_element_by_name("Company name").send_keys(user['Company name'])
    
    # Select Country of Residence
    select_country = Select(driver.find_element_by_name("Country"))
    select_country.select_by_visible_text(user['Kenya'])  # Assuming country names match exactly
    
    # Fill in the message
    driver.find_element_by_name("Message").send_keys(user['Message'])
    
    # Fill in the Organization ID
    driver.find_element_by_name("Organization ID").send_keys(user['Organization ID'])
    
    # Submit the form
    driver.find_element_by_css_selector("button[type='submit']").click()
    
    # Wait for a few seconds for the next form or observation (optional)
    time.sleep(3)

# Close the browser
driver.quit()
