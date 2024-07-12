from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os

app = Flask(__name__)
app.secret_key = 'aP9sB7xZ6kLmN8pQ3rT4uV5wX1yO2d'

# Setup Selenium WebDriver
chrome_driver_path = 'C:\\Users\\n\\Downloads\\chromedriver-win64\\chromedriver.exe'
service = Service(chrome_driver_path)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file part')
        return redirect(request.url)
    
    file = request.files['file']
    if file.filename == '':
        flash('No selected file')
        return redirect(request.url)
    
    if file:
        # Save the file to the uploads directory
        filepath = os.path.join('uploads', file.filename)
        file.save(filepath)
        
        # Load the Excel file
        df = pd.read_excel(filepath)
        users_data = df.to_dict(orient='records')
        
        # Open the webpage using Selenium
        driver = webdriver.Chrome(service=service)
        driver.get('https://openai.com/6a0bdd42abee7/')
        
        results = []
        for user in users_data:
            try:
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, "//input[@name='firstname']"))
                )

                # Fill out the form
                first_name = driver.find_element(By.XPATH, "//input[@name='firstname']")
                first_name.clear()
                first_name.send_keys(user['First name'])
                
                last_name = driver.find_element(By.XPATH, "//input[@name='lastname']")
                last_name.clear()
                last_name.send_keys(user['Last name'])
                
                email = driver.find_element(By.XPATH, "//input[@name='email']")
                email.clear()
                email.send_keys(user['Business email'])
                
                company = driver.find_element(By.XPATH, "//input[@name='company']")
                company.clear()
                company.send_keys(user['Company'])
                
                country = driver.find_element(By.XPATH, "//select[@name='country']")
                country.send_keys(user['Country'])
                
                existing_work = driver.find_element(By.XPATH, "//textarea[@name='existing_work']")
                existing_work.clear()
                existing_work.send_keys(user['Existing work'])
                
                openai_api_organization_id = driver.find_element(By.XPATH, "//input[@name='openai_api_organization_id']")
                openai_api_organization_id.clear()
                openai_api_organization_id.send_keys(user['OpenAI API Organization ID'])
                
                submit_button = driver.find_element(By.XPATH, "//button[@type='submit']")
                submit_button.click()

                WebDriverWait(driver, 20).until(EC.staleness_of(submit_button))

                results.append(f"An error occurred for user {user['First name']} {user['Last name']}: {e}")
            except Exception as e:
                results.append(f"Successfully submitted form for user: {user['First name']} {user['Last name']}")
        
        driver.quit()
        return render_template('results.html', results=results)

    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
