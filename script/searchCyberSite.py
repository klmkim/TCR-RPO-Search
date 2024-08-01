'''
searchCyberSite.py

Python script to scrape and parse data from a website

The script will ask for a URL and a name for an Excel file to output. 
Make sure not to accidentally overwrite existing Excel files.

The script will visit the URL, search the site for all company details, and extract them to an Excel file.

Written by Klimentiy Kim for Technical Consulting & Research Inc.
'''

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd

# Function to get user input for URL and Excel file name
def get_user_input():
    url = input("Please enter the URL of the website you want to scrape: ")
    excel_file_name = input("Please enter the name of the Excel file to save the data. (e.g., 'CMMC RPOs Eastern'): ") + '.xlsx'
    return url, excel_file_name

# Get user input
url, excel_file_name = get_user_input()

# Set up the WebDriver (this example uses Chrome)
driver = webdriver.Chrome()  # Ensure chromedriver is in your PATH or provide the path to chromedriver

# Open the website
driver.get(url)

companies = []

try:
    # Wait until all elements are present
    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.col-12.mb-3.pointer.ng-scope"))
    )

    # Find all the elements
    elements = driver.find_elements(By.CSS_SELECTOR, "div.col-12.mb-3.pointer.ng-scope")

    # Click each element and extract details
    for element in elements:
        element.click()

        # Wait for the new content to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div#results-list-sidebar.row.ng-isolate-scope"))
        )

        # Extract page source
        page_source = driver.page_source

        # Parse the page source with BeautifulSoup
        soup = BeautifulSoup(page_source, 'html.parser')

        # Find the relevant div containing company details
        cards = soup.select('div#results-list-sidebar.row.ng-isolate-scope')

        # Extract company details
        for card in cards:
            name = card.select_one('h3.mt-5').text.strip() if card.select_one('h3.mt-5') else "N/A"
            address = card.select_one('.contact-address-line .ng-binding').text.strip() if card.select_one('.contact-address-line .ng-binding') else "N/A"
            email = card.select_one('a[href^="mailto:"]').text.strip() if card.select_one('a[href^="mailto:"]') else "N/A"
            website = card.select_one('a[href^="http"]').text.strip() if card.select_one('a[href^="http"]') else "N/A"
            
            companies.append({
                'name': name,
                'address': address,
                'email': email,
                'website': website
            })
            print(name, website)

finally:
    # Close the WebDriver
    driver.quit()

# Create a DataFrame outside the loop
df = pd.DataFrame(companies)

# Write to an Excel file
df.to_excel(excel_file_name, index=False)

print(f"Data has been written to {excel_file_name}")
