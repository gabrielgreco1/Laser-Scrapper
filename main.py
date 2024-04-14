import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
import openpyxl
from openpyxl import Workbook
from datetime import datetime

def setup_driver():
    service = Service(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def log(message):
    """Helper function to print messages with a timestamp."""
    print(f"{datetime.now()}: {message}")

def scrape_data(driver, adid):
    url = f"https://www.coldlasers.org/search/details/index.cfm?adid={adid}"
    log(f'Testing URL ID: {adid}')
    driver.get(url)
    
    try:
        content_html = driver.find_element(By.CLASS_NAME, 'post-content').get_attribute('innerHTML')
        
        name_matches = re.findall(r'<strong>([^<]+)</strong>', content_html)
        name = next((m for m in name_matches if "Go back" not in m and "Laser System" not in m), '')

        full_address_match = re.search(r'\d+ .+? [A-Z]{2} \d{5}', content_html)
        phone_match = re.search(r'\(\d{3}\) \d{3}-\d{4}', content_html)
        website_match = re.search(r'http://www\.[^/:]+/?', content_html)
        laser_system_match = re.search(r'Laser System:[^<]*<[^>]*>([^<]+)', content_html)

        full_address = full_address_match.group(0) if full_address_match else ''
        phone = phone_match.group(0) if phone_match else ''
        website = website_match.group(0) if website_match else ''
        laser_system = laser_system_match.group(1).strip() if laser_system_match else ''
        
        # Split the address into components
        if full_address:
            address_parts = full_address.rsplit(' ', 3)
            street = ' '.join(address_parts[:-3])
            city = address_parts[-3].replace(',', '')
            state = address_parts[-2]
            zip_code = address_parts[-1]
        else:
            street, city, state, zip_code = '', '', '', ''

        log(f'Extracted: {name}, {street}, {city}, {zip_code}, {state}, {phone}, {website}, {laser_system}')
        
        return [name, street, city, zip_code, state, phone, website, laser_system]
        
    except NoSuchElementException:
        log('Element not found')
        return []
    except Exception as e:
        log(f"An error occurred: {e}")
        return []

def main(start_id, end_id):
    driver = setup_driver()
    wb = Workbook()
    ws = wb.active
    ws.append(['Name', 'Street', 'City', 'Zip-Code', 'State', 'Phone', 'Website', 'Laser System'])  # Excel headers
    
    for adid in range(start_id, end_id + 1):
        data = scrape_data(driver, adid)
        if data:  # If data is found, add to Excel
            ws.append(data)
    
    excel_file = "ColdLaserData.xlsx"
    wb.save(excel_file)
    driver.quit()
    
    return excel_file

if __name__ == "__main__":
    excel_file = main(3900, 3910)
    log(f"Scraping complete. File saved as {excel_file}")
