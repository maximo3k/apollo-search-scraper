import csv
import json
import re
import time
from datetime import datetime

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService


# How to use:
# virtual environment: python -m venv virt
# Create CSV that you want to write in.
# Put the time between actions high. Run the script and when the chrome window opens, log in. There are issues with Chrome profiles. Need to log in manually.
# Then you can cancel the application and finetune the timers, according to your Internet Connection

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
driver.implicitly_wait(2)


def login_to_site(driver, config):
    print('Starting login process...')
    driver.get('https://app.apollo.io/#/login')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "button[type='submit']")))

    email_field = driver.find_element(By.NAME, "email")
    password_field = driver.find_element(By.NAME, "password")
    email_field.send_keys(config['email'])
    password_field.send_keys(config['password'], Keys.RETURN)

    time.sleep(2)  # Increase the time, if you need to hande a security check
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    driver.get(config['start_url'])

def find_and_copy_email(driver):
    # There can be 2 email addresses
    # Copy button cannot be used, because it will close the popup, will have to open the popup again to get 2nd email.

    try:
        # wait till popup is loaded
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, "zp_SZG4_")))

        popup_container = driver.find_element(By.CLASS_NAME, "zp_SZG4_")

        # the emails are in span elements
        spans = popup_container.find_elements(By.TAG_NAME, "span")

        email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')

        emails = []

        for span in spans:
            if email_pattern.match(span.text):
                emails.append(span.text)

    except Exception as e:
        print(f"Failed to copy email: {str(e)}")
        return None
    return emails

def next_page(driver):
    # Handling pagination
    try:

        next_page_button = driver.find_element(By.CSS_SELECTOR, "[aria-label='right-arrow']")

        if next_page_button:
            next_page_button.click()
            print("Next page loading")
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "[data-cy-loaded='true']")))
        else:
            print("Reached the last page.")
            return False
    except NoSuchElementException:
        print("No more pages to navigate.")
        return False 

    return True

def split_name(name):
    parts = name.split()
    first_name = parts[0] if parts else ''
    last_name = ' '.join(parts[1:]) if len(parts) > 1 else ''
    return first_name, last_name

def write_to_csv(config, first_name, last_name, job_title, company_name, location, email_address1, email_address2, linkedin_url):
    csv_file_path = config['export_file_name']
    date_now = datetime.now().strftime("%Y-%m-%d")

    headers = ['first_name', 'last_name', 'job_title', 'company_name', 'location', 'email_address1', 'email_address2', 'linkedin_url', 'date']
    data_row = [first_name, last_name, job_title, company_name, location, email_address1, email_address2, linkedin_url, date_now ]

    try:
        with open(csv_file_path, 'x', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(headers)
            writer.writerow(data_row)
    except FileExistsError:
        with open(csv_file_path, 'a+', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(data_row)

def main(driver):
    with open('config.json', 'r') as file:
        config = json.load(file)

    login_to_site(driver, config)

    try:
        iteration_count = 0
        while True:
            iteration_count += 1
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "[data-cy-loaded='true']")))
            loaded_section = driver.find_element(By.CSS_SELECTOR, "[data-cy-loaded='true']")
            tbodies = loaded_section.find_elements(By.TAG_NAME, 'tbody')

            # this is the base for virgin & non-virgin button
            email_button_classes_universal = ["zp-button", "zp_zUY3r", "zp_n9QPr", "zp_MCSwB"]
            email_button_classes_copy = ["zp-button", "zp_zUY3r", "zp_r4MyT", "zp_LZAms", "zp_jRaZ5"]
            phone_classes = ["zp-link", "zp0otKe", "zp_vc37T"]
            
            if not tbodies:
                print("No data to process.")
                break
            
            # go through each item in tbodies
            for tbody in tbodies:

                # define all variables
                first_name = "",
                last_name = "",
                linkedin_url = "",
                job_title = "",
                company_name = "",
                location = "",
                email_address1 = "",
                email_address2 = "",


                # getting the name
                first_anchor_text = tbody.find_element(By.TAG_NAME, 'a').text
                first_name, last_name = split_name(first_anchor_text)

                # getting the linkedin, loop through the <a>, stop once we find a link that includes "linkedin"
                # maybe make this a separate function
                for link in tbody.find_elements(By.TAG_NAME, 'a'):
                    href = link.get_attribute('href')
                    if 'linkedin.com' in href:
                        linkedin_url = href
                        break

                job_title_location_elements = tbody.find_elements(By.CLASS_NAME, 'zp_Y6y8d')
                job_title = job_title_location_elements[0].text
                location = job_title_location_elements[1].text


                # loop through <a> again and only stop if the link includes 
                for link in tbody.find_elements(By.TAG_NAME, 'a'):
                    if 'accounts' or 'organizations' in link.get_attribute('href'):
                        company_name = link.text
                        break

                # Here we need to open the field for email
                # There can be 2 ways, a contact that has already been clicked (virgin) and one where you have to spend a credit to unveil the email (non-virgin)
                try:
                    # find the button that has the universal button classes
                    email_button = tbody.find_element(By.CSS_SELECTOR, "." + ".".join(email_button_classes_universal))

                    # check if the classes for virgin button are included. 
                    non_virgin_button = "zp_IYteB" in email_button.get_attribute("class")

                    # check button if non-virgin
                    if non_virgin_button:
                        print("non-virgin button detected")
                        # open the popup by a click
                        email_button.click()
                        # wait till the copy button loads, this ensures the email is also visible
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "." + ".".join(email_button_classes_copy))))
                        email_addresses = find_and_copy_email(driver)

                        # close the popup, since nothing has been loaded from the database, the DOM element is still active, it can be clicked.
                        email_button.click()

                        # scroll down, by 1 item
                        tbody_height = driver.execute_script("return arguments[0].offsetHeight;", tbody)
                        driver.execute_script("arguments[0].scrollBy(0, arguments[1]);", loaded_section, tbody_height)

                    # button is virgin
                    else:
                        print("virgin button detected")
                        email_button.click()
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "." + ".".join(email_button_classes_copy))))
                        email_addresses = find_and_copy_email(driver)

                        # need 2 clicks to close the popup, if it is virgin
                        # the DOM element becomes stale after it is loading the email address from database
                        # need to find the DOM element again to ensure the click can be done to close the popup
                        email_button = tbody.find_element(By.CSS_SELECTOR, "." + ".".join(email_button_classes_universal))
                        email_button.click()
                        email_button.click()

                        tbody_height = driver.execute_script("return arguments[0].offsetHeight;", tbody)
                        driver.execute_script("arguments[0].scrollBy(0, arguments[1]);", loaded_section, tbody_height)

                    email_address1 = email_addresses[0] if len(email_addresses) > 0 else ''
                    email_address2 = email_addresses[1] if len(email_addresses) > 1 else ''

                except Exception as e:
                    print('There was a problem with data extraction from the popup \n {e}')
                    time.sleep(2)
                    continue

                write_to_csv(config, first_name, last_name, job_title, company_name, location, email_address1, email_address2, linkedin_url)
                print(f"Processed: {first_name} {last_name}, {job_title}, {company_name}, {location}, {email_address1}, {email_address2}, {linkedin_url}")

            if not next_page(driver):
                return None

    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main(driver)