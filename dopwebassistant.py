import os
import io
import glob
import shutil
import json
import time
import logging
from datetime import datetime

import pandas as pd
import requests
from PIL import Image, ImageFilter

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import (
    NoSuchElementException, WebDriverException
)

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.firefox.options import Options as FirefoxOptions

from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager

from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.firefox.service import Service as FirefoxService

import browsers



# Configure the logging settings
logging.basicConfig(filename='doplogs.log' ,level=20, format='%(asctime)s - %(levelname)s - %(message)s')


# Custom Error Class
class LotTaskError(Exception):
    pass

# Custom Error Class
class LoginError(Exception):
    pass

# Custom Error Class
class OpenBrowserError(Exception):
    pass

# Custom Error Class
class ElementNotFound(Exception):
    pass

# Custom Error Class
class DownloadTaskError(Exception):
    pass

# Custom Error Class
class UpdateAslaasError(Exception):
    pass

# Custom Error Class
class UpdateTaskError(Exception):
    pass

# Custom Error Class
class NoSupportedBrowserFound(Exception):
    pass


# DOPWebAssistant class to perform tasks on web portal of DOP agent indiapost
class DOPWebAssistant:
    def __init__(self):
        self.user_id = "USERID"
        self.user_password = "PASSWORD"
        self.create_folder_if_not_exists('RDRecord')
        self.temp_download_dir = self.create_folder_if_not_exists('temp')
        self.driver = None
        self.ocr_apikey = "APIKEY"
        settings_data = {
            'user_id': "USERID",
            'user_password': "PASSWORD",
            'def_download_dir': "./",
            'agent_id': "AGENTID",
            'lisence_exp_date': "EXPDATE",
            'agent_address': "ADDRESS",
            'agent_name': "AGENTNAME",
            'agent_husband_name': "AGENTHUSBANDNAME",
            'ocr_apikey':"APIKEY",
            'ocr_apikey':"APIKEY",
            "theme": "Dark",
            "ascent": "blue",
            "scale": "0"
        }
        try:
            with open('settings.json', 'x') as file:
                json.dump(settings_data, file)
        except FileExistsError:
            pass

        try:
            with open('settings.json', 'r') as file:
                settings_data = json.load(file)
            self.ocr_apikey = settings_data.get('ocr_apikey','')
        except FileNotFoundError:
            # Handle the case where the file is not found (e.g., first-time setup)
            pass


    # Function to check if a browser is installed or not
    def is_browser_installed(self, browser_executable_name):
        """
        Checks if a browser is installed on the system by its executable name.

        Parameters:
        browser_executable (str): The name of the browser executable to check (e.g., 'chrome', 'msedge', 'firefox').

        Returns:
        bool: True if the browser is installed, False otherwise.
        """
        # Check if the specified browser executable exists in the browsers dictionary
        browser = browsers.get(browser_executable_name)

        # Return True if the browser exists, otherwise return False
        return browser is not None


    # Takes temporary download path and sets up the driver, then returns the driver variable
    def setup_driver(self, temp_download_dir):
        """
        Sets up the WebDriver for a supported browser and returns the driver instance.

        Parameters:
        temp_download_dir (str): The temporary download directory path to be used by the WebDriver.

        Returns:
        WebDriver: The initialized WebDriver instance for the supported browser.

        Raises:
        NoSupportedBrowserFound: If no supported browser is found on the system.
        """
        # Check if Firefox is installed and set up its WebDriver
        if self.is_browser_installed('firefox'):
            return self.setup_driver_firefox(temp_download_dir)
        
        # Check if Chrome is installed and set up its WebDriver
        elif self.is_browser_installed('chrome'):
            return self.setup_driver_chrome(temp_download_dir)
        
        # Check if Microsoft Edge is installed and set up its WebDriver
        elif self.is_browser_installed('msedge'):
            return self.setup_driver_edge(temp_download_dir)
        
        # Raise an exception if no supported browser is found
        else:
            raise NoSupportedBrowserFound("No supported browser found")


    def setup_driver_firefox(self, temp_download_dir):
        """
        Sets up the Firefox WebDriver with specific options and returns the driver instance.

        Parameters:
        temp_download_dir (str): The temporary download directory path to be used by Firefox.

        Returns:
        WebDriver: The initialized Firefox WebDriver instance.
        """
        # Initialize Firefox options
        options_firefox = FirefoxOptions()
        
        # Set the download directory to the specified temporary download directory
        options_firefox.set_preference("browser.download.dir", temp_download_dir)
        
        # Set the download folder list to use the custom directory
        options_firefox.set_preference("browser.download.folderList", 2)
        
        # Disable showing the download manager when starting a download
        options_firefox.set_preference("browser.download.manager.showWhenStarting", False)
        
        # Set the MIME types to save to disk without asking
        options_firefox.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/octet-stream,application/vnd.ms-excel")
        
        # Initialize the Firefox WebDriver with the specified options and service
        driver = webdriver.Firefox(options=options_firefox, service=FirefoxService(GeckoDriverManager().install()))
        
        return driver


    # Define the setup_driver_* functions using webdriver-manager (Chrome)
    def setup_driver_chrome(self, temp_download_dir):
        """
        Sets up the Chrome WebDriver with specific options and returns the driver instance.

        Parameters:
        temp_download_dir (str): The temporary download directory path to be used by Chrome.

        Returns:
        WebDriver: The initialized Chrome WebDriver instance.
        """
        # Initialize Chrome options
        options_chrome = ChromeOptions()
        
        # Set Chrome preferences for downloads
        prefs = {
            "download.default_directory": temp_download_dir,   # Set the default download directory
            "download.prompt_for_download": False,             # Disable download prompt
            "download.directory_upgrade": True,                # Allow directory upgrade
            "safebrowsing.enabled": True                       # Enable safe browsing
        }
        options_chrome.add_experimental_option("prefs", prefs)
        
        # Initialize the Chrome WebDriver with the specified options and service
        driver = webdriver.Chrome(options=options_chrome, service=ChromeService(ChromeDriverManager().install()))
        
        return driver


    # Define the setup_driver_* functions using webdriver-manager (Edge)
    def setup_driver_edge(self, temp_download_dir):
        """
        Sets up the Edge WebDriver with specific options and returns the driver instance.

        Parameters:
        temp_download_dir (str): The temporary download directory path to be used by Edge.

        Returns:
        WebDriver: The initialized Edge WebDriver instance.
        """
        # Initialize Edge options
        options = EdgeOptions()
        
        # Set Edge preferences for downloads
        prefs = {
            "download.default_directory": temp_download_dir,   # Set the default download directory
            "download.prompt_for_download": False,             # Disable download prompt
            "download.directory_upgrade": True                 # Allow directory upgrade
        }
        options.add_experimental_option("prefs", prefs)
        
        # Initialize the Edge WebDriver with the specified options and service
        driver = webdriver.Edge(options=options, service=EdgeService(EdgeChromiumDriverManager().install()))
        
        return driver


    # Takes Folder name and creates a folder of that name if it does not exist and returns full path
    def create_folder_if_not_exists(self, folder_name):
        """
        Creates a folder with the given name if it does not already exist, or
        deletes and recreates it if it does exist, and returns the full path.

        Parameters:
        folder_name (str): The name of the folder to create or recreate.

        Returns:
        str: The absolute path of the created folder.
        """
        # Get the absolute path of the folder
        full_path = os.path.abspath(folder_name)
        
        # Check if the folder already exists
        if os.path.exists(full_path):
            shutil.rmtree(full_path)  # Delete the folder and all its contents
        
        # Create the folder
        os.makedirs(full_path)
        
        # Return the full path of the created folder
        return full_path


    # Creates a driver and opens browser and goes to dop portal
    def open_browser_portal(self):
        """
        Creates a driver, opens the browser, and navigates to the DOP portal.

        This function sets up the appropriate web driver, configures browser settings,
        and opens the browser to navigate to the specified URL. In case of any 
        exceptions, it handles them appropriately, logs the errors, and raises a 
        custom error.

        Raises:
        OpenBrowserError: If an error occurs while opening the browser.
        """
        try:
            # Set up the driver with the specified download directory
            self.driver = self.setup_driver(self.temp_download_dir)
            self.driver.implicitly_wait(20)  # Set implicit wait time
            self.driver.maximize_window()  # Maximize the browser window
            self.driver.get("https://dopagent.indiapost.gov.in")  # Navigate to the specified URL
            logging.info("Opened the browser and navigated to the website successfully!")

        except WebDriverException as we:
            # Log the WebDriverException and raise a custom error
            logging.error(f"WebDriverException during opening browser: {we}")
            self.close_browser()  # Close the browser
            raise OpenBrowserError(f"Error while opening browser: {we}")

        except Exception as e:
            # Log any other exceptions and raise a custom error
            logging.error(f"Error during opening browser: {e}")
            self.close_browser()  # Close the browser
            raise OpenBrowserError(f"Error while opening browser: {e}")


    # Log in to the dop agent portal
    def login(self):
        """
        Logs into the DOP agent portal using provided credentials and CAPTCHA solving.

        This function attempts to log in by filling out the username, password,
        and CAPTCHA fields on the DOP agent portal login page. It retries until
        successful login or encounters an exception.

        Raises:
        LoginError: If an error occurs during the login process.
        """
        login_suc = False  # Flag to track successful login
        try:
            while not login_suc:
                # Find and clear the username field, then enter user_id
                username = self.driver.find_element(By.NAME, "AuthenticationFG.USER_PRINCIPAL")
                username.clear()
                username.send_keys(self.user_id)
                logging.info("Username inserted successfully!")

                # Find and clear the password field, then enter user_password
                password = self.driver.find_element(By.NAME, "AuthenticationFG.ACCESS_CODE")
                password.clear()
                password.send_keys(self.user_password)
                logging.info("Password inserted successfully!")

                # Find the CAPTCHA image element and capture its screenshot
                element = self.driver.find_element(By.ID, "IMAGECAPTCHA")
                png = element.screenshot_as_png

                # Process the CAPTCHA image using OCR to retrieve text
                image = Image.open(io.BytesIO(png))
                image = image.filter(ImageFilter.MedianFilter(size=3))
                image = image.convert('L')
                self.create_folder_if_not_exists('temp')
                image.save('./temp/image.png', format='PNG')

                try:
                    text = self.ocr_space_file(filename='./temp/image.png')
                except:
                    text = ""

                # Enter the CAPTCHA text into the input field
                captcha_input = self.driver.find_element(By.ID, 'AuthenticationFG.VERIFICATION_CODE')
                captcha_input.clear()
                captcha_input.send_keys(text)
                logging.info("CAPTCHA solved successfully!")

                # Wait for user to complete CAPTCHA solving (adjust time if needed)
                time.sleep(10)

                # Click the submit button to validate credentials and CAPTCHA
                submit_button = self.driver.find_element(By.ID, 'VALIDATE_RM_PLUS_CREDENTIALS_CATCHA_DISABLED')
                submit_button.click()
                logging.info("Submitted login details successfully!")

                try:
                    # Check if login was successful by finding the Accounts button
                    self.driver.find_element(By.ID, 'Accounts')
                    login_suc = True  # Set login_suc to True if Accounts button Not found
                    os.remove('./temp/image.png')
                    logging.info("Login successful!")
                except NoSuchElementException:
                    logging.info("Login unsuccessful. Trying to login again...")
                    continue  # Retry login process if Accounts button not found

        except Exception as login_error:
            # Log any errors during login and raise a custom LoginError
            logging.error(f"Error during login: {login_error}")
            self.close_browser()  # Close the browser in case of error
            raise LoginError("Error while logging in to account.")


    # Takes the image file and returns the solved captcha text using OCRSpace API
    def ocr_space_file(self, filename, overlay=False, language='eng'):
        """
        Performs OCR (Optical Character Recognition) on the given image file using the OCRSpace API.

        Args:
        filename (str): The path to the image file to be processed.
        overlay (bool, optional): Whether to overlay the result on the image. Defaults to False.
        language (str, optional): The language code for OCR processing. Defaults to 'eng' (English).

        Returns:
        str: The extracted text from the image after OCR processing.

        Raises:
        Exception: If there is an error during the API request or parsing the response.
        """
        payload = {
            'isOverlayRequired': overlay,
            'apikey': self.ocr_apikey,
            'language': language,
        }

        try:
            with open(filename, 'rb') as f:
                r = requests.post('https://api.ocr.space/parse/image',
                                files={filename: f},
                                data=payload)

            response = r.content.decode()
            data = json.loads(response)
            parsed_text = data["ParsedResults"][0]["ParsedText"].strip()
            return parsed_text

        except Exception as e:
            # Log any errors encountered during OCR processing and raise an exception
            logging.error(f"Error during OCR processing: {e}")
            raise Exception("Error during OCR processing.")

    # Close the Browser
    def close_browser(self):
        """
        Closes the web browser if it is currently open.

        This method safely quits the WebDriver session, releasing all associated resources.

        Raises:
        WebDriverException: If there is an issue while quitting the WebDriver session.
        """
        if self.driver:
            try:
                self.driver.quit()
                logging.info("Browser closed successfully.")
            except WebDriverException as e:
                logging.error(f"WebDriverException while closing browser: {e}")
                raise WebDriverException("Error occurred while closing the browser.")


    # Takes two arrays account numbers and no of installments and performs lot
    def perform_lot_task(self, acc_nos, no_installments):
        """
        Performs the LOT task on the DOP agent portal.

        Args:
        acc_nos (list): List of account numbers to perform LOT.
        no_installments (list): List of number of installments corresponding to each account.

        Returns:
        str: Reference number extracted after completing the LOT task.

        Raises:
        LotTaskError: If there is an error during any step of the LOT task.
        """
        try:
            # Navigate to Accounts page
            accounts_button = self.driver.find_element(By.ID, 'Accounts')
            accounts_button.click()
            logging.info("Navigated to Accounts page successfully!")

            # Calculate page navigation variables
            n = len(acc_nos)
            elem_last_page = n % 10
            no_full_pages = n // 10
            no_pages = no_full_pages + (1 if elem_last_page > 0 else 0)
            if elem_last_page == 0:
                elem_last_page = 10

            # Navigate to Agent Enquire & Update Screen
            agent_enquire_button = self.driver.find_element(By.ID, 'Agent Enquire & Update Screen')
            agent_enquire_button.click()
            logging.info("Navigated to Agent Enquire & Update Screen page successfully!")

            # Enter account numbers and fetch accounts
            text_box = self.driver.find_element(By.ID, 'CustomAgentRDAccountFG.ACCOUNT_NUMBER_FOR_SEARCH')
            text_box.send_keys(','.join(map(str, acc_nos)))
            fetch_button = self.driver.find_element(By.ID, 'Button3087042')
            fetch_button.click()
            logging.info("Accounts fetched successfully!")

            # Select Cash mode of Payment
            cash_button = self.driver.find_element(By.XPATH, '//input[@id="CustomAgentRDAccountFG.PAY_MODE_SELECTED_FOR_TRN"][@value="C"]')
            cash_button.click()
            logging.info("Cash mode of Payment selected successfully!")

            # Select all accounts and save the LOT
            for i in range(no_pages):
                page_limit = elem_last_page if i == no_pages - 1 else 10
                
                for j in range(page_limit):
                    try:
                        checkbox = self.driver.find_element(By.ID, f'CustomAgentRDAccountFG.SELECT_INDEX_ARRAY[{i * 10 + j}]')
                        checkbox.click()
                    except NoSuchElementException:
                        pass
                
                button_id = 'Button26553257' if i == no_pages - 1 else 'Action.AgentRDActSummaryAllListing.GOTO_NEXT__'
                self.driver.find_element(By.ID, button_id).click()
            
            logging.info("Selected all accounts and saved the LOT successfully!")
            

            # Change number of installments for accounts with more than one installment
            for i in range(n):
                if int(no_installments[i]) > 1:
                    element_found_on_current_page = False

                    for j in range(no_pages):
                        page_limit = elem_last_page if j == no_pages - 1 else 10

                        for k in range(page_limit):
                            index = (j * 10) + k
                            element_id = f'HREF_CustomAgentRDAccountFG.ACCOUNT_NUMBER_ARRAY[{index}]'

                            if self.driver.find_element(By.ID, element_id).text == str(acc_nos[i]):
                                radio_button = self.driver.find_element(By.XPATH,f'//input[@id="CustomAgentRDAccountFG.SELECTED_INDEX"][@value="{index}"]')
                                radio_button.click()

                                no_installments_box = self.driver.find_element(By.ID, 'CustomAgentRDAccountFG.RD_INSTALLMENT_NO')
                                no_installments_box.clear()
                                no_installments_box.send_keys(no_installments[i])

                                save_installments_button = self.driver.find_element(By.ID, 'Button11874602')
                                save_installments_button.click()

                                element_found_on_current_page = True
                                break

                        if not element_found_on_current_page:
                            next_page_button = self.driver.find_element(By.ID, 'Action.SelectedAgentRDActSummaryListing.GOTO_NEXT__')
                            next_page_button.click()

                        if element_found_on_current_page:
                            break

            logging.info("Changed number of installments successfully!")

            # Pay all saved installments
            pay_saved_installments_button = self.driver.find_element(By.ID, 'PAY_ALL_SAVED_INSTALLMENTS')
            pay_saved_installments_button.click()
            logging.info("Paid all saved installments successfully!")

            # Extract reference number from alert
            ref_no_alert = self.driver.find_element(By.XPATH, '//div[@class="greenbg"][@role="alert"]')
            reference_no = ref_no_alert.text.split()[7].split('.')[0]
            logging.info("Extracted reference number successfully!")

            # Return reference number
            return reference_no

        except NoSuchElementException as e:
            logging.error(f"Element not found: {e}")
            raise LotTaskError("Error while saving the LOT")
        except Exception as e:
            logging.error(f"Some error during LOT task: {e}")
            raise LotTaskError("Error while saving the LOT")


    # Takes referance number download path and downloads the Report xls file
    def perform_download_report_task(self, ref_no, download_path):
        """
        Performs the task to download the Report xls file using the reference number on the DOP agent portal.

        Args:
        ref_no (str): Reference number used to search and download the report.
        download_path (str): Path where the downloaded report should be saved.

        Returns:
        str: Full path of the downloaded report file.

        Raises:
        DownloadTaskError: If there is an error during any step of the download task.
        """
        try:
            # Navigate to Accounts page
            accounts_button = self.driver.find_element(By.ID, 'Accounts')
            accounts_button.click()
            logging.info("Navigated to Accounts page successfully!")

            # Navigate to Reports Section
            reports_button = self.driver.find_element(By.ID, 'Reports')
            reports_button.click()
            logging.info("Navigated to Reports Section successfully!")

            # Insert Reference Number
            ref_no_input = self.driver.find_element(By.ID, 'CustomAgentRDAccountFG.EBANKING_REF_NUMBER')
            ref_no_input.send_keys(ref_no)
            logging.info("Inserted Reference Number successfully!")

            # Select 'Success' from dropdown
            select_success = self.driver.find_element(By.ID, "CustomAgentRDAccountFG.INSTALLMENT_STATUS")
            select = Select(select_success)
            select.select_by_value('SUC')

            # Clear and set From Date to the first day of the current month
            from_date_field = self.driver.find_element(By.ID, "CustomAgentRDAccountFG.REPORT_DATE_FROM")
            from_date_field.clear()
            current_date = datetime.now()
            first_date_of_month = current_date.replace(day=1)
            formatted_date = first_date_of_month.strftime("%d-%b-%Y")
            from_date_field.send_keys(formatted_date)
            logging.info("Selected From Date successfully!")

            # Click Search button
            search_button = self.driver.find_element(By.ID, 'SearchBtn')
            search_button.click()
            logging.info("Clicked Search button successfully!")

            # Select output format (4 refers to the xls file)
            select_outformat = self.driver.find_element(By.ID, "CustomAgentRDAccountFG.OUTFORMAT")
            select = Select(select_outformat)
            select.select_by_value('4')

            # Click Download button
            download_button = self.driver.find_element(By.ID, 'GENERATE_REPORT')
            download_button.click()

            # Wait for block overlay to disappear
            wait = WebDriverWait(self.driver, 10)
            wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "blockUI.blockOverlay")))
            logging.info("Download completed successfully!")

            # Get default download directory
            default_download_dir = self.temp_download_dir

            # Get downloaded file name
            downloaded_file_name = os.listdir(default_download_dir)[0]

            # Get downloaded file path
            downloaded_file_path = os.path.join(default_download_dir, downloaded_file_name)

            # Format new file name
            new_file_name = f"RDReport-{current_date.strftime('%d-%m-%Y-%H-%M-%S')}-{ref_no}.xls"

            # Get new file path
            new_file_path = os.path.join(download_path, new_file_name)

            # Check if new file already exists
            if os.path.exists(new_file_path):
                # Delete old file
                os.remove(new_file_path)
                logging.info(f"Deleted old file at {new_file_path}")

            # Move and rename downloaded file to new directory
            shutil.move(downloaded_file_path, new_file_path)
            logging.info(f"Moved and renamed file to {new_file_path}")

            return new_file_path

        except NoSuchElementException as e:
            logging.error(f"Element not found: {e}")
            raise DownloadTaskError("Error while downloading report")
        except Exception as e:
            logging.error(f"Error during downloading: {e}")
            raise DownloadTaskError("Error while downloading report")


    # Takes two lists account numbers and aslaas numbers and update those
    def perform_update_aslaas_task(self, acc_nos, aslaas_nos):
        """
        Performs the task to update ASLAAS numbers for given account numbers on the DOP agent portal.

        Args:
        acc_nos (list): List of account numbers for which ASLAAS numbers need to be updated.
        aslaas_nos (list): List of ASLAAS numbers corresponding to each account number.

        Raises:
        UpdateAslaasError: If there is an error during any step of the update ASLAAS task.
        """
        try:
            # Navigate to Accounts page
            accounts_button = self.driver.find_element(By.ID, 'Accounts')
            accounts_button.click()
            logging.info("Navigated to Accounts page successfully!")

            # Navigate to Update ASLAAS Number Section
            update_aslaas_button = self.driver.find_element(By.ID, 'Update ASLAAS Number')
            update_aslaas_button.click()
            logging.info("Navigated to Update ASLAAS page successfully!")

            # Loop through account numbers and update ASLAAS number for each account
            for i in range(len(acc_nos)):
                # Enter account number
                acc_no_input = self.driver.find_element(By.ID, 'CustomAgentAslaasNoFG.RD_ACC_NO')
                acc_no_input.send_keys(acc_nos[i])

                # Enter ASLAAS number
                aslaas_no_input = self.driver.find_element(By.ID, 'CustomAgentAslaasNoFG.ASLAAS_NO')
                aslaas_no_input.send_keys(aslaas_nos[i])

                # Click Continue button
                continue_button = self.driver.find_element(By.ID, 'LOAD_CONFIRM_PAGE')
                continue_button.click()

                # Click Save button
                save_button = self.driver.find_element(By.ID, 'ADD_FIELD_SUBMIT')
                save_button.click()

                logging.info(f"Updated ASLAAS number for account number: {acc_nos[i]}")

        except NoSuchElementException as e:
            logging.error(f"Element not found: {e}")
            raise UpdateAslaasError("Error while updating ASLAAS number")
        except Exception as e:
            logging.error(f"Error during update ASLAAS task: {e}")
            raise UpdateAslaasError("Error while updating ASLAAS number")


    # Fetches Available Accounts from portal and saves to a csv file in RDRecord folder
    def download_accounts_list_task(self):
        """
        Fetches available accounts from the DOP agent portal and saves them to a CSV file in the 'RDRecord' folder.

        Raises:
        DownloadTaskError: If there is an error during any step of the download accounts list task.
        """
        try:
            # Navigate to Accounts page
            accounts_button = self.driver.find_element(By.ID, 'Accounts')
            accounts_button.click()
            logging.info("Navigated to Accounts page successfully!")

            # Navigate to Agent Enquire & Update Screen
            agent_enquire_button = self.driver.find_element(By.ID, 'Agent Enquire & Update Screen')
            agent_enquire_button.click()
            logging.info("Navigated to Agent Enquire & Update Screen page successfully!")

            # Click Print Preview button and switch to new tab
            while True:
                try:
                    print_preview_button = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.ID, 'HREF_printPreview')))
                    self.driver.execute_script("arguments[0].scrollIntoView();", print_preview_button)
                    print_preview_button.click()
                    break
                except:
                    logging.info("Error while clicking print preview button, trying again!")
                    time.sleep(1)
                    
            logging.info("Print Preview Button Clicked successfully!")

            while True:
                try:
                    new_tab_handle = self.driver.window_handles[-1]  # Assuming the new tab is the last in the list
                    self.driver.switch_to.window(new_tab_handle)
                    break
                except:
                    time.sleep(1)
            
            logging.info("Switched to Print Preview tab successfully!")

            # Define columns for DataFrame
            columns = ["ac_no", "acc_holder_name", "denomination", "no_of_installments", "next_rd_due_date"]

            # Create an empty DataFrame with defined columns and data types
            data_types = {
                'ac_no': str,
                'no_of_installments': int
            }
            df = pd.DataFrame(columns=columns).astype(data_types)
            acc_count = 0

            logging.info("Started Extracting Accounts data!")
            time.sleep(5)
            
            # Extract data from the print preview tab
            while acc_count >= 0:
                try:
                    acc_no_element = self.driver.find_element(By.ID, f'HREF_CustomAgentRDAccountFG.ACCOUNT_NUMBER_ALL_ARRAY[{acc_count}]')
                    acc_holder_element = self.driver.find_element(By.ID, f'HREF_CustomAgentRDAccountFG.ACCOUNT_NAME_ALL_ARRAY[{acc_count}]')
                    acc_denomination_element = self.driver.find_element(By.ID, f'HREF_CustomAgentRDAccountFG.DEPOSIT_AMOUNT_ALL_ARRAY[{acc_count}]')
                    no_install_element = self.driver.find_element(By.ID, f'HREF_CustomAgentRDAccountFG.MONTH_PAID_UPTO_ALL_ARRAY[{acc_count}]')
                    next_date_element = self.driver.find_element(By.ID, f'HREF_CustomAgentRDAccountFG.NEXT_RD_INSTALLMENT_DATE_ALL_ARRAY[{acc_count}]')

                    ac_no = str(acc_no_element.text)
                    acc_holder_name = acc_holder_element.text
                    denomination = acc_denomination_element.text.split('.')[0].replace(",", "")
                    no_installments = no_install_element.text
                    next_date = next_date_element.text

                    temp_df = pd.DataFrame([[ac_no, acc_holder_name, denomination, no_installments, next_date]], columns=columns)

                    df = pd.concat([df, temp_df])
                    acc_count += 1
                except Exception as e:
                    logging.info(f"Parsed all accounts successfully! Total {acc_count} accounts!")
                    acc_count = -1
            
            logging.info("Completed Extracting Accounts data!")

            # Close the current tab
            self.driver.close()
            logging.info("Closed accounts print preview tab!")

            # Switch back to the previous tab
            self.driver.switch_to.window(self.driver.window_handles[-1])
            logging.info("Switched back to the previous tab!")

            # Format Next RD due date to datetime
            df['next_rd_due_date'] = pd.to_datetime(df['next_rd_due_date'], format='%d-%b-%Y', errors='coerce')

            # Determine active status based on Next RD due date
            df.loc[df['next_rd_due_date'].isnull(), 'is_active'] = 0
            df.loc[df['next_rd_due_date'].notnull(), 'is_active'] = 1
            df['is_active'] = df['is_active'].astype(int)

            # Calculate Account Opening Date
            df['next_rd_due_date'] = pd.to_datetime(df['next_rd_due_date'], errors='coerce')
            df['acc_opening_date'] = df.apply(lambda row: row['next_rd_due_date'] - pd.DateOffset(months=int(row['no_of_installments'])), axis=1)

            # Drop Next RD due date column
            df = df.drop('next_rd_due_date', axis=1)

            logging.info("DataFrame Edited Successfully!")

            # Create 'RDRecord' folder if it doesn't exist
            folder_name = "RDRecord"
            if not os.path.exists(folder_name):
                os.makedirs(folder_name)

            # Generate file name with current timestamp
            date_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_name = f"{folder_name}/RDRecord_{date_stamp}.csv"

            # Save DataFrame to CSV file
            df.to_csv(file_name, index=False)
            logging.info("Download Data Successfully Completed!")

        except NoSuchElementException as e:
            logging.error(f"Element not found: {e}")
            raise DownloadTaskError(f"Error during downloading: Element not found - {e}")
        except Exception as e:
            logging.error(f"Error during downloading: {e}")
            raise DownloadTaskError(f"Error during downloading: {e}")


    # Downloads the xlsx file of aslaas details
    def download_aslaas_csv(self):
        """
        Downloads the Excel file of ASLAAS details from the DOP agent portal, processes it,
        and saves the relevant data to a CSV file.

        Raises:
        DownloadTaskError: If there is an error during any step of downloading ASLAAS details.
        """
        try:
            # Navigate to Accounts page
            accounts_button = self.driver.find_element(By.ID, 'Accounts')
            accounts_button.click()
            logging.info("Navigated to Accounts page successfully!")

            # Navigate to ASLAAS Number Report
            aslaas_report_button = self.driver.find_element(By.ID, 'ASLAAS Number Report')
            aslaas_report_button.click()
            logging.info("Navigated to ASLAAS Number Report page successfully!")

            # Click Search button and wait for overlay to disappear
            search_button = self.driver.find_element(By.ID, "SEARCH_ASLAAS_NUMBER")
            search_button.click()
            wait = WebDriverWait(self.driver, 20)
            wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "blockUI.blockOverlay")))
            logging.info("Search Button Clicked successfully!")

            # Select output format - 4 refers to the xls file
            select_outformat = self.driver.find_element(By.ID, "CustomAgentAslaasNoFG.OUTFORMAT")
            select = Select(select_outformat)
            select.select_by_value('4')
            logging.info("Excel Format Selection successful!")

            # Click Download Button and wait for download to complete
            download_button = self.driver.find_element(By.ID, "GENERATE_REPORT")
            download_button.click()
            wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "blockUI.blockOverlay")))
            logging.info("Download Button Clicked successfully!")

            # Find the latest downloaded Excel file
            old_path = self.find_latest_xls(self.temp_download_dir)

            # Read Excel data into DataFrame
            df = pd.read_excel(old_path)
            logging.info("Data Loaded successfully!")

            # Extract relevant columns from DataFrame
            df = df.iloc[7:, [2, 8]]  # Assuming the columns containing account and ASLAAS numbers
            df = df.rename(columns={"Unnamed: 2": "ac_no", "Unnamed: 8": "aslaas_no"})  # Rename columns
            df['ac_no'] = df['ac_no'].astype(str)  # Convert account number to string
            df['aslaas_no'] = df['aslaas_no'].astype(str)  # Convert ASLAAS number to string

            # Define new path for saving CSV file
            new_path = os.path.join(self.temp_download_dir, "aslaas_report.csv")

            # Save DataFrame to CSV file
            df.to_csv(new_path, index=False)
            logging.info("CSV File Saved successfully!")

            # Remove the old Excel file
            os.remove(old_path)

        except Exception as e:
            logging.error(f"Error during Downloading: {e}")
            raise DownloadTaskError("Error While Downloading ASLAAS Details")


    # Finds the latest csv file in a directory 
    def find_latest_xls(self, directory):
        """
        Finds the latest modified Excel file (.xls) in the specified directory.

        Args:
        directory (str): The directory path where Excel files are located.

        Returns:
        str: The file path of the latest modified Excel file (.xls).

        Raises:
        ValueError: If no Excel files (.xls) are found in the directory.
        """
        # Get list of all Excel files (.xls) in the directory
        files = glob.glob(os.path.join(directory, "*.xls"))

        # Check if any Excel files (.xls) were found
        if not files:
            raise ValueError(f"No Excel files (.xls) found in directory: {directory}")

        # Return the file with the latest modification time
        return max(files, key=os.path.getmtime)

