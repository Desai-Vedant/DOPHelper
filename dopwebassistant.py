import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoSuchElementException,WebDriverException
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
from PIL import Image,ImageFilter
import io
from datetime import datetime
import os
import glob
import shutil
import json
import requests


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
            'ocr_apikey':"APIKEY"
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
        
    # Takes temperory download path and setup the driver and returns driver variable
    def setup_driver(self,temp_download_dir):
        options = Options()
        options.set_preference("browser.download.dir", temp_download_dir)
        options.set_preference("browser.download.folderList", 2)
        options.set_preference("browser.download.manager.showWhenStarting", False)
        options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/octet-stream,application/vnd.ms-excel")

        driver = webdriver.Firefox(options=options)
        return driver

    
    # Takes Folder name and creates a folder of that name if it does not exist and returns full path
    def create_folder_if_not_exists(self, folder_name):
        full_path = os.path.abspath(folder_name)
        if os.path.exists(full_path):
            shutil.rmtree(full_path)  # Delete the folder and all its contents
        os.makedirs(full_path)  # Create the folder
        return full_path

    # Creates a driver and opens browser and goes to dop portal
    def open_browser_portal(self):
        try:
            self.driver = self.setup_driver(self.temp_download_dir)
            self.driver.implicitly_wait(20)
            self.driver.maximize_window()
            self.driver.get("https://dopagent.indiapost.gov.in")
            logging.info("Opened the Browser and navigated to website Suceessfully !")

        except WebDriverException as we:
            logging.error(f"WebDriverException during Opening Browser: {we}")
            self.close_browser()
            raise OpenBrowserError(f"Error While Opening Browser : {we}")
        except Exception as e:
            logging.error(f"Error during Opening Browser: {e}")
            self.close_browser()
            raise OpenBrowserError(f"Error While Opening Browser : {e}")

    # Creates a driver and opens browser in hidden mode and goes to dop portal
    def open_browser_portal_headless(self):
        try:
            # Set up Firefox options
            options = Options()
            # Headless mode (optional, you can remove this line if you want to see the browser)
            options.add_argument("-headless")

            self.driver = webdriver.Firefox(options=options)
            self.driver.implicitly_wait(10)
            self.driver.maximize_window()
            self.driver.get("https://dopagent.indiapost.gov.in")
            logging.info("Opened the Browser and navigated to website Suceessfully !")

        except WebDriverException as we:
            logging.error(f"WebDriverException during Opening Browser: {we}")
            self.close_browser()
            raise OpenBrowserError(f"Error While Opening Browser : {we}")
        except Exception as e:
            logging.error(f"Error during Opening Browser: {e}")
            self.close_browser()
            raise OpenBrowserError(f"Error While Opening Browser : {e}")

    # Log in to the dop agent portal
    def login(self):
        login_suc = False
        try:
            while not login_suc:
                username = self.driver.find_element(By.NAME, "AuthenticationFG.USER_PRINCIPAL")
                username.clear()
                username.send_keys(self.user_id)

                password = self.driver.find_element(By.NAME, "AuthenticationFG.ACCESS_CODE")
                password.clear()
                password.send_keys(self.user_password)
                logging.info("Username and Password inserted Suceessfully !")

                # Find the element
                element = self.driver.find_element(By.ID, "IMAGECAPTCHA")

                # Get the element screenshot
                png = element.screenshot_as_png

                # Create an Image object and crop it to the element's size
                image = Image.open(io.BytesIO(png))

                image = image.filter(ImageFilter.MedianFilter(size=3))
                image = image.convert('L')
                self.create_folder_if_not_exists('temp')
                image.save('./temp/image.png', format='PNG')

                try:
                    text = self.ocr_space_file(filename='./temp/image.png')
                except:
                    text = "  "

                captcha_input = self.driver.find_element(By.ID, 'AuthenticationFG.VERIFICATION_CODE')
                captcha_input.clear()
                captcha_input.send_keys(text)
                logging.info("Solved CAPTCHA Suceessfully !")

                # wait for user to solve CAPTCHA
                time.sleep(10)

                submit_button = self.driver.find_element(By.ID, 'VALIDATE_RM_PLUS_CREDENTIALS_CATCHA_DISABLED')
                submit_button.click()
                logging.info("Details Submission Suceessful !")

                try:
                    # Find the Accounts button
                    self.driver.find_element(By.ID, 'Accounts')
                    login_suc = True
                    logging.info("Login Suceessful !")
                except:
                    logging.info(f"Login Unsucessful Trying to login again !")
                    continue

        except Exception as login_error:
            logging.error(f"Error during login: {login_error}")
            self.close_browser()
            raise LoginError("Error While Logging in to Account.")
        
    def ocr_space_file(self, filename, overlay=False , language='eng'):
        payload = {'isOverlayRequired': overlay,
                'apikey': self.ocr_apikey,
                'language': language,
                }
        with open(filename, 'rb') as f:
            r = requests.post('https://api.ocr.space/parse/image',
                            files={filename: f},
                            data=payload,
                            )
        response =  r.content.decode()
        data = json.loads(response)
        return data["ParsedResults"][0]["ParsedText"].strip()

    # Close the browser if its open
    def close_browser(self):
        if self.driver:
            self.driver.quit()

    # Takes two arrays account numbers and no of installments and performs lot
    def perform_lot_task(self, acc_nos, no_installments):
        try:
            # Find the Accounts button and click it
            accounts_button = self.driver.find_element(By.ID, 'Accounts')
            accounts_button.click()
            logging.info("Navigated to Accounts page Suceessful !")

            # Derived Variables
            n = len(acc_nos)
            elem_last_page = n % 10
            no_full_pages = n // 10
            no_pages = no_full_pages + (1 if elem_last_page > 0 else 0)
            if elem_last_page==0:
                elem_last_page=10

            # Find the Agent Enquire button and click it
            agent_enquire_button = self.driver.find_element(By.ID, 'Agent Enquire & Update Screen')
            agent_enquire_button.click()
            logging.info("Navigated to Agent Enquire & Update Screen page Suceessful !")
            

            # Find the Cash button and click it
            cash_button = self.driver.find_element(By.XPATH, '//input[@id="CustomAgentRDAccountFG.PAY_MODE_SELECTED_FOR_TRN"][@value="C"]')
            cash_button.click()
            logging.info("Cash mode of Payment Selected Suceessful !")

            # Find Text Box and insert numbers into it and fetch
            text_box = self.driver.find_element(By.ID, 'CustomAgentRDAccountFG.ACCOUNT_NUMBER_FOR_SEARCH')
            text_box.send_keys(','.join(map(str, acc_nos)))
            fetch_button = self.driver.find_element(By.ID, 'Button3087042')
            fetch_button.click()
            logging.info("Accounts Fetched Suceessful !")

            # Select all Numbers and Save the Lot
            for i in range(no_pages):
                page_limit = elem_last_page if i == no_pages - 1 else 10
                
                for j in range(page_limit):
                    try:
                        # Find Checkbox and select it
                        checkbox = self.driver.find_element(By.ID, f'CustomAgentRDAccountFG.SELECT_INDEX_ARRAY[{i * 10 + j}]')
                        checkbox.click()
                    except NoSuchElementException:
                        pass

                # Find Save or Next Button and Click it
                button_id = 'Button26553257' if i == no_pages - 1 else 'Action.AgentRDActSummaryAllListing.GOTO_NEXT__'
                self.driver.find_element(By.ID, button_id).click()
            logging.info("Selected All Accounts and Saved the Lot Suceessful !")

            # Search for accounts with more than one installments and search them on different pages and then change their no of installments
            for i in range(n):
                
                if int(no_installments[i]) > 1:
                    element_found_on_current_page = False
                    
                    for j in range(no_pages):
                        page_limit = elem_last_page if j == no_pages - 1 else 10
                        
                        for k in range(page_limit):
                            index = (j * 10) + k
                            element_id = f'HREF_CustomAgentRDAccountFG.ACCOUNT_NUMBER_ARRAY[{index}]'

                            if self.driver.find_element(By.ID, element_id).text == str(acc_nos[i]):
                                # Find the radio button by its id and value
                                radio_button = self.driver.find_element(By.XPATH,f'//input[@id="CustomAgentRDAccountFG.SELECTED_INDEX"][@value="{index}"]')
                                radio_button.click()

                                # Find no of installments Box and insert the number of installment into it and save
                                no_installments_box = self.driver.find_element(By.ID, 'CustomAgentRDAccountFG.RD_INSTALLMENT_NO')
                                no_installments_box.clear()
                                no_installments_box.send_keys(no_installments[i])

                                save_installments_button = self.driver.find_element(By.ID, 'Button11874602')
                                save_installments_button.click()

                                # Set the flag to True and exit the inner loop
                                element_found_on_current_page = True
                                break

                        # Check if we found the element on the current page
                        if not element_found_on_current_page:
                            # Go to Next Page
                            next_page_button = self.driver.find_element(By.ID, 'Action.SelectedAgentRDActSummaryListing.GOTO_NEXT__')
                            next_page_button.click()
                        
                        if element_found_on_current_page:
                            break

            logging.info("Changing no of Installments Suceessful !")

            # After setting all the no of installments click on pay all saved installments
            pay_saved_installments_button = self.driver.find_element(By.ID, 'PAY_ALL_SAVED_INSTALLMENTS')
            pay_saved_installments_button.click()
            logging.info("Pay all Saved Installments Suceessful !")

            # Extract the referance number of the lot from alert on page
            ref_no_alert = self.driver.find_element(By.XPATH,f'//div[@class="greenbg"][@role="alert"]')
            referance_no = ref_no_alert.text.split()[7].split('.')[0]
            logging.info("Referance Number Extracted Suceessful !")

            # Return Referance Number
            return referance_no

        except NoSuchElementException as e:
            logging.error(f"Element not found: {e}")
            raise LotTaskError("Error while Saving the Lot")
        except Exception as e:
            logging.error(f"Some Error during lot task: {e}")
            raise LotTaskError("Error while Saving the Lot")

    # Takes referance number download path and downloads the Report xls file
    def perform_download_report_task(self, ref_no, download_path):
        try:
            # Find the Accounts button and click it
            accounts_button = self.driver.find_element(By.ID, 'Accounts')
            accounts_button.click()
            logging.info("Navigated to Accounts page Successful!")

            # Go to reports section
            reports_button = self.driver.find_element(By.ID, 'Reports')
            reports_button.click()
            logging.info("Navigated to Reports Section Successful!")

            # Enter the reference number
            ref_no_input = self.driver.find_element(By.ID, 'CustomAgentRDAccountFG.EBANKING_REF_NUMBER')
            ref_no_input.send_keys(ref_no)
            logging.info("Inserting Reference Number Successful!")

            # Select success from dropdown
            select_success = self.driver.find_element(By.ID, "CustomAgentRDAccountFG.INSTALLMENT_STATUS")
            select = Select(select_success)
            select.select_by_value('SUC')

            # Select the from date field
            from_date_field = self.driver.find_element(By.ID, "CustomAgentRDAccountFG.REPORT_DATE_FROM")
            from_date_field.clear()

            # Get the current date
            current_date = datetime.now()

            # Get the first day of the current month
            first_date_of_month = current_date.replace(day=1)

            # Format the date in "dd-MMM-yyyy" format
            formatted_date = first_date_of_month.strftime("%d-%b-%Y")

            # Insert First date of current month
            from_date_field.send_keys(formatted_date)
            logging.info("Selected Form Date Successful!")

            # click on Search button
            search_button = self.driver.find_element(By.ID, 'SearchBtn')
            search_button.click()
            logging.info("Searched For Lot Successful!")

            # Select output format - 4 refers to the xls file
            select_outformat = self.driver.find_element(By.ID, "CustomAgentRDAccountFG.OUTFORMAT")
            select = Select(select_outformat)
            select.select_by_value('4')

            # click on Download button
            download_button = self.driver.find_element(By.ID, 'GENERATE_REPORT')
            download_button.click()
            
            # Wait for block overlay to disappear
            wait = WebDriverWait(self.driver, 10)
            wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "blockUI.blockOverlay")))
            logging.info("Download Successful!")

            # Get the default download directory
            default_download_dir = self.temp_download_dir

            # Get the downloaded file name
            downloaded_file_name = os.listdir(default_download_dir)[0]

            # Get the downloaded file path
            downloaded_file_path = os.path.join(default_download_dir, downloaded_file_name)

            # Format the new file name
            new_file_name = f"RDReport-{current_date.strftime('%d-%m-%Y-%H-%M-%S')}-{ref_no}.xls"

            # Get the new file path
            new_file_path = os.path.join(download_path, new_file_name)
            
            # Check if the new file already exists
            if os.path.exists(new_file_path):
                # Delete the old file
                os.remove(new_file_path)
                logging.info(f"Deleted the old file at {new_file_path}")

            # Move and rename the downloaded file to the new directory
            shutil.move(downloaded_file_path, new_file_path)
            logging.info(f"Moved and renamed the file to {new_file_path}")
            
            return new_file_path
            
        except NoSuchElementException as e:
            logging.error(f"Element not found: {e}")
            raise DownloadTaskError("Error While Downloading Report")
        except Exception as e:
            logging.error(f"Error during Downloading: {e}")
            raise DownloadTaskError("Error While Downloading Report")
        
    # Takes two lists account numbers and aslaas numbers and update those
    def perform_update_aslaas_task(self, acc_nos, aslaas_nos):
        try:
            # Find the Accounts button and click it
            accounts_button = self.driver.find_element(By.ID, 'Accounts')
            accounts_button.click()
            logging.info("Navigated to Accounts page Suceessful !")

            # Go to update aslaas number Section
            update_aslaas_button = self.driver.find_element(By.ID, 'Update ASLAAS Number')
            update_aslaas_button.click()
            logging.info("Navigated to Update Aslaas page Suceessful !")

            # loop through account numbers and change aslaas number for each account
            for i in range(0,len(acc_nos)):

                # Enter the account number
                acc_no_input = self.driver.find_element(By.ID, 'CustomAgentAslaasNoFG.RD_ACC_NO')
                acc_no_input.send_keys(acc_nos[i])

                # Enter the aslaas number
                aslaas_no_input = self.driver.find_element(By.ID, 'CustomAgentAslaasNoFG.ASLAAS_NO')
                aslaas_no_input.send_keys(aslaas_nos[i])

                # Click on continue button
                continue_button = self.driver.find_element(By.ID, 'LOAD_CONFIRM_PAGE')
                continue_button.click()

                # Click on save button
                save_button = self.driver.find_element(By.ID, 'ADD_FIELD_SUBMIT')
                save_button.click()

                logging.info(f"Updated Aslaas number for account number : {acc_nos[i]}")

        except NoSuchElementException as e:
            logging.error(f"Element not found: {e}")
            raise UpdateAslaasError("Error While Updating Aslaas Number")
        except Exception as e:
            logging.error(f"Error during update aslaas task: {e}")
            raise UpdateAslaasError("Error While Updating Aslaas Number")

    
    # Fetches Available Accounts from portal and saves to a csv file in RDRecord folder
    def download_accounts_list_task(self):
        try:
            # Find the Accounts button and click it
            accounts_button = self.driver.find_element(By.ID, 'Accounts')
            accounts_button.click()
            logging.info("Navigated to Accounts page Suceessful !")

            # Find the Agent Enquire button and click it
            agent_enquire_button = self.driver.find_element(By.ID, 'Agent Enquire & Update Screen')
            agent_enquire_button.click()
            logging.info("Navigated to Agent Enquire & Update Screen page Suceessful !")

            while True:
                try:
                    # Wait for the Print Preview button to be clickable and click it
                    print_privew_button = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.ID, 'HREF_printPreview')))
                    self.driver.execute_script("arguments[0].scrollIntoView();", print_privew_button)
                    print_privew_button.click()
                    break
                except:
                    logging.info("Error while clicking print privew button, trying again !")
                    time.sleep(1)
                    
            logging.info("Print Privew Button Clicked Suceessful !")

            while True:
                try:
                    new_tab_handle = self.driver.window_handles[-1]  # Assuming the new tab is the last in the list
                    self.driver.switch_to.window(new_tab_handle)
                    break
                except:
                    time.sleep(1)
            
            logging.info("Switch to Print Privew tab Suceessful !")

            columns = ["ac_no", "acc_holder_name", "denomination", "no_of_installments", "next_rd_due_date"]

            # Define the data types
            data_types = {
                'ac_no': str,
                'no_of_installments': int
            }
            df = pd.DataFrame(columns = columns).astype(data_types)
            acc_count = 0

            logging.info("Started Extracting Books data !")
            time.sleep(5)
            while acc_count>=0:
                try:
                    acc_no_element = self.driver.find_element(By.ID, f'HREF_CustomAgentRDAccountFG.ACCOUNT_NUMBER_ALL_ARRAY[{acc_count}]')
                    acc_holder_element = self.driver.find_element(By.ID, f'HREF_CustomAgentRDAccountFG.ACCOUNT_NAME_ALL_ARRAY[{acc_count}]')
                    acc_denomination_element = self.driver.find_element(By.ID, f'HREF_CustomAgentRDAccountFG.DEPOSIT_AMOUNT_ALL_ARRAY[{acc_count}]')
                    no_install_element = self.driver.find_element(By.ID, f'HREF_CustomAgentRDAccountFG.MONTH_PAID_UPTO_ALL_ARRAY[{acc_count}]')
                    next_date_element = self.driver.find_element(By.ID, f'HREF_CustomAgentRDAccountFG.NEXT_RD_INSTALLMENT_DATE_ALL_ARRAY[{acc_count}]')

                    ac_no = str(acc_no_element.text)
                    acc_holder_name = acc_holder_element.text
                    denomination = acc_denomination_element.text.split('.')[0].replace(",","")
                    no_installments = no_install_element.text
                    next_date = next_date_element.text

                    temp_df = pd.DataFrame([[ac_no,acc_holder_name,denomination,no_installments,next_date]], columns=columns)

                    df = pd.concat([df,temp_df])
                    acc_count +=1
                except Exception as e:
                    logging.info(f"Parsed all books sucessfuly ! total {acc_count} books !")
                    acc_count = -1
            
            logging.info("Completed Extracting Books data !")
            
            # Close the current tab
            self.driver.close()
            logging.info("Closed accounts print privew tab !")

            # Switch to the previous tab
            self.driver.switch_to.window(self.driver.window_handles[-1])
            logging.info("Switched to the previous tab !")

            # Changing Date Format of Next RD due date
            df['next_rd_due_date'] = pd.to_datetime(df['next_rd_due_date'], format='%d-%b-%Y', errors='coerce')

            # Choose Active Status
            df.loc[df['next_rd_due_date'].isnull(), 'is_active'] = 0
            df.loc[df['next_rd_due_date'].notnull(), 'is_active'] = 1
            df['is_active'] = df['is_active'].astype(int)

            # Calculate Account Opening Date
            df['next_rd_due_date'] = pd.to_datetime(df['next_rd_due_date'], errors='coerce')
            df['acc_opening_date'] = df.apply(lambda row: row['next_rd_due_date'] - pd.DateOffset(months=int(row['no_of_installments'])), axis=1)

            # Drop Next RD due date column
            df = df.drop('next_rd_due_date', axis=1)

            logging.info("Dataframe Edited Sucessfuly !")
            
            # Create a folder named "RDRecord" if it doesn't exist
            folder_name = "RDRecord"
            if not os.path.exists(folder_name):
                os.makedirs(folder_name)

            # Get the current date and time for the file name
            date_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            # Create the file name with the date stamp
            file_name = f"{folder_name}/RDRecord_{date_stamp}.csv"

            # Save the DataFrame to the CSV file
            df.to_csv(file_name, index=False)
            logging.info("Download Data Sucessfuly Completed!")

        except NoSuchElementException as e:
            logging.error(f"Element not found: {e}")
            raise DownloadTaskError(f"Error during Downloading: Element not found - {e}")
        except Exception as e:
            logging.error(f"Error during Downloading: {e}")
            raise DownloadTaskError(f"Error during Downloading: {e}")
    
    
    # Downloads the xlsx file of aslaas details
    def download_aslaas_csv(self):
        try:
            # Find the Accounts button and click it
            accounts_button = self.driver.find_element(By.ID, 'Accounts')
            accounts_button.click()
            logging.info("Navigated to Accounts page Suceessful !")

            # Find the ASLAAS Number Report and click it
            aslaas_report_button = self.driver.find_element(By.ID, 'ASLAAS Number Report')
            aslaas_report_button.click()
            logging.info("Navigated to ASLAAS Number Report page Suceessful !")
            
            search_button = self.driver.find_element(By.ID, "SEARCH_ASLAAS_NUMBER")
            search_button.click()
            wait = WebDriverWait(self.driver, 20)
            wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "blockUI.blockOverlay")))
            logging.info("Search Button Clicked Suceessful !")
            
            # Select output format - 4 refers to the xls file
            select_outformat = self.driver.find_element(By.ID, "CustomAgentAslaasNoFG.OUTFORMAT")
            select = Select(select_outformat)
            select.select_by_value('4')
            logging.info("Excel Format Selection Suceessful !")
            
            # Find and click Download Button
            download_button = self.driver.find_element(By.ID, "GENERATE_REPORT")
            download_button.click()
            # Wait for block overlay to disappear
            wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "blockUI.blockOverlay")))
            logging.info("Download Button Clicked Sucessful !")
            
            old_path = self.find_latest_xls(self.temp_download_dir)
            
            df = pd.read_excel(old_path)
            logging.info("Load Data Sucessful !")
            
            df = df.iloc[7:,[2,8]]
            df = df.rename(columns={"Unnamed: 2": "ac_no", "Unnamed: 8": "aslaas_no"})
            df['ac_no'] = df['ac_no'].astype(str)
            df['aslaas_no'] = df['aslaas_no'].astype(str)
            
            new_path = os.path.join(self.temp_download_dir,"aslaas_report.csv")
        
            df.to_csv(new_path, index=False)
            logging.info("CSV File Saved Sucessful !")
            os.remove(old_path)
            
        except Exception as e :
            logging.error(f"Error during Downloading: {e}")
            raise DownloadTaskError("Error While Downloading Aslaas Details")
            
            
            
    # Finds the latest csv file in a directory 
    def find_latest_xls(self,directory):
        # Get list of all CSV files in directory
        files = glob.glob(os.path.join(directory, "*.xls"))
        # Return the file with the latest modification time
        return max(files, key=os.path.getmtime)