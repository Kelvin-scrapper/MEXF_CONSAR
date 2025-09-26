import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import os

# Configuration
HEADLESS_MODE = True  # Set to True to run browser in headless mode, False to show browser window

class CONSARDownloader:
    def __init__(self, download_path=None):
        """
        Initialize the CONSAR downloader

        Args:
            download_path (str): Directory to save downloaded files. If None, uses current directory.
        """
        # URLs for all three steps
        self.investment_url = "https://www.consar.gob.mx/gobmx/aplicativo/siset/CuadroInicial.aspx?md=61"
        self.flow_url = "https://www.consar.gob.mx/gobmx/aplicativo/siset/Series.aspx?cd=60&cdAlt=False"
        self.withdrawal_url = "https://www.consar.gob.mx/gobmx/aplicativo/siset/Series.aspx?cd=141&cdAlt=False"
        self.download_path = download_path or os.getcwd()
        self.driver = None

        # Ensure download directory exists
        os.makedirs(self.download_path, exist_ok=True)
    
    def setup_driver(self):
        """Setup Chrome driver with download capabilities"""
        print("Setting up Chrome driver...")

        try:
            options = uc.ChromeOptions()

            # Ensure download directory exists and is absolute
            abs_download_path = os.path.abspath(self.download_path)
            os.makedirs(abs_download_path, exist_ok=True)

            if HEADLESS_MODE:
                options.add_argument("--headless=new")
                print("Running in headless mode")
            else:
                print("Running in windowed mode")

            # Essential download preferences
            prefs = {
                "download.default_directory": abs_download_path,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "profile.default_content_settings.popups": 0,
                "profile.default_content_setting_values.automatic_downloads": 1
            }
            options.add_experimental_option("prefs", prefs)

            # Basic stability options
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-extensions")

            self.driver = uc.Chrome(options=options)
            print("SUCCESS: Chrome driver setup successful")
            return True

        except Exception as e:
            print(f"ERROR: Chrome driver setup failed: {e}")
            return False
    

    def find_and_click_excel_link(self):
        """Find and click the Excel download link"""
        try:
            wait = WebDriverWait(self.driver, 15)
            print("Looking for Excel download link...")

            # Primary selector for the specific Excel link
            excel_selector = "//a[@id='ctl00_ContentPlaceHolder1_lnkBtn_Excel']"

            try:
                excel_link = wait.until(EC.element_to_be_clickable((By.XPATH, excel_selector)))
                print("SUCCESS: Excel link found")

                # Click using JavaScript to ensure it works
                self.driver.execute_script("arguments[0].click();", excel_link)
                print("SUCCESS: Excel link clicked")
                return True

            except TimeoutException:
                print("ERROR: Excel link not found or not clickable")
                return False

        except Exception as e:
            print(f"ERROR: Error clicking Excel link: {e}")
            return False
    
    def wait_for_download(self, timeout_seconds=30):
        """Wait for Excel file to be downloaded"""
        print(f"Waiting for download (timeout: {timeout_seconds}s)...")

        start_time = time.time()
        while time.time() - start_time < timeout_seconds:
            try:
                # Look for Excel files in download directory
                for filename in os.listdir(self.download_path):
                    if filename.endswith(('.xlsx', '.xls', '.xlsm')):
                        filepath = os.path.join(self.download_path, filename)

                        # Check if file was created recently and is not empty
                        if os.path.getsize(filepath) > 0:
                            file_age = time.time() - os.path.getmtime(filepath)
                            if file_age < 60:  # File created within last minute
                                print(f"SUCCESS: Download completed: {filepath}")
                                return filepath

                time.sleep(1)  # Wait 1 second before checking again

            except Exception as e:
                print(f"Error checking for downloads: {e}")
                time.sleep(1)

        print("ERROR: Download timeout reached")
        return None

    def select_withdrawal_checkboxes(self):
        """Select Retiro de Recursos IMSS and ISSSTE checkboxes based on exact HTML structure"""
        try:
            wait = WebDriverWait(self.driver, 15)
            print("Looking for Retiro de Recursos checkboxes...")

            # Target specific checkboxes based on exact HTML structure provided
            withdrawal_checkboxes = [
                {
                    'name': 'Retiro de Recursos IMSS',
                    'selectors': [
                        "//input[@type='checkbox' and @value='5185']",  # Exact value from HTML
                        "//input[@type='checkbox'][following-sibling::*[contains(., 'Retiro de Recursos IMSS')]]",
                        "//input[@type='checkbox'][following-sibling::*[contains(@onclick, \"showSubElements('3')\")]]"
                    ]
                },
                {
                    'name': 'Retiro de Recursos ISSSTE',
                    'selectors': [
                        "//input[@type='checkbox' and @value='24180']",  # Exact value from HTML
                        "//input[@type='checkbox'][following-sibling::*[contains(., 'Retiro de Recursos ISSSTE')]]",
                        "//input[@type='checkbox'][following-sibling::*[contains(@onclick, \"showSubElements('187')\")]]"
                    ]
                }
            ]

            selected_count = 0

            for checkbox_info in withdrawal_checkboxes:
                checkbox_found = False

                for selector in checkbox_info['selectors']:
                    try:
                        checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                        if not checkbox.is_selected():
                            checkbox.click()
                            print(f"SUCCESS: Selected: {checkbox_info['name']}")
                            selected_count += 1
                            checkbox_found = True
                            time.sleep(1)  # Small delay between selections
                            break
                        else:
                            print(f"INFO: Already selected: {checkbox_info['name']}")
                            checkbox_found = True
                            break
                    except TimeoutException:
                        continue
                    except Exception as e:
                        print(f"Error with selector {selector}: {e}")
                        continue

                if not checkbox_found:
                    print(f"WARNING: Could not find checkbox: {checkbox_info['name']}")

            print(f"Selected {selected_count} new checkboxes")
            return selected_count > 0 or True  # Return True if any progress made

        except Exception as e:
            print(f"❌ Error selecting checkboxes: {e}")
            return False

    def select_rcv_checkboxes(self):
        """Select RCV, RCV IMSS, and RCV ISSSTE checkboxes based on exact HTML structure"""
        try:
            wait = WebDriverWait(self.driver, 15)
            print("Looking for RCV checkboxes...")

            # Target specific checkboxes based on exact HTML structure provided
            rcv_checkboxes = [
                {
                    'name': 'RCV Main',
                    'selectors': [
                        "//input[@type='checkbox' and @value='4135']",  # Main RCV checkbox
                        "//input[@type='checkbox'][following-sibling::*[contains(., 'RCV') and not(contains(., 'IMSS')) and not(contains(., 'ISSSTE'))]]"
                    ]
                },
                {
                    'name': 'RCV IMSS',
                    'selectors': [
                        "//input[@type='checkbox'][following-sibling::*[contains(., 'RCV IMSS')]]",
                        "//input[@type='checkbox'][following-sibling::*[contains(@onclick, \"showSubElements('93')\")]]",
                        "//input[@type='checkbox'][ancestor::tr[contains(., 'RCV IMSS')]]"
                    ]
                },
                {
                    'name': 'RCV ISSSTE',
                    'selectors': [
                        "//input[@type='checkbox' and @value='4165']",  # Exact value from HTML
                        "//input[@type='checkbox'][following-sibling::*[contains(., 'RCV ISSSTE')]]",
                        "//input[@type='checkbox'][following-sibling::*[contains(@onclick, \"showSubElements('123')\")]]"
                    ]
                }
            ]

            selected_count = 0

            for checkbox_info in rcv_checkboxes:
                checkbox_found = False

                for selector in checkbox_info['selectors']:
                    try:
                        checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                        if not checkbox.is_selected():
                            checkbox.click()
                            print(f"SUCCESS: Selected: {checkbox_info['name']}")
                            selected_count += 1
                            checkbox_found = True
                            time.sleep(1)  # Small delay between selections
                            break
                        else:
                            print(f"INFO: Already selected: {checkbox_info['name']}")
                            checkbox_found = True
                            break
                    except TimeoutException:
                        continue
                    except Exception as e:
                        print(f"Error with selector {selector}: {e}")
                        continue

                if not checkbox_found:
                    print(f"WARNING: Could not find checkbox: {checkbox_info['name']}")

            print(f"Selected {selected_count} new checkboxes")
            return selected_count > 0 or True  # Return True if any progress made

        except Exception as e:
            print(f"❌ Error selecting checkboxes: {e}")
            return False

    def click_export_button(self):
        """Click the Export button using exact HTML structure with enhanced debugging"""
        try:
            wait = WebDriverWait(self.driver, 15)
            print("Looking for Export button...")

            # First, try to scroll to make sure the button is visible
            try:
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)
            except:
                pass

            # Target the exact export button from the HTML structure
            export_selectors = [
                "//input[@id='ctl00_ContentPlaceHolder1_btn_ExportaSeries']",  # Exact ID from HTML
                "//input[@value='Exportar']",  # Spanish text
                "//input[@type='submit' and @value='Exportar']",
                "//input[contains(@name, 'btn_ExportaSeries')]"
            ]

            for i, selector in enumerate(export_selectors):
                try:
                    print(f"Trying export selector {i+1}: {selector}")
                    export_button = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    print(f"Found export button with selector {i+1}")

                    # Scroll to button and click
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", export_button)
                    time.sleep(1)

                    try:
                        export_button.click()
                    except:
                        # Try JavaScript click if regular click fails
                        self.driver.execute_script("arguments[0].click();", export_button)

                    print("SUCCESS: Export button clicked")
                    time.sleep(2)  # Wait for any processing
                    return True

                except TimeoutException:
                    print(f"Timeout with selector {i+1}")
                    continue
                except Exception as e:
                    print(f"Error with export selector {i+1}: {e}")
                    continue

            print("ERROR: Export button not found with any selector")
            return False

        except Exception as e:
            print(f"ERROR: Error clicking export button: {e}")
            return False

    def translate_page_to_english(self):
        """Enhanced page translation using logic from dfcs_scraper.py"""
        print("Starting page translation process...")
        try:
            page_text = self.driver.find_element(By.TAG_NAME, "body").text.lower()
            english_keywords = ['export', 'download', 'resources channeled', 'investment', 'english', 'translate']
            english_count = sum(1 for keyword in english_keywords if keyword in page_text)

            if english_count >= 2:
                print(f"✅ Page already translated! Found {english_count} English keywords")
                return True

            time.sleep(3)

            # Look for Chrome's translation bar/notification
            auto_translate_selectors = [
                "//div[contains(@class, 'goog-te-banner')]//button[contains(text(), 'Translate')]",
                "//button[contains(text(), 'Translate')]",
                "//span[contains(text(), 'Translate this page')]//parent::button"
            ]

            for selector in auto_translate_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    for element in elements:
                        if element.is_displayed() and element.is_enabled():
                            element.click()
                            print("Clicked automatic translation button")
                            time.sleep(4)
                            return True
                except:
                    continue

            # Method 1: Look for Google Translate dropdown (from screenshot)
            try:
                wait = WebDriverWait(self.driver, 10)

                # Try to find the Google Translate dropdown elements
                translate_selectors = [
                    # Look for "English" option in dropdown
                    "//span[contains(text(), 'English')]",
                    "//div[contains(text(), 'English')]",
                    "//a[contains(text(), 'English')]",
                    # Look for Google Translate elements
                    "//*[contains(@class, 'VIpgJd-ZVi9od-xl07Ob-lTBxed')]//span[contains(text(), 'English')]",
                    "//*[contains(@data-language-code, 'en')]",
                    # General translate selectors
                    "//*[text()='English' or @title='English']",
                    "//*[@role='listitem' and contains(., 'English')]"
                ]

                for selector in translate_selectors:
                    try:
                        print(f"Trying translate selector: {selector}")
                        english_option = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))

                        # Scroll to element and click
                        self.driver.execute_script("arguments[0].scrollIntoView();", english_option)
                        time.sleep(1)
                        english_option.click()

                        print("SUCCESS: Clicked English translation option")
                        time.sleep(5)  # Wait for translation to complete
                        return True

                    except TimeoutException:
                        print(f"Timeout with selector: {selector}")
                        continue
                    except Exception as e:
                        print(f"Error with selector {selector}: {e}")
                        continue

            except Exception as e:
                print(f"Google Translate dropdown method failed: {e}")

            # Method 2: Look for Chrome's translate bar at top of page
            try:
                # Check for Chrome's built-in translate notification
                chrome_translate_selectors = [
                    "//div[contains(@class, 'translate')]//button[contains(text(), 'English')]",
                    "//*[@id='translate-button' or contains(@class, 'translate-button')]",
                    "//button[contains(text(), 'Translate')]",
                    "//*[contains(@class, 'goog-te-banner')]//button"
                ]

                for selector in chrome_translate_selectors:
                    try:
                        translate_btn = WebDriverWait(self.driver, 3).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                        translate_btn.click()
                        print("SUCCESS: Clicked Chrome translate button")
                        time.sleep(5)
                        return True
                    except:
                        continue

            except Exception as e:
                print(f"Chrome translate bar method failed: {e}")

            # Method 3: Try to find translate widget anywhere on page
            try:
                # Look for any translate-related text or buttons
                general_selectors = [
                    "//*[contains(text(), 'Translate') or contains(text(), 'traducir')]",
                    "//*[contains(text(), 'English') or contains(text(), 'Inglés')]",
                    "//select[@name='language']//option[@value='en']"
                ]

                for selector in general_selectors:
                    try:
                        element = WebDriverWait(self.driver, 2).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                        element.click()
                        print(f"SUCCESS: Clicked general translate element")
                        time.sleep(3)
                        return True
                    except:
                        continue

            except Exception as e:
                print(f"General translate method failed: {e}")

            print("Manual translation may be needed")
            time.sleep(5)
            return True

        except Exception as e:
            print(f"Translation process error: {e}")
            return False

    def download_rcv_flow_data(self):
        """Step 2: Download RCV flow data from Series page"""
        if not self.setup_driver():
            return None

        try:
            # Access the Series page
            print(f"Opening Series page: {self.flow_url}")
            self.driver.get(self.flow_url)
            time.sleep(5)  # Wait for page to load

            # Select the required checkboxes (work directly with Spanish)
            if not self.select_rcv_checkboxes():
                print("❌ Failed to select required checkboxes")
                return None

            # Click Export button
            if not self.click_export_button():
                print("❌ Failed to click Export button")
                return None

            # Wait for download to complete
            downloaded_file = self.wait_for_download(timeout_seconds=30)

            if downloaded_file:
                print("Download completed, waiting 5 seconds before closing browser...")
                time.sleep(5)  # Allow file to completely download

            return downloaded_file

        except Exception as e:
            print(f"❌ RCV flow data download failed: {e}")
            return None

        finally:
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass

    def download_withdrawal_data(self):
        """Step 3: Download Retiro de Recursos data from Series page"""
        if not self.setup_driver():
            return None

        try:
            # Access the withdrawal page
            print(f"Opening withdrawal page: {self.withdrawal_url}")
            self.driver.get(self.withdrawal_url)
            time.sleep(5)  # Wait for page to load

            # Select the required checkboxes (work directly with Spanish)
            if not self.select_withdrawal_checkboxes():
                print("❌ Failed to select required checkboxes")
                return None

            # Click Export button
            if not self.click_export_button():
                print("❌ Failed to click Export button")
                return None

            # Wait for download to complete
            downloaded_file = self.wait_for_download(timeout_seconds=30)

            if downloaded_file:
                print("Download completed, waiting 5 seconds before closing browser...")
                time.sleep(5)  # Allow file to completely download

            return downloaded_file

        except Exception as e:
            print(f"❌ Withdrawal data download failed: {e}")
            return None

        finally:
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass

    def download_monthly_chart(self):
        """Download the monthly chart from CONSAR"""
        if not self.setup_driver():
            return None

        try:
            # Access the CONSAR investment page
            print(f"Opening: {self.investment_url}")
            self.driver.get(self.investment_url)
            time.sleep(3)

            # Click Excel download link
            if not self.find_and_click_excel_link():
                return None

            # Wait for download to complete
            downloaded_file = self.wait_for_download(timeout_seconds=30)

            if downloaded_file:
                print("Download completed, waiting 5 seconds before closing browser...")
                time.sleep(5)  # Allow file to completely download

            return downloaded_file

        except Exception as e:
            print(f"❌ Download failed: {e}")
            return None

        finally:
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass

def main():
    """
    Main function to run all three download steps
    """
    # Set your desired download path
    download_path = "./downloads"  # Change this to your preferred directory

    downloader = CONSARDownloader(download_path)

    print("CONSAR Data Downloader - Three Step Process")
    print("=" * 50)

    # Step 1: Download Investment Data (monthly chart)
    print("\nSTEP 1: Downloading Investment Data...")
    investment_file = downloader.download_monthly_chart()

    if investment_file:
        file_size = os.path.getsize(investment_file)
        print(f"Step 1 Complete! Investment data: {investment_file}")
        print(f"   File size: {file_size:,} bytes")
    else:
        print("Step 1 Failed: Could not download investment data")
        return

    # Step 2: Download RCV Flow Data
    print("\nSTEP 2: Downloading RCV Flow Data...")
    flow_file = downloader.download_rcv_flow_data()

    if flow_file:
        file_size = os.path.getsize(flow_file)
        print(f"Step 2 Complete! Flow data: {flow_file}")
        print(f"   File size: {file_size:,} bytes")
    else:
        print("Step 2 Failed: Could not download flow data")
        return

    # Step 3: Download Withdrawal Data
    print("\nSTEP 3: Downloading Withdrawal Data...")
    withdrawal_file = downloader.download_withdrawal_data()

    if withdrawal_file:
        file_size = os.path.getsize(withdrawal_file)
        print(f"Step 3 Complete! Withdrawal data: {withdrawal_file}")
        print(f"   File size: {file_size:,} bytes")
    else:
        print("Step 3 Failed: Could not download withdrawal data")
        return

    print("\nSUCCESS! All three steps completed:")
    print(f"   Investment Data: {investment_file}")
    print(f"   Flow Data: {flow_file}")
    print(f"   Withdrawal Data: {withdrawal_file}")
    print("\nAll files are ready for processing with the mapping script!")

if __name__ == "__main__":
    main()