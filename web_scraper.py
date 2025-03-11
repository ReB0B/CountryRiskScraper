from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from dotenv import load_dotenv
import os
import time

class WebScraper:
    def __init__(self):
        """Initialize the web scraper by loading environment variables and setting up the browser."""
        load_dotenv()  # Load environment variables

        # Load configuration from .env
        self.chromedriver_location = os.getenv("CHROMEDRIVER_LOCATION")
        self.document_checklist_website = os.getenv("DOCUMENT_CHECKLIST_WEBSITE")
        self.passport_field = os.getenv("PASSPORT_FIELD")
        self.education_provider_field = os.getenv("EDUCATION_PROVIDER_FIELD")
        self.submit_button_id = "btnSubmitEvidence"

        # Setup Selenium Chrome driver
        service = Service(self.chromedriver_location)
        chrome_options = Options()
        # chrome_options.add_argument("--headless")  # Uncomment for headless mode
        self.driver = webdriver.Chrome(service=service, options=chrome_options)

        self.driver.implicitly_wait(5)  # Apply implicit wait globally

    def open_website(self):
        """Open the target website."""
        self.driver.get(self.document_checklist_website)
        print("Website loaded successfully.")

    def process_countries(self, countries, excel_handler, provider):
        """
        Loops through all countries, performs checks, and updates the ExcelHandler.
        
        Args:
          - countries: List of country names from the ExcelHandler.
          - excel_handler: Instance of ExcelHandler to store UNI1 and UNI2 results.
          - provider: The education provider to select (value from .env for UNI1 or UNI2).
        """
        uni1 = os.getenv("UNI1")
        uni2 = os.getenv("UNI2")
        institution_selected = False  # Track if the institution is already selected

        for index, country in enumerate(countries):
            print(f"Processing country: {country} with provider: {provider}")
            self.select_country(country)

            if not institution_selected:  # Select institution only once
                self.select_education_provider(provider)
                institution_selected = True  # Mark institution as selected

            self.select_radio_option()
            self.click_display_evidence()
            contains_evidence = self.check_evidence_on_page()

            # Update the corresponding provider value based on the evidence check
            if provider == uni1:
                uni1_value = "Y" if contains_evidence else "N"
                excel_handler.update_country_data(country, uni1_value, excel_handler.country_data[country][uni2])
            elif provider == uni2:
                uni2_value = "Y" if contains_evidence else "N"
                excel_handler.update_country_data(country, excel_handler.country_data[country][uni1], uni2_value)

        print(f"All countries processed for provider: {provider}")

    def select_country(self, country_name):
        """Selects a country from the dropdown."""
        try:
            dropdown = self.driver.find_element(By.ID, self.passport_field)
            dropdown.click()
            time.sleep(1)  # Short wait for dropdown to activate

            search_box = self.driver.find_element(By.CLASS_NAME, "select2-search__field")
            search_box.send_keys(country_name)
            search_box.send_keys(Keys.ENTER)
            time.sleep(2)  # Allow selection to process

            print(f"Country '{country_name}' selected successfully.")
        except Exception as e:
            print(f"Error selecting country {country_name}: {e}")

    def select_education_provider(self, provider_name):
        """Selects an education provider from the dropdown. This is only done once per loop."""
        try:
            dropdown = self.driver.find_element(By.ID, self.education_provider_field)
            dropdown.click()
            time.sleep(1)  # Allow dropdown to open

            search_box = self.driver.find_element(By.CLASS_NAME, "select2-search__field")
            search_box.send_keys(provider_name)
            time.sleep(1)  # Allow options to populate

            # Select the first result containing provider name
            options = self.driver.find_elements(By.CLASS_NAME, "select2-results__option")
            for option in options:
                if provider_name in option.text:
                    option.click()
                    print(f"Selected {provider_name}.")
                    break

            time.sleep(1)  # Allow selection to process
        except Exception as e:
            print(f"Error selecting provider {provider_name}: {e}")

    def select_radio_option(self):
        """Selects the radio button from the image (for='01')."""
        try:
            radio_button = self.driver.find_element(By.ID, "01")  # ID is '01' on the website
            radio_button.click()
            print("Radio button selected successfully.")
            time.sleep(1)  # Allow time for selection
        except Exception as e:
            print(f"Error selecting the radio button: {e}")

    def click_display_evidence(self):
        """Clicks the 'Display Evidence' button."""
        try:
            button = self.driver.find_element(By.ID, self.submit_button_id)
            button.click()
            time.sleep(3)  # Wait for the new data to load
            print("Clicked 'Display Evidence' button.")
        except Exception as e:
            print(f"Error clicking the Display Evidence button: {e}")

    def check_evidence_on_page(self):
        """Checks if 'Evidence of financial capacity' or 'Evidence of English language ability' appears inside <h3> tags."""
        try:
            h3_elements = self.driver.find_elements(By.XPATH, "//h3")
            h3_texts = [h3.text.lower() for h3 in h3_elements]

            contains_financial = any("evidence of financial capacity" in text for text in h3_texts)
            contains_english = any("evidence of english language ability" in text for text in h3_texts)

            if contains_financial or contains_english:
                print("Evidence found on the page.")
                return True
            else:
                print("No evidence found on the page.")
                return False
        except Exception as e:
            print(f"Error checking evidence on page: {e}")
            return False

    def close_browser(self):
        """Close the browser session."""
        self.driver.quit()
        print("Browser closed.")
