from web_scraper import WebScraper
from excel_handler import ExcelHandler
from gui_handler import GUIHandler
import os
import sys
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Get user inputs through GUI
print("Starting Country Risk Scraper...")
print("Please configure the settings in the GUI window that will open...")

user_inputs = GUIHandler.run_gui()

# Check if user cancelled the operation
if user_inputs is None:
    print("Operation cancelled by user.")
    sys.exit(0)

print(f"Configuration received:")
print(f"  Input file: {user_inputs['input_filename']}")
print(f"  Sheet name: {user_inputs['sheet_name']}")
print(f"  Website URL: {user_inputs['website_url']}")
print(f"  University 1: {user_inputs['uni1_name']}")
print(f"  University 2: {user_inputs['uni2_name']}")
print(f"  Output file: {user_inputs['output_filename']}")

# Initialize the Excel handler with user-provided values
excel_loader = ExcelHandler(
    filename=user_inputs['input_filename'], 
    sheetname=user_inputs['sheet_name'],
    uni1_name=user_inputs['uni1_name'],
    uni2_name=user_inputs['uni2_name']
)

# Initialize the web scraper with user-provided website URL
scraper = WebScraper(website_url=user_inputs['website_url'])
scraper.open_website()

# Process all countries for UNI1 (first provider)
scraper.process_countries(excel_loader.countries, excel_loader, provider=user_inputs['uni1_name'])

# Process all countries again for UNI2 (second provider)
scraper.process_countries(excel_loader.countries, excel_loader, provider=user_inputs['uni2_name'])

# Save updated values to the original Excel file
excel_loader.save_to_excel()

# Export final dictionary to a new Excel file with user-specified filename
excel_loader.export_to_excel(user_inputs['output_filename'])

# Close the browser when done
scraper.close_browser()

print("\n" + "="*50)
print("SCRAPING COMPLETED SUCCESSFULLY!")
print("="*50)
print(f"✓ Updated original file: {user_inputs['input_filename']}")
print(f"✓ Exported results to: {user_inputs['output_filename']}")
print("The scraping process has finished.")
