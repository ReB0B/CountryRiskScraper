from web_scraper import WebScraper
from excel_handler import ExcelHandler
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Initialize the Excel handler
file_name = os.getenv("FILE_NAME")
sheet_name = os.getenv("SHEET_NAME")
excel_loader = ExcelHandler(filename=file_name, sheetname=sheet_name)

# Initialize the web scraper
scraper = WebScraper()
scraper.open_website()

# Process all countries for University of Newcastle (UON)
scraper.process_countries(excel_loader.countries, excel_loader, provider="University of Newcastle")

# Process all countries again for The Peach Institute (PEACH)
scraper.process_countries(excel_loader.countries, excel_loader, provider="The Peach Institute")

# Save updated values to original Excel file
excel_loader.save_to_excel()

# Export final dictionary to a new Excel file if needed
# excel_loader.export_to_excel("country.xlsx")

# Close the browser when done
scraper.close_browser()
