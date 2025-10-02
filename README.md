# Country Risk Scraper
Web scraper using selenium to check whether a student has increased evidentiary requirements based on the student's education provider and country of citizenship.

## Use Case
This web scraper helps universities and education providers determine visa requirements for international students under the Simplified Student Visa Framework (SSVF). It checks whether students from specific countries qualify for streamlined visa processing or need to provide additional evidence (e.g., financial capacity, English proficiency) based on their Country Risk Rating and the institution's Risk Rating.

## Prerequisite
Risk ratings of a two different universities with a different risk rating should be known.

### Who can use it:

1. Universities: Assess visa requirements for applicants.
2. Education Agents: Guide students on visa documentation.

The tool automates the process and saves time for a mundane task

## Dependencies
1. Selenium
2. python-dotenv
3. openpyxl
4. tkinter (included with most Python installations)

## How to use

### Method 1: Interactive Mode (Recommended)
1. Set up your .env file with the required configuration (see below)
2. Run the program: `python3 main.py`
3. The program will automatically detect if GUI is available:
   - **GUI Mode**: A window will open where you can select files and configure settings
   - **Console Mode**: If GUI is not available, a console interface will guide you through the configuration
4. Configure your settings:
   - Select your input Excel file
   - Choose or enter the sheet name
   - Enter the document checklist website URL
   - Specify the names of the two universities to compare
   - Specify the output file location
   - Confirm and start the scraping process

### Method 2: Test Interface Only
To test the interface without running the full scraper:
```bash
python3 test_gui.py
```

### Method 3: Legacy Mode (Environment Variables)
The traditional method using environment variables is still supported for backward compatibility.

### What to use in the .env
1. CHROMEDRIVER_LOCATION
    - absolute path to the chromedriver location (e.g. /usr/local/bin/chromedriver)
    - can be downloaded from (https://googlechromelabs.github.io/chrome-for-testing/)
    - browser has to match the driver version
2. DOCUMENT_CHECKLIST_WEBSITE=https://immi.homeaffairs.gov.au/visas/web-evidentiary-tool
3. PASSPORT_FIELD=XXX
    - div id for the passport field
    - passport_field
4. EDUCATION_PROVIDER_FIELD=XXX
    - div id for the education provider field
5. CRICOS_CODE_FIELD=XXX
    - can be left as is unless the education provider needs to be accesed using cricos code (either one is fine as the other field populates itself)
6. FILE_NAME=./XXX
    - file name of the excel
    - at least requires the names of countries
    - If there are a limited number of countries, the row number needs to be changed
7. SHEET_NAME=XXX
    - sheet name for the excel
8. UNI1/UNI2 = Name of the universities to be compared

## Running the Program

### GUI Mode (Recommended)
```bash
python3 main.py
```

### Test GUI Only
```bash
python3 test_gui.py
```

### Features of the Interactive Interface

#### GUI Mode (when available):
- **File Browser**: Easily select input Excel files
- **Sheet Detection**: Automatically detect and select available sheets
- **Website Configuration**: Enter the document checklist website URL
- **University Setup**: Configure names for both universities to compare
- **Validation**: Input validation to prevent common errors
- **User-Friendly**: Clear interface with status messages
- **Flexible Output**: Choose custom output file names and locations

#### Console Mode (fallback):
- **File Path Input**: Enter file paths directly
- **Auto-Detection**: Automatically detect and suggest available sheets
- **Website Configuration**: Enter website URL with default suggestion
- **University Setup**: Enter names for both universities to compare
- **Validation**: Input validation and error handling
- **Configuration Summary**: Review all settings before proceeding

Note: If accessing through a mamba/conda env, you can run main.py in your code editor

## GUI Components

### Main Window Features:
1. **Input Excel File Selection**: Browse and select your Excel file containing country data
2. **Sheet Name Configuration**: 
   - Manual entry of sheet name
   - Auto-detection of available sheets with dropdown selection
3. **Website URL Configuration**: Enter the document checklist website URL (pre-filled with default)
4. **University Configuration**: Enter names for both universities to compare risk ratings
5. **Output File Configuration**: Specify where to save the results
6. **Validation**: Comprehensive input validation before starting the scraper
7. **Status Messages**: Real-time feedback on configuration status