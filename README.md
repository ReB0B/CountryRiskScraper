# Country Risk Scraper
Web scraper using Selenium to check whether a student has increased evidentiary requirements based on the student's education provider and country of citizenship.

## Use Case
This web scraper helps universities and education providers determine visa requirements for international students under the Simplified Student Visa Framework (SSVF). It checks whether students from specific countries qualify for streamlined visa processing or need to provide additional evidence (e.g., financial capacity, English proficiency) based on their Country Risk Rating and the institution's Risk Rating.

### Who can use it:
- **Universities**: Assess visa requirements for applicants
- **Education Agents**: Guide students on visa documentation

The tool automates the process and saves time for repetitive manual checks.

## Dependencies
- Selenium
- python-dotenv  
- openpyxl
- tkinter (included with most Python installations)

## Setup
1. Set up your `.env` file with the required configuration (see Configuration section below)
2. Ensure you have ChromeDriver installed and configured

## How to Use

### Running the Scraper
```bash
python3 main.py
```

The program will automatically detect if GUI is available:
- **GUI Mode**: A window opens where you can configure all settings
- **Console Mode**: If GUI is unavailable, a console interface guides you through configuration

### Configuration Steps:
1. **Select Input Excel File**: Browse and choose your country data file
2. **Choose Sheet Name**: Auto-detected or manually entered
3. **Enter Website URL**: Document checklist website (pre-filled with default)
4. **Configure Universities**: Enter names for both universities to compare
5. **Specify Output File**: Choose location and name for results
6. **Start Scraping**: Confirm settings and begin the process

### Testing the Interface
To test the configuration interface without running the full scraper:
```bash
python3 test_gui.py
```

## Configuration

### Required .env File Settings
Create a `.env` file in the project directory with the following settings:

```env
# ChromeDriver Configuration
CHROMEDRIVER_LOCATION=/path/to/chromedriver
# Download from: https://googlechromelabs.github.io/chrome-for-testing/
# Ensure Chrome browser version matches driver version

# Web Element Selectors (for the immigration website)
PASSPORT_FIELD=select2-drpWebEvtCountryPassport-container
EDUCATION_PROVIDER_FIELD=select2-drpWebEvtProvider-container
CRICOS_CODE_FIELD=select2-drpWebEvtCRICOS-container
```

**Note**: The website URL, university names, input/output files are now configured through the GUI interface and don't need to be in the .env file.

## Features

### Interface Modes
- **GUI Mode**: User-friendly graphical interface (automatically selected when available)
- **Console Mode**: Text-based interface (fallback when GUI is unavailable)

### Key Capabilities
- **File Browser**: Easy Excel file selection
- **Sheet Detection**: Automatic detection and selection of available sheets
- **Website Configuration**: Configurable document checklist website URL
- **University Comparison**: Setup for comparing two different universities
- **Input Validation**: Comprehensive validation to prevent configuration errors
- **Flexible Output**: Custom naming and location for result files
- **Status Feedback**: Real-time feedback during configuration and processing

## Input Requirements

### Excel File Format
- **Column A**: Country names (starting from row 3)
- **Sheet**: Should contain country data for visa risk assessment
- **Format**: .xlsx or .xls files supported

### University Information
- Names of two universities with different risk ratings for comparison
- Will be used to check evidentiary requirements for each institution

## Output
- **Updated Original File**: Original Excel file with results added
- **New Results File**: Separate file with processed data (custom naming)
- **Processing Status**: Console output showing progress and completion

## Notes
- Compatible with mamba/conda environments
- ChromeDriver version must match your installed Chrome browser
- Internet connection required for website scraping