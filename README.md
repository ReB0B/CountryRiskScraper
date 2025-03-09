# Country Risk Scraper
Web scraper using selenium to check whether a student has increased evidentiary requirements based on the student's education provider and country of citizenship.

## Use Case
This web scraper helps universities and education providers determine visa requirements for international students under the Simplified Student Visa Framework (SSVF). It checks whether students from specific countries qualify for streamlined visa processing or need to provide additional evidence (e.g., financial capacity, English proficiency) based on their Country Risk Rating and the institution’s Risk Rating.

### Who can use it:

1. Universities: Assess visa requirements for applicants.

2. Education Agents: Guide students on visa documentation.

3. Students: Self-check visa application requirements.

The tool automates the process and saves time for a mundane task

## Dependencies
1. Selenium
2. python-dotenv
3. openpyxl

## How to use
1. Change the env file to .env and populate data
2. Change the university name in main.py

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

## Running the Program
```python3 main.py```

Note: If accessing through a mamba/conda env, you can run main.py on your code editor