from openpyxl import load_workbook, Workbook

class ExcelHandler:
    def __init__(self, filename, sheetname):
        """Initialize the Excel handler by loading the workbook and extracting country names."""
        self.filename = filename
        self.wb = load_workbook(filename=self.filename)
        self.sheet = self.wb[sheetname]

        # Read country names from Column A (Rows 3 to 238)
        self.countries = [self.sheet[f"A{row}"].value for row in range(3, 239)]
        
        # Initialize nested dictionary with default values
        self.country_data = {country: {"UON": "N", "PEACH": "N"} for country in self.countries}

    def update_country_data(self, country, uon, peach):
        """Update the UON and PEACH values for a specific country."""
        if country in self.country_data:
            self.country_data[country]["UON"] = uon
            self.country_data[country]["PEACH"] = peach
        else:
            print(f"Warning: {country} not found in Excel list!")

    def save_to_excel(self):
        """Write updated UON and PEACH values back to the original Excel file and save."""
        for i, country in enumerate(self.countries, start=3):  # Row 3 to 238
            self.sheet[f"B{i}"] = self.country_data[country]["UON"]
            self.sheet[f"C{i}"] = self.country_data[country]["PEACH"]

        self.wb.save(self.filename)
        print(f"Updated data saved to {self.filename}")

    def export_to_excel(self, output_filename="country.xlsx"):
        """Exports the country data dictionary to a new Excel file."""
        wb = Workbook()
        ws = wb.active
        ws.title = "Country Data"

        # Write headers
        ws.append(["Country", "UON", "PEACH"])

        # Write data
        for country, values in self.country_data.items():
            ws.append([country, values["UON"], values["PEACH"]])

        # Save the file
        wb.save(output_filename)
        print(f"Data successfully exported to {output_filename}")
