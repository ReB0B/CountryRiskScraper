from openpyxl import load_workbook

wb = load_workbook(filename="./XXXXX.xlsx")

sheet = wb["XXXX"]

countries = [sheet[f"A{row}"].value for row in range(3, 239)]

country_data = {country: {"UON": "N", "PEACH": "N"} for country in countries}

print(country_data)
# print(sheet["A2"].value)