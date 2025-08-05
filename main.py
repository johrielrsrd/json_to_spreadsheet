import json
from openpyxl import Workbook

file_path = 'data.json'

def read_json_file():
    json_list = []
    try:
        with open(file_path, 'r') as file:
            json_data = json.load(file)

        for item in json_data:
            item_list = [item["profile"]["name"], item["email"], item["profile"]["address"], item["profile"]["company"]]
            json_list.append(item_list)

        return json_list

    except FileNotFoundError:
        print("File not found.")
    except json.JSONDecodeError:
        print("File is not valid.")

def create_spreadsheet(json_list_data):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sample Sheet"

    sheet["A1"] = "Name"
    sheet["B1"] = "Email"
    sheet["C1"] = "Address"
    sheet["D1"] = "Company"

    sheet_data = json_list_data

    for row in sheet_data:
        print(row)
        sheet.append(row)

    file_name = "example.xlsx"
    workbook.save(file_name)

    print("Excel created.")

created_list_from_json = read_json_file()
create_spreadsheet(created_list_from_json)