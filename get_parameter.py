import pandas as pd
import xlwings as xw
from datetime import datetime


def read_parameters_from_excel(file_path, sheet_name, parameter_column):
    try:
        parameters = ""
        # Read the Excel file
        ws = xw.Book(file_path).sheets[sheet_name]

        # Extract the parameter column
        parameters = ws.range(parameter_column).value

        if (
            parameter_column == "M2"
            or parameter_column == "N2"
            or parameter_column == "O2"
            or parameter_column == "P2"
        ):
            date = datetime.date(parameters)

            # Parse the date string
            date_obj = datetime.strptime(str(date), "%Y-%m-%d")

            # Format the date as "28-Mar-2024"
            parameters = date_obj.strftime("%d-%b-%Y")

        return parameters
    except Exception as e:
        print("Error:", e)


# excel_file_path = (
#     r"C:\Users\bipin\OneDrive\Desktop\bipinmsit\tradetools\option_chain.xlsx"
# )
# test = read_parameters_from_excel(excel_file_path, "option_chain", "M2")
# print(test)
