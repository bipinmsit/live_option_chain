import pandas as pd


def read_parameters_from_excel(file_path, sheet_name, parameter_column):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Extract the parameter column
        parameters = df[parameter_column].tolist()

        return parameters
    except Exception as e:
        print("Error:", e)
