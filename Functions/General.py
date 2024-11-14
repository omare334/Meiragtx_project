# library of function for the most general commans like saving df and putting data into excel file 
import xlwings as xw
import pandas as pd

def write_data_to_excel(data, sheet_name, start_cell='A1'):
    # Get the caller workbook and the specified sheet
    wb = xw.Book.caller()
    ws = wb.sheets[sheet_name]
    # Convert data to list format if it's a DataFrame
    data = data.reset_index(drop=True)
    if hasattr(data, 'values'):  
        data = data.values.tolist()
    ws.range(start_cell).value = data

def read_excel_to_df(sheet_name, header_range, data_range_start):
    # Reference the workbook and the specified sheet
    wb = xw.Book.caller()
    sheet = wb.sheets[sheet_name]
    # Read headers
    headers = sheet.range(header_range).value
    # Read the data below the headers
    data_range = sheet.range(data_range_start).expand('down').expand('right')
    data = data_range.value
    # Create the DataFrame
    df = pd.DataFrame(data, columns=headers)
    return df

def format_column(df, columns, format_str):
 #convert columns to certain notation
    for col in columns:
        df[col] = df[col].apply(lambda x: format_str.format(x))

