import xlwings as xw
from Functions.Standards_class import SampleAnalysis 
from Functions.Samples_class import calculate_vg_rx
from Functions.Standards_class import calculate_dilution_summary
from Functions.Samples_class import QPCRAnalysis
from Functions.General import write_data_to_excel

def main(ct_file_path='RawDataFile.xlsx', ct_sheet_name=0, ct_column_name='Cq', use_existing_values=False):
    # Open workbook in Excel (assuming function is called from within Excel)
    wb = xw.Book.caller()
    standards_ws = wb.sheets['Standards']
    sheet_name = 'PlateLayout' 
    table_range = 'E3:P26'  
    # Process data with Ct values
    analysis = SampleAnalysis(sheet_name, table_range)
    analysis.process_data_ct(ct_file_path, ct_sheet_name, ct_column_name)
    # Get DataFrame with Ct values
    df_with_ct = analysis.get_dataframe()
    if not use_existing_values:
        # Get DataFrame with Ct values
        df_with_ct = analysis.get_dataframe()
        dilution_summary = calculate_dilution_summary(df_with_ct)
        # Initating pcr analysis to caclulate the slot and intercept
        qpcr = QPCRAnalysis(dilution_summary)
        qpcr.calculate_rsq_efficiency()
        slope, intercept = qpcr.get_slope_intercept()
    else:
        # finds the slope and intercept 
        slope = None
        intercept = None
        for cell in standards_ws.used_range:
            if cell.value == "Slope":
                slope = cell.offset(0, 1).value  
            elif cell.value == "Y-Intercept":
                intercept = cell.offset(0, 1).value  
            if slope is not None and intercept is not None:
                break
        # assertion
        if slope is None or intercept is None:
            raise ValueError("Slope or Y-intercept not found in the expected format in the Standards sheet, please put Slope and Y-intercept value to the right of the label.")
    df_sample_only = calculate_vg_rx(df_with_ct, intercept, slope)
    # Plot samples in order
    df = df_sample_only[['Sample Name', 'Dilution', 'Replicate', 'Ct/Cq', 'VG/rx well', 'VG/ml']]
    write_data_to_excel(df, sheet_name='Samples', start_cell='A5')

