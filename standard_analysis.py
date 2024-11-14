from Functions.Standards_class import SampleAnalysis
from Functions.Standards_class import calculate_dilution_summary
from Functions.General import write_data_to_excel

def main(ct_file_path, ct_sheet_name = 0 , ct_column_name = 'Cq'):
    # Open the workbook
    sheet_name = 'PlateLayout'
    table_range = 'E3:P26'  
    # Create an instance of SampleAnalysis
    analysis = SampleAnalysis(sheet_name, table_range)
    analysis.process_data_ct(ct_file_path, ct_sheet_name, ct_column_name)
    # Getting the dataframe
    df_with_ct = analysis.get_dataframe()
    # diltuion summary calculation 
    dilution_summary = calculate_dilution_summary(df_with_ct)
    dilution_summary = dilution_summary.reset_index(drop=True)
    # placing the data in the right area 
    write_data_to_excel(dilution_summary, sheet_name='Standards', start_cell='A3')
    df_with_ct.to_csv('df_with_ct.csv')
    return df_with_ct,dilution_summary



