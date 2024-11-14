import pandas as pd
import numpy as np
import xlwings as xw
import pandas as pd
import numpy as np

class SampleAnalysis:
    """"class for extracting sample and standard data based on the well that has been setup"""
    def __init__(self, sheet_name, table_range, wb=None):
        """
        Initialize the SampleAnalysis class.
        Parameters:
        wb (xw.Book or None): use self.wb = wb and specify it for testing 
        """
        if wb is None:  # If wb is not provided (e.g., in testing), use xw.Book.caller()
            self.wb = xw.Book.caller()  # This gets the workbook that called the script
        else:
            self.wb = wb  # Use the workbook provided explicitly (for testing purposes)
        
        self.sheet_name = sheet_name
        self.table_range = table_range
        self.df = None

    def open_excel(self):
        self.ws = self.wb.sheets[self.sheet_name]
        self.table_range = self.ws.range(self.table_range)

    def extract_data(self):
        sample_names = []
        replicates = []
        sample_types = []
        dilutions = []

        for col in range(0, self.table_range.columns.count):  # Iterate over each column
            for row in range(0, self.table_range.rows.count, 3):  # Step by 3 to capture each sample's data
                sample_name_full = self.ws.range(self.table_range[row, col]).value
                sample_type = self.ws.range(self.table_range[row + 1, col]).value
                dilution = self.ws.range(self.table_range[row + 2, col]).value
                
                if sample_name_full is None or sample_name_full == '':  # If empty cell
                    sample_names.append(np.nan)
                    replicates.append(np.nan)
                    sample_types.append(np.nan)
                    dilutions.append(np.nan)
                else:
                    if '_r' in sample_name_full:
                        sample_name, replicate_raw = sample_name_full.split('_r')
                        replicate = int(replicate_raw)
                    else:
                        sample_name = sample_name_full
                        replicate = np.nan
                    
                    sample_names.append(sample_name)
                    replicates.append(replicate)
                    sample_types.append(sample_type)
                    dilutions.append(dilution)
        
        self.df = pd.DataFrame({
            'Sample Name': sample_names,
            'Replicate': replicates,
            'Sample Type': sample_types,
            'Dilution': dilutions
        })
        
    def generate_well_labels(self):
        num_rows = len(self.df)
        columns = np.array([chr(65 + (i % 8)) for i in range(num_rows)])  # 'A' to 'H' based on the index
        rows = np.array([str(i // 8 + 1).zfill(2) for i in range(num_rows)])  # Row numbers (01 to 08, etc.)
        well_labels = columns + rows
        self.df['Well'] = well_labels
        
    def organize_dataframe(self):
        self.df = self.df[['Well', 'Sample Name', 'Replicate', 'Sample Type', 'Dilution']]

    def match_ct_values(self, ct_file_path, ct_sheet_name, ct_column_name='Cq'):
        """"Process data to reference file and match CT on wells"""
     # function finds the ct values that match the data based on well 
        ct_wb = xw.Book(ct_file_path)
        ct_ws = ct_wb.sheets[ct_sheet_name]
        # Extract data from the sheet into a DataFrame
        ct_data = ct_ws.used_range.value
        ct_df = pd.DataFrame(ct_data[1:], columns=ct_data[0])
        # Check if the necessary columns are present
        if 'Well' not in ct_df.columns or ct_column_name not in ct_df.columns:
            raise ValueError(f"The Ct file must contain 'Well' and '{ct_column_name}' columns.")

        # Merge Ct/Cq values with the main DataFrame on the 'Well' column
        self.df = pd.merge(self.df, ct_df[['Well', ct_column_name]], on='Well', how='left')
        
        # Rename the column to 'Ct/Cq' regardless of its original name
        self.df.rename(columns={ct_column_name: 'Ct/Cq'}, inplace=True)
        
        # Close the workbook
        ct_wb.close()
    
    def process_data(self):
        self.open_excel()
        self.extract_data()
        self.generate_well_labels()
        self.organize_dataframe()
    
    def process_data_ct(self, ct_file_path, ct_sheet_name, ct_column_name='Cq'):
        self.process_data()  # Perform the standard processing steps
        self.match_ct_values(ct_file_path, ct_sheet_name, ct_column_name)  # Include Ct values

    def get_dataframe(self):
        return self.df
    
def calculate_dilution_summary(df):
    """
    Calculates the average and standard deviation of Ct values for each dilution and sample name,
    and adds a log-transformed dilution column, rounded to two decimal places.

    Parameters:
    df (pd.DataFrame): The DataFrame containing sample data, including Ct/Cq values and Dilution.

    Returns:
    pd.DataFrame: A summary DataFrame with columns ['Sample Name', 'Dilution', 'Log Dilution', 'CT', 'CT SD'].
    """
    # Filter for "Standard" sample type and use .loc to prevent SettingWithCopyWarning
    standard_df = df.loc[df['Sample Type'] == 'Standard'].copy()

    # Ensure Dilution is numeric for log transformation
    standard_df['Dilution'] = pd.to_numeric(standard_df['Dilution'], errors='coerce')

    # Group by 'Sample Name' and 'Dilution', calculate mean and std of Ct/Cq, and add log Dilution
    dilution_summary = (
        standard_df.groupby(['Sample Name', 'Dilution'])['Ct/Cq']
        .agg(avg_ct_cq='mean', std_ct_cq='std')
        .reset_index()
    )
    dilution_summary['Dilution'] = dilution_summary['Dilution'].map(lambda x: "{:.2E}".format(x))
    dilution_summary['log_dilution'] = np.log10(dilution_summary['Dilution'].astype(float).replace(0, np.nan)).round(2)
    # show column in the correct way 
    dilution_summary = dilution_summary[['Sample Name', 'Dilution', 'log_dilution', 'avg_ct_cq', 'std_ct_cq']]
    dilution_summary.columns = ['Sample Name', 'Template Conc. (copies/rxn)', 'Log template Conc.', 'CT', 'Ct SD']  # Rename columns
    return dilution_summary
