import pandas as pd
import xlwings as xw
from Functions.General import read_excel_to_df,format_column

def main():
    # exlcuding the first row if Exclude? is not in the header 
    df = read_excel_to_df(sheet_name="Samples", header_range="A4:G4", data_range_start="A5")
    wb = xw.Book.caller()
    standards_ws = wb.sheets['Samples']
    df_filtered = df[df['Exclude?'] != 'Yes']
    # Grouping data 
    grouped = df_filtered.groupby(['Sample name', 'Dilution']).agg(
        Replicate_mean_VG_conc=('VG/mL', 'mean'),
        Replicate_SD=('VG/mL', 'std')
    ).reset_index()
    # Calculate Replicate CV as (Replicate SD / Replicate mean VG conc.) * 100
    grouped['Replicate_CV'] = (grouped['Replicate_SD'] / grouped['Replicate_mean_VG_conc']) * 100
    titre_stats = df_filtered.groupby('Sample name').agg(
        VG_titre=('VG/mL', 'mean'),
        VG_titre_SD=('VG/mL', 'std')).reset_index()
    # Calculate CV (%) as (VG titre SD / VG titre) * 100
    titre_stats['CV_percent'] = (titre_stats['VG_titre_SD'] / titre_stats['VG_titre']) * 100
    final_df = pd.merge(grouped, titre_stats, on='Sample name', how='left')
    # column rename 
    final_df = final_df.rename(columns={'Sample name': 'Sample ID','Replicate_mean_VG_conc': 'Replicate mean VG conc. (VG/mL)',
        'Replicate_SD': 'Replicate SD','Replicate_CV': 'Replicate CV','VG_titre': 'VG titre (VG/mL)',
        'VG_titre_SD': 'VG titre SD','CV_percent': 'CV (%)'
    })
    # Format the values in scientific notation where needed
    scientific_columns = ['Replicate mean VG conc. (VG/mL)','Replicate SD','VG titre (VG/mL)','VG titre SD']
    percentage_columns = ['Replicate CV','CV (%)']
    format_column(final_df, scientific_columns, "{:.2e}")
    format_column(final_df, percentage_columns, "{:.2f}%")
    # Placing data as list 
    data_list = final_df.values.tolist()
    start_cell = standards_ws.range('J5')
    standards_ws.range(start_cell, start_cell.offset(len(data_list)-1, len(data_list[0])-1)).value = data_list


