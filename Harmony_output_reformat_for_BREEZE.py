import os
import pandas as pd
import csv

# Define folder paths for input and output files
input_folder_txt = r"C:\Users\kun.qian\Desktop\Projects\U2OS phospholipidoses assay\Image analysis\hit confirmation and re-screen\txt"
output_folder_csv = r'C:\Users\kun.qian\Desktop\Projects\U2OS phospholipidoses assay\Image analysis\hit confirmation and re-screen\csv'
output_folder_excel = r'C:\Users\kun.qian\Desktop\Projects\U2OS phospholipidoses assay\Image analysis\hit confirmation and re-screen\excel'
conversion_file = r'C:\Users\kun.qian\Desktop\Projects\U2OS phospholipidoses assay\Image analysis\hit confirmation and re-screen\plate_conversion.xlsx'
platemap = r'C:\Users\kun.qian\Desktop\Projects\U2OS phospholipidoses assay\Image analysis\Big screen\PLATEMAP_CC02444_screen_plates.xlsx'
matched_file_path = r'C:\Users\kun.qian\Desktop\Projects\U2OS phospholipidoses assay\Image analysis\Big screen2\matched.xlsx'
combined_file_path = r'C:\Users\kun.qian\Desktop\Projects\U2OS phospholipidoses assay\Image analysis\Big screen2\combined.xlsx'
output_file_path = r'C:\Users\kun.qian\Desktop\Projects\U2OS phospholipidoses assay\Image analysis\Big screen2\Data for BREEZE.xlsx'
row_start = 9

def convert_txt_to_csv(input_folder, output_folder):
    files = [f for f in os.listdir(input_folder) if f.endswith('.txt')]
    for file in files:
        input_file = os.path.join(input_folder, file)
        output_file = os.path.join(output_folder, f"{os.path.splitext(file)[0]}.csv")
        with open(input_file, 'r') as txt_file:
            data = txt_file.readlines()
        with open(output_file, 'w', newline='') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerows([line.strip().split('\t') for line in data])

def rename_files_based_on_content(input_folder):
    for file in os.listdir(output_folder_csv):
        csv_file_path = os.path.join(output_folder_csv, file)
        with open(csv_file_path, 'r') as file:
            csv_data = list(csv.reader(file))
            new_filename = csv_data[3][1]
        new_csv_file_path = os.path.join(output_folder_csv, new_filename + '.csv')
        os.rename(csv_file_path, new_csv_file_path)

def truncate_csv_and_save_as_xlsx(csv_folder_path, excel_folder_path, row_start):
    csv_files = [file for file in os.listdir(csv_folder_path) if file.endswith('.csv')]
    for csv_file in csv_files:
        csv_file_path = os.path.join(csv_folder_path, csv_file)
        rows = []
        with open(csv_file_path, 'r') as file:
            csv_reader = csv.reader(file)
            for _ in range(row_start - 1):
                next(csv_reader)
            for row in csv_reader:
                rows.append(row[:16])
        df = pd.DataFrame(rows)
        excel_file_path = os.path.join(excel_folder_path, os.path.splitext(csv_file)[0] + '.xlsx')
        df.to_excel(excel_file_path, index=False, header=False)

def add_well_id_and_plate_id(directory, conversion_file):
    conversion_df = pd.read_excel(conversion_file)
    well_ids = [f"{row}{'{:02d}'.format(col)}" for row in 'ABCDEFGHIJKLMNOP' for col in range(1, 25)]
    for filename in os.listdir(directory):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(directory, filename)
            excel_df = pd.read_excel(file_path)
            barcode = os.path.splitext(filename)[0]
            if barcode in conversion_df['Barcode'].values:
                plate_id = conversion_df.loc[conversion_df['Barcode'] == barcode, 'PlateID'].values[0]
                excel_df.insert(0, 'WellID', well_ids[:len(excel_df)])
                excel_df.insert(1, 'PlateID', plate_id)
                excel_df.to_excel(file_path, index=False)

def combine_files(folder_path, combined_file_path):
    combined_data = pd.DataFrame()
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            df = pd.read_excel(os.path.join(folder_path, filename))
            combined_data = combined_data.append(df[['WellID', 'PlateID', 'Cell Selected - Number of Objects']].rename(columns={'Cell Selected - Number of Objects': 'WELL_SIGNAL'}), ignore_index=True)
    combined_data.to_excel(combined_file_path, index=False)
    
def merge_and_sort_data(combined_file_path, platemap, matched_file_path):
    combined_data = pd.read_excel(combined_file_path)
    matching_data = pd.read_excel(platemap)
    matching_data = matching_data.rename(columns={'Platt ID': 'PlateID', 'Well': 'WellID'})
    merged_data = pd.merge(combined_data, matching_data, on=['PlateID', 'WellID'], how='left')
    merged_data = merged_data.sort_values(by='PlateID')
    plate_mapping = {plate: i + 1 for i, plate in enumerate(merged_data['PlateID'].unique()[:25])}
    merged_data['PLATE'] = merged_data['PlateID'].map(plate_mapping)
    merged_data['Batch nr'] = merged_data.apply(lambda row: row['Compound ID'] if row['Batch nr'] in ['DMSO', 'Water'] else row['Batch nr'], axis=1)
    merged_data.to_excel(matched_file_path, index=False)

def format_for_breeze(matched_file_path, output_file_path):
    matched_data = pd.read_excel(matched_file_path).rename(columns={'WellID': 'WELL','Batch nr': 'DRUG_NAME', 'Conc (mM)': 'CONCENTRATION'})
    matched_data['CONCENTRATION'] = 10000
    matched_data['SCREEN_NAME'] = 'KQ_U2OS_PL_10uM_screen'
    matched_data['DRUG_NAME'] = matched_data['DRUG_NAME'].replace('TAM', 'POS')
    matched_data[['WELL', 'PLATE', 'DRUG_NAME', 'CONCENTRATION', 'SCREEN_NAME', 'WELL_SIGNAL']].to_excel(output_file_path, index=False)

# Ensure the output folders exist before running the script
os.makedirs(output_folder_csv, exist_ok=True)
os.makedirs(output_folder_excel, exist_ok=True)

# Execute all steps
convert_txt_to_csv(input_folder_txt, output_folder_csv)
rename_files_based_on_content(output_folder_csv)
truncate_csv_and_save_as_xlsx(output_folder_csv, output_folder_excel, row_start)
add_well_id_and_plate_id(output_folder_excel, conversion_file)
combine_files(output_folder_excel, combined_file_path)
merge_and_sort_data(combined_file_path, platemap, matched_file_path)
format_for_breeze(matched_file_path, output_file_path)
