import os
import pandas as pd
from tabula import read_pdf
from PyPDF2 import PdfReader
import openpyxl

# Set JAVA_HOME environment variable if needed
os.environ["JAVA_HOME"] = "C:/Program Files/Java/jdk-22"

pdf_path = "enter your path"

pages = input("Enter the page numbers to process (e.g., '1,2,3' for specific pages or 'all' for all pages): ").strip()

# Define the coordinates of the area containing the table [top, left, bottom, right]
area = [50, 20, 1000, 900]

output_dir = r"C:\Users\DEV\Desktop"
output_file_path_excel = os.path.join(output_dir, "sample output.xlsx")

def process_table(data, page_num, writer):
    data = data.dropna(axis=0, how="all").dropna(axis=1, how="all")

    data = data.apply(lambda x: x.str.split('\n').str[0] if x.dtype == object else x)

    expected_columns = 19 
    current_columns = len(data.columns)

    if current_columns < expected_columns:
        data = pd.concat([data, pd.DataFrame(columns=[None] * (expected_columns - current_columns))], axis=1)

    data.iloc[:, [3, 4]] = data.iloc[:, [4, 3]].values

    headers = ["SL. NO.", "COMMODITIES", "UNIT", "QUANTITY", "VALUE", "VALUE", "QUANTITY", 
               "VALUE", "VALUE", "QUANTITY", "VALUE", "VALUE", "VALUE", "VALUE", "QUANTITY", 
               "VALUE", "VALUE", "VALUE", "VALUE"] 

    data.columns = headers

    data = data.drop(data.tail(1).index)

    data.reset_index(drop=True, inplace=True)

    def split_and_expand(column):
        split_col = column.str.split(' ', expand=True)
        max_splits = 2  
        if split_col.shape[1] < max_splits:
            for _ in range(max_splits - split_col.shape[1]):
                split_col[split_col.shape[1]] = None
        elif split_col.shape[1] > max_splits:
            split_col[1] = split_col.iloc[:, 1:].apply(lambda x: ' '.join(x.dropna().astype(str)), axis=1)
            split_col = split_col.iloc[:, :2]
        return split_col

    split_values_15 = split_and_expand(data.iloc[:, 14])
    split_values_16 = split_and_expand(data.iloc[:, 15])

    split_values_15.columns = ['QUANTITY_RUPEES', 'QUANTITY_DOLLARS'] 
    split_values_16.columns = ['VALUE_RUPEES_1', 'VALUE_DOLLARS_1']  

   
    data = pd.concat([data.iloc[:, :14], split_values_15, split_values_16, data.iloc[:, 16:]], axis=1)

    
    data["PAGE_NUM"] = page_num

    
    data.to_excel(writer, sheet_name=f'Page_{page_num}', index=False, header=True)

if pages.lower() == 'all':
    
    reader = PdfReader(pdf_path)
    num_pages = len(reader.pages)

    with pd.ExcelWriter(output_file_path_excel, engine='openpyxl') as writer:
        for page_num in range(1, num_pages + 1):
            try:
                print(f"Processing page {page_num}...")
                tables = read_pdf(pdf_path, pages=str(page_num), area=area, multiple_tables=True)
                if not tables:
                    raise ValueError(f"No tables found on page {page_num}.")
                for table in tables:
                    process_table(table, page_num, writer)
            except Exception as e:
                print(f"Error processing page {page_num}: {e}")
else:
    specified_pages = pages.split(',')
    with pd.ExcelWriter(output_file_path_excel, engine='openpyxl') as writer:
        for page in specified_pages:
            try:
                print(f"Processing page {page}...")
                tables = read_pdf(pdf_path, pages=page.strip(), area=area, multiple_tables=True)
                if not tables:
                    raise ValueError(f"No tables found on page {page}.")
                for table in tables:
                    process_table(table, page, writer)
            except Exception as e:
                print(f"Error processing page {page}: {e}")

print(f"Final output saved to {output_file_path_excel}")
