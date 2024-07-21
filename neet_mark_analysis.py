import requests
import tabula
import pandas as pd
import os

base_url="https://neetfs.ntaonline.in/NEET_2024_Result/"

#variables user to adjust based on the state and center
start_center_count=270001
end_center_count=273105

#path where download files
output_file_path = os. getcwd()+"/neet_data/"

#mark cutoff , program will check how many students scored >= given mark.
mark_cutoff=685

#mark True if you want to print rows - for debugging.
print_marks_above_cutoff=False

#By default code will skip downloading PDF files if we found >10 files in given directory. Mark in True if code to redownload.
force_download=False

#derived variables
count=end_center_count-start_center_count
base_filename=start_center_count

if not os.path.exists(output_file_path):
    os.mkdir(output_file_path)

def pdf_to_excelv2(pdf_path):
    try:
        tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)

        if tables:
            combined_df = pd.concat(tables, ignore_index=True)

            combined_df.to_excel(excel_path, index=False)

            print(f"PDF has been successfully converted to Excel. The file is saved as {excel_path}")
        else:
            print("No tables were found in the PDF file.")
    except Exception as e:
        print(f"An error occurred: {e}")

def read_excel_and_check_marks(file_path, base_column_name):
    
    try:
        df = pd.read_excel(file_path)

        marks_columns = [col for col in df.columns if col.startswith(base_column_name)]

        if not marks_columns:
            print(f"No columns starting with '{base_column_name}' found in the Excel file.")
            return

        high_marks_df = pd.DataFrame()

        for col in marks_columns:
            high_marks_df = pd.concat([high_marks_df, df[df[col] >= mark_cutoff]])
        
        high_marks_count = len(high_marks_df)
        if print_marks_above_cutoff:
            print(high_marks_df)
        return high_marks_count

    except Exception as e:
        print(f"An error occurred: {e}")

loop_count=0
files = os.listdir(output_file_path)
xlsx_files = [file for file in files if file.endswith('.xlsx')]
print(f"Total centers downloaded {len(xlsx_files)}")

if len(xlsx_files) < 10 or force_download:
    print("Start downloading ... Set force_download to False to skip")
    while loop_count < count:
        loop_count=loop_count+1
        filename=base_filename+loop_count
        url=base_url+str(filename)+".pdf"
        print("Checking for "+ str(url))
        
        output_path=output_file_path+str(filename)+".pdf"
        excel_path=output_file_path+str(filename)+".xlsx"
        if os.path.exists(output_path):
                print(f"The file {output_path} already exists.")
                continue
        response = requests.get(url)
        if response.status_code == 200:
            with open(output_path, 'wb') as file:
                file.write(response.content)
            print(f"PDF downloaded successfully: {output_path}")

            pdf_to_excelv2(output_path)
        else:
            print(f"Failed to download PDF. Status code: {response.status_code}")
        # time.sleep(1)

loop_count=0
overall_higher_mark_count=0
while loop_count < count:
    loop_count=loop_count+1
    filename=base_filename+loop_count
    excel_path=output_file_path+str(filename)+".xlsx"
    
    if not os.path.exists(excel_path):
            continue
    print("Checking for center data " + excel_path)
    column_name = 'Marks'
    high_marks_count=read_excel_and_check_marks(excel_path, column_name)
    overall_higher_mark_count=overall_higher_mark_count+high_marks_count

xlsx_files = [file for file in files if file.endswith('.xlsx')]
print(f"Number of students got {overall_higher_mark_count} for center number from {start_center_count} to center no {end_center_count}, Total centers checked {len(xlsx_files)}")
