import pandas as pd

def import_text_file_to_excel(file_path, save_path):
    # Initialize lists to store names and overviews
    names = []
    overviews = []
    
    # Read the text file
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    
    # Initialize variables
    name_column = ""
    overview_column = ""
    name_start = False
    overview_start = False
    
    # Parse the text file
    for line in lines:
        line = line.strip()
        if line.startswith("%"):
            name_column = line[1:].strip()
            name_start = True
            overview_start = False
        elif line.startswith("&"):
            overview_column = line[1:].strip()
            overview_start = True
            name_start = False
        elif name_start:
            names.append(name_column)
            overviews.append('')
            name_start = False
        elif overview_start:
            overviews[-1] = overview_column
            overview_start = False
    
    # Create a DataFrame
    df = pd.DataFrame({
        '이름': names,
        '개요': overviews
    })
    
    # Save the DataFrame to an Excel file
    df.to_excel(save_path, index=False)

# Define file paths
file_path = 'C:/Path/To/Your/File.txt'
save_path = 'C:/Path/To/Save/Your/ExcelFile.xlsx'

# Run the function
import_text_file_to_excel(file_path, save_path)
