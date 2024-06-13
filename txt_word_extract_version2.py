import pandas as pd

def text_to_excel(txt_file_path, xlsx_file_path):
    # Initialize lists to store names and summaries
    names = []
    summaries = []
    
    # Read the text file
    with open(txt_file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    
    # Initialize variables
    current_name = ""
    current_summary = ""
    
    # Parse the text file
    for line in lines:
        line = line.strip()
        if "&" in line:
            current_name = line.split("&")[1]
            names.append(current_name)
            summaries.append('')  # Add an empty string to maintain alignment
        if "%" in line:
            current_summary = line.split("%")[1]
            summaries[-1] = current_summary  # Update the latest summary
    
    # Create a DataFrame
    df = pd.DataFrame({
        '이름': names,
        '개요': summaries
    })
    
    # Save the DataFrame to an Excel file
    df.to_excel(xlsx_file_path, index=False)

    print("변환 작업이 성공적으로 완료되었습니다!")

# Define file paths
txt_file_path = 'C:/경로/텍스트파일.txt'  # 텍스트 파일 경로로 변경하세요
xlsx_file_path = 'C:/경로/스프레드시트.xlsx'  # 저장할 엑셀 파일 경로로 변경하세요

# Run the function
text_to_excel(txt_file_path, xlsx_file_path)
