import pandas as pd
import numpy as np

# function to convert the excel file specified through the file path into a 2d array


def excel_to_2d_array(file_path, sheet_name=0):
    # Read the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    # Convert the DataFrame to a 2D array (list of lists)
    array_2d = df.values.tolist()
    array_2d = array_2d[2:]
    return array_2d

# returns the 2d array with all characters being lowercase


def array_2d_lowercase(array_2d):
    k = 0
    j = 0
    for row in array_2d:
        j = 0
        for element in row:
            array_2d[k][j] = str(element).lower()
            j = j+1
        if array_2d[k][0] == '' and array_2d[k][1] == '':
            array_2d.pop(k)  # removes blank rows
        k = k+1
    i = 1
    k = 0
    j = 0
    BoPass = True
    return array_2d

# returns the 2d array with all rows being filtered


def filter_array_2d(array_2d, job_title_keywords, company_name_keywords):
    i = 0
    array = [array_2d[0]]
    for row in array_2d[1:]:
        BoPass = False
        for element in job_title_keywords:
            if (element in str(row[5])):
                BoPass = True
        for element in company_name_keywords:
            if (element in str(row[4])) and BoPass:
                array.append(row)
        i = i + 1
    return array


if __name__ == "__main__":
    # !Specify the input and output paths of your Excel file!
    input_path = 'C:\\Users\\count\\Desktop\\Croton Capital\\Files\\Excel\\connections.xlsx'
    output_path = 'C:\\Users\\count\\Desktop\\Croton Capital\\Files\\Excel\\output.xlsx'
    job_title_keywords = []
    company_name_keywords = []
    # Specify job title and company keywords
    job_title_keywords = input(
        "Enter job title keywords separated by commas: ")
    company_name_keywords = input(
        "Enter company name keywords separated by commas: ")
    job_title_keywords = [job_title_keywords.strip()
                          for keyword in job_title_keywords.split(',')]
    company_name_keywords = [company_name_keywords.strip()
                             for keyword in company_name_keywords.split(',')]
    print(job_title_keywords)
    print(company_name_keywords)
    # Convert the Excel sheet to a 2D array
    array_2d = excel_to_2d_array(input_path)
    # Turns the 2d array lowercase
    array_2d = array_2d_lowercase(array_2d)
    # Filters the 2d array with the keywords
    array_2d_filtered = filter_array_2d(
        array_2d, job_title_keywords, company_name_keywords)
    # Creates and saves an output excel
    df = pd.DataFrame(array_2d_filtered, columns=array_2d_filtered[0])
    df.to_excel(output_path, index=False, sheet_name='Sheet1')
