import datetime as dt
import openpyxl as xl
import pandas as pd
from os import remove


def clean_sheet(wb: xl.Workbook, sheet_name: str)->(dt.datetime, str):
    """
    :param wb: OpenPyXL workbook
    :param sheet_name: String representation of the sheet
    :return: the name of the sheet, and the datetime object
    Takes an OpenPyXL workbook and sheet representation and cleans the data for processing
    Returns the datetime object and sheet name to be later processed in Pandas
    """

    date_cell = 'C5'  # cell that contains the date of the report
    worksheet = wb[sheet_name]  # worksheet to be cleaned

    date = worksheet[date_cell].value  # date of the report - should be a datetime.datetime object

    worksheet.unmerge_cells('C2:F2')  # un-merge the title cells
    worksheet.unmerge_cells('C3:F4')  # un-merge the title cells

    worksheet.delete_rows(15, 16)  # delete the in-between rows
    worksheet.delete_rows(1,13)  # delete the title rows

    worksheet.delete_cols(1)  # delete the bumper column

    return date, sheet_name


if __name__ == '__main__':

    # Load the DID spreadsheet from NHS Digital

    file = xl.load_workbook('data\\Tables-1a-1l-2017-18-Modality-Provider-Counts-XLSX-222KB (1).xlsx')

    # Initialise a list to contain the datetime/sheet_name tuples

    date_sheet_pairs = []

    # loop through the sheets (ignoring the title sheet) and clean the sheet using clean_sheet()

    for sheet in file.sheetnames[1:]:
        date_sheet_pairs.append(clean_sheet(file, sheet))

    # save the file as a temp file

    file.save('data\\temp_file.xlsx')

    # Close file

    file.close()

    # initialised the DataFrame list

    df_list = []

    # Loop through the date/sheet_name pairs, and create a DataFrame for each
    # Append each DataFrame to the list

    for date, sheet in date_sheet_pairs:
        temp_df = pd.read_excel('data\\temp_file.xlsx', sheet_name=sheet)
        # Use the date part from the tuple to add a new field 'Period' to the DataFrame
        temp_df['Period'] = date
        df_list.append(temp_df)

    # Combine all the dataframes

    df = pd.concat(df_list, axis=0)

    # Remove any missing numbers

    df.dropna(subset=['Provider Name'], inplace=True)

    # Fill any NA fields with 0

    df.fillna(0, inplace=True)

    # save the temporary DataFrame as an excel file

    df.to_excel('data\\df.xlsx')

    # delete the temporary excel file

    remove('data\\temp_file.xlsx')



