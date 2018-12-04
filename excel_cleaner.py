import datetime as dt
import openpyxl as xl
import pandas as pd


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

    file = xl.load_workbook('data\\Tables-1a-1l-2017-18-Modality-Provider-Counts-XLSX-222KB (1).xlsx')

    date_sheet_pairs = []

    for sheet in file.sheetnames[1:]:
        date_sheet_pairs.append(clean_sheet(file, sheet))

    file.save('data\\temp_file.xlsx')
    file.close()

    df_list = []

    for date, sheet in date_sheet_pairs:
        temp_df = pd.read_excel('data\\temp_file.xlsx', sheet_name=sheet)
        temp_df['Period'] = date
        df_list.append(temp_df)

    df = pd.concat(df_list, axis=0)
    df.dropna(subset=['Provider Name'])
    df.to_excel('data\\df.xlsx')

