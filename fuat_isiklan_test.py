import os
import pandas as pd


# Function for checking nan values

def isNaN(num):
    return num != num

# Function for importing and reconstructing excel data in tabular form

def import_excel():
    # pd.set_option('display.max_rows', None, 'display.max_columns', None)
    working_directory = os.getcwd()
    file_name = '4- BOBİ FRS Nakit Akış Tablosu - Dolaylı Yöntem (Konsolide).xlsx'
    path = working_directory + '/data/' + file_name  # Specifying director
    Data_Base = pd.read_excel(path)   # Reading excel file

    aciklamalar_column = Data_Base.values.tolist()  # Forming list for all data

    new_aciklamalar_column = []

    # Extracting valuable data for aciklamalar column

    for i in range(len(aciklamalar_column)):
        for j in range(len(aciklamalar_column[i])):
            if not isNaN(aciklamalar_column[i][j]) and aciklamalar_column[i][j] != 0:
                new_aciklamalar_column.append(aciklamalar_column[i][j])

    # Forming new data frame/table with extracted valued data and zeros

    data = {'ACIKLAMALAR': new_aciklamalar_column[4:],
          'CARI_DONEM':[0 for i in range(69)],
          'ONCEKI_DONEM':[0 for i in range(69)]}

    return pd.DataFrame(data)

DataBase = import_excel()