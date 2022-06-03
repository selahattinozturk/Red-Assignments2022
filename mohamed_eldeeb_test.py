# This script requires having the excel file in the same working directory as the script
# The resulting database will be created in the Samples project path in your IDEA Managed Projects directory
# By default, the path is C:\Users\<user>\Documents\My IDEA Documents\IDEA Projects\Samples

import os
import pandas as pd

# Load the IDEA Client in python
import win32com.client
client = win32com.client.gencache.EnsureDispatch("Idea.IdeaClient")

# Constants
FILE_NAME = '4- BOBİ FRS Nakit Akış Tablosu - Dolaylı Yöntem (Konsolide).xlsx'
SHEET_NAME = 'BOBİ FRS NAT Dolaylı Konsolide'
DATABASE_NAME = 'mohamed_eldeeb_test.IMD'

FIELDS = ['ACIKLAMALAR', 'CARI_DONEM','ONCEKI_DONEM']

def Test_ImportExcel():
    # Create a new table to add the fields to
    table = client.NewTableDef()
    
    # Create a new field
    field0 = table.NewField()
    field0.Name = FIELDS[0]
    field0.Type = 3         # WI_CHAR_FIELD
    field0.Length = 130
    
    # Add the field to the table
    table.AppendField(field0)
    
    # Do the same for the other fields
    field1 = table.NewField()
    field1.Name = FIELDS[1]
    field1.Type = 4         # WI_NUM_FIELD
    table.AppendField(field1)
    
    field2 = table.NewField()
    field2.Name = FIELDS[2]
    field2.Type = 4         # WI_NUM_FIELD
    table.AppendField(field2)
    
    # Turn off protection to allow modifying the fields/records
    table.Protect = False
    
    db = client.NewDatabase(DATABASE_NAME, '', table)
    
    # Make a record set to add the records to
    rs = db.RecordSet()
    
    # Read the Excel file as a pandas DataFrame
    # Skip rows 1 and 2, read rows 4 to 72
    df = pd.read_excel(os.path.join(os.path.curdir, FILE_NAME), sheet_name=SHEET_NAME, skiprows=2)
    
    # Get the first valid element in each row (each row contains a string in only one of the columns)
    for d in map(lambda i: df.at[i, df.loc[i].first_valid_index()], range(len(df.index))):
        # Make a new record for each row
        rec = rs.NewRecord()
        
        # Set the values for each field
        rec.SetCharValue(FIELDS[0], d)
        rec.SetNumValue(FIELDS[1], 0)
        rec.SetNumValue(FIELDS[2], 0)
        
        # Append the record to the record set
        rs.AppendRecord(rec)
    
    # Turn protection back on
    table.Protect = True
    
    # Commit the changes to the database
    db.CommitDatabase()

if __name__ == '__main__':
    Test_ImportExcel()
