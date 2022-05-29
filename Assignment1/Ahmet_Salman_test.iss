Begin Dialog NewDialog 50,50,150,150,"NewDialog", .NewDialog
End Dialog
Dim FilePath As String 'variable to hold the filename

Sub Main
	
Call createDatabase ()
Call insertData()

End Sub

Function createDatabase()
	' Create a table.
	Dim NewTable As Table
	Set NewTable = Client.NewTableDef
	
	' Define a field for the table.
	Dim AddedField As Field
	Set AddedField = NewTable.NewField
	AddedField.Name = "ACIKLAMALAR"
	AddedField.Type = WI_CHAR_FIELD
	AddedField.Length = 200
	
	' Add the field to the table.
	NewTable.AppendField AddedField
	
	' Perform the same steps for a second field.
	Set AddedField = NewTable.NewField
	AddedField.Name = "CARI_DONEM"
	AddedField.Type = WI_NUM_FIELD
	AddedField.Decimals = 0
	NewTable.AppendField AddedField
	
	Set AddedField = NewTable.NewField
	AddedField.Name = "ONCEKI_DONEM"
	AddedField.Type = WI_NUM_FIELD
	AddedField.Decimals = 0
	NewTable.AppendField AddedField
	
	' Change the table settings to allow writing.
	NewTable.Protect = False
	
	' Create the database.
	Dim db As Database
	Set db = Client.NewDatabase("Ahmet_Salman_test.IMD", "", NewTable)
	
	' Commit the database.
	db.CommitDatabase
	
	Set db = Nothing
	Set AddedField = Nothing
	Set NewTable = Nothing
End Function

Function insertData()

	Dim NewTable As Table
	Set NewTable = Client.NewTableDef
	
	' Create the database.
	Dim db As Database
	Set db = Client.OpenDatabase("Ahmet_Salman_test.IMD")
	
	' Obtain the recordset.
	Dim rs As RecordSet
	Set rs = db.RecordSet
	
	' Obtain a new record.
	Dim rec As Record
	Set rec = rs.NewRecord
	
	' Use the field name method to add data.
	rec.SetCharValue"ACIKLAMALAR", "new value"
	rec.SetCharValue "CARI_DONEM", 0
	rec.SetCharValue "ONCEKI_DONEM", 0
	rs.AppendRecord rec
	
	db.CommitDatabase
	db.close
	Client.OpenDatabase "Ahmet_Salman_test.IMD"
	
	Set db = Nothing
	Set AddedField = Nothing
	Set NewTable = Nothing


End Function