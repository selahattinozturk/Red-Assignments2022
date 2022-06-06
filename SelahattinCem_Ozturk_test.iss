Begin Dialog NewDialog 50,50,150,150,"NewDialog", .NewDialog
End Dialog
Dim FilePath As String 'variable to hold the filename
Sub Main
	
	Call Create()
	Call Insert()
End Sub

Function Create()
	' Creating the table and defining the fields.
	Dim NewTable As Table
	Set NewTable = Client.NewTableDef
	
	
	Dim AddedField As Field
	Set AddedField = NewTable.NewField
	AddedField.Name = "ACIKLAMALAR"
	AddedField.Type = WI_CHAR_FIELD
	AddedField.Length = 200
	NewTable.AppendField AddedField

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
	
	' Making the created table writable
	NewTable.Protect = False
	
	Dim db As Database
	Dim Path As String
	Set db = Client.NewDatabase("SelahattinCemOzturk_test.IMD", "", NewTable)
	db.CommitDatabase
	
	Set db = Nothing
	Set AddedField = Nothing
	Set NewTable = Nothing
End Function

Function Insert()
	Dim Path As String
	Path = "C:\Users\loose\Documents\My IDEA Documents\IDEA Projects\RedAssignment\"
	Dim dbName As String
	dbName = Client.LocateInputFile (Path + "data.xlsx")
	Dim task As ImportExcel
	Set task = Client.GetImportTask("ImportExcel")
	task.FileToImport = dbName
	task.SheetToImport = "Sheet1"
	task.OutputFilePrefix = "Excel"
	task.FirstRowIsFieldName = "False"
	task.EmptyNumericFieldAsZero = "False"
	task.PerformTask
	Path = "C:\Users\loose\Documents\My IDEA Documents\IDEA Projects\RedAssignment\"
	
	
	Set task = Nothing
	
	Dim db1 As Database
	
	Set db1 = Client.OpenDatabase( "C:\Users\loose\Documents\My IDEA Documents\IDEA Projects\RedAssignment\Excel-Sheet1.IMD")
	
		
	
	Dim RS1 As RecordSet
	Set RS1 = db1.RecordSet
	' Move to the first record.
	RS1.ToFirst
	
	
	Dim currentText As String
	Dim col2Val As String
	Dim col3Val As String
	Dim col4Val As String
	
	Dim NewTable As Table
	Set NewTable = Client.NewTableDef
	
	
	Dim db As Database
	Set db = Client.OpenDatabase("SelahattinCemOzturk_test.IMD")
	
	' Obtain the recordset.
	Dim rs As RecordSet
	Set rs = db.RecordSet
	
	Dim rec As Record
	Set rec = rs.NewRecord
	RS1.Next
	RS1.Next
	RS1.Next
	RS1.Next
	For Count = 0 To 67
		col3Val = RS1.ActiveRecord.GetCharValue("COL3")
		col4Val = RS1.ActiveRecord.GetCharValue("COL4")
		col5Val = RS1.ActiveRecord.GetCharValue("COL5")
		If col3Val <> "" Then
			rec.SetCharValue"ACIKLAMALAR", col3Val
		ElseIf col4Val <> "" Then
			rec.SetCharValue"ACIKLAMALAR", col4Val
		ElseIf col5Val <> "" Then
			rec.SetCharValue"ACIKLAMALAR", col5Val
		End If
		rec.SetCharValue "CARI_DONEM", 0
		rec.SetCharValue "ONCEKI_DONEM", 0
		rs.AppendRecord rec
		RS1.Next
	Next
	db.CommitDatabase
	db.close
	Client.OpenDatabase "SelahattinCemOzturk_test.IMD"
	
	Set db = Nothing
	Set AddedField = Nothing
	Set NewTable = Nothing

End Function


