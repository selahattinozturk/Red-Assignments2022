Dim path As String
Dim sheet As String

Sub Main

path = "C:\Users\Ahmet Ay\Documents\My IDEA Documents\IDEA Projects\CI104 - IDEAScript for Analysts1_v10 - A4\Source Files.ILB\4- BOBI FRS Nakit Akis Tablosu - Dolayli Yontem (Konsolide).xlsx"
sheet = "BOBI FRS NAT Dolayli Konsolide"
               	Call ImportFrExcel()
	Call DirectExtraction()	'syntax deneme-BOBI FRS NAT Dolayli Konsolide.IMD
	Call DirectExtraction1()	'syntax deneme-BOBI FRS NAT Dolayli Konsolide.IMD
	Call DirectExtraction2()	'syntax deneme-BOBI FRS NAT Dolayli Konsolide.IMD
	Call ModifyField()	                'EXTRACTION1.IMD
	Call ModifyField1()	'EXTRACTION2.IMD
	Call ModifyField2()	'EXTRACTION3.IMD
	Call AppendDatabase()	'EXTRACTION3.IMD
	Call createDatabase()
	Call AppendDatabase1()	'Append Databases1.IMD

	


End Sub

Function ImportFrExcel
Set task = Client.GetImportTask("ImportExcel")
dbName = path
task.FileToImport = dbName
task.SheetToImport = sheet
task.OutputFilePrefix = "AhmetAy"
task.FirstRowIsFieldName = "TRUE"
task.EmptyNumericFieldAsZero = "TRUE"
task.PerformTask
dbName = task.OutputFilePath(sheet)
Set task = Nothing
Client.OpenDatabase(dbName)
End Function

' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("AhmetAy-BOBI FRS NAT Dolayli Konsolide.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "COL3"
	dbName = "EXTRACTION1.IMD"
	task.AddExtraction dbName, "", "@IsBlank(COL3)=0"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Data: Direct Extraction
Function DirectExtraction1
	Set db = Client.OpenDatabase("AhmetAy-BOBI FRS NAT Dolayli Konsolide.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "BOBI_FRS_NAKIT_AK_TABLOSU_DOLAYL_YÖNTEM_"
	dbName = "EXTRACTION2.IMD"
	task.AddExtraction dbName, "", "@IsBlank( BOBI_FRS_NAKIT_AK_TABLOSU_DOLAYL_YÖNTEM_ )=0"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Data: Direct Extraction
Function DirectExtraction2
	Set db = Client.OpenDatabase("AhmetAy-BOBI FRS NAT Dolayli Konsolide.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "COL5"
	dbName = "EXTRACTION3.IMD"
	task.AddExtraction dbName, "", "@IsBlank(COL5)=0"
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Modify Field
Function ModifyField
	Set db = Client.OpenDatabase("EXTRACTION1.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "ACIKLAMALAR"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 88
	task.ReplaceField "COL3", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Modify Field
Function ModifyField1
	Set db = Client.OpenDatabase("EXTRACTION2.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "ACIKLAMALAR"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 123
	task.ReplaceField "BOBI_FRS_NAKIT_AK_TABLOSU_DOLAYL_YÖNTEM_", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Modify Field
Function ModifyField2
	Set db = Client.OpenDatabase("EXTRACTION3.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "ACIKLAMALAR"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 110
	task.ReplaceField "COL5", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' File: Append Databases
Function AppendDatabase
	Set db = Client.OpenDatabase("EXTRACTION1.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "EXTRACTION2.IMD"
	task.AddDatabase "EXTRACTION3.IMD"
	dbName = "Append Databases1.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

Function createDatabase()
	' Create a table.
	Dim NewTable As Table
	Set NewTable = Client.NewTableDef
	
	' Define a field for the table.
	Dim AddedField As Field
	Set AddedField = NewTable.NewField
	AddedField.Name = "ACIKLAMALAR"
	AddedField.Type = WI_CHAR_FIELD
	AddedField.Length = 100
	
	' Add the field to the table.
	NewTable.AppendField AddedField
	
	' Perform the same steps for a second field.
	Set AddedField = NewTable.NewField
	AddedField.Name = "CARI_DONEM"
	AddedField.Type = WI_NUM_FIELD
	AddedField.Decimals = 2
	NewTable.AppendField AddedField
	
	
	Set AddedField = NewTable.NewField
	AddedField.Name = "ONCEKI_DONEM"
	AddedField.Type = WI_NUM_FIELD
	AddedField.Decimals = 2
	NewTable.AppendField AddedField
	
	' Change the table settings to allow writing.
	NewTable.Protect = False
	
	' Create the database.
	Dim db As Database
	Set db = Client.NewDatabase("SampleData.IMD", "", NewTable)
	
	' Obtain the recordset.
	Dim rs As RecordSet
	Set rs = db.RecordSet
	
	' Obtain a new record.
	Dim rec As Record
	Set rec = rs.NewRecord
	
	' Use the field name method to add data.
	rec.SetCharValue "ACIKLAMALAR"," "
	rec.SetCharValue "CARI_DONEM", 0
	rec.SetCharValue "ONCEKI_DONEM", 0

	rs.AppendRecord rec
	
	' Protect the table before you commit it.
	NewTable.Protect = True
	
	' Commit the database.
	db.CommitDatabase
	' Open the database.
	Client.OpenDatabase "SampleData.IMD"
	' Clear the memory.
	
	Set db = Nothing
	Set AddedField = Nothing
	Set NewTable = Nothing
End Function

Function AppendDatabase1
	Set db = Client.OpenDatabase("SampleData.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "Append Databases1.IMD"
	task.Criteria = "@IsBlank(ACIKLAMALAR)=0"
	dbName = "Ahmet_Ay_Test.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function



