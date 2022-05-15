Begin Dialog NewDialog 50,50,150,150,"NewDialog", .NewDialog
End Dialog
Dim sFilename As String 'variable to hold the filename

Sub Main
	
Call createDatabase ()

End Sub

Function createDatabase()
	sFilename = "D:\Bilkent Uni\MED_IDEA Internship\Red-Assignments2022\Assignment1/test.xlsx"
	Call importExcelFile()
	'Call ImportExcel("4- BOBI FRS Nakit Akis Tablosu - Dolayli Yöntem (Konsolide).xlsx", "BOBI FRS NAT Dolayli Konsolide")

End Function

Function importExcelFile() 'Import the Excel file
	Dim task As task	
	Dim dbName As String
	Set task = Client.GetImportTask("ImportExcel")
	dbName = sFilename
	task.FileToImport = dbName
	task.SheetToImport = "Sheet1"
	task.OutputFilePrefix = "Ahmet_Salman_test"
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask 
	'dbName = task.OutputFilePath("Sheet1")
	'Set task = Nothing
	'Client.OpenDatabase(dbName)

End Function
