Dim DosyaAdi As String
DosyaAdi = "C:\Users\ufukc\Downloads\4- BOBI FRS Nakit Akis Tablosu - Dolayli Yöntem (Konsolide).xlsx"
Dim TabloAdi As String
TabloAdi = "BOBI FRS NAT Dolayli Konsolide"

Sub Main
Call ExceldenImport()
End Sub

Function ExceldenImport()

	Dim dbAdi As String

	Dim task As Task
	
	dbAdi = DosyaAdi
	
	Set task = Client.GetImportTask("ImportExcel")
	
	task.FileToImport = dbAdi
	
	task.SheetToImport = TabloAdi
	
	task.OutputFilePrefix = "BOBI FRS Nakit Akis Tablosu"
	
	task.FirstRowIsFieldName = "True"
	
	task.EmptyNumericFieldAsZero = "True"
	
	task.PerformTask

	dbAdi = task.OutputFilePath("TableSheet")

	Set task = Nothing
	
	Client.OpenDatabase(dbAdi)
End Function
