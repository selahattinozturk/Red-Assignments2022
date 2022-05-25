Dim dosyaAdi As String
Dim tabloAdi As String

Sub Main
	
	dosyaAdi = "C:\Users\ufukc\Documents\My IDEA Documents\IDEA Projects\Red-Assignment\Source Files.ILB\4- BOBI FRS Nakit Akis Tablosu - Dolayli Yöntem (Konsolide).xlsx"
	
	tabloAdi = "BOBI FRS NAT Dolayli Konsolide"
	
	Call ExceldenImport()

End Sub

Function ExceldenImport()
	
	Dim dbAdi As String
	
	Dim task As task	
	
	dbAdi = dosyaAdi
	
	Set task = Client.GetImportTask("ImportExcel")
	
	task.FileToImport = dbAdi
	
	task.SheetToImport = tabloAdi
	
	task.OutputFilePrefix = "Ufuk Can Adanir"
	
	task.FirstRowIsFieldName = "TRUE"
	
	task.EmptyNumericFieldAsZero = "TRUE"
	
	task.PerformTask
	
	dbAdi = task.OutputFilePath("BOBI FRS NAT Dolayli Konsolide")
	
	Set task = Nothing
	
	Client.OpenDatabase(dbAdi)

	End Function

