Begin Dialog NewDialog 50,50,150,150,"NewDialog", .NewDialog
End Dialog
Dim FilePath As String 'variable to hold the filename

Sub Main
	
Call createDatabase ()

End Sub

Function createDatabase()
	FilePath = "D:\Bilkent Uni\MED_IDEA Internship\Red-Assignments2022\Assignment1\Ahmet_Salman_test.imd"
	Open FilePath For Output As FileNum
	Close FileNum
End Function