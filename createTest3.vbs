Option Explicit	
'Set fso = CreateObject("Scripting.FileSystemObject")
Dim fileLocation, objWorkbook, objExcel, rowCount, i, foundTest, currentTest
currentTest = "Test8"
Set objExcel = CreateObject("Excel.Application")

fileLocation = "C:\Users\ajbel\Documents\Merged_Data_Sheet.xlsx"
	objExcel.displayalerts = False
	Set objWorkbook = objExcel.Workbooks.Open(fileLocation)

objExcel.Application.Visible = False
objWorkbook.WorkSheets(1).Activate
	'check how many rows there are
	rowCount = objWorkbook.WorkSheets(1).UsedRange.Rows.Count
		
	For i = 2 to rowCount
		If objWorkbook.WorkSheets(1).Cells(i, 1).Value = currentTest Then
			i = rowCount
			foundTest = 1
		Else
			foundTest = 0
		End If
	Next
		
	If foundTest = 0 Then
		objWorkbook.WorkSheets(1).Cells(rowCount+1, 1).Value = currentTest
		objWorkbook.WorkSheets(1).Cells(rowCount+1, 2).Value = 1		
	End If
	foundTest = 0		
		
	objExcel.ActiveWorkbook.Save
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit
	WScript.Quit
