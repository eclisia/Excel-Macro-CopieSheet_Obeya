Attribute VB_Name = "CR_Copy"
'***************************************
' Author : FTA, based on Formation VBA pdf file
' Summarize :   Macro which permits to duplicate the active worksheet
'               then set its name
'***************************************

Sub CopyWorksheet()
    
    'This call takes the last worksheet of the workbook (thanks to ThisWorkbook.Sheets.Count) and then pastes it at the end of the workbook
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    
    Dim Name As String
    'Popup a input dialobox to set the Week number. By default, text type is send back.
    Name = Application.InputBox("Saisir le N° de semaine sous la forme WYYxx", "N° de semaine", "W16xx")
    ' Rename the new worksheet with the result of the Input dialog box.
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = Name
    
    ' Fill the cell "E7" with the right Week Number
    ActiveSheet.Cells(7, 5).Value = Name

    
End Sub

