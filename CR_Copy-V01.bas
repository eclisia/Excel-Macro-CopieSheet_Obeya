Attribute VB_Name = "CR_Copy"
'***************************************
' Author : FTA, based on Formation VBA pdf file
' Summarize :   Macro which permits to duplicate the active worksheet
'               then set its name
' Evolution :   00 - Creation
'               01 - Remove of the classic Inputbox for the benefit of Userform
'                    All the code, is now contained inside the "Valide" Button
'***************************************



Sub CopyWorksheet()


    'New Code - Evolution 01
    'Display the UserFormChoix
    'Then, all the code is contains inside the "Valide" Button.
    UserFormChoix.Show

    
    'This call takes the last worksheet of the workbook (thanks to ThisWorkbook.Sheets.Count) and then pastes it at the end of the workbook
'    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
'
'
'
'    'Popup a input dialobox to set the Week number. By default, text type is send back.
'    'Name = Application.InputBox("Saisir le N° de semaine sous la forme WYYxx", "N° de semaine", "W16xx")
'    'TypeObeya = Application.InputBox("le type d'onglet de report", "Type-Obeya", "LPCB-B ou ObeyaClient")
    

    
'    Name = "W16" & UserFormChoix.TextBox_WeekNumber.Text
'
'    'Get the caption of the OptionButton selected inside the frame (Frame which groupes the Optionbuttons)
'    For Each myctrl In UserFormChoix.FrameChoix.Controls
'        If myctrl.Object.Value = True Then
'            TypeObeya = myctrl.Object.Caption
'        End If
'    Next myctrl
    
    
'    ' Rename the new worksheet with the result of the Input dialog box.
'    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = Name & "-" & TypeObeya
'
'
'    ' Fill the cell "E8" with the right Week Number
'    ActiveSheet.Cells(8, 5).Value = Name
'
'    ' Modify the cell E4 & E5 color
'    ActiveSheet.Cells(4, 5).Interior.ColorIndex = 15    'white
'    ActiveSheet.Cells(5, 5).Interior.ColorIndex = 15    'white
'
'    ' Erase the cell E4 & E5 value
'    ActiveSheet.Cells(4, 5).Value = ""
'    ActiveSheet.Cells(5, 5).Value = ""
    
End Sub

