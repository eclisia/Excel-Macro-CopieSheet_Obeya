VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormChoix 
   Caption         =   "Paramétrage de la copie"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5625
   OleObjectBlob   =   "UserFormChoix-V01.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormChoix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    'Local variable
    Dim Name As String
    Dim TypeObeya As String
    Dim myNameWorksheet As String
    Dim myctrl As Control
    Dim myworkSheet As Worksheet
    

    
    
    'Popup a input dialobox to set the Week number. By default, text type is send back.
    Name = "W16" & TextBox_WeekNumber.Text
     
    'Get the caption of the OptionButton selected inside the frame (Frame which groupes the Optionbuttons)
    For Each myctrl In FrameChoix.Controls
        If myctrl.Object.Value = True Then
            TypeObeya = myctrl.Object.Caption
        End If
    Next myctrl

    
    ' Rename the new worksheet with the result of the Input dialog box.
    myNameWorksheet = Name & "-" & TypeObeya
    For Each myworkSheet In ThisWorkbook.Sheets
        If myworkSheet.Name = myNameWorksheet Then
            MsgBox "Il existe déjà un onglet avec ce nom", vbCritical, "Erreur de Nom"
            Exit Sub 'exit the macro
        Else
            
        End If
    Next myworkSheet
    
    'This call takes the last worksheet of the workbook (thanks to ThisWorkbook.Sheets.Count) and then pastes it at the end of the workbook
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    
    'This rename the sheet
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = myNameWorksheet
    
    
    
    ' Fill the cell "E8" with the right Week Number
    ActiveSheet.Cells(8, 5).Value = Name
    
    ' Modify the cell E4 & E5 color
    ActiveSheet.Cells(4, 5).Interior.ColorIndex = 15    'white
    ActiveSheet.Cells(5, 5).Interior.ColorIndex = 15    'white

    ' Erase the cell E4 & E5 value
    ActiveSheet.Cells(4, 5).Value = ""
    ActiveSheet.Cells(5, 5).Value = ""
    
    
    
    Unload Me
    'close the userform
End Sub



Private Sub TextBox_WeekNumber_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If InStr("1234567890", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep
    'Line to force only Number
End Sub
