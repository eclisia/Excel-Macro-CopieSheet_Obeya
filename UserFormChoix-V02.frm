VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormChoix 
   Caption         =   "Paramétrage de la copie"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5625
   OleObjectBlob   =   "UserFormChoix-V02.frx":0000
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
    'By default,the text box has already beeing filled by the call of the function "RecuperationDateString"
    Name = "W" & TextBox_WeekNumber.Text
     
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




Private Function RecuperationDateString(ByVal ChoixSortie As Integer) As String
'   Function wich permits to get the actual system Date as String value
'   From the date, the function extracts the Year as YY
'   From the date, the function extracts the Week number as WW (from 01 to 52)
'   And the function proposes various String format output

    Dim myDate As Date
    Dim myVariante As Integer
    Dim tmp As String

    'Get the system DATE
    myDate = Date
    



Select Case ChoixSortie
    Case 0
        'Format YY
        myVariante = Right(DatePart("yyyy", myDate), 2)
        RecuperationDateString = myVariante
    Case 1
        'Format WW - Week Number
        
        'Get the Week Number from the current Date
        myVariante = DatePart("ww", myDate, vbUseSystemDayOfWeek, vbFirstFullWeek)
        
        'Format the date to have '0' for value under '10' --> 01, 02, 09, 10, ...52,
        If (myVariante < 10) Then
            RecuperationDateString = "0" & myVariante
        Else
            RecuperationDateString = myVariante
        End If
        
    Case 2
        'Format YYWW -
        
        'Get the Week Number from the current Date
        myVariante = DatePart("ww", myDate, vbUseSystemDayOfWeek, vbFirstFullWeek)
        
        'Format the date to have '0' for value under '10' --> 01, 02, 09, 10, ...52,
        If (myVariante < 10) Then
            tmp = "0" & myVariante
        Else
            tmp = myVariante
        End If
        
        'Format and concatenate the YYWW output
        RecuperationDateString = Right(DatePart("yyyy", myDate), 2) & tmp
    Case Else
        RecuperationDateString = myDate
        
End Select

Debug.Print "Format de sortie choisi pour la fonction : "; RecuperationDateString



End Function




Private Sub UserForm_Initialize()
    'Remplissage des champs par défaut
    
    Me.TextBox_WeekNumber.Value = RecuperationDateString(2)
    
End Sub
