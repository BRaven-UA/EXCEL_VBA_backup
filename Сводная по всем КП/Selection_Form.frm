VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Selection_Form 
   Caption         =   "Список КП в текущей папке"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   OleObjectBlob   =   "Selection_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Selection_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Button_OK_Click()   ' по нажатию кнопки проверяем выбрало ли хотябы одно КП
'    Dim found As Boolean
'
'    For item = 0 To Selection_Form.ListBox1.ListCount - 1
'        If Selection_Form.ListBox1.Selected(item) Then found = True
'    Next item
'
'    With Button_OK
'        If found Then
'            .Caption = "Готово"
'            .ForeColor = RGB(0, 0, 0)
'            .Font.Bold = False
'            Selection_Form.Hide
'        Else
'            .Caption = "Не выбрано элементов"
'            .ForeColor = RGB(255, 0, 0)
'            .Font.Bold = True
'        End If
'    End With
'
'End Sub

Private Sub Button_OK_Click()
    Selection_Form.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Main.break_execution = True
End Sub
