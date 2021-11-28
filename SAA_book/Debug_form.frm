VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Debug_form 
   Caption         =   "Отладочное окно"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   840
   ClientWidth     =   4815
   OleObjectBlob   =   "Debug_form.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "Debug_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub format_Click()
format.WordWrap = Not format.WordWrap
format.Height = IIf(format.WordWrap, 48, 12)
format.ZOrder (IIf(format.WordWrap, 0, 1))
format.BackStyle = IIf(format.WordWrap, 1, 0)
format.BorderStyle = IIf(format.WordWrap, 1, 0)
End Sub


'Private Sub OnTopBtn_Click() 'кнопка-триггер делает форму поверх всех окон
'Dim hwnd As Long
'Select Case OnTopBtn.Caption
'Case "Go Low"
'hwnd = SetWindowPos(DF_hwnd, -2, 0, 0, 0, 0, 3)
'If hwnd <> 0 Then OnTopBtn.Caption = "Go Top"
'Case "Go Top"
'hwnd = SetWindowPos(DF_hwnd, -1, 0, 0, 0, 0, 3)
'If hwnd <> 0 Then OnTopBtn.Caption = "Go Low"
'End Select
'End Sub




''делает неактивным форму при движении мыши
'Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'Dim hwnd As Long
''Dim Last_title As String
''Last_title = String(255, vbNullChar)
'hwnd = GetActiveWindow
'If hwnd <> DF_hwnd Then Last_hwnd = hwnd
''e = GetWindowText(Last_hwnd, Last_title, 255)
''SendMessageSTRING Last_hwnd, &HD, 255, Last_title
''Last_title = Left(Last_title, InStr(1, Last_title, vbNullChar) - 1)
''If Last_title <> "" And hwnd <> Last_hwnd Then AppActivate (Last_title)
'CommandButton2.Caption = DF_hwnd '& " - " & e
'CommandButton1.Caption = Last_hwnd '& " - " & Last_title
'If hwnd <> Last_hwnd Then SetActiveWindow Last_hwnd
'End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Hide
Cancel = True
End Sub
