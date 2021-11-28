VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppWithEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'новый класс для отлавливания событий в других книгах
Public WithEvents App_Ev As Application
Attribute App_Ev.VB_VarHelpID = -1
Public Last_Page As String

Private Sub App_Ev_SheetDeactivate(ByVal Sh As Object)  'сохранение названия последнего активного листа
    Last_Page = Sh.Name
End Sub

Private Sub App_Ev_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
If Not Debug_form.Visible Then Exit Sub
On Error Resume Next
Debug_form.Label1 = Target.Address(False, False, xlA1)
If Target.Rows.Count > 1 Then Debug_form.Label2 = Target.Row & ":" & Target.Rows.Count + Target.Row - 1 Else Debug_form.Label2 = Target.Row
If Target.Columns.Count > 1 Then Debug_form.Label3 = Target.Column & ":" & Target.Columns.Count + Target.Column - 1 Else Debug_form.Label3 = Target.Column
Debug_form.format = "Format: " & Target.NumberFormat
Debug_form.Formula = "Formula: " & Target.Formula
End Sub



'Private Sub App_Ev_WorkbookOpen(ByVal Wb As Workbook)
'On Error Resume Next
'ThisWorkbook.VBProject.References.AddFromFile (Wb.FullName)
'
'End Sub

'Private Sub App_Ev_WorkbookOpen(ByVal Wb As Workbook)
'
'For Each компонент In Wb.VBProject.VBComponents
'    For i = 1 To компонент.CodeModule.CountOfLines
'        код = компонент.CodeModule.Lines(i, 1)
'        If код = "         ar_s1_(1 To 85) As my_s1" Then компонент.CodeModule.ReplaceLine i, "         ar_s1_(1 To 200) As my_s1"
'        If код = "KOL_GROUP = 85" Then компонент.CodeModule.ReplaceLine i, "KOL_GROUP = 200"
'        'Else компонент.CodeModule.ReplaceLine i, Replace(код, ".Show ?", ".Show vbModeless", , , vbTextCompare)
'    Next i
'Next компонент
'
'
'End Sub


'Private Sub App_Ev_WorkbookOpen(ByVal Wb As Workbook)
'
'For Each компонент In Wb.VBProject.VBComponents
'    For i = 1 To компонент.CodeModule.CountOfLines
'        код = компонент.CodeModule.Lines(i, 1)
'        If InStr(код, Macro_directory) Then компонент.CodeModule.ReplaceLine i, Replace(код, "C:\" & Macro_directory, RAM_drive & "\" & Macro_directory)
'    Next i
'Next компонент
'
''Application.EnableEvents = True
'
'End Sub
''

'Private Sub App_Ev_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
'For Each ref In ThisWorkbook.VBProject.References
'If Right(ref.FullPath, Len(Wb.Name)) = Wb.Name Then ThisWorkbook.VBProject.References.Remove ref
'Next
'
'End Sub

