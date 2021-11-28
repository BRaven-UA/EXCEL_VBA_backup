Attribute VB_Name = "saa_module"
Public App As New AppWithEvents     'подключение нового класса
'Public Declare Function GetActiveWindow Lib "user32" () As Long
'Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
'''Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Public Declare Function SendMessageSTRING Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
'Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'''Public Declare Function Getasynckeystate Lib "user32" Alias "GetAsyncKeyState" (ByVal VKEY As Long) As Integer
'Public DF_hwnd As Long, Last_hwnd As Long
Public RAM_drive As String
Public Const Macro_directory = "BAZI DLYA MACROSOV SUPER NEW 21.11.16"

Public Sub test() 'точка входа в тестируемый код
    r = Selection.Row
    Selection.EntireRow.Cut
    Rows(r - 2).Insert Shift:=xlDown
    Debug.Print r
End Sub

Public Sub Номер_узла()
    Selection.Formula = "=MAX(OFFSET(INDIRECT(""A1""),,,ROW()-1,1))+1"
End Sub

Public Sub продолжение_нумерации()  ' для стандарной нумерации вида "1,1,1"
On Error GoTo ExitSub
Dim Номер_узла, адрес_узла, посл_номер, маска_номера, позиция As String, c As Range, кол As Integer
Номер_узла = Application.WorksheetFunction.Max(Range(Cells(5, Selection.Column), Cells(Selection.Row - 1, Selection.Column)), 0)
адрес_узла = Range(Cells(5, Selection.Column), Cells(Selection.Row - 1, Selection.Column)).Find(what:=Номер_узла, LookAt:=xlWhole, LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Address(RowAbsolute:=False, ColumnAbsolute:=False)
посл_номер = Range(Cells(5, Selection.Column), Cells(Selection.Row - 1, Selection.Column)).Find(what:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Text
посл_номер = Right(посл_номер, Len(посл_номер) - Application.Max(InStrRev(посл_номер, "-"), InStrRev(посл_номер, " ")))
If StrComp(посл_номер, Номер_узла, vbTextCompare) = 0 Then посл_номер = посл_номер & ",1,0" 'для первой позиции внутри узла
посл_номер = Split(посл_номер, ",")
If Getasynckeystate(vbKeyControl) < 0 Then посл_номер(1) = посл_номер(1) + 1: посл_номер(2) = 0  ' если нажата клавиша Ctrl начинаем новую нумерацию
маска_номера = адрес_узла & " & " & Chr(34) & "," & посл_номер(1) & ","
For Each c In Selection.Cells
    c.ClearContents
    кол = Val(Cells(c.Row, 5).Text)
    If кол > 0 Then
        посл_номер(2) = посл_номер(2) + 1
        c = "=" & маска_номера & посл_номер(2) & Chr(34)
        If кол > 1 Then
            посл_номер(2) = посл_номер(2) + кол - 1
            c = c.Formula & " & " & Chr(34) & "-" & Chr(34) & " & " & маска_номера & посл_номер(2) & Chr(34)
            c.Rows.AutoFit      'корректируем высоту строки если номер получился длинным
        End If
    End If
Next c
ExitSub:
End Sub

Public Sub рус_символы()
Dim a As Range
Application.ScreenUpdating = False
'l = Cells.Find(what:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
'For Each a In Range(Cells(1, 1), Cells(l, 8))
For Each a In Selection
    For i = 192 To 255
        pos = 0
        Do
            pos = InStr(pos + 1, a, Chr(i))
            If pos > 0 Then a.Characters(pos, 1).Font.ColorIndex = 3
        Loop Until pos < 1
    Next i
'    Call Progress_Bar(Round(a.Row * 100 / l))
    If a.Formula <> a.Text Then a.Font.ColorIndex = 43
Next
Application.ScreenUpdating = True
'Application.StatusBar = False
End Sub

Public Sub Progress_Bar(ByVal percent As Long)
    Dim progress As String
    progress = String(100 - percent, Chr(1))
    Application.StatusBar = percent & " % " & progress
End Sub

Public Sub shiftgotoend()           'обработка Shift+Ctrl+End
gotoend (True)
End Sub

Public Sub gotoend(Optional ByVal Sh As Boolean = False)    ' обработка Ctrl+End
On Error Resume Next
c = Application.Selection.Column
r2 = Cells.Find(what:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
r1 = IIf(Sh, Selection.Row, r2)
Range(Cells(r1, c), Cells(r2, c)).Select
End Sub

Public Sub hideemptyrows()          'прячет все пустые строки
Attribute hideemptyrows.VB_ProcData.VB_Invoke_Func = "q\n14"
Application.ScreenUpdating = False
For i = 5 To Cells.Find(what:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
     If Not Rows(i).Hidden Then Rows(i).AutoFit     'автоподбор высоты строки
     If Cells(i, 5).Value = 0 Then If Cells(i, 11) > 0 Then Rows(i).Hidden = True 'прячет оборудование если количество = 0
     If Rows(i).Text = "" Then Rows(i).Hidden = True
Next i
Application.ScreenUpdating = True
End Sub

Public Sub TABpages()               'обработка Ctrl+TAB
On Error Resume Next
If App.Last_Page <> "" Then ActiveWorkbook.Sheets(App.Last_Page).Activate
End Sub

Public Sub Debug_window()
If Debug_form.Visible Then Debug_form.Hide Else Debug_form.Show vbModeless
'DF_hwnd = GetActiveWindow
AppActivate (Application.Caption)
End Sub

Public Function F(ByRef cell As Range) As String
    Application.Volatile
    F = cell.Formula
End Function

Sub долгий_расчет(имя_процедуры As String)
'If Getasynckeystate(&H42) <> 0 Then имя_процедуры = "avto_ochenka_start" 'клавиша B
'If Getasynckeystate(&H57) <> 0 Then имя_процедуры = "avto_ves" 'клавиша W
'If Getasynckeystate(&H4F) <> 0 Then имя_процедуры = "avto_koef_opis" 'клавиша O
'If Getasynckeystate(&H52) <> 0 Then имя_процедуры = "PRESS_ctrl_r" 'клавиша R
'If Getasynckeystate(&H50) <> 0 Then имя_процедуры = "avto_mownost" 'клавиша P
'Application.EnableEvents = False
Application.ScreenUpdating = False
'calc_temp = Application.Calculation
'Application.Calculation = xlCalculationManual
'Application.Visible = False
On Error Resume Next
Application.Run "'" & ActiveWorkbook.Name & "'!" & имя_процедуры
Application.EnableEvents = True
Application.ScreenUpdating = True
'Application.Calculation = calc_temp
Application.Visible = True
End Sub


Sub количество_графобъектов_в_книге()
For Each ch In ActiveSheet.ChartObjects
Debug.Print ch.Name
Next ch
End Sub

Sub Разбивка_имп_рос()
    Application.ScreenUpdating = False
    ActiveWorkbook.Worksheets("спецификация").Copy After:=Worksheets("спецификация")
    ActiveSheet.Name = "спецификация имп"
    ActiveSheet.Tab.ColorIndex = 22
    ActiveWorkbook.Worksheets("спецификация").Cells.Copy
    ActiveWorkbook.Worksheets("спецификация имп").Cells.Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats

    Cells(2, 2).Value = Left(Cells(2, 2).Value, Len(Cells(2, 2).Value) - 1) & " (импортное оборудование)"
    If ActiveWorkbook.Worksheets("ЭЛЕВАТОР").Cells(7, 3) = "Россия" Then c = 15 Else c = 14
    
    'удаление неимпортного оборудования. Обратный порядок перебора чтобы удаление строк не сбивало переменную цикла
    For i = Cells.Find(what:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row - 1 To 7 Step -1
        If Cells(i, 5) <> 0 And (Cells(i, c) + Cells(i, c + 1) > 0) Then Rows(i).Delete
        If Cells(i, 5) = 0 And Cells(i, 11) <> 0 Then Rows(i).Delete
    Next i
    
    'наводим красоту
    узел = 0: пустая_строка = 0: оборудование = 0
    For i = 7 To Cells.Find(what:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row - 1
        If Cells(i, 9).Interior.ColorIndex = 39 Then
            If оборудование < узел And узел > 0 Then Range(Rows(узел), Rows(i - 1)).EntireRow.Hidden = True
            узел = i
        End If
        If Rows(i).Text = "" Then
            If пустая_строка = i - 1 Then Rows(i - 1).Hidden = True
            пустая_строка = i
        End If
        If Cells(i, 5) > 0 Then оборудование = i
    Next i
    If оборудование < узел And узел > 0 Then Rows(узел).Hidden = True
    
        Application.ScreenUpdating = True
End Sub


Sub delete_all() ' запускать только вручную! ОПАСНО
Application.ScreenUpdating = False
For i = Cells.Find(what:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row To 2 Step -1
     'Rows(i).AutoFit     'автоподбор высоты строки
     If Cells(i, 5).Value = "" Then Rows(i).Delete
     If Rows(i).Text = "" Then Rows(i).Delete
Next i
Application.ScreenUpdating = True
End Sub

Sub Summ_to_Clipboard() ' помещает в буфер обмена сумму выделенных ячеек
    On Error Resume Next
    Dim DataObj As New MSForms.DataObject
    Dim c As Range
    For Each c In Selection.Cells
        Sum = Sum + c.Value
    Next c
    DataObj.SetText Sum
    DataObj.PutInClipboard
    Set DataObj = Nothing
End Sub

Sub Formula_to_Clipboard()  ' помещает в буфер обмена формулы выделенного диапазона ячеек
    On Error Resume Next
    Dim DataObj As New MSForms.DataObject
    For Row = Selection.Row To Selection.Row + Selection.Rows.Count - 1
        For Col = Selection.Column To Selection.Column + Selection.Columns.Count - 1
            Arr = Arr & Cells(Row, Col).FormulaLocal & Chr(9)
        Next Col
        Arr = Left(Arr, Len(Arr) - 1) & Chr(10)
    Next Row
    Arr = Left(Arr, Len(Arr) - 1)
    DataObj.SetText Arr
    DataObj.PutInClipboard
    Set DataObj = Nothing
End Sub

Sub Формирование_приложения()
    ' вспешке откорректированный записанный макрос
    If MsgBox("Сформировать спецификацию для договора ?", vbOKCancel) = vbCancel Then Exit Sub
    OriginSheet = ActiveSheet.Name
    ActiveSheet.Copy After:=ActiveSheet
    ActiveSheet.Name = "приложение"
    ActiveSheet.Tab.ColorIndex = 28
    Sheets(OriginSheet).Cells.Copy
    ActiveSheet.Cells.Select
    ActiveSheet.Paste
    Range("G3:G4").FormulaR1C1 = "Цена ед. оборудо вания с доставкой без НДС"
    Range("H3:H4").FormulaR1C1 = "Стоимость с доставкой без НДС"
    LastRow = Cells.Find(what:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    Range(Cells(6, 7), Cells(LastRow - 1, 7)).FormulaR1C1 = "=ROUND(" & OriginSheet & "!RC*_k_EXW_DDP_RU/1.2,2)"
    Range(Cells(6, 8), Cells(LastRow - 1, 8)).FormulaR1C1 = "=IF(ISNUMBER(RC[-1]),RC[-1]*RC[-3], """")"
    Application.CutCopyMode = False
    Range("H" & LastRow + 1).FormulaR1C1 = "=R[-1]C*0.2"
    Range("H" & LastRow + 2).FormulaR1C1 = "=R[-2]C+R[-1]C"
    Range(Cells(6, 7), Cells(LastRow + 2, 8)).Select
    Selection.NumberFormat = "#,##0.00$"
    Range(Cells(LastRow, 1), Cells(LastRow, 7)).Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Итого без НДС"
    Selection.Copy
    Range("A" & LastRow + 1).Select
    ActiveSheet.Paste
    Range("A" & LastRow + 2).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A" & LastRow + 1).FormulaR1C1 = "НДС"
    Range("A" & LastRow + 2).FormulaR1C1 = "Итого с НДС"
    Range(Cells(LastRow, 1), Cells(LastRow + 2, 8)).Select
    With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .Font.Bold = True
    End With
'    If MsgBox("Сформировать спецификацию для договора ?", vbOKCancel) = vbCancel Then Exit Sub
'    OriginSheet = ActiveSheet.Name
'    ActiveSheet.Copy After:=ActiveSheet
'    ActiveSheet.Name = "приложение"
'    ActiveSheet.Tab.ColorIndex = 28
'    Sheets(OriginSheet).Cells.Copy
'    ActiveSheet.Cells.Select
'    ActiveSheet.Paste
'    Range("G3:G4").FormulaR1C1 = "Цена ед. оборудо вания с доставкой без НДС"
'    Range("H3:H4").FormulaR1C1 = "Стоимость с доставкой без НДС"
'    LastRow = Cells.Find(what:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
'    Range("G6").Select
'    Selection.FormulaR1C1 = "=ROUND(" & OriginSheet & "!RC*_k_EXW_DDP_RU/1.18,2)"
'    Selection.Copy
'    Application.Run "shiftgotoend"
'    ActiveSheet.Paste
'    Range("H6").Select
'    Selection.FormulaR1C1 = "=IF(ISNUMBER(RC[-1]),RC[-1]*RC[-3], """")"
'    Selection.Copy
'    Range(Cells(LastRow, 1), Cells(LastRow + 2, 8)).Select
'    ActiveSheet.Paste
'    Application.CutCopyMode = False
'    Range("H" & LastRow + 1).FormulaR1C1 = "=R[-1]C*0.18"
'    Range("H" & LastRow + 2).FormulaR1C1 = "=R[-2]C+R[-1]C"
'    Range(Cells(LastRow, 1), Cells(LastRow, 7)).Select
'    Selection.Merge
'    ActiveCell.FormulaR1C1 = "Итого без НДС"
'    Selection.Copy
'    Range("A" & LastRow + 1).Select
'    ActiveSheet.Paste
'    Range("A" & LastRow + 2).Select
'    ActiveSheet.Paste
'    Application.CutCopyMode = False
'    Range("A" & LastRow + 1).FormulaR1C1 = "НДС"
'    Range("A" & LastRow + 2).FormulaR1C1 = "Итого с НДС"
'    Range(Cells(LastRow, 1), Cells(LastRow + 2, 8)).Select
'    With Selection.Borders
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With Selection
'        .HorizontalAlignment = xlRight
'        .VerticalAlignment = xlTop
'        .Font.Bold = True
'    End With
'    Range(Cells(6, 7), Cells(LastRow + 2, 8)).Select
'    Selection.NumberFormat = "#,##0.00$"
End Sub

Sub доллары()
    Selection.NumberFormat = "[$$-409]#,##0"
End Sub
Sub рубли()
    Selection.NumberFormat = "#,##0$"
End Sub
Sub евро()
    Selection.NumberFormat = "[$€-2] #,##0"
End Sub

Sub Alt_1()
    Selection.EntireRow.Delete
End Sub

Sub Alt_2()
    Selection.EntireRow.Hidden = True
End Sub

Sub Alt_3()
    Selection.EntireRow.Hidden = False
End Sub

'CustomerName = Replace(Sheets("спецификация").Range("B2").Text, "Спецификация.", "", 1, 1, vbTextCompare)
'CustomerName = LTrim(Left(CustomerName, InStr(1, CustomerName, ".", vbTextCompare) - 1))
'Filename = ThisWorkbook.Path & "\" & Replace(Replace(Replace(ThisWorkbook.Name, CustomerName, "(" & CustomerName & ")"), "КП ", "Комм предл "), ".xls", ".doc")
'If Dir(Filename) > "" Then Filename = Replace(Filename, ".doc", " +.doc")
'wobj.Application.ActiveDocument.SaveAs Filename:=Filename

Sub RAM_files_update()
    
    On Error Resume Next    'игнорируем ошибку "диск не готов", т.к. Drive.IsReady работает некорректно
    Set FSO = CreateObject("Scripting.FileSystemObject")
    For Each Drive In FSO.Drives                                        'Определяем наличие RAM-диска в системе через имя диска, т.к.
        If InStr(Drive.VolumeName, "RAM") Then RAM_drive = Drive.Path   ' DriveType определяет RAM-диск как обычный стационарный
    Next
    On Error GoTo 0 'снова отслеживаем ошибки
    If RAM_drive = "" Then
        Result = MsgBox("Работа будет продолжена с использованием файлов на диске С", vbOKOnly, "RAM-диск не найден")
        Exit Sub
    End If
    If Not FSO.FolderExists(RAM_drive & "\" & Macro_directory) Then FSO.CreateFolder (RAM_drive & "\" & Macro_directory)    'проверяем наличие рабочей папки на RAM-диске
    For Each File In FSO.GetFolder("C:\" & Macro_directory).Files   'перебираем все файлы в рабочей папке на диске С
        If (Now - File.DateLastModified) < 730 Then 'работаем только с файлами, измененными за последние 2 года
            td = 0  'флаг отсутствия копии файла
            t = Replace(File.Path, "C:", RAM_drive) 'имя копии файла на RAM-диске
            If FSO.FileExists(t) Then   'если такой файл уже есть на RAM-диске, то получаем ссылку на файл
                Set tf = FSO.GetFile(t)
                td = tf.DateLastModified
            End If
            If td < File.DateLastModified Then File.Copy RAM_drive & "\" & Macro_directory & "\", True  'создаем или обновляем файл
        End If
    Next
    Set FSO = Nothing

End Sub

Sub Shift_selected(direction As String) ' смещает выделенный диапазон ячеек на одну строку/столбец в выбранном направлении
    r = Selection.Row
    h = Selection.Rows.Count
    c = Selection.Column
    w = Selection.Columns.Count
    Dim dx, dy
    
    Select Case direction
    Case "up"
        If r > 1 Then
            Selection.EntireRow.Cut
            Rows(r - 1).Insert Shift:=xlDown
            dy = -1
        End If
    Case "down"
        Selection.EntireRow.Cut
        Rows(r + h + 1).Insert Shift:=xlDown
        dy = 1
    Case "left"
        If c > 1 Then
            Selection.EntireColumn.Cut
            Columns(c - 1).Insert Shift:=xlRight
            dx = -1
        End If
    Case "right"
        Selection.EntireColumn.Cut
        Columns(c + w + 1).Insert Shift:=xlRight
        dx = 1
    End Select
    
    Range(Cells(r + dy, c + dx), Cells(r + dy + h - 1, c + dx + w - 1)).Select  ' перевыделяем смещенный диапазон
        
End Sub

Sub Duplicate_rows()    ' Дублирует строки из выделенного диапазона
    Selection.EntireRow.Copy
    Rows(Selection.Row + Selection.Rows.Count).Insert Shift:=xlDown
    Application.CutCopyMode = False ' убираем рамку скопированного диапазона
    Selection.Offset(Selection.Rows.Count, 0).Select    ' выделяем смещенный диапазон
End Sub
