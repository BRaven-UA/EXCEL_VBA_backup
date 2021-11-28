Attribute VB_Name = "Макросы"
Public Const remote_folder_path As String = "\\chief\Коммерция\ЗАЯВКИ НА КП\"
Public Const remote_filename As String = "запросы менеджеров.xlsm"
Public Const extension As String = ".xls"
' перечисление директорий, в которых производится поиск файлов КП находится в модуле FileSearch, функция FileSearch
Public Const first_row As Integer = 2   ' Первая строка, с которой начинаются заявки
Public Const last_column As Integer = 13    ' Последний столбец, подлежащий копированию



Sub Fill_range(manager)

    ' копируем данные из файла менеджера
    reference = "'" & remote_folder_path & manager & "\[" & remote_filename & "]Лист1'!A" & CStr(first_row)
    With Sheets(manager).range(Sheets(manager).Cells(first_row, 1), Sheets(manager).Cells(Sheets(manager).UsedRange.Rows.Count, last_column))
     .FormulaLocal = "=ЕСЛИ(" & reference & "=""""; """"; " & reference & ")"   ' избавляемся от нулей в пустых ячейках
     .Value = .Value2
    End With

' ---- проверка готового КП по имени файла и заполнение итоговой суммы КП ----

'    updated = False ' признак что есть данные, которые нужно внести в файл менеджера
'    Const filename_column = 11
'
'    ' проверяем есть ли имя файла готового КП и заполняем итоговую сумму КП по импортному и российскому оборудованию
'    With ThisWorkbook.Sheets(manager)
'        For i = first_row To .UsedRange.Rows.Count
'            filename = .Cells(i, filename_column).Text
'            If filename <> "" And (.Cells(i, filename_column + 1).Text & .Cells(i, filename_column + 2).Text) = "" Then
'                filename = filename + extension
'                full_path = Replace(FileSearch.FileSearch(filename), filename, "")
'                reference = "'" & full_path & "[" & filename & "]" & "ДОСТАВКА" & "'!"
'                .Cells(i, filename_column + 1).Formula = "=" & reference & "F36"    ' в долларах
'                .Cells(i, filename_column + 1).Value = .Cells(i, filename_column + 1).Value2
'                .Cells(i, filename_column + 2).Formula = "=" & reference & "G36*" & reference & "I2"    ' сразу переводим в рубли
'                .Cells(i, filename_column + 2).Value = .Cells(i, filename_column + 2).Value2
'                updated = True
'            End If
'        Next i
'    End With
'
'    ' если были изменения в данных по сумме КП, обновляем оригинальный файл запроса
'    If updated = True Then
'        full_name = remote_folder_path & manager & "\" & remote_filename
'        Set new_excel = CreateObject("Excel.Application")
'        new_excel.Workbooks.Open (full_name)
''        new_excel.Visible = False
'        Set sel = new_excel.Selection   ' запоминаем позицию выделения в удаленном файле
'        ThisWorkbook.Sheets(manager).range(ThisWorkbook.Sheets(manager).Cells(first_row, filename_column + 1), _
'            ThisWorkbook.Sheets(manager).Cells(ThisWorkbook.Sheets(manager).UsedRange.Rows.Count, filename_column + 2)).Copy
'        new_excel.Worksheets("Лист1").Cells(first_row, filename_column + 1).PasteSpecial Paste:=xlPasteValues
'        Application.CutCopyMode = False ' убираем рамку скопированного диапазона
'        new_excel.Worksheets("Лист1").Cells(sel.Row, 1).Select  ' выделяем первую ячейку из запомненного положения (чтобы область видимости не смещалась от начала листа)
'        new_excel.Application.ActiveWorkbook.Save
'        new_excel.Application.Quit
'        Set new_excel = Nothing
'    End If

' ---- конец проверки ----

End Sub

Sub run_Fill_ranges()
  
    For i = 2 To ActiveWorkbook.Worksheets.Count
        Call Fill_range(ActiveWorkbook.Worksheets(i).Name)
    Next i

End Sub

Sub Open_my_folder()
  folder_name = ActiveSheet.Name
  If folder_name = "ВСЕ" Then folder_name = range("C" & Selection.Row).Text
  Set fs = CreateObject("Scripting.FileSystemObject")
  If Not fs.folderexists(remote_folder_path & folder_name) Then folder_name = "_Курские менеджеры"
  Shell "explorer.exe " & remote_folder_path & folder_name, vbNormalFocus
End Sub

Sub Open_my_file()
  folder_name = ActiveSheet.Name
  If folder_name = "ВСЕ" Then folder_name = range("C" & Selection.Row).Text
  Set fs = CreateObject("Scripting.FileSystemObject")
  If Not fs.folderexists(remote_folder_path & folder_name) Then folder_name = "_Курские менеджеры"
  Workbooks.Open filename:=remote_folder_path & folder_name & "\" & remote_filename
End Sub


