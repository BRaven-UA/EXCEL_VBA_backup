Attribute VB_Name = "Main"
    Public break_execution As Boolean, debug_mode As Boolean
    Declare PtrSafe Function OemToCharBuff Lib "user32" Alias "OemToCharBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
    Dim sPath As String, sFile As String
    Const MainSheet As String = "Спецификация"
    Const TempSheet As String = "Raw data"
    Const ExclSheet As String = "Исключения"
    Dim FirstRow As Long, LastRow As Long
'    Dim Exclusions As Variant   ' массив с именами файлов для исключения из поиска (ранние версии, копии, для внутреннего пользования и т.д.)
    Dim Filename_Collection As New Collection   ' коллекция (массив) имен файлов, которые уже занесены в сводную спецификацию
    Dim OBJ As Object, Folder As Object, File As Object, SubFolder As Object


Sub KP_scan()
Attribute KP_scan.VB_Description = "запуск сканирования папок"
Attribute KP_scan.VB_ProcData.VB_Invoke_Func = "r\n14"
    debug_mode = False   ' закоментировать в нормальном режиме работы
    If debug_mode Then Debug.Print Now & " - Начало программы"
    break_execution = False
    Set Filename_Collection = Nothing   ' Очищаем коллекцию на случай если код запускается повторно
    
    With Application.FileDialog(msoFileDialogFolderPicker) 'выбераем папку для сканирования
        .Show
        sPath = .SelectedItems(1)
    End With
    Set OBJ = CreateObject("Scripting.FileSystemObject")
    Set Folder = OBJ.GetFolder(sPath)

    Application.DisplayAlerts = 0
    Application.ScreenUpdating = False

    LastRow = Sheets(MainSheet).Columns(12).EntireColumn.Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    On Error Resume Next
    For Each item In Sheets(MainSheet).Range("L2:L" & LastRow)
       Filename_Collection.Add "Внесено", Right(item.Value2, Len(item.Value2) - InStrRev(item.Value2, "\"))   ' чтобы не писать дополнительный код на поиск дубликатов, используется коллекция, которая автоматом отфильтровует дубликаты при добавлении
    Next item
    For Each item In Sheets(ExclSheet).Range(Sheets(ExclSheet).Cells(1, 1), Sheets(ExclSheet).Columns(1).EntireColumn.Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows))
       If item.Value2 Then Filename_Collection.Add "Исключено", item.Value2 ' в списке только названия файлов
    Next item
    Err.Clear
    If debug_mode Then Debug.Print Now & " - Конец создания коллекции"
    
'    Exclusions = Sheets(ExclSheet).Range("A1:A" & Sheets(ExclSheet).Cells.Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row)
    
    Call ListFiles(Folder)

For Each SubFolder In Folder.SubFolders 'сканируем содержимое подкаталогов
    Call ListFiles(SubFolder)
    If break_execution Then Exit For
    Call GetSubFolders(SubFolder)
Next SubFolder
    
    'Call Модель_сводная
    'Debug.Print Now & " - Конец заполнения моделей"
    'Rows(2).Copy
    'Range(Cells(3, 1), Cells(LastRow, 1)).EntireRow.PasteSpecial Paste:=xlPasteFormats 'форматируем все строки по шаблону второй строки
    
    Application.StatusBar = False
    Application.DisplayAlerts = 1
    Application.ScreenUpdating = True
End Sub

Sub GetSubFolders(ByRef SubFolder As Object) 'перебор вложенных каталогов

Dim FolderItem As Object

For Each FolderItem In SubFolder.SubFolders
    Call ListFiles(FolderItem)
    Call GetSubFolders(FolderItem)
Next FolderItem

End Sub

Sub ListFiles(ByRef Folder As Object) 'основной цикл перебора файлов в каталоге
    Dim data_changed As Boolean
    Dim new_items As Integer
    Dim Item_List As New Collection ' список для показа пользователяю, введен для повышения производительности
    Set Item_List = Nothing
    If debug_mode Then Debug.Print Now & " - Папка " & Folder.Name
    new_items = 0
    FirstRow = Sheets(TempSheet).Cells.Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row + 1  ' первая пустая строка
    
    For Each File In Folder.Files
'        Dim found As Boolean
        Application.StatusBar = File.Path 'визуализация процесса сканирования
        If (Left(File.Name, 3) = "КП ") And (Right(File.Path, 4) = ".xls") Then ' отбираем только файлы КП
            On Error Resume Next
            Var = Filename_Collection(File.Name)
            If Err.Number = 0 Then
                Item_List.Add Array(Var, File.Name), File.Name
            Else
                Err.Clear
                Item_List.Add Array("", File.Name), File.Name
                new_items = new_items + 1
            End If
        End If
    Next File
    
    If new_items > 0 Then
        Selection_Form.Directory.Caption = Folder.Path
        With Selection_Form.ListBox1
            .Clear
            For i = 0 To Item_List.Count - 1
                Var = Item_List(i + 1)
                .AddItem
                .Column(0, i) = Var(0)
                .Column(1, i) = Var(1)
            Next i
        End With
        Application.ScreenUpdating = True
        Selection_Form.Show
        Application.ScreenUpdating = False
        
        If break_execution Then GoTo Skip
        
        With Selection_Form.ListBox1
            For item = 0 To .ListCount - 1
                If .List(item, 0) = "" Then ' новый файл
                    If .Selected(item) Then  ' выделен пользователем - вносим в спецификацию
                        Set File = Folder.Files(.List(item, 1))
                        If debug_mode Then Debug.Print Now & " - Чтение файла " & File.Name
                        With Sheets(TempSheet).Range(Cells(FirstRow, 1), Cells(FirstRow + 4000, 11))
                            .Formula = "='" & File.parentFolder.Path & "\[" & File.Name & "]" & MainSheet & "'!" & "A6" & Chr(32) & Chr(38) & Chr(32) & Chr(34) & Chr(34) 'копируем содержание листа "спецификация"
                            .Value2 = .Value2
                        End With
                        If debug_mode Then Debug.Print Now & " - Список оборудования заполнен"
                        LastRow = Sheets(TempSheet).Cells.Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row 'последняя заполнення строка
                        Sheets(TempSheet).Range(Cells(FirstRow, 12), Cells(LastRow, 12)).Value = File.Path
                        Sheets(TempSheet).Range(Cells(FirstRow, 13), Cells(LastRow, 13)).Value = File.DateCreated
                        Sheets(TempSheet).Range(Cells(FirstRow, 14), Cells(LastRow, 14)).Value = File.DateLastModified
                        'Range(Cells(FirstRow, 15), Cells(LastRow, 15)).Value = "новый" 'это маркер чтобы знать какие КП добавились с момента последнего сканирования
                        If debug_mode Then Debug.Print Now & " - Атрибуты файла заполнены"
                        FirstRow = LastRow + 1
                        data_changed = True
                        Filename_Collection.Add "Внесено", File.Name
                    Else    ' файл не выделен пользователем - вносим его в исключения
                        Sheets(ExclSheet).Columns(1).EntireColumn.Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Offset(1, 0) = .List(item, 1)
                        Filename_Collection.Add "Исключено", .List(item, 1)
                    End If
                End If
            Next item
        End With
    End If

'
'    For Each File In Folder.Files
'        Dim found As Boolean
'        Application.StatusBar = File.Path 'визуализация процесса сканирования
'        If (Left(File.Name, 3) = "КП ") And (Right(File.Path, 4) = ".xls") Then ' отбираем только файлы КП
'            Selection_Form.ListBox1.AddItem
'            Item_Position = Selection_Form.ListBox1.ListCount - 1
'            Selection_Form.ListBox1.Column(1, Item_Position) = File.Name
'            For i = 1 To Filename_Collection.Count
'                If Filename_Collection.item(i) = File.Name Then
'                    Selection_Form.ListBox1.Column(0, Item_Position) = "Обработано"
''                    Selection_Form.ListBox1.Selected(Item_Position) = True
'                    found = True
'                    Exit For
'                End If
'            Next i
''            If debug_mode Then Debug.Print Now & " - Обработан файл " & File.Name
'            If Not found Then new_items = new_items + 1
'        End If
'
'    Next File
'
'    If new_items > 0 Then
'        'If new_items > 1 Then Selection_Form.Show
'        Selection_Form.Show
'        If break_execution Then GoTo Skip
'        For item = 0 To Selection_Form.ListBox1.ListCount - 1
'            If Selection_Form.ListBox1.Selected(item) And IsNull(Selection_Form.ListBox1.List(item, 0)) Then
'                Set File = Folder.Files(Selection_Form.ListBox1.List(item, 1))
'                If debug_mode Then Debug.Print Now & " - Чтение файла " & File.Name
'                With Sheets(TempSheet).Range(Cells(FirstRow, 1), Cells(FirstRow + 4000, 11))
'                    .Formula = "='" & File.parentFolder.Path & "\[" & File.Name & "]" & MainSheet & "'!" & "A6" & Chr(32) & Chr(38) & Chr(32) & Chr(34) & Chr(34) 'копируем содержание листа "спецификация"
'                    .Value2 = .Value2
'                End With
'                If debug_mode Then Debug.Print Now & " - Список оборудования заполнен"
'                LastRow = Sheets(TempSheet).Cells.Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row 'последняя заполнення строка
'                Sheets(TempSheet).Range(Cells(FirstRow, 12), Cells(LastRow, 12)).Value = File.Path
'                Sheets(TempSheet).Range(Cells(FirstRow, 13), Cells(LastRow, 13)).Value = File.DateCreated
'                Sheets(TempSheet).Range(Cells(FirstRow, 14), Cells(LastRow, 14)).Value = File.DateLastModified
'                'Range(Cells(FirstRow, 15), Cells(LastRow, 15)).Value = "новый" 'это маркер чтобы знать какие КП добавились с момента последнего сканирования
'                If debug_mode Then Debug.Print Now & " - Атрибуты файла заполнены"
'                FirstRow = LastRow + 1
'                data_changed = True
'                Filename_Collection.Add File.Name, File.Name
'            End If
'        Next item
'    End If

Skip:
    If data_changed Or break_execution Then
    '    If debug_mode Then Debug.Print Now & " - Конец перебора файлов"
        LastRow = Sheets(MainSheet).Cells.Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row + 1
        Sheets(TempSheet).Cells.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=Sheets(TempSheet).Range("A1:T2"), CopyToRange:=Sheets(MainSheet).Cells(LastRow, 1), Unique:=False
        If debug_mode Then Debug.Print Now & " - Конец копирования"
        If LastRow > 1 Then Sheets(MainSheet).Rows(LastRow).Delete 'удаляем повторяющуюся строку с названиями столбцов
        Sheets(TempSheet).Range(Cells(3, 1), Cells(Sheets(TempSheet).Cells.Rows.Count, 1)).EntireRow.Delete 'очищаем рабочую область листа "Raw data"
        If debug_mode Then Debug.Print Now & " - Конец очистки рабочей области"
    End If
    
End Sub



Sub temp()  'раскрасить выделенный диапазон для визуализации отличий в строках
Dim c As Range
'Cells(Selection.Row, Selection.Column).Interior.ColorIndex = 37
For Each c In Selection
    'If Cells(c.Row + 1, c.Column).Value = c.Value Then Cells(c.Row + 1, c.Column).Interior.ColorIndex = c.Interior.ColorIndex Else Cells(c.Row + 1, c.Column).Interior.ColorIndex = c.Interior.ColorIndex Xor 10
    t1 = c.Value
    t2 = Cells(c.Row + 1, c.Column).Value
    If t1 <> t2 Then
        For i = 1 To Len(t2)
            If Mid(t1, i, 1) <> Mid(t2, i, 1) Then Exit For
        Next i
        Cells(c.Row + 1, c.Column).Characters(i, Len(t2)).Font.ColorIndex = 3
    End If

Next c

'For i = 1 To 10023
'    Cells(i, 1) = ChrW(i)
'Next i
'Debug.Print Worksheets(MainSheet).Columns(19).Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row

'Dim results As Object
'Dim r1 As String, r2 As String
'
'    With Application.FileDialog(msoFileDialogFolderPicker) 'выбераем папку для сканирования
'        .Show
'        sPath = .SelectedItems(1)
'    End With
'
'Set results = CreateObject("WScript.Shell").Exec("CMD /C DIR """ & sPath & "\КП*.xl*"" /S /B /A:-D").StdOut
'While Not results.AtEndOfStream
'    Debug.Print ToAnsi(results.readline)
'Wend
'r1 = results.readall
'r2 = ToAnsi(r1)

'Debug.Print "CMD /C DIR """ & sPath & " /S /B /A:-D" 'results
'Range("A4").Resize(UBound(Split(results, vbCrLf)), 1).Value = WorksheetFunction.Transpose(Split(results, vbCrLf))

End Sub

Function ToAnsi(s As String) As String
    Dim Buffer As String
    Buffer = Space(Len(s) + 1)
    OemToCharBuff s, Buffer, Len(s)
    ToAnsi = Left(Buffer, Len(s))
End Function

Sub Дополнительные_колонки()
    LastRow = Worksheets(MainSheet).Cells.Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row  'последняя заполнення строка
'    Worksheets(MainSheet).Range(Cells(2, 16), Cells(LastRow, 16)).Formula = "=IF(J2="""",C2&"""",J2)" 'выбераем итоговую производительность
'    Worksheets(MainSheet).Range(Cells(2, 17), Cells(LastRow, 17)).Formula = "=IF(I2="""",D2&"""",I2)" 'выбераем курскую или импортную модель
'    Worksheets(MainSheet).Range(Cells(2, 18), Cells(LastRow, 18)).Formula = "=YEAR(M2)"    'проставляем год
     Worksheets(MainSheet).Range(Cells(Worksheets(MainSheet).Columns(19).Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row + 1, 19), Cells(LastRow, 19)).Formula = "=YEAR(M2)"  'проставляем год
'    For i = 2 To LastRow
'    If IsEmpty(Cells(i, 9)) Then Cells(i, 16) = Cells(i, 4) Else Cells(i, 16) = Cells(i, 9) 'выбераем курскую или импортную модель
'    Call progress_bar(0.5, Fix(i / LastRow * 100))
'    Next i
    Application.StatusBar = False
End Sub

Sub progress_bar(Interval As Single, Persent As Byte)
    Static t As Single
    s = String(50, ChrW(9601))  'чтобы  не прописывать этот символ юникода вручную много-много раз
    If Timer > t + Interval Then
        t = Timer
        b = Replace(s, ChrW(9601), ChrW(9609), , Fix(Persent / 2) + 1)  ' +1 добавленно из-за некорректного отображения исходной строки s, не могу понять с чем это связано
        Application.StatusBar = b & " " & Persent & "%"
    End If
End Sub
