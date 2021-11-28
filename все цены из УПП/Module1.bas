Attribute VB_Name = "Module1"
Sub Set_updating(ByVal enabled As Boolean)    ' вкл/выкл отпимизации времени выполнения расчетов
    Application.ScreenUpdating = enabled
    Application.Calculation = IIf(enabled, xlCalculationAutomatic, xlCalculationManual)
End Sub

Sub рассчитать()
    ' процедура просто заполняет рабочие ячейки формулами, считает их один раз и сохраняет как значения
    ' процедура нужна для того чтобы формулы на листе не пересчитывались постоянно, так как это долгий процесс
    Set_updating (False)
    
    last_row = Cells.Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    Range("L2:L" & last_row).FormulaLocal = "=ДАТАЗНАЧ(ПСТР(K2;НАЙТИ("" от "";K2)+4;10))"
    Range("M2:M" & last_row).FormulaLocal = "=ЕСЛИ(И(H2=D2;D2=I2);E2;ЕСЛИ(I2=D2;E2/J2;0)) * ЕСЛИ(F2=""нет"";1.2;1)"
    Range("N2:N" & last_row).FormulaLocal = "=НЕ(И(H2<>I2;J2=1))"
    Range("O2:O" & last_row).FormulaLocal = "=ЕСЛИ(C2=""руб"";M2;ЕСЛИ(ЕЧИСЛО(M2);ВПР(C2;'форма расчета'!$D$1:$E$3;2;ЛОЖЬ)*M2;""""))"
    'Range("P2:P" & last_row).FormulaLocal = "=ЕСЛИ(O2<>0;B2;"""")"
    'Range("Q2:Q" & last_row).FormulaLocal = "=СУММЕСЛИ(P:P;B2;O:O)/СЧЁТЕСЛИ(P:P;B2)"
    Application.Calculate
    Range("L2:N" & last_row).Copy
    Range("L2:N" & last_row).PasteSpecial xlPasteValues
    
    ' добавляем в список аналоги номенклатуры из соответствующего листа
    Скопировать_аналоги
    
End Sub

Sub Скопировать_аналоги()   ' добавляет в таблицу с ценами из УПП данные об аналогах номенклатур с листа "Аналоги"
    
    Set_updating (False)
    
    Set лист_полный_список = Worksheets("полный список")
    Set лист_аналоги = Worksheets("Аналоги")
    
    конец_таблицы_из_УПП = лист_полный_список.Columns(1).Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    
    ' удаляем лишнее после таблицы с ценами из УПП
    лист_полный_список.Rows(конец_таблицы_из_УПП + 1 & ":" & 1048576).EntireRow.ClearContents
    
    ' копируем столбцы аналогов в конец соотв. столбцов таблицы из УПП
    LastRow = лист_аналоги.Cells.Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    лист_аналоги.Range("A2:A" & LastRow).Copy (лист_полный_список.Cells(конец_таблицы_из_УПП + 1, 2))
    лист_аналоги.Range("B2:B" & LastRow).Copy (лист_полный_список.Cells(конец_таблицы_из_УПП + 1, 16))
    
    Set_updating (True)
    
End Sub


' Добавляет значение текущей ячейки покупного изделия из листа "форма расчета" в таблицу аналогов номенклатуры на листе "Аналоги"
Sub Добавить_аналог()

    Set_updating (False)
    
    номенклатура = Cells(Selection.Row, 1).Text
    покупное_название = Cells(Selection.Row, 9).Text
    Set лист_аналоги = Worksheets("Аналоги")
    Set лист_полный_список = Worksheets("полный список")
    
    If покупное_название = "#Н/Д" Or покупное_название = "" Then покупное_название = " - НЕТ ДАННЫХ - "
    
    LastRow = лист_аналоги.Cells.Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
    дубли = 0
    For i = 2 To LastRow
        If (лист_аналоги.Cells(i, 1).Text = номенклатура) And (лист_аналоги.Cells(i, 2).Text = покупное_название) Then
            дубли = MsgBox("Такой аналог уже был добавлен ранее", vbInformation, "Дублирование")
            Exit For
        End If
    Next i
    
    If дубли = 0 Then   ' если дублирования нет, добавляем аналог в таблицу аналогов
        лист_аналоги.Cells(LastRow + 1, 1) = номенклатура
        лист_аналоги.Cells(LastRow + 1, 2) = покупное_название
        
'        ' тормозит, если сразу добавлять в полный список номенклатур
        
        LastRow = лист_полный_список.Cells.Find(What:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
        лист_полный_список.Cells(LastRow + 1, 2) = номенклатура
        лист_полный_список.Cells(LastRow + 1, 16) = покупное_название
        
        Cells(5, 9).Copy (Cells(Selection.Row, 9))  ' восстанавливаем шаблонную формулу аналога
    End If
        
    Set_updating (True)
    
End Sub

Function InsertExchangeRateFromCBRF(inpDate As Date, Optional curr As String = "USD") As Double
'Функция запрашивает на необходимую дату курс доллара в архиве ЦБРФ
Dim sURI As String
Dim oHttp As Object
Dim htmlcode, outstr As String
Dim date_, month_, year_ As Integer

'разбираем дату на составляющие
day_ = Format(inpDate, "dd")
month_ = Format(inpDate, "mm")
year_ = Format(inpDate, "yyyy")

'формируем строку для веб-запроса
'sURI = "http://old.cbr.ru/currency_base/daily.aspx?C_month=" & month_ & "&C_year=" & year_ & "&date_req=" & day_ & "%2F" & month_ & "%2F" & year_
sURI = "http://cbr.ru/currency_base/daily/"
'On Error Resume Next
Set oHttp = CreateObject("MSXML2.XMLHTTP")
'If Err.Number <> 0 Then
'    Set oHttp = CreateObject("MSXML.XMLHTTPRequest")
'End If
'On Error GoTo 0

'On Error Resume Next
oHttp.Open "GET", sURI, False
oHttp.Send
If Err.Number <> 0 Then
    MsgBox "Не удалось подключиться к базе данных ЦБРФ!", vbExclamation
    Err.Clear
    InsertExchangeRateFromCBRF = 0
    Exit Function
End If
'получаем HTML страницы с курсами и извлеккаем курс доллара
htmlcode = oHttp.responseText
outstr = Mid(htmlcode, InStr(InStr(1, htmlcode, curr), htmlcode, ",") - 2, 7)
Set oHttp = Nothing

InsertExchangeRateFromCBRF = Val(Replace(outstr, ",", "."))

End Function


Sub find_similar()
    
    Set_updating (False)
    Application.EnableCancelKey = xlInterrupt
    max_comparision_value = Range("search_precision")
    Set First_cell = Cells(Selection.Row, 9)
    Source = Cells(Selection.Row, 1).Text
    Dim full_list As Variant
    Dim result_list As New Collection
    Set result_list = Nothing
    fast_check_string = ""
    result_column = 11
    full_list = Application.Transpose(Range("Номенклатура").Value2)
    validation_list = "=" & Cells(First_cell.Row, result_column).Address
    
    For Each entry In full_list
        comparision_result = KRITERIJ_BLIZOSTI_STROK(Source, entry)
        If comparision_result > max_comparision_value Then
            If InStr(1, fast_check_string, entry, vbTextCompare) = 0 Then
                fast_check_string = fast_check_string & entry
                For Index = 1 To result_list.Count
                    If comparision_result > result_list(Index)(0) Then Exit For
                Next Index
                Content = Array(comparision_result, entry)
                If Index > result_list.Count Then result_list.Add Content Else result_list.Add Content, Before:=Index
            End If
       End If
    Next entry
    
    If result_list.Count > 0 Then
        
        First_cell.Value = Source
        
        For Index = 1 To result_list.Count
            Cells(First_cell.Row, result_column) = result_list(Index)(1)
            result_column = result_column + 1
        Next Index
       
        With First_cell.Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:=validation_list & ":" & Cells(First_cell.Row, result_column - 1).Address
        End With
    
    End If
    
    Set_updating (True)
    
End Sub


Function KRITERIJ_BLIZOSTI_STROK(ByVal s1 As String, ByVal s2 As String) As Double
  ' Взято у Вадима и переделано под текущую задачу
  '
  ' Находим количество би-грамм в строке меньшей длины: N = len(s_min) + 1 - 2.
  ' Для каждой би-граммы из короткой строки производим поиск - входит ли эта би-граммы во вторую строку.
  ' Критерий похожести двух строк находим по формуле J = 100 * [количество биграмм из короткой строки, входящие в первую] / N
  ' Критерий похожести двух строк может принимать значения от 0 до 100.
  '

  s1 = UCase(s1)
  s2 = UCase(s2)
  s = vse_v_angliu(s)
  Ls1 = Len(s1)
  Ls2 = Len(s2)
      
'-------------------------------------------------------------------
      
If (Ls1 <= 4) Or (Ls2 <= 4) Then
      
  If (Ls1 <= 1) Or (Ls2 <= 1) Then
    KRITERIJ_BLIZOSTI_STROK = 0
    GoTo vihod
  End If
      
  If Ls1 < Ls2 Then
  
    If (Ls1 >= 2) And (Ls1 <= 4) Then
        s1 = " " & s1 & " "
        s2 = " " & s2 & " "
      If InStr(s2, s1) <> 0 Then
       KRITERIJ_BLIZOSTI_STROK = 100
      Else
       KRITERIJ_BLIZOSTI_STROK = 0
      End If
    GoTo vihod
    End If
    
  Else
  
    If (Ls2 >= 2) And (Ls2 <= 4) Then
        s1 = " " & s1 & " "
        s2 = " " & s2 & " "
      If InStr(s1, s2) <> 0 Then
       KRITERIJ_BLIZOSTI_STROK = 100
      Else
       KRITERIJ_BLIZOSTI_STROK = 0
      End If
     GoTo vihod
    End If
    
  End If
End If

'-------------------------------------------------------------------

   ss1 = Replace(s1, " ", "")
   ss2 = Replace(s2, " ", "")

  If Len(ss1) < Len(ss2) Then
   If InStr(ss2, ss1) <> 0 Then
      KRITERIJ_BLIZOSTI_STROK = 99
    GoTo vihod
   End If
  Else
   If InStr(ss1, ss2) <> 0 Then
      KRITERIJ_BLIZOSTI_STROK = 99
    GoTo vihod
   End If
  End If
   
  s1 = " " & s1 & " "
  s2 = " " & s2 & " "
  Ls1 = Len(s1)
  Ls2 = Len(s2)
  
  If Ls1 > Ls2 Then
   sbuf = s1
   s1 = s2
   s2 = sbuf
   Lbuf = Ls1
   Ls1 = Ls2
   Ls2 = Lbuf
  End If
  
   ' s1 - короткая строка
   ' s2 - длинная строка
  
  N = Len(s1) - 2 + 1 ' количество би-грамм в короткой строке
  
  Dim arg_sovp1() As Integer
  Dim arg_sovp2() As Integer
  
  ReDim arg_sovp1(1 To Ls1)
  ReDim arg_sovp2(1 To Ls2)
  
  N_sovp = 0
  
  i_end = Ls1 - 2 + 1
  j_end = Ls2 - 2 + 1

  For i = 1 To i_end
    smid1 = Mid(s1, i, 2)
    For j = 1 To j_end
         If StrComp(smid1, Mid(s2, j, 2)) = 0 Then
          If arg_sovp2(j) = 0 Then ' би-грама из короткой строки есть в длинной строке
             arg_sovp2(j) = 1
             N_sovp = N_sovp + 1
             Exit For
          End If
        End If
    Next j
  Next i
 
 KRITERIJ_BLIZOSTI_STROK = 100 * N_sovp / N
    
vihod:
  

End Function


Function vse_v_angliu(s) As String

s = Trim(s)

s_bez_prob = s
len_s = Len(s_bez_prob)

s_bez_prob = UCase(s_bez_prob)
For bukva = 1 To len_s

bukva_now = Mid(s_bez_prob, bukva, 1)

Select Case bukva_now
Case "О"
bukva_now = "O"
Case "М"
bukva_now = "M"
Case "К"
bukva_now = "K"
Case "Р"
bukva_now = "P"
Case "Н"
bukva_now = "H"
Case "Т"
bukva_now = "T"
Case "С"
bukva_now = "C"
Case "А"
bukva_now = "A"
Case "Е"
bukva_now = "E"
Case "В"
bukva_now = "B"
Case "У"
bukva_now = "Y"
Case "Х"
bukva_now = "X"
End Select

Mid(s_bez_prob, bukva, 1) = bukva_now

Next bukva

vse_v_angliu = s_bez_prob

End Function
