Attribute VB_Name = "Module1"
Dim расчет As Worksheet, структура As Worksheet, цены As Worksheet, трудозатраты As Worksheet
Dim все_изделия As Variant, все_цены As Variant, все_трудозатраты As Variant
Dim все_изделия_область As Variant, все_цены_область As Variant, все_трудозатраты_область As Variant
Dim виды, изделия, номенклатуры, родители, количества, общие_количества, массы    ' названия столбцов с данными
Dim стартовая_колонка, последняя_колонка, временные_колонки
Dim текущая_строка  ' глобальная позиция текущей строки
Dim количество_нулевых_цен As Long, количество_нулевых_трудозатрат As Long  ' счетчики для ошибок

Sub Сформировать_структуру()
    
    стартовая_строка = 15
    стартовая_колонка = 2
    текущая_строка = стартовая_строка
    временные_колонки = стартовая_колонка + 30  ' пока неизвестна глубина вложенности спецификации записываем дополнительные данные во временную область
    
    'start_time = Timer
    Application.ScreenUpdating = False  ' Отключаем обновление экрана (для повышения производительности)
    Application.Calculation = xlCalculationManual   ' Устанавливаем ручной пересчет ячеек (для повышения производительности)

    Set расчет = Worksheets("Расчет")   ' лист с результатами расчета
    Set структура = Worksheets("Структура")    ' лист со структурой изделия и массами
    Set цены = Worksheets("Покупные_изделия")   ' лист в ценами на покупные изделия
    Set трудозатраты = Worksheets("Трудозатраты")   ' лист в трудозатратами на изделие
    
    ' проверка наличия посторонних данных в зоне построения
    If расчет.UsedRange.Rows.Count > стартовая_строка Then
        result = MsgBox("В рабочей области имеются данные, которые будут удалены. Продолжить?", vbYesNoCancel, "Внимание !")
        If result <> vbYes Then GoTo Выход
    End If
    
    ' очищаем зону построения
    расчет.Cells(стартовая_строка, 1).Resize(1048576 - стартовая_строка).EntireRow.Delete
    
    ' задаем области для ссылок в ГИПЕРССЫЛКА и массивы этих данных для быстроты поиска
    Set все_изделия_область = структура.Cells.Find("Изделие").CurrentRegion
    все_изделия = Application.Transpose(все_изделия_область.Value2) ' заполняем массив исходных данных
    Set все_цены_область = цены.Cells.Find("Последняя цена").CurrentRegion
    все_цены = Application.Transpose(все_цены_область.Value2) ' заполняем массив цен на покупные изделия
    Set все_трудозатраты_область = трудозатраты.Cells.Find("Изделие").CurrentRegion
    все_трудозатраты = Application.Transpose(все_трудозатраты_область.Value2) ' заполняем массив трудозатрат на изделие
    
    ' определяем названия и номера столбцов с данными
    For i = LBound(все_изделия, 1) To UBound(все_изделия, 1)
        заголовок = все_изделия(i, 1)
        Select Case заголовок
            Case "Вид элемента": виды = i
            Case "Изделие": изделия = i
            Case "Номенклатура": номенклатуры = i
            Case "Куда входит": родители = i
            Case "Кол.": количества = i
            Case "Общ. кол.": общие_количества = i
            Case "Масса изделия": массы = i
        End Select
    Next i
    
    ' перебор и разузлование всех изделий без родителя
    For i = LBound(все_изделия, 2) + 1 To UBound(все_изделия, 2)
        If все_изделия(родители, i) = "" Then
            текущая_строка = текущая_строка + 1
            Call Разузлование(i, стартовая_колонка)
        End If
    Next i
    
    ' обработка полученной таблицы
    With расчет
        
        ' настройка отображения группировок с "плюсом" вверху и без итогов
        With .Outline
            .SummaryRow = xlAbove
            .SummaryColumn = xlLeft
            
            For i = 8 To 1 Step -1  ' скрываем все группировки, кроме первой
                .ShowLevels i
            Next i
        End With
    
        ' удаляем промежуточные колонки чтобы временные колонки пододвинулись к изделиям
        .range(.Cells(стартовая_строка, последняя_колонка + 1), .Cells(стартовая_строка, временные_колонки - 1)).EntireColumn.Delete
        
'        .Cells(стартовая_строка + 1, последняя_колонка + 1).Value2 = 1   ' исправляем количество главного изделия
        .range(.Cells(стартовая_строка, стартовая_колонка), .Cells(стартовая_строка, последняя_колонка - 1)).ColumnWidth = 3 ' уплотняем колонки со структурой
        .Columns(последняя_колонка).ColumnWidth = 50   ' раздвигаем последнюю колонку чтобы не перекрывать названия изделий
        
        ' пересчитываем ячейки для правильного автоподгонки размеров таблицы
        Application.Calculate
        
        ' создаем заголовки для таблицы
        .Cells(стартовая_строка, стартовая_колонка).Value = "Изделие"
        .Cells(стартовая_строка, последняя_колонка + 1).Value = "Кол."
        .Cells(стартовая_строка, последняя_колонка + 2).Value = "Масса"
        .Cells(стартовая_строка, последняя_колонка + 3).Value = "Закупка за ед."
        .Cells(стартовая_строка, последняя_колонка + 4).Value = "Закупка общая"
        .Cells(стартовая_строка, последняя_колонка + 5).Value = "Время на ед., мин*"
        .Cells(стартовая_строка, последняя_колонка + 6).Value = "Время общее, мин*"
        .Cells(стартовая_строка, последняя_колонка + 7).Value = "Трудозатраты"
        .Cells(стартовая_строка, последняя_колонка + 8).Value = "Итого"
        .Cells(стартовая_строка, последняя_колонка + 9).Value = "* время указано с учетом коэффициента пробивки (см. настройки на листе ""Трудозатраты"")"
        
        ' форматируем заголовки
        Set заголовки = .range(.Cells(стартовая_строка, стартовая_колонка), .Cells(стартовая_строка, последняя_колонка + 8))
        заголовки.Interior.Color = 16247773 ' цвет заголовков
        заголовки.VerticalAlignment = xlCenter  ' вертикальное центрирование
        Set заголовки_2 = заголовки.Offset(, последняя_колонка - 1).Resize(, заголовки.Columns.Count - последняя_колонка + 1) ' отдельно форматируем остальные заголовки
        заголовки_2.Borders.LineStyle = xlContinuous
        Call Внешние_границы(заголовки, xlMedium)   ' толстые границы нужны чтобы не сливаться с разделительной чертой при закреплении областей экрана
        заголовки_2.WrapText = True ' перенос по словам
        заголовки_2.HorizontalAlignment = xlCenter  ' горизонтальное центрирование
        заголовки_2.EntireRow.AutoFit   ' корректировка высоты
        заголовки_2.Columns.AutoFit ' атоматически подгоняем ширину столбцов
        Set таблица_2 = заголовки_2.Resize(.UsedRange.Rows.Count - стартовая_строка + 1)  ' часть таблицы без изделий
        таблица_2.EntireColumn.AutoFit
        
        ' форматируем данные в таблице
        .Columns(последняя_колонка + 2).NumberFormat = "#,##0.0" ' массы
        .Columns(последняя_колонка + 3).Resize(, 2).NumberFormat = "#,##0$" ' цены в рублях
        .Columns(последняя_колонка + 5).Resize(, 2).NumberFormat = "#,##0" ' время в минутах
        .Columns(последняя_колонка + 7).Resize(, 2).NumberFormat = "#,##0$" ' трудозатраты и итоговая сумма в рублях
    End With
    
    'скрываем настройки расчета чтобы не занимали место на экране
    расчет.Cells(1, 1).Resize(стартовая_строка - 1).EntireRow.Hidden = True
    
    ' закрепляем заголовки
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    
Выход:
    ' восстанавливаем исходные значения
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' предупреждение об ошибках
    текст_сообщения = ""
    If количество_нулевых_цен > 0 Then текст_сообщения = текст_сообщения & vbNewLine & "В спецификации " & Str(количество_нулевых_цен) & " покупных изделий без цены"
    If количество_нулевых_трудозатрат > 0 Then текст_сообщения = текст_сообщения & vbNewLine & "В спецификации " & Str(количество_нулевых_трудозатрат) & " изделий без трудозатрат"
    If текст_сообщения <> "" Then result = MsgBox(текст_сообщения, vbOKOnly, "Внимание !")
    
    End ' очищаем глобальные переменные
End Sub

' вносит в лист структуры текущее изделие и рекурсивно обрабатывает все его компоненты
' вложения реализованы в отдельных колонках из-за ограничения Экселя в 8 уровней вложенности для группировок
' сами группировки оставлены для удобства пользования
Sub Разузлование(позиция, текущая_колонка)
    
    эта_строка = текущая_строка ' локальная позиция текущей строки
    
    последняя_колонка = WorksheetFunction.Max(текущая_колонка, последняя_колонка)    ' орбновляем номер самой последней колонки
    
    изделие = все_изделия(изделия, позиция) ' получаем название текущего изделия
    
    ' ищем дочерние элементы (если они есть)
    For i = LBound(все_изделия, 2) + 1 To UBound(все_изделия, 2)
        If все_изделия(родители, i) = изделие Then  ' изделие является родителем
            текущая_строка = текущая_строка + 1 ' глобальный счетчик текущей строки
            Call Разузлование(i, текущая_колонка + 1)   ' рекурсивное разузлование дочернего элемента
        End If
    Next i
    
    ' вносим данные по текущему изделию
    расчет.Cells(эта_строка, текущая_колонка).Value2 = изделие    ' название
    
    If IsEmpty(все_изделия(количества, позиция)) Then   ' для основных изделий корректируем количество на 1 шт.
        все_изделия_область(позиция, количества) = 1
        все_изделия_область(позиция, общие_количества) = 1
    End If
    ссылка = структура.Name & "!" & все_изделия_область(позиция, количества).Address(False, False)
    расчет.Cells(эта_строка, временные_колонки).FormulaLocal = "=ГИПЕРССЫЛКА(""#" & ссылка & """;" & ссылка & ")"    ' количество
    
    ссылка = структура.Name & "!" & все_изделия_область(позиция, массы).Address(False, False)
    расчет.Cells(эта_строка, временные_колонки + 1).FormulaLocal = "=ГИПЕРССЫЛКА(""#" & ссылка & """;" & _
        IIf(IsEmpty(все_изделия(массы, позиция)), Chr(34) & Chr(34), ссылка) & ")" ' масса
    
    ' задаем уровень группировки на основании номера колонки
    расчет.Rows(эта_строка).OutlineLevel = WorksheetFunction.Median(1, текущая_колонка - стартовая_колонка + 1, 8) ' Excel ограничен максимум 8 уровнями
    
    Select Case все_изделия(виды, позиция)  ' дальше действуем с учетом вида изделия
        
        Case "Сборочные единицы", "Детали", "Комплекты"
            
            Set временный_диапазон = расчет.range(расчет.Cells(эта_строка, текущая_колонка), расчет.Cells(эта_строка, временные_колонки + 7))
            временный_диапазон.Interior.Color = 10086143  ' цвет для детали/сборки
            Call Внешние_границы(временный_диапазон)   ' границы ячеек
            
            ' общий диапазон для всех дальнейших формул с СУММЕСЛИ
            Set диапазон_критерий = расчет.range(расчет.Cells(эта_строка + 1, текущая_колонка + 1), расчет.Cells(текущая_строка, текущая_колонка + 1))
            
            ' при отсутствии данных о массе сборки вычисляем ее как сумму всех составляющих сборки
            If IsEmpty(все_изделия(массы, позиция)) Then
                Set диапазон_суммирования = расчет.range(расчет.Cells(эта_строка + 1, временные_колонки + 1), расчет.Cells(текущая_строка, временные_колонки + 1))
               ' расчет.Cells(эта_строка, временные_колонки + 1).FormulaLocal = "=ГИПЕРССЫЛКА(""#" & ссылка & """;СУММЕСЛИ(" & _
                    диапазон_критерий.Address(False, False) & ";""*"";" & диапазон_суммирования.Address(False, False) & "))"
                With расчет.Cells(эта_строка, временные_колонки + 1)
                    .FormulaArray = "=SUMPRODUCT(IF(ISTEXT(" & диапазон_критерий.Address(False, False) & _
                    "),1,0)," & диапазон_суммирования.Offset(0, -1).Address(False, False) & "," & диапазон_суммирования.Address(False, False) & ")"
                    ' выделяем красным шрифтом без подчеркивания (на данный момент подчеркивание установлено автоматически из-за наличия ГИПЕРССЫЛКА)
                    .Font.Color = RGB(255, 0, 0)
                    .Font.Underline = xlUnderlineStyleNone
                End With
            End If
            
            ' прописываем в сборочные единицы формулу для суммарной стоимости
            Set диапазон_суммирования = расчет.range(расчет.Cells(эта_строка + 1, временные_колонки + 3), расчет.Cells(текущая_строка, временные_колонки + 3))
            расчет.Cells(эта_строка, временные_колонки + 3).FormulaLocal = _
                "=СУММЕСЛИ(" & диапазон_критерий.Address(False, False) & ";""*"";" & диапазон_суммирования.Address(False, False) & ")*" & расчет.Cells(эта_строка, временные_колонки).Address(False, False)
            
            ' прописываем формулу стоимости за 1 шт исходя из общей стоимости (общая стоимость должна быть простой формулой чтобы не заморачиваться с формулами массива, а нужной формулы СУММПРОИЗВЕСЛИ в Экселе нет)
            расчет.Cells(эта_строка, временные_колонки + 2).FormulaR1C1 = "=RC[1]/RC[-2]"
            
            ' ищем трудозатраты на текущее изделие
            труд = 0
            ссылка = "0"
            For i = LBound(все_трудозатраты, 2) + 1 To UBound(все_трудозатраты, 2)
                If все_трудозатраты(1, i) = изделие Then
                    труд = все_трудозатраты(2, i) ' берем итоговое значение из второй колонки
                    ссылка = трудозатраты.Name & "!" & все_трудозатраты_область(i, 2).Address(False, False)
                    Exit For
                End If
            Next i
                        
            If труд = 0 Then ' если не найдено трудозатрат для данного изделия
                количество_нулевых_трудозатрат = количество_нулевых_трудозатрат + 1 ' увеличиваем глобальный счетчик
                расчет.Cells(эта_строка, временные_колонки + 4).Interior.Color = RGB(255, 0, 0)  ' красим ячейку в красный цвет
            End If
            
            расчет.Cells(эта_строка, временные_колонки + 5).FormulaR1C1 = "=RC[-1]*RC[-5]"
            
            Set диапазон_суммирования = расчет.range(расчет.Cells(эта_строка + 1, временные_колонки + 5), расчет.Cells(текущая_строка, временные_колонки + 5))
            расчет.Cells(эта_строка, временные_колонки + 4).FormulaLocal = "=ГИПЕРССЫЛКА(""#" & ссылка & """;" & ссылка & _
                    "+СУММЕСЛИ(" & диапазон_критерий.Address(False, False) & ";""*"";" & диапазон_суммирования.Address(False, False) & "))"
            
            расчет.Cells(эта_строка, временные_колонки + 6).FormulaR1C1 = "=RC[-1]/60*(1+запас_труд)*стоимость_трудочаса"
            
            расчет.Cells(эта_строка, временные_колонки + 7).FormulaR1C1 = "=RC[-4]*(1+наценка_покупное)+RC[-1]"
        
        Case Else   ' покупные изделия
        
            ' ищем стоимость покупного изделия
            цена = 0
            For i = LBound(все_цены, 2) + 1 To UBound(все_цены, 2)
                If все_цены(1, i) = все_изделия(номенклатуры, позиция) Then
                    цена = все_цены(4, i)
                    ссылка = цены.Name & "!" & все_цены_область(i, 4).Address(False, False)
                    Exit For
                End If
            Next i
            
            If цена = 0 Then ' если нет цены
                количество_нулевых_цен = количество_нулевых_цен + 1 ' увеличиваем глобальный счетчик
                расчет.Cells(эта_строка, временные_колонки + 2).Interior.Color = RGB(255, 0, 0)  ' красим ячейку в красный цвет
            Else
                расчет.Cells(эта_строка, временные_колонки + 2).FormulaLocal = "=ГИПЕРССЫЛКА(""#" & ссылка & """;" & ссылка & ")"  ' цена за единицу
            End If
            расчет.Cells(эта_строка, временные_колонки + 3).FormulaR1C1 = "=RC[-3]*RC[-1]"  ' цена с учетом количества
            
            ' цвет для покупных изделий
            расчет.range(расчет.Cells(эта_строка, текущая_колонка), расчет.Cells(эта_строка, временные_колонки + 7)).Font.Italic = True
    End Select
    
End Sub

' устанавливает только внешние границы для области
Sub Внешние_границы(ByRef область As Variant, Optional толщина = xlThin)
    With область
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = толщина
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = толщина
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = толщина
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = толщина
    End With
End Sub

Sub Корректировка_структуры()
    Application.ScreenUpdating = False
    With Worksheets("Структура")
        первая_строка = .Cells.Find("Изделие").Row + 1
        последняя_колонка = .Cells.Find(what:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column
    
        ' устанавливаем количество основного изделия в 1 шт
        .Cells(первая_строка, .Cells.Find("Кол.").Column).Value = 1
        
        'удаление повторяющихся строк. Обратный порядок перебора чтобы удаление строк не сбивало переменную цикла
        предыдущая_строка = ""
        For i = .Cells.Find(what:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row To первая_строка Step -1
            текущая_строка = ""
            For j = 1 To последняя_колонка
                текущая_строка = текущая_строка + .Cells(i, j).Text
            Next j
            If текущая_строка = предыдущая_строка Then .Rows(i).EntireRow.Hidden = True
            предыдущая_строка = текущая_строка
        Next i
    End With
    Application.ScreenUpdating = True
End Sub

Sub Корректировка_трудозатрат() ' берем сводную таблицу трудозатрат на изделие целиком и делим все трудозатраты на количество одинаковых сборок в изделии
    
    Dim тех_операции As Variant
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    With Worksheets("Трудозатраты")
    
        If Not IsEmpty(.Cells(2, 1)) Then
            result = MsgBox("Корректировка уже была проведена", vbInformation, "Отмена")
            GoTo Выход
        End If
        
        Set начало_таблицы = .Cells.Find("Изделие")
        
        .Rows(начало_таблицы.Row + 1).EntireRow.Delete  ' ненужная строка из отчета PLM
        
        последняя_строка = .Cells.Find(what:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
        последняя_колонка = .Cells.Find(what:="*", LookIn:=xlValues, SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column
        
        Set temp = .Cells.Find("Пробивная")
        If temp Is Nothing Then пробивная = 0 Else пробивная = temp.Column - начало_таблицы.Column - 1 ' начиная от первой техоперации
            
        ' прописываем формулу суммы по каждой техоперации
        For Each current_cell In range(начало_таблицы.Offset(1, 2), .Cells(начало_таблицы.Row + 1, последняя_колонка)).Cells
            current_cell.FormulaLocal = "=СУММ(" & current_cell.Offset(1).Address(False, False) & _
                         ":" & .Cells(последняя_строка, current_cell.Column).Address(False, False) & ")"
        Next current_cell
        
        ' прописываем формулу суммы по каждому изделию
        For Each current_cell In range(начало_таблицы.Offset(1, 1), .Cells(последняя_строка, начало_таблицы.Column + 1)).Cells
            current_cell.FormulaLocal = "=СУММ(" & current_cell.Offset(0, 1).Address(False, False) & _
                         ":" & current_cell.Offset(0, последняя_колонка - 2).Address(False, False) & ")"
        Next current_cell
        
        ' переносим исходные данные по составу изделия в массив для ускорения поиска
        все_изделия = Application.Transpose(Worksheets("Структура").Cells.Find("Изделие").CurrentRegion.Value2)
        
        ' для удобства сопоставляем колонкам имена
        For i = LBound(все_изделия, 1) To UBound(все_изделия, 1)
            заголовок = все_изделия(i, 1)
            Select Case заголовок
                Case "Вид элемента": виды = i
                Case "Изделие": изделия = i
                Case "Куда входит": родители = i
                Case "Кол.": количества = i
                Case "Общ. кол.": общие_количества = i
                Case "Масса изделия": массы = i
            End Select
        Next i
        
        ' построчный перебор всех изделий
        For i = начало_таблицы.Row + 2 To последняя_строка
            изделие = .Cells(i, начало_таблицы.Column).Value2
            
            общее_количество = 0
            For j = LBound(все_изделия, 2) + 2 To UBound(все_изделия, 2)
                If все_изделия(изделия, j) = изделие Then общее_количество = общее_количество + все_изделия(общие_количества, j)
            Next j
            
            тех_операции = range(.Cells(i, начало_таблицы.Column + 2), .Cells(i, последняя_колонка)).Value2 ' для ускорения перебора
            For j = LBound(тех_операции, 2) To UBound(тех_операции, 2)
                If (Not IsEmpty(тех_операции(1, j))) Then
                    формула = "=" & тех_операции(1, j) & "/ЕСЛИ(Общие_или_заединицу=""Общие трудозатраты"";1;" & IIf(общее_количество > 0, общее_количество, 1) & ")"
                    If j = пробивная Then формула = формула & "*ЕСЛИ(Пробивка=""Стандартная пробивная"";1;3.5)"
                    тех_операции(1, j) = формула
                End If
            Next j
            range(.Cells(i, начало_таблицы.Column + 2), .Cells(i, последняя_колонка)).FormulaLocal = тех_операции    ' записываем обратно
        Next i
      
      .Cells(2, 1) = "Корректировка проведена"
        
      .Activate
      With ActiveWindow
          .SplitColumn = начало_таблицы.Column
          .SplitRow = начало_таблицы.Row + 1
          .FreezePanes = True
      End With
    
    End With
    
Выход:
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    End

End Sub

Public Function Стт() As String
    Application.Volatile
    Стт = UserForm.TextBox1.Text
End Function

