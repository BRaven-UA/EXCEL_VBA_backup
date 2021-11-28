Attribute VB_Name = "Module1"
Sub лист_соответствия()
    ThisWorkbook.IsAddin = Not ThisWorkbook.IsAddin
    If ThisWorkbook.IsAddin Then ThisWorkbook.Save
End Sub


Sub Сформировать_маршрутный_лист()
    ' процедура рассчитана на запуск из листа с технологической схемой, полученной путем экспорта из 1С:PLM + ERP
    ' для корректной работы также требуется открытая книга с трудозатратами, полученная оттуда же
    ' (C) Синица Анатолий 2020, MIT license :)
    On Error Resume Next
    Application.ScreenUpdating = False  ' Отключаем обновление экрана (для повышения производительности)
    calc = Application.Calculation
    Application.Calculation = xlCalculationManual   ' Устанавливаем ручной пересчет ячеек (для повышения производительности)

    название_маршрутного_листа = "Маршрутный лист"    ' название листа по-умолчанию

    ' проверяем существование такого листа
    If WorksheetExists(название_маршрутного_листа) Then
        result = MsgBox("Такой лист уже существует. Переименуйте или удалите его перед формированием нового листа", vbCritical + vbOKOnly)
        Exit Sub
    End If

    ' копируем первый лист (вводные данные) и переименовываем
    Worksheets(1).Copy after:=Worksheets(Worksheets.Count)
    Dim маршрутный_лист As Worksheet
    Dim лист_трудозатрат As Worksheet
    Dim лист_соответствия As Worksheet
    Set маршрутный_лист = Worksheets(Worksheets.Count)
    маршрутный_лист.UsedRange.UnMerge  ' убираем объединение ячеек для корректной работы
    маршрутный_лист.name = название_маршрутного_листа
    
    If ActiveWindow.TabRatio = 0 Then ActiveWindow.TabRatio = 0.5   ' сдвигаем горизонтальную прокрутку чтобы видеть листы
    
    Dim номенклатура As Range, технология As Range, вид_воспроизводства As Range, количество As Range, основной_материал As Range, краска As Range, техоперация As Range    ' перечень необходимых полей
    ' ищем необходимые заголовки таблицы и делаем на них ссылки (первый аргумент функции передается по ссылке)
    Call FindCell(номенклатура, "Номенклатура", маршрутный_лист)
    Call FindCell(технология, "Технология", маршрутный_лист)
    Call FindCell(вид_воспроизводства, "Вид воспроизводства", маршрутный_лист)
    Call FindCell(количество, "Кол. (норма)", маршрутный_лист)
    Call FindCell(техоперация, "Характеристика", маршрутный_лист)
    
    ' добавляем колонки с основным материалом и краской
    Set основной_материал = техоперация.End(xlToRight).Offset(0, 1)
    Set краска = техоперация.End(xlToRight).Offset(0, 2)
    техоперация.Copy
    основной_материал.Resize(1, 2).PasteSpecial (xlPasteFormats) ' переносим форматирование на две ячейки
    основной_материал.Value = "Основной материал"
    краска.Value = "Краска"
    
    ' перемещаем колонку с характеристиками в самый конец
    техоперация.EntireColumn.Cut
    техоперация.End(xlToRight).Offset(0, 1).EntireColumn.Insert Shift:=xlToRight

    ' создаем новые заголовки для техопераций
    Set заголовок_для_копирования = техоперация.Resize(1, 3)
    техоперация.Value = "Техоперация1"    ' переименовываем характеристику в техоперацию
    техоперация.AutoFill Destination:=заголовок_для_копирования    ' копируем ячейку на две соседние ячейки
    техоперация.Offset(0, 1).Value = "Норма на 1 шт, мин"
    техоперация.Offset(0, 2).Value = "Норма на партию, мин"
    
    Set изделие = номенклатура.End(xlDown) ' название изделия берем из первой номенклатуры в списке
    маршрутный_лист.Cells(изделие.Row, количество.Column) = 1 ' принудительно ставим количество изделий в единицу
    
    ' удаляем покупные изделия и промежуточные строки
    For Row = маршрутный_лист.UsedRange.Rows.Count To номенклатура.Row + 1 Step -1 ' перебираем с конца чтобы при удалении строки не сбивался итератор
        ' если текущая номенклатура является краской, прописывем её в ближайшую технологию
        If InStr(маршрутный_лист.Cells(Row, номенклатура.Column), "Эмаль") > 0 Then _
            маршрутный_лист.Cells(маршрутный_лист.Cells(Row, технология.Column).End(xlUp).Row, краска.Column) = маршрутный_лист.Cells(Row, номенклатура.Column)
        ' если текущая номенклатура является основным материалом, прописывем его в ближайшую технологию
        If InStr(маршрутный_лист.Cells(Row - 1, 1).Text, "Основных материалов") > 0 Then _
            маршрутный_лист.Cells(маршрутный_лист.Cells(Row, технология.Column).End(xlUp).Row, основной_материал.Column) = маршрутный_лист.Cells(Row, номенклатура.Column)
        If маршрутный_лист.Cells(Row, вид_воспроизводства.Column) <> "Производство" Then маршрутный_лист.Rows(Row).EntireRow.Delete
    Next Row

    For Each wb In Application.Workbooks    ' перебор всех открытых книг
        If InStr(wb.name, ".xls") > 0 And Not (wb Is ActiveWorkbook) Then   ' интересуют только обычные книги, за исключением текущей
            ' если в активном листе книги есть упоминание нашего изделия, то считаем что это лист с трудозатратами
            If Not (wb.ActiveSheet.UsedRange.Find(what:=изделие.Text, lookat:=xlWhole, LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlNext) Is Nothing) Then
                Set лист_трудозатрат = wb.ActiveSheet
                лист_трудозатрат.UsedRange.UnMerge  ' убираем объединение ячеек для корректной работы
                Set лист_соответствия = ThisWorkbook.Worksheets("Лист соответствия")    ' лист соответствия названий характеристик с техоперациями. Находится в самой настройке и скрыт по-умолчанию. Для отображения использовать процедуру "лист_соответствия"
                
                Set таблица_соответствия = CreateObject("Scripting.Dictionary") ' для удобства будем хранить данные из листа в словаре
                Set количество_изделий = CreateObject("Scripting.Dictionary") ' сюда занесем суммарное количество каждого изделия
                
                Set характеристика_УПП = лист_соответствия.UsedRange.Find(what:="Характеристика из УПП", lookat:=xlWhole, LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlNext)
                Set техоперация_PLM = лист_соответствия.UsedRange.Find(what:="Техоперация из PLM", lookat:=xlWhole, LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlNext)
                
                первая_строка = характеристика_УПП.Row + 1
                последняя_строка = лист_соответствия.UsedRange.Rows.Count
                For Row = первая_строка To последняя_строка
                    ' заполняем словарь. Ключ - характеристика УПП, значение - техоперация PLM
                    таблица_соответствия.Add лист_соответствия.Cells(Row, характеристика_УПП.Column).Text, лист_соответствия.Cells(Row, техоперация_PLM.Column)
                Next Row
                
                ' перебираем технологическую схему и выставляем трудозатраты рядом с техоперацией
                    ' ссылка на начало таблицы трудозатрат
                    Set первая_ячейка = лист_трудозатрат.UsedRange.Find(what:="Изделие", lookat:=xlWhole, LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlNext)
                    
                    ' диапазон ячеек с названиями техопераций
                    Set техоперации = первая_ячейка.Resize(1, лист_трудозатрат.UsedRange.Columns.Count)
                    
                    ' диапазон ячеек с названиями изделий
                    Set изделия = первая_ячейка.Resize(лист_трудозатрат.UsedRange.Rows.Count, 1)
                    
                    колонка_трудозатрат = техоперация.Column + 1 ' колонка, в которую будем записывать трудозатраты
                    
                    ' перебираем все строки с изделиями для подсчета суммарного количества каждого изделия с техоперацией
                    For Row = номенклатура.Row + 1 To маршрутный_лист.UsedRange.Rows.Count
                        текущее_изделие = Без_СБ(маршрутный_лист.Cells(Row, номенклатура.Column).Text)
                        'последняя_техоперация = технология.EntireColumn.Find(what:=текущее_изделие, after:=маршрутный_лист.Cells(Row, технология.Column), SearchDirection:=xlPrevious).Row
                            ' из-за неточностей с названиями изделий и технологий приходится делать это:
                            For r = Row To номенклатура.Row + 1 Step -1
                                If текущее_изделие = Без_СБ(маршрутный_лист.Cells(r, технология.Column).Text) Then
                                    последняя_техоперация = r
                                    Exit For
                                End If
                            Next r
                        количество_в_партии = IIf(Row = последняя_техоперация, 1, маршрутный_лист.Cells(последняя_техоперация, количество.Column).Value)
                        текущая_техоперация = маршрутный_лист.Cells(Row, техоперация.Column).Text
                        текущее_количество = маршрутный_лист.Cells(Row, количество.Column).Value
                        изделие_с_техоперацией = текущее_изделие + текущая_техоперация
    
                        ' отдельно создавать записи не нужно, так как метод Item создает новую запись, если не находит указанный ключ
                        количество_изделий.Item(изделие_с_техоперацией) = количество_изделий.Item(изделие_с_техоперацией) + текущее_количество * количество_в_партии
                     Next Row
    
                     ' перебираем все строки с изделиями для заполнения трудозатрат
                     For Row = номенклатура.Row + 1 To маршрутный_лист.UsedRange.Rows.Count
                        Set текущее_изделие = маршрутный_лист.Cells(Row, номенклатура.Column)
                        текущая_технология = Без_СБ(маршрутный_лист.Cells(Row, технология.Column).Text)
                        текущая_техоперация = маршрутный_лист.Cells(Row, техоперация.Column).Text
                        текущее_количество = маршрутный_лист.Cells(Row, количество.Column).Value
                        изделие_с_техоперацией = Без_СБ(текущее_изделие.Text) + текущая_техоперация
                        текущие_трудозатраты = лист_трудозатрат.Cells(изделия.Find(текущее_изделие.Text).Row, техоперации.Find(таблица_соответствия.Item(текущая_техоперация).Text).Column).Value
                        
                        Set ячейка_трудозатрат = маршрутный_лист.Cells(Row, колонка_трудозатрат)
    
                        If текущие_трудозатраты = 0 Then
                            ячейка_трудозатрат.Value = 0    ' ноль нужен для корректной работы метода "end" дальше по коду
                            ячейка_трудозатрат.Interior.Color = RGB(255, 0, 0) ' если не найдено трудозатрат красим ячейку в красный цвет
                        Else
                            ' вычисляем трудозатраты на единицу изделия
                            ячейка_трудозатрат.Value = Round(текущие_трудозатраты / количество_изделий.Item(изделие_с_техоперацией), 3)
                            ячейка_трудозатрат.Offset(0, -1).Interior.Color = таблица_соответствия.Item(текущая_техоперация).Interior.Color
                        End If
                        
                        ' добавляем формулу общего количества на партию
                        ячейка_трудозатрат.Offset(0, 1).FormulaR1C1Local = "=RC" + CStr(количество.Column) + "*RC[-1]"
                    Next Row
                    
                техоперация.Resize(1, 3).EntireColumn.HorizontalAlignment = xlCenter    ' центрируем рабочие колонки

                ' переносим техоперации в один ряд для каждого изделия
                For Row = маршрутный_лист.UsedRange.Rows.Count To номенклатура.Row + 1 Step -1
                    Set текущее_изделие = маршрутный_лист.Cells(Row, номенклатура.Column)
                    
                    If Без_СБ(маршрутный_лист.Cells(Row, технология.Column).Text) <> Без_СБ(текущее_изделие.Text) Then ' это не последняя техоперация для изделия
                        Set рабочий_диапазон = маршрутный_лист.Range(маршрутный_лист.Cells(Row, техоперация.Column), маршрутный_лист.Cells(Row, техоперация.Column).End(xlToRight))   ' диапазон всех техопераций и трудозатрат в этом ряду
                        ' ищем последнюю техоперацию выше по спецификации
                        'Set последняя_техоперация = технология.EntireColumn.Find(what:=Без_СБ(текущее_изделие.Text), after:=маршрутный_лист.Cells(Row, технология.Column), lookat:=xlWhole, SearchDirection:=xlPrevious)
                            ' из-за неточностей с названиями изделий и технологий приходится делать это:
                            For r = Row To номенклатура.Row + 1 Step -1
                                If Без_СБ(текущее_изделие.Text) = Без_СБ(маршрутный_лист.Cells(r, технология.Column).Text) Then
                                    Set последняя_техоперация = маршрутный_лист.Cells(r, технология.Column)
                                    Exit For
                                End If
                            Next r
                        рабочий_диапазон.Copy  ' используем "copy" чтобы копировать вместе с форматами
                        ' вставляем в начало рабочего диапазона последней техоперации
                        маршрутный_лист.Cells(последняя_техоперация.Row, техоперация.Column).Insert Shift:=xlShiftToRight
                        маршрутный_лист.Rows(Row).EntireRow.Delete
                        'маршрутный_лист.Rows(Row).Hidden = True
                   End If
                Next Row
                
                ' удаляем ненужные столбцы
                технология.EntireColumn.Delete
                'маршрутный_лист.Columns(1).EntireColumn.Delete
                номенклатура.EntireColumn.Delete
                вид_воспроизводства.EntireColumn.Delete

                ' форматируем итоговую таблицу
                Set все_рабочие_заголовки = заголовок_для_копирования.Resize(1, маршрутный_лист.UsedRange.Columns.Count - заголовок_для_копирования.Column + 1)
                заголовок_для_копирования.AutoFill Destination:=все_рабочие_заголовки   ' заполняем заголовки для всех техопераций
                все_рабочие_заголовки.ColumnWidth = 5   ' ставим минимальную ширину столбцов
                Set заголовки = маршрутный_лист.UsedRange.Resize(1) ' оставляем только первую строку
                заголовки.WrapText = True ' перенос по словам
                заголовки.EntireRow.AutoFit   ' корректировка высоты
                маршрутный_лист.Columns.AutoFit ' атоматически подгоняем ширину столбцов
                маршрутный_лист.UsedRange.Borders.LineStyle = xlContinuous  ' добавляем границы для всей таблицы
                количество.EntireColumn.NumberFormat = "General"    ' корректируем формат чисел для удобства восприятия
                
                'маршрутный_лист.Copy after:=Worksheets(Worksheets.Count)    ' дублируем лист с группировками без фильтра
                'маршрутный_лист.Cells.ClearOutline  ' убираем все группировки
                маршрутный_лист.UsedRange.AutoFilter    ' применяем фильтр ко всей таблице
                
            End If
        End If
    Next wb

    ' восстанавливаем исходные значения
    Application.Calculation = calc
    Application.Calculate   ' на случай если стоял ручной пересчет
    Application.ScreenUpdating = True
    
End Sub

' проверяет существование листа с заданным именем
' https://stackoverflow.com/a/6688482
Function WorksheetExists(ByVal shtName As String) As Boolean
    Dim sht As Worksheet
    On Error Resume Next
    Set sht = ActiveWorkbook.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

' поиск ячейки с текстом среди всех ячеек листа и прекращение работы при неудаче
Sub FindCell(ByRef cell As Range, ByVal name As String, ByRef sheet As Worksheet)
    Set cell = sheet.UsedRange.Find(what:=name, lookat:=xlWhole, LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If cell Is Nothing Then
        result = MsgBox("Поле " + name + " не найдено. Проверьте правильность исходной таблицы", vbInformation + vbOKOnly)
        End ' прекращаем работу после первой ошибки
    End If
End Sub

Function Без_СБ(str As String) As String
    Без_СБ = Replace(str, " СБ", "")
End Function


Sub temp()

End Sub


