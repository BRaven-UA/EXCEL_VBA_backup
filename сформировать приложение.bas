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
    Range(Cells(6, 7), Cells(LastRow - 1, 7)).FormulaR1C1 = "=ROUND(" & OriginSheet & "!RC*_k_EXW_DDP_RU/1.2,0)"
    Range(Cells(6, 8), Cells(LastRow - 1, 8)).FormulaR1C1 = "=IF(ISNUMBER(RC[-1]),RC[-1]*RC[-3], """")"
    Application.CutCopyMode = False
    Range("H" & LastRow + 1).FormulaR1C1 = "=R[-1]C*0.2"
    Range("H" & LastRow + 2).FormulaR1C1 = "=R[-2]C+R[-1]C"
    Range(Cells(6, 7), Cells(LastRow + 2, 8)).Select
    Selection.NumberFormat = "#,##0$"
    Range(Cells(LastRow, 8), Cells(LastRow + 2, 8)).NumberFormat = "#,##0.00$"
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
End Sub