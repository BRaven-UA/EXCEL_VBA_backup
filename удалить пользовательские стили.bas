Sub StyleKiller()
    Dim N As Long, i As Long

    With ActiveWorkbook
        N = .Styles.Count
        For i = N To 1 Step -1
            If Not .Styles(i).BuiltIn Then .Styles(i).Delete
        Next i
    End With
End Sub