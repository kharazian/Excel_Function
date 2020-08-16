Attribute VB_Name = "Module1"
Sub SplitWorkbook(Optional colLetter As String, Optional SavePath As String)
If colLetter = "" Then colLetter = "I"
Dim lastValue As String
Dim hasHeader As Boolean
Dim wsb As Worksheet
Dim ws As Worksheet
Dim c As Range
Dim currentRow As Long
hasHeader = True

Set wsb = ThisWorkbook.Worksheets(1)
wsb.Sort.SortFields.Add Key:=Range(colLetter & ":" & colLetter), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With wsb.Sort
    .SetRange Cells
    If hasHeader Then ' Was a header indicated?
        .Header = xlYes
    Else
        .Header = xlNo
    End If
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

For Each c In wsb.Range(colLetter & ":" & colLetter)
    If c.Value = "" Then Exit For
    If c.Row = 1 And hasHeader Then
    Else
        If lastValue <> c.Value Then
            lastValue = c.Value
            currentRow = 1
            Set ws = ThisWorkbook.Sheets.Add
            ws.Name = lastValue
            ws.DisplayRightToLeft = True
            wsb.Cells.Copy
            ws.Cells.PasteSpecial Paste:=xlPasteColumnWidths
            wsb.Rows(currentRow & ":" & currentRow).Copy
            ws.Cells(Rows.Count, 1).End(xlUp).Select
            ws.Paste
            currentRow = currentRow + 1
        End If
        'ThisWorkbook.Sheets(1).Rows(c.Row & ":" & c.Row).Copy
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        wsb.Rows(c.Row & ":" & c.Row).Copy Destination:=ws.Range(currentRow & ":" & currentRow)
        currentRow = currentRow + 1
    End If
Next

End Sub







