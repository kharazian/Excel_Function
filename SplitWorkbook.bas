Attribute VB_Name = "Module1"
Sub start()
    SplitWorkbook
End Sub
Sub SplitWorkbook(Optional colLetter As String, Optional SavePath As String)
If colLetter = "" Then colLetter = "A"
Dim lastValue As String
Dim hasHeader As Boolean
Dim wb As Workbook
Dim c As Range
Dim currentRow As Long
hasHeader = True 'Indicate true or false depending on if sheet  has header row.

If SavePath = "" Then SavePath = ThisWorkbook.Path
'Sort the workbook.
'ThisWorkbook.Worksheets
ThisWorkbook.Worksheets(1).Sort.SortFields.Add Key:=Range(colLetter & ":" & colLetter), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ThisWorkbook.Worksheets(1).Sort
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

For Each c In ThisWorkbook.Sheets(1).Range("A:A")
    If c.Value = "" Then Exit For
    If c.Row = 1 And hasHeader Then
    Else
        If lastValue <> c.Value Then
            If Not (wb Is Nothing) Then
                wb.SaveAs SavePath & "\" & lastValue & ".xls"
                wb.Close
            End If
            lastValue = c.Value
            currentRow = 1
            Set wb = Application.Workbooks.Add
            ThisWorkbook.Sheets(1).Rows(currentRow & ":" & currentRow).Copy
            wb.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Select
            wb.Sheets(1).Paste
            currentRow = currentRow + 1
        End If
        'ThisWorkbook.Sheets(1).Rows(c.Row & ":" & c.Row).Copy
        Dim lastrow As Long
        lastrow = wb.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
        ThisWorkbook.Sheets(1).Rows(c.Row & ":" & c.Row).Copy Destination:=Sheets(1).Range(currentRow & ":" & currentRow)
        currentRow = currentRow + 1
    End If
Next
If Not (wb Is Nothing) Then
    wb.SaveAs SavePath & "\" & lastValue & ".xls"
    wb.Close
End If
End Sub







