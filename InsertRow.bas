Attribute VB_Name = "Module1"
Sub InsertRowsAtIntervals()
'Updateby20150707
Dim Rng As Range
Dim xInterval As Integer
Dim xRows As Integer
Dim xRowsCount As Integer
Dim xNum1 As Integer
Dim xNum2 As Integer
Dim WorkRng As Range
Dim xWs As Worksheet
xTitleId = "AmirAkh"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
Set WorkRngFooter = Application.InputBox("Footer Range", xTitleId, WorkRng.Address, Type:=8)
xRowsCount = WorkRng.Rows.Count
xInterval = Application.InputBox("Enter row interval. ", xTitleId, 1, Type:=1)
xRows = WorkRngFooter.Rows.Count
xNum1 = WorkRng.Row + xInterval
xNum2 = xRows + xInterval
Set xWs = WorkRng.Parent
For i = 1 To Int(xRowsCount / xInterval)
    WorkRngFooter.Copy
    xWs.Range(xWs.Cells(xNum1, WorkRng.Column), xWs.Cells(xNum1 + xRows - 1, WorkRng.Column)).Select
    Application.Selection.EntireRow.Insert
    xWs.Range(xWs.Cells(xNum1, WorkRng.Column), xWs.Cells(xNum1 + xRows - 1, WorkRng.Column)).PasteSpecial Paste:=xlPasteAll
    xNum1 = xNum1 + xNum2
Next
End Sub
