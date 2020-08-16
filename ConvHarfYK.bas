Attribute VB_Name = "Module1"
Sub ConvHarfYK()
Dim ws As Worksheet
    
    For Each ws In Worksheets
       
        ws.Name = Replace(ws.Name, ChrW(1610), ChrW(1740), vbTextCompare)
        ws.Name = Replace(ws.Name, ChrW(1603), ChrW(1705), vbTextCompare)
    Next ws
         
        
    Cells.Replace What:=ChrW(1610), Replacement:=ChrW(1740), LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Cells.Replace What:=ChrW(1603), Replacement:=ChrW(1705), LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub
