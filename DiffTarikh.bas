Attribute VB_Name = "Module1"
' ************************************************** ***********
Public Function Rooz(F_Date As Long) As Byte
'«Ì‰  «»⁄ ⁄œœ „—»Êÿ »Â —Ê“ Ìﬂ  «—ÌŒ —« »—„ê—œ«‰œ
Rooz = F_Date Mod 100
End Function
'*******************************************
Function Mah(F_Date As Long) As Byte
'«Ì‰  «»⁄ ⁄œœ „—»Êÿ »Â „«Â Ìﬂ  «—ÌŒ —« »—„ê—œ«‰œ
Mah = Int((F_Date Mod 10000) / 100)
End Function
'*******************************************
Public Function Sal(F_Date As Long) As Integer
'«Ì‰  «»⁄ ⁄œœ „—»Êÿ »Â ”«· Ìﬂ  «—ÌŒ —« »—„ê—œ«‰œ
Sal = Int(F_Date / 10000)
End Function
'*******************************************
Public Function Kabiseh(ByVal OnlySal As Variant) As Byte
'Ê—ÊœÌ  «»⁄ ⁄œœ œÊ—ﬁ„Ì «” 
'«Ì‰  «»⁄ ﬂ»Ì”Â »Êœ‰ ”«· —« »—„Ìê—œ«‰œ
'«ê— ”«· ﬂ»Ì”Â »«‘œ ⁄œœ Ìﬂ Ê œ—€Ì— «Ì‰’Ê—  ’›— —« »— „Ìê—œ«‰œ
Kabiseh = 0
If OnlySal >= 1375 Then
If (OnlySal - 1375) Mod 4 = 0 Then
Kabiseh = 1
Exit Function
End If
ElseIf OnlySal <= 1370 Then
If (1370 - OnlySal) Mod 4 = 0 Then
Kabiseh = 1
Exit Function
End If
End If

End Function
'*******************************************
Public Function AddDay(ByVal F_Date As Long, ByVal Add As Integer) As Long
Dim K, M, R, Days As Byte
Dim s As Integer
R = Rooz(F_Date)
M = Mah(F_Date)
s = Sal(F_Date)
K = Kabiseh(s)

' »œÌ· —Ê“ »Â ⁄œœ 1 ÃÂ  «œ«„Â „Õ«”»«  Ê Ì« « „«„ „Õ«”»Â
Days = MahDays(s, M)
If Add > Days - R Then
Add = Add - (Days - R + 1)
R = 1
If M < 12 Then
M = M + 1
Else
M = 1
s = s + 1
End If
Else
R = R + Add
Add = 0
End If

While Add > 0
K = Kabiseh(s) 'ﬂ»Ì”Â: 1 Ê €Ì— ﬂ»Ì”Â: 0
Days = MahDays(s, M) ' ⁄œ«œ —Ê“Â«Ì „«Â ›⁄·Ì
Select Case Add
Case Is < Days
'«ê—  ⁄œ«œ —Ê“Â«Ì «›“Êœ‰Ì ﬂ„ — «“ Ìﬂ „«Â »«‘œ
R = R + Add
Add = 0
Case Days To IIf(K = 0, 365, 366) - 1
'«ê—  ⁄œ«œ —Ê“Â«Ì «›“Êœ‰Ì »Ì‘ — «“ Ìﬂ „«Â Ê ﬂ„ — «“ Ìﬂ ”«· »«‘œ
Add = Add - Days
If M < 12 Then
M = M + 1
Else
s = s + 1
M = 1
End If
Case Else
'«ê—  ⁄œ«œ —Ê“Â«Ì «›“Êœ‰Ì »Ì‘ — «“ Ìﬂ ”«· »«‘œ
s = s + 1
Add = Add - IIf(K = 0, 365, 366)
End Select
Wend
'AddDay = (s * 10000) + (M * 100) + (R)
AddDay = CLng(s & Format(M, "00") & Format(R, "00"))
End Function

'***********************************************
Public Static Function Shamsi() As Long
' «—ÌŒ Ã«—Ì ”Ì” „ —« »Â  «—ÌŒ ÂÃ—Ì ‘„”Ì  »œÌ· „Ì ﬂ‰œ
Dim Shamsi_Mabna As Long
Dim Miladi_mabna As Date
Dim Dif As Long
'œ— «Ì‰Ã« 78/10/11 »« 2000/01/01 „⁄«œ· ﬁ—«—œ«œÂ ‘œÂ
Shamsi_Mabna = 13781011
Miladi_mabna = #1/1/2000#
Dif = DateDiff("d", Miladi_mabna, Date)
If Dif < 0 Then
MsgBox " «—ÌŒ Ã«—Ì ”Ì” „ ‘„« ‰«œ—”  «”  , ¬‰—« «’·«Õ ﬂ‰Ìœ."
Else
Shamsi = AddDay(Shamsi_Mabna, Dif)
End If
End Function
Function MahDays(ByVal Sal As Integer, ByVal Mah As Byte) As Byte
'«Ì‰  «»⁄  ⁄œ«œ —Ê“Â«Ì Ìﬂ „«Â —« »—„Ì ê—œ«‰œ
Select Case Mah
Case 1 To 6
MahDays = 31
Case 7 To 11
MahDays = 30
Case 12
If Kabiseh(Sal) = 1 Then
MahDays = 30
Else
MahDays = 29
End If
End Select

End Function

'***********************************************
Public Function Diff2(ByVal FromDate As Long, ByVal To_Date As Long) As Long
Dim S1, M1, r1, S2, M2, r2, rs, rm, rr As Integer

If FromDate = 0 Or IsNull(FromDate) = True Or To_Date = 0 Or IsNull(To_Date) = True Or FromDate > To_Date Then
Diff2 = 0
Exit Function
End If

r1 = Rooz(FromDate)
M1 = Mah(FromDate)
S1 = Sal(FromDate)
r2 = Rooz(To_Date)
M2 = Mah(To_Date)
S2 = Sal(To_Date)
'--------------------------------------------------------------------------------------
rr = r2 - r1
rm = M2 - M1
rs = S2 - S1
'--------------------------------------------------------------------------------------
If rr < 0 Then
    If M2 > 1 Then
        rm = rm - 1
        rr = MahDays(S2, M2 - 1) + rr
    Else
        rm = 12
        rs = rs - 1
        rr = MahDays(S2 - 1, 12) + rr
    End If
End If

If rm < 0 Then
    rs = rs - 1
    rm = 12 + rm
End If

Diff2 = (rs * 100 + rm) * 100 + rr
End Function

'***********************************************
Public Function muldate(ByVal F_Date As Long) As Long
Dim S1, M1, r1, rs, rm, rr As Integer

r1 = Rooz(F_Date)
M1 = Mah(F_Date)
S1 = Sal(F_Date)

'--------------------------------------------------------------------------------------
rr = 2 * r1
rm = 2 * M1
rs = 2 * S1
'--------------------------------------------------------------------------------------
If rr > 30 Then
    rr = rr - 30
    rm = rm + 1
End If

If rm > 12 Then
    rs = rs + 1
    rm = rm - 12
End If

muldate = (rs * 100 + rm) * 100 + rr
End Function

