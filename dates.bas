Attribute VB_Name = "Module1"
'??? ?????? ????? ???? ??? ?? ??? ????? ?? ????? ?????? ?? ???? ????? ?????
Public strDate As String
'//////////////////////////////////////
'www.exceliha.ir
' 1- ????? ???? Number(Long) ??? ?? ????? Date ???????? ?? ??? ????
' 2- ??? ?????? ?? ????? 0000/00/00 ????? ???? InputMask ?????
' ????? 8 ???? ?? ??? ????? ???? ????? ? ??? ????? ?? ??? 1999 ?????? ????
' ...
' ????? ???? ????? ?? ?? ???? ???? ????? ?? ??? Shamsi() ????
' ???? ????? Now() ?? ?? ?????? ?? ??????? ???? ???? Dat() ????
' :???? ??????? ?? ???? ????? ??? ?? ???? ?? ???? ?????? ??? ??? ??????
' :???? ??? ???? ????? ValidationRule ?? ?? ????? ValidDate() ????
' ...
' ************************************************** ***********
Public Function Rooz(F_Date As Long) As Byte
'??? ???? ??? ????? ?? ??? ?? ????? ?? ?????????
Rooz = F_Date Mod 100
End Function
'*******************************************
Function Mah(F_Date As Long) As Byte
'??? ???? ??? ????? ?? ??? ?? ????? ?? ?????????
Mah = Int((F_Date Mod 10000) / 100)
End Function
'*******************************************
Public Function Sal(F_Date As Long) As Integer
'??? ???? ??? ????? ?? ??? ?? ????? ?? ?????????
Sal = Int(F_Date / 10000)
End Function
'*******************************************
Public Function Kabiseh(ByVal OnlySal As Variant) As Byte
'????? ???? ??? ?????? ???
'??? ???? ????? ???? ??? ?? ??????????
'??? ??? ????? ???? ??? ?? ? ????? ??????? ??? ?? ?? ????????
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
Function ValidDate(F_Date As Long) As Boolean
Dim M, s, R As Byte
' ??? ???? ?????? ?? ??? ????? ?? ?? ??? ????? ???? ???? ????? ?? ???
' ?? ???? ?????? False ???? ??????? ???? True ??? ????? ????? ????
ValidDate = True

s = Sal(F_Date)
M = Mah(F_Date)
R = Rooz(F_Date)
'********
If F_Date < 10000101 Then
ValidDate = False
Exit Function
End If

If M > 12 Or M = 0 Or R = 0 Then
ValidDate = False
Exit Function
End If

If R > MahDays(s, M) Then
ValidDate = False
Exit Function
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

'????? ??? ?? ??? 1 ??? ????? ??????? ? ?? ????? ??????
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
K = Kabiseh(s) '?????: 1 ? ??? ?????: 0
Days = MahDays(s, M) '????? ?????? ??? ????
Select Case Add
Case Is < Days
'??? ????? ?????? ??????? ???? ?? ?? ??? ????
R = R + Add
Add = 0
Case Days To IIf(K = 0, 365, 366) - 1
'??? ????? ?????? ??????? ????? ?? ?? ??? ? ???? ?? ?? ??? ????
Add = Add - Days
If M < 12 Then
M = M + 1
Else
s = s + 1
M = 1
End If
Case Else
'??? ????? ?????? ??????? ????? ?? ?? ??? ????
s = s + 1
Add = Add - IIf(K = 0, 365, 366)
End Select
Wend
'AddDay = (s * 10000) + (M * 100) + (R)
AddDay = CLng(s & Format(M, "00") & Format(R, "00"))
End Function

'***********************************************
Public Static Function Shamsi() As Long
'????? ???? ????? ?? ?? ????? ???? ???? ????? ?? ???
Dim Shamsi_Mabna As Long
Dim Miladi_mabna As Date
Dim Dif As Long
'?? ????? 78/10/11 ?? 2000/01/01 ????? ???????? ???
Shamsi_Mabna = 13781011
Miladi_mabna = #1/1/2000#
Dif = DateDiff("d", Miladi_mabna, Date)
If Dif < 0 Then
MsgBox "????? ???? ????? ??? ?????? ??? , ???? ????? ????."
Else
Shamsi = AddDay(Shamsi_Mabna, Dif)
End If
End Function
'***********************************************
Public Function DayWeek(F_Date As Long) As String
Dim a As String
Dim N As Byte
N = DayWeekNo(F_Date)
Select Case N
Case 0
a = "????"
Case 1
a = "??????"
Case 2
a = "??????"
Case 3
a = "???????"
Case 4
a = "????????"
Case 5
a = "????????"
Case 6
a = "????"
End Select
DayWeek = a
End Function

'***********************************************
Public Function Dat()
Dim d As Long
d = Shamsi
Dat = DayWeek(d) & Sal(d) & "/" & Mah(d) & "/" & Rooz(d)
End Function

'***********************************************
Public Function Diff(ByVal FromDate As Long, ByVal To_Date As Long) As Long
'??? ???? ????? ?????? ??? ?? ????? ?? ????? ?? ???
Dim Tmp As Long
Dim S1, M1, r1, S2, M2, r2 As Integer
Dim Sumation As Single
Dim Flag As Boolean
Flag = False
If FromDate = 0 Or IsNull(FromDate) = True Or To_Date = 0 Or IsNull(To_Date) = True Then
Diff = 0
Exit Function
End If

If FromDate > To_Date Then
'??? ????? ???? ?? ????? ????? ?????? ???? ???? ????? ????? ?? ????
Flag = True
Tmp = FromDate
FromDate = To_Date
To_Date = Tmp
End If
r1 = Rooz(FromDate)
M1 = Mah(FromDate)
S1 = Sal(FromDate)
r2 = Rooz(To_Date)
M2 = Mah(To_Date)
S2 = Sal(To_Date)
Sumation = 0

Do While S1 < S2 - 1 Or (S1 = S2 - 1 And (M1 < M2 Or (M1 = M2 And r1 <= r2)))
'??? ?? ??? ?? ????? ?????? ???
If Kabiseh((S1)) = 1 Then
If M1 = 12 And r1 = 30 Then
Sumation = Sumation + 365
r1 = 29
Else
Sumation = Sumation + 366
End If
Else
Sumation = Sumation + 365
End If
S1 = S1 + 1
Loop

Do While S1 < S2 Or M1 < M2 - 1 Or (M1 = M2 - 1 And r1 < r2)
'??? ?? ??? ?? ????? ?????? ???
Select Case M1
Case 1 To 6
If M1 = 6 And r1 = 31 Then
Sumation = Sumation + 30
r1 = 30
Else
Sumation = Sumation + 31
End If
M1 = M1 + 1
Case 7 To 11
If M1 = 11 And r1 = 30 And Kabiseh(S1) = 0 Then
Sumation = Sumation + 29
r1 = 29
Else
Sumation = Sumation + 30
End If
M1 = M1 + 1
Case 12
If Kabiseh(S1) = 1 Then
Sumation = Sumation + 30
Else
Sumation = Sumation + 29
End If
S1 = S1 + 1
M1 = 1
End Select
Loop

If M1 = M2 Then
Sumation = Sumation + (r2 - r1)
Else
Select Case M1
Case 1 To 6
Sumation = Sumation + (31 - r1) + r2
Case 7 To 11
Sumation = Sumation + (30 - r1) + r2
Case 12
If Kabiseh(S1) = 1 Then
Sumation = Sumation + (30 - r1) + r2
Else
Sumation = Sumation + (29 - r1) + r2
End If
End Select
End If

If Flag = True Then
Sumation = -Sumation
End If
Diff = Sumation
End Function

Public Function DayWeekNo(F_Date As Long) As String
'??? ???? ?? ????? ?? ?????? ???? ? ???? ?? ??? ?? ???? ?? ???? ???
'??? ???? ???? ??? 0
'??? 1???? ???? ??? 1
'......
'??? ???? ???? ??? 6
Dim day As String
Dim Shmsi_Mabna As Long
Dim Dif As Long
'???? 80/10/11
Shmsi_Mabna = 13801011
Dif = Diff(Shmsi_Mabna, F_Date)
If Shmsi_Mabna > F_Date Then
Dif = -Dif
End If
'?? ???? ?? ????? 80/10/11 3???? ??? ?????? ????? day ?????
day = (Dif + 3) Mod 7
If day < 0 Then
DayWeekNo = day + 7
Else
DayWeekNo = day
End If
End Function


Function MahName(ByVal Mah_no As Byte) As String
Select Case Mah_no
Case 1
MahName = "???????"
Case 2
MahName = "????????"
Case 3
MahName = "?????"
Case 4
MahName = "???"
Case 5
MahName = "?????"
Case 6
MahName = "??????"
Case 7
MahName = "???"
Case 8
MahName = "????"
Case 9
MahName = "???"
Case 10
MahName = "??"
Case 11
MahName = "????"
Case 12
MahName = "?????"
End Select
End Function

Function SalMah(ByVal F_Date As Long) As Long
'?? ??? ??? ????? ?? ???? ??? ? ??? ??? ?? ???? ??????
SalMah = Val(Left$(F_Date, 6))
End Function

Function MahDays(ByVal Sal As Integer, ByVal Mah As Byte) As Byte
'??? ???? ????? ?????? ?? ??? ?? ???? ??????
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

Function Make_Date(ByVal F_Date As Long) As String
'?? ????? ?? ????? ?? ???? 10 ???? ?? ??? ???? ??? ???? ??? ????? ?? ???
Dim d As String
d = Trim(Str(F_Date))
If IsNull(F_Date) = True Or F_Date = 0 Then
Make_Date = ""
Else
Make_Date = Mid(d, 1, 4) & "/" & Mid(d, 5, 2) & "/" & Mid(d, 7, 2)
End If
End Function

Function NextMah(ByVal Sal_Mah As Long) As Long
If (Sal_Mah Mod 100) = 12 Then
NextMah = (Int(Sal_Mah / 100) + 1) * 100 + 1
Else
NextMah = Sal_Mah + 1
End If
End Function

Function PreviousMah(ByVal Sal_Mah As Long) As Long
If (Sal_Mah Mod 100) = 1 Then
PreviousMah = (Int(Sal_Mah / 100) - 1) * 100 + 12
Else
PreviousMah = Sal_Mah - 1
End If
End Function


Function SubtractDay(ByVal F_Date As Long, ByVal Subtract As Long) As Long
'?? ????? ??? ????? ?? ?? ????? ?? ???? ? ????? ????? ?? ????? ?????
Dim K, M, s, R, Days As Byte

R = Rooz(F_Date)
M = Mah(F_Date)
s = Sal(F_Date)
K = Kabiseh(s)

'????? ??? ?? ??? 1 ??? ????? ??????? ? ?? ????? ??????
If Subtract >= R - 1 Then
Subtract = Subtract - (R - 1)
R = 1
Else
R = R - Subtract
Subtract = 0
End If

While Subtract > 0
K = Kabiseh(s - 1) '?????: 1 ? ??? ?????: 0
Days = MahDays(IIf(M >= 2, s, s - 1), IIf(M >= 2, M - 1, 12)) '????? ?????? ??? ????
Select Case Subtract
Case Is < Days
'??? ????? ?????? ???? ???? ?? ?? ??? ????
R = Days - Subtract + 1
Subtract = 0
If M >= 2 Then
M = M - 1
Else
s = s - 1
M = 12
End If
Case Days To IIf(K = 0, 365, 366) - 1
'??? ????? ?????? ???? ????? ?? ?? ??? ? ???? ?? ?? ??? ????
Subtract = Subtract - Days
If M >= 2 Then
M = M - 1
Else
s = s - 1
M = 12
End If
Case Else
'??? ????? ?????? ???? ????? ?? ?? ??? ????
s = s - 1
Subtract = Subtract - IIf(K = 0, 365, 366)
End Select
Wend
SubtractDay = (s * 10000) + (M * 100) + (R)

End Function


'????? ????? ??? ???
Public Function Firstday(Sal As Integer, Mah As Byte) As Long
Dim strfd As Long
strfd = Sal & Format(Mah, "00") & Format(1, "00")
Firstday = DayWeekNo(strfd)
End Function
