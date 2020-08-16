Attribute VB_Name = "Module1"
Function Numeric_Check(Text)
Application.Volatile

For i = 1 To Len(Text)
  
  My_Str = Mid(Text, i, 1)
  If My_Str = 0 Or Val(My_Str) > 0 Then
    Numeric_Check = True
  Else
    Numeric_Check = False
    Exit Function
  End If

Next i

End Function
Function Code_Melli_Check(Code_Melli)

Dim Phase_1, Phase_2, Phase_3


If Numeric_Check(Code_Melli) = "True" And Len(Code_Melli) = 10 Then
        
    Phase_1 = (Mid(Code_Melli, 1, 1) * 10) + (Mid(Code_Melli, 2, 1) * 9) + (Mid(Code_Melli, 3, 1) * 8) + _
    (Mid(Code_Melli, 4, 1) * 7) + (Mid(Code_Melli, 5, 1) * 6) + (Mid(Code_Melli, 6, 1) * 5) + _
    (Mid(Code_Melli, 7, 1) * 4) + (Mid(Code_Melli, 8, 1) * 3) + (Mid(Code_Melli, 9, 1) * 2)
'     MsgBox "Phase1   " & Phase_1
        
        Phase_2 = Phase_1 - (Int(Phase_1 / 11) * 11)
'        MsgBox "Phase2_1   " & Phase_2
            
            If Phase_2 > 1 Then
                Phase_2 = 11 - Phase_2
            Else
                Phase_2 = Phase_2
            End If
'            MsgBox "Phase2_2   " & Phase_2

                Phase_3 = Mid(Code_Melli, 10, 1)
'                MsgBox "Phase3   " & Phase_2
                    
                    If Val(Phase_2) = Val(Phase_3) Then
                        Code_Melli_Check = "True"
                    Else
                        Code_Melli_Check = "False"
                    End If
'                    MsgBox "Phase3-phase2   " & Phase_3 - Phase_2
Else
   Code_Melli_Check = "Melli Code Error"
End If

End Function
