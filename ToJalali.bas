Attribute VB_Name = "Module1"

'tabdil as date object be arraye jalali
Public Function toJalaaliFromDateObject(gDate As Date)
    Dim jalaliDateArray1
    jalaliDateArray1 = toJalaali(Year(gDate), Month(gDate), day(gDate))

   toJalaaliFromDateObject = jalaliDateArray1(0) & "-" & Right("0" & jalaliDateArray1(1), 2) & "-" & Right("0" & jalaliDateArray1(2), 2)
End Function
' tabdild to object date gregorian
Private Function toGregorianDateObject(jy As Long, jm As Long, jd As Long)
    Dim result
    result = toGregorian(jy, jm, jd)
    toGregorianDateObject = DateValue(result(0) & "-" & result(1) & "-" & result(2))
End Function
'tabdil jalali be miladi ba daryaf sale , mahe , rooz jalali
' yek arraye barmigardanad ke index(0)= sal, index(1)=mah, index(2)=rooz
Private Function toGregorian(jy As Long, jm As Long, jd As Long)
      toGregorian = d2g(j2d(jy, jm, jd))
End Function



' tabdile tarikh miladi be jalali ba daryaft sale, mah , rooz miladi
' yek arraye barmigardanad ke index(0)= sal, index(1)=mah, index(2)=rooz
Private Function toJalaali(gy As Long, gm As Long, gd As Long)
    toJalaali = d2j(g2d(gy, gm, gd))
End Function
  




' check valid bodan jalali date
Private Function isValidJalaaliDate(jy As Long, jm As Long, jd As Long)
    isValidJalaaliDate = jy >= -61 And jy <= 3177 And jm >= 1 And jm <= 12 And jd >= 1 And jd <= jalaaliMonthLength(jy, jm)
End Function
  

' tedade rooz haye mah ra baraye sale , mahe jalali bar migardanad
Private Function jalaaliMonthLength(jy As Long, jm As Long)
    If (jm <= 6) Then
        jalaaliMonthLength = 31
        Exit Function
    End If
    If (jm <= 11) Then
     jalaaliMonthLength = 30
     Exit Function
    End If
    If (isLeapJalaaliYear(jy)) Then
        jalaaliMonthLength = 30
        Exit Function
        
    End If
  jalaaliMonthLength = 29
End Function
' check in ke sale jalali kabise as ya na
Private Function isLeapJalaaliYear(jy As Long)
    Dim leap As Long
    leap = jalCal(jy)(0)
    
    If (leap = 0) Then
        isLeapJalaaliYear = True
    Else
        isLeapJalaaliYear = False
    End If
    
End Function





 

' function haye paeeen baraye amaliate dakheli ast va nabayad estefade shavad
Private Function j2d(jy As Long, jm As Long, jd As Long)
    Dim r As Long
    Dim rgy As Long
    Dim rmarch As Long
    
    rgy = jalCal(jy)(1)
    rmarch = jalCal(jy)(2)
    j2d = g2d(rgy, 3, rmarch) + ((jm - 1) * 31) - ((jm \ 7) * (jm - 7)) + jd - 1
    
End Function



Private Function d2j(jdn As Long)
    Dim gy As Long
    gy = d2g(jdn)(0) ' Calculate Gregorian year (gy)
    Dim jy As Long
    jy = gy - 621
    Dim rmarch  As Long
    jalCal (jy)
    rmarch = jalCal(jy)(2)
    Dim rleap  As Long
    rleap = jalCal(jy)(0)
    
    Dim jdn1f As Long
    jdn1f = g2d(gy, 3, rmarch)  'r.march
    Dim jd As Long
    Dim jm As Long
    Dim k As Long

    ' Find number of days that passed since 1 Farvardin.
    k = jdn - jdn1f
    
    Dim result(3)
     
    If (k >= 0) Then
        If (k <= 185) Then
          ' The first 6 months.
          jm = 1 + (k \ 31)
          jd = (k Mod 31) + 1
          result(0) = jy
          result(1) = jm
          result(2) = jd
             
          d2j = result
          Exit Function
          
          
        Else
          ' The remaining months.
          k = k - 186
        End If
    Else
        ' Previous Jalaali year.
        jy = jy - 1
        k = k + 179
        If (rleap = 1) Then 'r.leap
          k = k + 1
        End If
    End If
    
    
    jm = 7 + (k \ 30)
    jd = (k Mod 30) + 1
    
    result(0) = jy
    result(1) = jm
    result(2) = jd
             
    d2j = result
    
End Function


Private Function d2g(jdn As Long)
    Dim j As Long
    Dim i As Long
    Dim gd As Long
    Dim gm As Long
    Dim gy As Long
    j = 4 * jdn + 139361631
    j = j + (((((4 * jdn + 183187720) \ 146097) * 3) \ 4) * 4) - 3908
    
    i = (((j Mod 1461) \ 4) * 5) + 308
    gd = ((i Mod 153) \ 5) + 1
    gm = ((i \ 153) Mod 12) + 1
    gy = (j \ 1461) - 100100 + ((8 - gm) \ 6)
    
    Dim result(3)

    result(0) = gy
    result(1) = gm
    result(2) = gd
  
    d2g = result

End Function



Private Function g2d(gy As Long, gm As Long, gd As Long)

    Dim d As Long
    d = (((gy + ((gm - 8) \ 6) + 100100) * 1461) \ 4) + ((153 * ((gm + 9) Mod 12) + 2) \ 5) + gd - 34840408
    d = d - ((((gy + 100100 + ((gm - 8) \ 6)) \ 100) * 3) \ 4) + 752
    g2d = d
End Function





Private Function jalCal(jy As Long)

    Dim breaks
    breaks = Array(-61, 9, 38, 199, 426, 686, 756, 818, 1111, 1181, 1210, 1635, 2060, 2097, 2192, 2262, 2324, 2394, 2456, 3178)

    Dim bl As Long
    bl = 20
    Dim gy As Long
    
    gy = jy + 621
    Dim leapJ  As Long
    leapJ = -14
    Dim jp As Long
    jp = breaks(0)
    Dim jm As Long
    Dim jump As Long
    Dim leap As Long
    Dim leapG As Long
    Dim march As Long
    Dim n As Long
    Dim i As Long
    

    If (jy < jp Or jy >= breaks(bl - 1)) Then
        MsgBox "Invalid Jalaali year " & jy
    End If
   

   'Find the limiting years for the Jalaali year jy.
   For i = 1 To (bl - 1) Step 1
        jm = breaks(i)
        jump = jm - jp
        If (jy < jm) Then Exit For
        
        leapJ = leapJ + (jump \ 33) * 8 + ((jump Mod 33) \ 4)
        jp = jm
   Next
   
  
   n = jy - jp

  ' Find the number of leap years from AD 621 to the beginning
  ' of the current Jalaali year in the Persian calendar.
  
  leapJ = leapJ + (n \ 33) * 8 + (((n Mod 33) + 3) \ 4)
  If ((jump Mod 33) = 4 And (jump - n) = 4) Then
    leapJ = leapJ + 1
  End If

  ' And the same in the Gregorian calendar (until the year gy).
  leapG = (gy \ 4) - ((((gy \ 100) + 1) * 3) \ 4) - 150

  ' Determine the Gregorian date of Farvardin the 1st.
  march = 20 + leapJ - leapG

  ' Find how many years have passed since the last leap year.
  If ((jump - n) < 6) Then
    n = n - jump + ((jump + 4) \ 33) * 33
  End If
  
  leap = ((((n + 1) Mod 33) - 1) Mod 4)
  If (leap = -1) Then
    leap = 4
  End If
  
  Dim result(3)

  result(0) = leap
  result(1) = gy
  result(2) = march
  
  jalCal = result


End Function





