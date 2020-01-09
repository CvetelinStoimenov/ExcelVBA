'
' Copyright (C) 2006, Ceco Vasilev
'
Option Explicit
Dim warray(9) As String
Dim edin, edna, dwa, dwe, sto, dwesta, trista, stotin, deset, mil, mila, hil, hili, lw, st, i, jn, na

Sub Init()
    warray(1) = "åä"
    warray(2) = "äâ"
    warray(3) = "òðè"
    warray(4) = "÷åòèðè"
    warray(5) = "ïåò"
    warray(6) = "øåñò"
    warray(7) = "ñåäåì"
    warray(8) = "îñåì"
    warray(9) = "äåâåò"


    jn = "èí"
    na = "íà"
    sto = "ñòî"
    dwesta = "äâåñòà"
    trista = "òðèñòà"
    stotin = "ñòîòèí"
    deset = "äåñåò"
    mil = "ìèëèîí"
    mila = " ìèëèîíà"
    hil = "õèëÿäà"
    hili = " õèëÿäè"
    lw = " ëâ."
    st = " ñò."
    i = " è "
End Sub

Function Conv(sstr As String) As String
Dim rstr, astrc, bstrc, cstrc
Dim apos, bpos, cpos

   If Len(sstr) = 3 Then
     apos = Asc(Mid(sstr, 3, 1)) - Asc("0")
     bpos = Asc(Mid(sstr, 2, 1)) - Asc("0")
     cpos = Asc(Mid(sstr, 1, 1)) - Asc("0")
   Else
     apos = Asc(Mid(sstr, 2, 1)) - Asc("0")
     bpos = Asc(Mid(sstr, 1, 1)) - Asc("0")
     cpos = 0
   End If

   If apos = 1 Then
     If Len(sstr) = 3 Then
       astrc = warray(apos) + jn
     Else
       astrc = warray(apos) + na
     End If
   Else
     If apos = 2 Then
       If Len(sstr) = 3 Then
         astrc = warray(apos) + "à"
       Else
         astrc = warray(apos) + "e"
       End If
     Else
       If (apos >= 3) And (apos <= 9) Then astrc = warray(apos)
     End If
   End If
   
   If bpos = 1 Then
      If apos = 1 Then
         bstrc = astrc + "à" + deset
         astrc = ""
      Else
         If apos = 0 Then
            bstrc = deset
            astrc = ""
         Else
            If (apos >= 2) And (apos <= 9) Then bstrc = astrc + na + deset
            astrc = ""
         End If
      End If
   Else
      If bpos = 2 Then
        bstrc = warray(bpos) + "à" + deset
      Else
        If (bpos >= 3) And (bpos <= 9) Then bstrc = warray(bpos) + deset
      End If
   End If
    
  Select Case cpos
      Case 1
        cstrc = sto
      Case 2
        cstrc = dwesta
      Case 3
        cstrc = trista
   Case Else
      If (cpos >= 4) And (cpos <= 9) Then cstrc = warray(cpos) + stotin
   End Select
 
   rstr = astrc
   If Len(cstrc) > 0 Then
      If Len(astrc) > 0 Then
         If Len(bstrc) > 0 Then
            rstr = cstrc + " " + bstrc + i + rstr
         Else
            rstr = cstrc + i + rstr
         End If
      Else
         If Len(bstrc) > 0 Then
            rstr = cstrc + i + bstrc
         Else
            rstr = cstrc
         End If
      End If
   Else
      If Len(bstrc) > 0 Then
         If Len(astrc) > 0 Then
            rstr = bstrc + i + rstr
         Else
            rstr = bstrc
         End If
      End If
   End If

   Conv = rstr
   
End Function

Function Dig2Txt(instring As String) As String
Dim LastDelimiter As String, wstr As String, fstr As String, ostring As String, cstrc As String
Dim c1, c2

Init


wstr = instring
LastDelimiter = InStr(1, wstr, ".")

If LastDelimiter > 0 Then
  fstr = Mid(wstr, LastDelimiter + 1, 2)
  wstr = Mid(wstr, 1, LastDelimiter - 1)
End If

If Len(fstr) < 2 Then
  fstr = Mid("00", 1, 2 - Len(fstr)) + fstr
End If

ostring = ""
c1 = 0
Do While wstr <> ""

  If Len(wstr) < 3 Then
    wstr = Mid("000", 1, 3 - Len(wstr)) + wstr
  End If
  cstrc = Conv(Mid(wstr, Len(wstr) - 2, 3))

  Select Case c1
  Case 0
     If cstrc <> "" Then
     If (wstr = "001") Then
         cstrc = edin
       ElseIf (wstr = "002") Then
         cstrc = dwa
       End If
     End If
  Case 1
     If (wstr = "001") Then
       cstrc = hil
     Else
       cstrc = cstrc + hili
     End If
  Case 2
     If (wstr = "001") Then
       cstrc = mil
     Else
       cstrc = cstrc + mila
     End If
  End Select
  
  If cstrc <> "" Then
    If c1 > 0 Then
'      If c1 = 1 Then
'        ostring = cstrc + i + ostring
'      Else
        ostring = cstrc + " " + ostring
'      End If
    Else
      ostring = cstrc
    End If
  End If
  If Len(wstr) >= 3 Then
    wstr = Mid(wstr, 1, Len(wstr) - 3)
  Else
    wstr = ""
  End If
  c1 = c1 + 1
  
Loop

  If Len(fstr) > 0 Then
    cstrc = Conv(fstr)
    If ostring <> "" Then
      Dig2Txt = ostring + lw + " è " + cstrc + st
    Else
      Dig2Txt = cstrc + st
    End If
  Else
    Dig2Txt = ostring + lw
  End If

End Function
