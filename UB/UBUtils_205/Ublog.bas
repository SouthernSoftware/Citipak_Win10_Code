Attribute VB_Name = "Module1"

Static Sub UBLog(Text$)

'  IF NOT BeenDone THEN
'    BeenDone = True
    Today$ = Date$
    Today$ = LEFT$(Today$, 2) + MID$(Today$, 4, 2) + RIGHT$(Today$, 2)
'  END IF

  TheTime$ = Time$
  If LEFT$(TheTime$, 1) = "0" Then
    Hour = VAL(MID$(TheTime$, 2, 1))
  Else
    Hour = VAL(MID$(TheTime$, 1, 2))
  End If

  Select Case Hour
  Case Is > 11
    Hour = Hour - 12
    If Hour = 0 Then Hour = 12
    AmPm$ = "pm"
  Case 1 To 12
    AmPm$ = "am"
  Case 0
    Hour = 12
    AmPm$ = "am"
  End Select
  Select Case Hour
    Case 1 To 9
      Hour$ = "0" + QPTrim$(STR$(Hour))
    Case Else
      Hour$ = QPTrim$(STR$(Hour))
  End Select
  TheTime$ = Hour$ + ":" + MID$(TheTime$, 4) + AmPm$
  LogFile = FREEFILE
  Open "UBLOG.DAT" For Append Shared As LogFile Len = 255
  Print #LogFile, "UB: "; Today$; " @"; TheTime$
  Print #LogFile, "    "; Text$
  Close #LogFile
  Text$ = "": TheTime$ = ""
End Sub

