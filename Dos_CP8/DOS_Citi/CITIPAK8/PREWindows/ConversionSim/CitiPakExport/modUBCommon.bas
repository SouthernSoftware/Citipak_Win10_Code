Attribute VB_Name = "modUBCommon"
Option Explicit
DefInt A-Z
Dim UBSetUpRec(1) As UBSetupRecType


Public Sub Main()
  Dim UTILNAME As String, DEFCITY As String, TCity As String
  Dim ch As Integer, zz As Integer, TUtil As String
'  If Exist("HollySpringsUBExport.EXE") = False Then
'    GoTo NotTodayBrad
'  End If

  TUtil = "+3*4[7?V?;9>*P" + Chr(34) + ">=%#-8"
  TCity = Mid$(TUtil, 9)
'  If Not Exist("glsetup.dat") Then
'    GoTo NotTodayBrad
'  End If
  For zz = 1 To Len(TUtil)
    Mid$(TUtil, zz, 1) = Chr$((Asc(Mid$(TUtil, zz, 1)) Xor 126) Xor zz)
  Next
  For zz = 1 To Len(TCity)
    Mid$(TCity, zz, 1) = Chr$((Asc(Mid$(TCity, zz, 1)) Xor 126) Xor zz + 8)
  Next
  LoadUBSetUpFile UBSetUpRec(), 0
  UTILNAME = QPTrim$(UBSetUpRec(1).UTILNAME)
  DEFCITY = QPTrim$(UBSetUpRec(1).DEFCITY)

'  If TUtil <> UtilName Then
'    GoTo NotTodayBrad
'  End If
'  If DefCity <> TCity Then
'    GoTo NotTodayBrad
'  End If
'  If FileSize("UBTRANS.DAT") < 440000000 Then
'    GoTo NotTodayBrad
'  End If
'  If FileSize("UBCUST.DAT") < 39000000 Then
'    GoTo NotTodayBrad
'  End If

  'Form1.Show

NotTodayBrad:
End Sub

Public Static Function Round#(ByVal N#)
  Round# = (Int(N# * 100 + 0.5000001)) / 100
End Function

Public Static Sub UBLog(Text$)
  Dim Today As String, TheTime As String
  Dim AmPm As String, Hour As String
  Dim ThisHour As Integer, LogFile As Integer

  Today$ = Date$
  Today$ = Left$(Today$, 2) + Mid$(Today$, 4, 2) + Right$(Today$, 2)

  TheTime$ = Time$
  If Left$(TheTime$, 1) = "0" Then
    ThisHour = Val(Mid$(TheTime$, 2, 1))
  Else
    ThisHour = Val(Mid$(TheTime$, 1, 2))
  End If

  Select Case ThisHour
  Case Is > 11
    ThisHour = ThisHour - 12
    If ThisHour = 0 Then ThisHour = 12
    AmPm$ = "pm"
  Case 1 To 12
    AmPm$ = "am"
  Case 0
    Hour = 12
    AmPm$ = "am"
  End Select
  Select Case ThisHour
    Case 1 To 9
      Hour$ = "0" + QPTrim$(Str$(ThisHour))
    Case Else
      Hour$ = QPTrim$(Str$(ThisHour))
  End Select
  TheTime$ = Hour$ + ":" + Mid$(TheTime$, 4) + AmPm$
  LogFile = FreeFile
  Open UBPath$ + "UBLOG.DAT" For Append Shared As LogFile Len = 255
  Print #LogFile, "UB: "; Today$; " @"; TheTime$
  Print #LogFile, "    "; Text$
  Close #LogFile
  If WDflag Then
    LogFile = FreeFile
    Open UBPath$ + "UBTRSPEC.DAT" For Append Shared As LogFile Len = 255
    Print #LogFile, "UB: "; Today$; " @"; TheTime$
    Print #LogFile, "    "; Text$
    Close #LogFile
  End If
  Text$ = "": TheTime$ = ""
End Sub

'!!! populates the combo box with revenues

'!!! from gl common for date check on report screens

Public Static Function Num2Date$(intDate%)
  On Error GoTo BadNum2Date
  Dim Year  As Long

  If intDate% <= -29218 Then
    Num2Date$ = ""
  Else
    Num2Date$ = Format(DateAdd("d", (intDate%), "12-31-1979"), "mm/dd/yyyy")
  End If

  If Num2Date$ = "12/31/1979" Then
    Num2Date$ = ""
  End If

  Exit Function
BadNum2Date:
  On Error GoTo 0
  Num2Date = ""
End Function

'Public Function Date2Num%(txtDate$)
'  On Error GoTo BadDate2Num
'  If Len(QPTrim$(txtDate$)) = 10 Then
'    Date2Num% = DateDiff("d", "12/31/1979", txtDate$)
'  Else
'    Date2Num% = -32767
'  End If
'  Exit Function
'
'BadDate2Num:
'  On Error GoTo 0
'  Date2Num% = -32767
'End Function


Public Function IsDeleted%(AcctNum&)
  ReDim UBCustRec(1) As NewUBCustRecType
  Dim Handle As Integer
  Dim UBCustRecLen As Integer

  UBCustRecLen = Len(UBCustRec(1))
  Handle = FreeFile
  Open UBCustFile For Random Shared As Handle Len = UBCustRecLen
  Get Handle, AcctNum&, UBCustRec(1)
  Close Handle

  If UBCustRec(1).DelFlag <> 0 Then
    IsDeleted% = True
  Else
    IsDeleted% = False
  End If
  Erase UBCustRec

End Function

'This function returns the number of customer records
Public Function GetNumOfCust&()
  ReDim TCustRec(1) As NewUBCustRecType
  Dim RecLen As Integer
  RecLen = Len(TCustRec(1))
  GetNumOfCust = FileSize(UBCustFile) \ RecLen
  Erase TCustRec
End Function

Public Sub LoadUBSetUpFile(UBSetUpRec() As UBSetupRecType, UBSetupLen)
  Dim Handle As Integer
  UBSetupLen = Len(UBSetUpRec(1))            'use the length as an error flag
  If Exist("UBSETUP.DAT") Then
    Handle = FreeFile
    Open "UBSETUP.DAT" For Random Shared As Handle Len = UBSetupLen     'open data file
    If LOF(Handle) > 0 Then
      Get #Handle, 1, UBSetUpRec(1)
    End If
    Close Handle
  End If
End Sub

Public Function FileSize(FileName$) As Long
  Dim FileHandle As Integer
  If Exist(FileName$) Then
    FileHandle = FreeFile
    Open FileName$ For Binary As FileHandle
    FileSize = LOF(FileHandle)
    Close FileHandle
  Else
    FileSize = 0
  End If
End Function
Public Static Function UBUsing$(ByVal Number As Double, ByVal fmt As String, Optional LeadZeroFlag As Boolean)
  Dim TempNumber As String
  Dim FmtNumber As String
  Dim TempLen As Integer
  Dim BuckPos As Integer, FmtLen As Integer
  FmtLen = Len(fmt)
  BuckPos = InStr(fmt, "$")
  If BuckPos = 1 Then
    fmt = Right$(fmt, FmtLen - 1)
  ElseIf BuckPos > 1 Then
    fmt = Left$(fmt, BuckPos - 1) + Mid$(fmt, BuckPos + 1)
  End If
  FmtNumber = Space$(Len(fmt))
  TempNumber = Format(Number, fmt)
  TempLen = Len(TempNumber)
  If TempLen = 0 Then
    TempNumber = "0"
    GoTo GotZero
  End If
  If TempLen >= 2 Then
    If Mid$(TempNumber, (TempLen - 1), 1) = "." Then
      TempNumber = TempNumber + "0"
    End If
  End If
  If Right$(TempNumber, 1) = "." Then
    TempNumber = TempNumber + "00"
  End If
  If BuckPos > 0 Then
    TempNumber = "$" + TempNumber
  End If
GotZero:
'  If LeadZeroFlag Then
    If TempNumber = ".00" Then
      TempNumber = "0.00"
    End If
'  End If

  RSet FmtNumber = TempNumber

  UBUsing = FmtNumber

'Number = 5: Fmt = "$##,##0.00": Print Right(String(Len(Fmt), " ") & Format(Number, Fmt), Len(Fmt))
End Function


