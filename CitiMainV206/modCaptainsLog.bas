Attribute VB_Name = "modCaptainsLog"
Option Explicit

Public Sub MainLog(Info As String)
  Dim AcctLogFileName As String, AcctLogFile As Integer
  AcctLogFileName = "CitiLog.dat"
  AcctLogFile = FreeFile
  Open AcctLogFileName$ For Append As AcctLogFile
  Print #AcctLogFile, Date$; " @ "; Time$; " USER: "; PWUser$; " ON: "; ComputerName$; " "; Info$
  Close AcctLogFile
End Sub

Public Function QTR$(Text As String)
  Dim StrLen As Long
  Dim cnt As Long
  Dim ThisChar As Integer
  StrLen = Len(Text)
  For cnt = 1 To StrLen
    ThisChar = Asc(Mid$(Text, cnt, 1))
    If ThisChar = 0 Then
      Mid$(Text$, cnt, 1) = " "
    End If
  Next
  QTR$ = Trim$(Text)
End Function
