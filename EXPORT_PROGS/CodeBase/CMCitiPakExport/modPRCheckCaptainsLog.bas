Attribute VB_Name = "modPRCheckCaptainsLog"
Option Explicit
Public Sub MainLog(Info As String)
  Dim AcctLogFileName As String, AcctLogFile As Integer
  AcctLogFileName = "PRLog.dat"
  AcctLogFile = FreeFile
  Open AcctLogFileName$ For Append As AcctLogFile
  Print #AcctLogFile, Date$; " @ "; Time$; " "; ComputerName$; " "; Info$
  Close AcctLogFile
End Sub

