Attribute VB_Name = "modW2CaptainsLog"
Option Explicit
Public Sub MainLog(Info As String)
  Dim AcctLogFileName As String, AcctLogFile As Integer
  AcctLogFileName = "PRLog.dat"
  AcctLogFile = FreeFile
  Open AcctLogFileName$ For Append As AcctLogFile
  Print #AcctLogFile, Date$; " @ "; Time$; " "; ComputerName$; " "; Info$
  Close AcctLogFile
End Sub

