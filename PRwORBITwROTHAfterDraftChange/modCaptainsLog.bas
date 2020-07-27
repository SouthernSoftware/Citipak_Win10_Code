Attribute VB_Name = "modCaptainsLog"
Option Explicit
Public Sub MainLog(Info As String)
  Dim AcctLogFileName As String, AcctLogFile As Integer
  
  AcctLogFileName = "PRLog.dat"
  AcctLogFile = FreeFile
  Open AcctLogFileName$ For Append As AcctLogFile
  Print #AcctLogFile, Date$; " @ "; Time$; " USER: "; PWUser$; " ON: "; ComputerName$; " "; Info$; AcctLogFileName = "PRLog.dat"
  Close AcctLogFile
End Sub

