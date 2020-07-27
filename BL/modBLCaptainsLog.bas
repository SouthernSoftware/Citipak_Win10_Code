Attribute VB_Name = "modBLCaptainsLog"
Option Explicit
Public Sub MainLog(Info As String)
  Dim AcctLogFileName As String, AcctLogFile As Integer
  
  AcctLogFileName = "arlog.dat"
  AcctLogFile = FreeFile
  Open AcctLogFileName$ For Append As AcctLogFile
  Print #AcctLogFile, Date$; " @ "; Time$; " USER: "; PWUser$; " ON: "; ComputerName$; " "; Info$; ; " "; AcctLogFileName
  Close AcctLogFile
End Sub



