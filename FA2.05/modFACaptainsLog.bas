Attribute VB_Name = "modFACaptainsLog"
Option Explicit
Public Sub MainLog(Info As String)
  Dim AcctLogFileName As String, AcctLogFile As Integer
  
  AcctLogFileName = "FALog.dat"
  AcctLogFile = FreeFile
  Open AcctLogFileName$ For Append As AcctLogFile
  Print #AcctLogFile, Date$; " @ "; Time$; " USER: "; PWUser$; " ON: "; ComputerName$; " "; Info$; ; " "; AcctLogFileName
  Close AcctLogFile
End Sub


