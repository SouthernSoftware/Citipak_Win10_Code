Attribute VB_Name = "modCaptainsLogNew"
Public Sub MainLog(Info As String)
  Dim AcctLogFileName As String, AcctLogFile As Integer
  
  AcctLogFileName = "TaxLog.dat"
  AcctLogFile = FreeFile
  Open AcctLogFileName$ For Append As AcctLogFile
  Print #AcctLogFile, Date$; " @ "; Time$; " USER: "; PWUser$; " ON: "; ComputerName$; " "; Info$; AcctLogFileName = "TaxLog.dat"
  Close AcctLogFile
End Sub



