VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptLedHist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A/P Ledger History"
   ClientHeight    =   4380
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   8628
   Icon            =   "ARptLedHist.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   15219
   _ExtentY        =   7726
   SectionData     =   "ARptLedHist.dsx":08CA
End
Attribute VB_Name = "ARptLedHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Private hFile As Integer
Dim rpt As ActiveReport
Public Sub GetName(RName As String)
  ReportFile$ = RName$
End Sub

Private Sub ActiveReport_DataInitialize()
   
    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    Fields.Add "TransNum"
    Fields.Add "VendCode"
    Fields.Add "Vendor"
    Fields.Add "Date1"
    Fields.Add "Bank"
    Fields.Add "G1"
    Fields.Add "G2"
    Fields.Add "G3"
    Fields.Add "G4"
    Fields.Add "G5"
    Fields.Add "G6"
    Fields.Add "G7"
    Fields.Add "G8"
    Fields.Add "DistAcct"
    Fields.Add "DistAmt"
    Fields.Add "DistRec"
    
End Sub
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmLoadingRpt
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sLine As String
Dim arr() As String
On Error GoTo ERRORSTUFF
'    ' We reached the end of the file we exit leaving the
'    ' eof parameter as True (default except on first call) that will
'    ' tell AR that we are done feeding data
'    ' otherwise we have to set the eof parameter to False so that
'    ' AR continues fetching data, until we're done
'    ' if the report had a data control, the value of the parameter
'    ' will be ignored, AR will always follow the data control's recordset
'    ' EOF property
    If VBA.eof(hFile) Then
        eof = True
        Exit Sub
    Else
        eof = False
    End If

    Line Input #hFile, sLine
    arr = Split(sLine, "~")
'    ' Here we set the values of the fields that we defines as unbound
'    ' or user defined.
    Fields("TransNum").Value = arr(0)
    Fields("VendCode").Value = arr(1)
    Fields("Vendor").Value = arr(2)
    Fields("Date1").Value = arr(3)
    Fields("Bank").Value = arr(4)
    Fields("G1").Value = arr(5)
    Fields("G2").Value = arr(6)
    Fields("G3").Value = arr(7)
    Fields("G4").Value = arr(8)
    Fields("G5").Value = arr(9)
    Fields("G6").Value = arr(10)
    Fields("G7").Value = arr(11)
    Fields("G8").Value = arr(12)
    Fields("DistAcct").Value = arr(13)
    Fields("DistAmt").Value = arr(14)
    Fields("DistRec").Value = arr(15)
'If something wrong in file give message instead of crashing
Exit Sub
ERRORSTUFF:
      Unload frmLoadingRpt
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "ARptBudVAct", "Fetch Data", Erl)
    Case emrExitProc:
      Resume Proc_Exit
    Case emrResume:
      Resume
    Case emrResumeNext:
      Resume Next
    Case Else
      '--- Technically, this should never happen.
      Resume Proc_Exit
  End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    Unload Me
End Sub

Public Sub startrpt()
  Me.Run
End Sub

Private Sub ActiveReport_Initialize()
  Me.Toolbar.Tools.Add "&Close"
  Me.Toolbar.Tools.Add "Save/&Excel"
  Me.Toolbar.Tools.Add "/&Text"
End Sub

Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
    KeyCode = 0
  End If
  If Shift = 4 Then
    If KeyCode = vbKeyC Then
      Unload Me
      KeyCode = 0
    End If
    If KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - Ledhist.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - Ledhist.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If

End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
'  KillFile ReportFile$
End Sub
Private Sub ActiveReport_ReportEnd()
    If hFile <> 0 Then
        Close #hFile
    End If
  Unload frmLoadingRpt
  Me.Show 1
End Sub
Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
  If Tool = "&Close" Then
    Unload Me
  End If
  If Tool = "Save/&Excel" Then
    Screen.MousePointer = vbHourglass
    ExportReport 1
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - Ledhist.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - Ledhist.txt, created in the Citipak Directory.", vbOKOnly
  End If
End Sub
Private Sub ExportReport(x As Integer)
  Dim oEXL As ActiveReportsExcelExport.ARExportExcel
  Dim oTXT As ActiveReportsTextExport.ARExportText
  Dim outfile As String
  If Right$(StartPath, 1) = ":" Then
    outfile = StartPath
  Else
    outfile = StartPath & "\"
  End If
  Select Case x
    Case 1   '"Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & "Ledhist.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "Ledhist.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

