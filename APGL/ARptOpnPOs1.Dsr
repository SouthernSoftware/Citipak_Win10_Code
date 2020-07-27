VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptOpnPOs1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open Purchase Orders"
   ClientHeight    =   5352
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   9240
   Icon            =   "ARptOpnPOs1.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   16298
   _ExtentY        =   9440
   SectionData     =   "ARptOpnPOs1.dsx":08CA
End
Attribute VB_Name = "ARptOpnPOs1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Public SubFile As String
Private hFile As Integer
Dim rpt As ActiveReport
Dim tempTot As Double, temptotE As Double
Dim vendcnt As Integer, TotE As Double, TotL As Double
Public Sub GetName(RName As String, SName As String)
  ReportFile$ = RName$
  SubFile$ = SName$
End Sub
Private Sub ActiveReport_DataInitialize()
   
    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    Fields.Add "OPO1"
    Fields.Add "OPO2"
    Fields.Add "Date"
    Fields.Add "POAmt"
    Fields.Add "Desc"
    Fields.Add "DistAcct"
    Fields.Add "DistAmt"
    
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
    Fields("OPO1").Value = arr(0)
    Fields("OPO2").Value = arr(1)
    Fields("Date").Value = arr(2)
    Fields("POAmt").Value = arr(3)
    Fields("Desc").Value = arr(4)
    Fields("DistAcct").Value = arr(5)
    Fields("DistAmt").Value = arr(6)
'If something wrong in file give message instead of crashing
Exit Sub
ERRORSTUFF:
      Unload frmLoadingRpt
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "ARptVendHist", "Fetch Data", Erl)
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
    End If
    If KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - OpenPORpt.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - OpenPORpt.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If

End Sub


Private Sub Detail_Format()
If Me.Fields("Desc").Value = "Liquidated" Then
    tempTot = Round(tempTot + Me.Fields("DistAmt").Value)
Else
    temptotE = Round(temptotE + Me.Fields("DistAmt").Value)
End If
End Sub


Private Sub GroupFooter3_AfterPrint()
  tempTot = 0
  temptotE = 0
End Sub

Private Sub GroupFooter3_Format()
  SubLiq.DataValue = tempTot
  SubEnc.DataValue = temptotE
  TotE = Round(TotE + temptotE)
  TotL = Round(TotL + tempTot)
End Sub


Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
  KillFile ReportFile$
  KillFile SubFile$
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
    MsgBox "File - OpenPORpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - OpenPORpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "OpenPORpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "OpenPORpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ReportFooter_Format()
  Set Me.ARSubFundTot.object = New ARSubFundTot
  TotEnc.DataValue = TotE
  TotLiq.DataValue = TotL
End Sub

