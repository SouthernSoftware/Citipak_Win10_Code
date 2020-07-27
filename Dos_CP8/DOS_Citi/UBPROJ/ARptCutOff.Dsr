VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptCutOff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Cut-Off Report"
   ClientHeight    =   4380
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   8952
   Icon            =   "ARptCutOff.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   15790
   _ExtentY        =   7726
   SectionData     =   "ARptCutOff.dsx":08CA
End
Attribute VB_Name = "ARptCutOff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Private hFile As Integer
Dim rpt As ActiveReport
Dim headers(1 To 28) As String
Dim cnt As Integer
Dim TotType As Integer
Dim NoMtrInf As Boolean
'TotType is 1 for total only and 2 is for all 3 totals
'Current, Previous and Total
'NoMtrInf is true if summary only and no meter info listed
Public Sub GetName(RName As String, RptTyp As Integer, Mtr As Boolean)
  ReportFile$ = RName$
  TotType = RptTyp
  NoMtrInf = Mtr
End Sub

Private Sub ActiveReport_DataInitialize()
   
    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
    headers(1) = "Location"
    headers(2) = "Acct"
    headers(3) = "CustName"
    headers(4) = "ServAddr"
    headers(5) = "Past"
    headers(6) = "Curr"
    headers(7) = "RealB"
    headers(8) = "Met1"
    headers(9) = "Last1"
    headers(10) = "Read1"
    headers(11) = "Met2"
    headers(12) = "Last2"
    headers(13) = "Read2"
    headers(14) = "Met3"
    headers(15) = "Last3"
    headers(16) = "Read3"
    headers(17) = "Met4"
    headers(18) = "Last4"
    headers(19) = "Read4"
    headers(20) = "Met5"
    headers(21) = "Last5"
    headers(22) = "Read5"
    headers(23) = "Met6"
    headers(24) = "Last6"
    headers(25) = "Read6"
    headers(26) = "Met7"
    headers(27) = "Last7"
    headers(28) = "Read7"

'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    For cnt = 1 To 28
      Fields.Add headers(cnt)
    Next
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
    frmLoadingRpt.ShowHowMuch
    Line Input #hFile, sLine
    arr = Split(sLine, "~")
'    ' Here we set the values of the fields that we defines as unbound
'    ' or user defined.
    For cnt = 1 To 28
      Fields(headers(cnt)) = arr(cnt - 1)
    Next
'If something wrong in file give message instead of crashing
Exit Sub
ERRORSTUFF:
      Unload frmLoadingRpt
'  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "ARptVendHist", "Fetch Data", Erl)
'    Case emrExitProc:
'      Resume Proc_Exit
'    Case emrResume:
'      Resume
'    Case emrResumeNext:
'      Resume Next
'    Case Else
'      '--- Technically, this should never happen.
'      Resume Proc_Exit
'  End Select
   MsgBox "Err.Number, Err.Description, Err.Source", vbOKOnly, "Error"
   GoSub Proc_Exit
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    Unload Me
End Sub

Public Sub startrpt()
  Me.Run
End Sub
Private Sub ActiveReport_ReportStart()
  If NoMtrInf = True Then
    Label12.Visible = False
    Label13.Visible = False
    Label14.Visible = False
  End If
  If TotType = 1 Then
    Me.Field41.Top = 0
    Label19.Visible = False
    Label20.Visible = False
    totCurr.Visible = False
    totPast.Visible = False
    Me.Label5.Visible = False
    Me.Label6.Visible = False
  End If
  Me.Label16 = Me.txtRptParm1
  Me.Label18 = Me.txtRptParm2
End Sub
Private Sub PageHeader_Format()
If Me.pageNumber = 1 Then
  Label15.Visible = True
  Shape1.Visible = True
  txtRptParm1.Visible = True
  txtRptParm2.Visible = True
  If TotType = 2 Then
    Me.PageHeader.Height = 1704
  Else
    If NoMtrInf = True Then
      Me.PageHeader.Height = 1440
      Me.Label7.Top = 1170
    Else
      Me.PageHeader.Height = 1704
      Me.Label7.Top = 1170
    End If
  End If
Else
  Label15.Visible = False
  Shape1.Visible = False
  txtRptParm1.Visible = False
  txtRptParm2.Visible = False
  If TotType = 2 Then
    Me.PageHeader.Height = 1176
    Me.Label1.Top = 630
    Me.Label2.Top = 630
    Me.Label3.Top = 630
    Me.Label4.Top = 630
    Me.Label5.Top = 900
    Me.Label6.Top = 900
    Me.Label7.Top = 900
    Me.Label12.Top = 900
    Me.Label13.Top = 900
    Me.Label14.Top = 900
  Else
    If NoMtrInf = True Then
      Me.PageHeader.Height = 864
      Me.Label1.Top = 630
      Me.Label2.Top = 630
      Me.Label3.Top = 630
      Me.Label4.Top = 630
      Me.Label7.Top = 630
    Else
      Me.PageHeader.Height = 1176
      Me.Label1.Top = 630
      Me.Label2.Top = 630
      Me.Label3.Top = 630
      Me.Label4.Top = 630
      Me.Label7.Top = 630
      Me.Label12.Top = 900
      Me.Label13.Top = 900
      Me.Label14.Top = 900
    End If
  End If
End If
End Sub
Private Sub Detail_Format()
  If NoMtrInf = True Then
    If TotType = 1 Then
      Detail.Height = 0
    Else
      Detail.Height = 270
    End If
  Else
    Detail.Height = 0
    If Me.Fields("Last1").Value <> " " Then
      'Detail.Height = Detail.Height + 270
      Detail.Height = 270
    End If
    If Me.Fields("Last2").Value <> " " Then
      'Detail.Height = Detail.Height + 270
      Detail.Height = 540
    End If
    If Me.Fields("Last3").Value <> " " Then
      'Detail.Height = Detail.Height + 270
      Detail.Height = 810
    End If
    If Me.Fields("Last4").Value <> " " Then
      'Detail.Height = Detail.Height + 270
      Detail.Height = 1080
    End If
    If Me.Fields("Last5").Value <> " " Then
      'Detail.Height = Detail.Height + 270
      Detail.Height = 1350
    End If
    If Me.Fields("Last6").Value <> " " Then
      'Detail.Height = Detail.Height + 270
      Detail.Height = 1620
    End If
    If Me.Fields("Last7").Value <> " " Then
      'Detail.Height = Detail.Height + 270
      Detail.Height = 1890
    End If
'
   End If
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
      MsgBox "File - CutOff.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - CutOff.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If

End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
 ' KillFile ReportFile$
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
    MsgBox "File - CutOff.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - CutOff.txt, created in the Citipak Directory.", vbOKOnly
  End If
End Sub
Private Sub ExportReport(x As Integer)
  Dim oEXL As ActiveReportsExcelExport.ARExportExcel
  Dim oTXT As ActiveReportsTextExport.ARExportText
  Dim outfile As String
  If Right$(UBPath, 1) = ":" Then
    outfile = UBPath
  Else
    outfile = UBPath & "\"
  End If
  Select Case x
    Case 1   '"Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & "CutOff.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "CutOff.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

