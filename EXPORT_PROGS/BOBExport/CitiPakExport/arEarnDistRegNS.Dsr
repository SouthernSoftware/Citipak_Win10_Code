VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arEarnDistRegNS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Earnings Distribution"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arEarnDistRegNS.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arEarnDistRegNS.dsx":08CA
End
Attribute VB_Name = "arEarnDistRegNS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private hFile As Integer
Dim MultAccts As Integer
Dim EndReport As Boolean
Dim NoEscape As Boolean
Private Sub ActiveReport_Initialize()
  Me.ToolBar.Tools.Add "&Close"
  Me.ToolBar.Tools.Add "Save/&Excel"
  Me.ToolBar.Tools.Add "&Text"
End Sub

Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  If KeyCode = vbKeyEscape And NoEscape = False Then
    NoEscape = True
    DoEvents
    Unload Me
    DoEvents
    KeyCode = 0
  End If
  If Shift = 4 Then
    If KeyCode = vbKeyC Then
      Unload Me
      KeyCode = 0
    ElseIf KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - DistRegisterbyAcctNum.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - DistRegisterbyAcctNum.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
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
    MsgBox "File - DistRegisterbyAcctNum.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - DistRegisterbyAcctNum.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "DistRegisterbyAcctNum.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "DistRegisterbyAcctNum.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  EndReport = False
  NoEscape = False
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmLoadingRpt
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True 'removes the error message
End Sub
Private Sub ActiveReport_DataInitialize()
  hFile = FreeFile
  Open StartPath & "\PRRPTS\DISTRIBUNSG.RPT" For Input As #hFile
  
  Fields.Add "fldENum" '(0)
  Fields.Add "fldEmployee" '(1)
  Fields.Add "fldBaseRate" '(2)
  Fields.Add "fldOTRate" '(3)
  Fields.Add "fldAcctCnt" '(4)
  Fields.Add "fldEmployer" '(5)
  Fields.Add "fldDate" '(6)
  
  Fields.Add "fldEAcctNum" '(7)
  Fields.Add "fldSalPct" '(8)
  Fields.Add "fldRegHrs" '(9)
  Fields.Add "fldOTHrs" '(10)
  Fields.Add "fldRegPay" '(11)
  Fields.Add "fldOTPay" '(12)
  Fields.Add "fldETother" '(13)
  Fields.Add "fldGrsPy" '(14)
  Fields.Add "fldSocSec" '(15)
  Fields.Add "fldMed" '(16)
  Fields.Add "fldRet" '(17)
  End Sub
Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  
  If VBA.eof(hFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  
  Line Input #hFile, sLine
  arr = Split(sLine, "~")
  
  Fields("fldENum").Value = arr(0)
  Fields("fldEmployee").Value = arr(1)
  Fields("fldBaseRate").Value = arr(2)
  Fields("fldOTRate").Value = arr(3)
  Fields("fldAcctCnt").Value = arr(4)
  Fields("fldEmployer").Value = arr(5)
  Fields("fldDate").Value = arr(6)
  Fields("fldEAcctNum").Value = arr(7)
  Fields("fldSalPct").Value = arr(8)
  Fields("fldRegHrs").Value = arr(9)
  Fields("fldOTHrs").Value = arr(10)
  Fields("fldRegPay").Value = arr(11)
  Fields("fldOTPay").Value = arr(12)
  Fields("fldETother").Value = arr(13)
  Fields("fldGrsPy").Value = arr(14)
  Fields("fldSocSec").Value = arr(15)
  Fields("fldMed").Value = arr(16)
  Fields("fldRet").Value = arr(17)
End Sub
Private Sub ActiveReport_ReportEnd()
  Unload frmLoadingRpt
  If hFile <> 0 Then
    Close #hFile
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  ReportHeader.Height = 0
  Label47.Visible = False 'Summary
  Me.fldTimeDate.Text = Now
  Me.Zoom = -1
End Sub

Private Sub Detail_Format()
  If Fields("fldAcctCnt").Value > 1 Then
    GroupFooter1.Visible = True
  Else
    GroupFooter1.Visible = False
  End If

End Sub

Private Sub PageHeader_Format()
  If EndReport = True Then
    Label47.Visible = True
  End If
End Sub

Private Sub ReportFooter_Format()
  Set SubReport1.object = New arEarnDistRegTotalsNS
  EndReport = True
  Label48.Visible = False
  Label49.Visible = False
  Label50.Visible = False
'  Detail.Height = 0
  GroupHeader1.Height = 0
End Sub

Private Sub ReportHeader_Format()
  ReportHeader.Height = 0
End Sub




