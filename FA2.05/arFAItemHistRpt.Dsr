VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arFAItemHistRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Asset Item History Report"
   ClientHeight    =   8868
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arFAItemHistRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arFAItemHistRpt.dsx":08CA
End
Attribute VB_Name = "arFAItemHistRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsFATextBoxOverRider
Private Temp_Class As Resize_Class
Private HFile As Integer

Private Sub ActiveReport_Initialize()
  Me.ToolBar.Tools.Add "&Close"
  Me.ToolBar.Tools.Add "Save/&Excel"
  Me.ToolBar.Tools.Add "&Text"
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
    ElseIf KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - FAItemHist.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - FAItemHist.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
End Sub
Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
  If Tool = "&Close" Then
    Close
    Unload Me
  End If
  If Tool = "Save/&Excel" Then
    Screen.MousePointer = vbHourglass
    ExportReport 1
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - FAItemHist.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - FAItemHist.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "FAItemHist.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "FAItemHist.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub
Private Sub ActiveReport_DataInitialize()
  HFile = FreeFile
  Open StartPath & "\FARPTS\FAITEMHIST.RPT" For Input As #HFile
  Fields.Add "fldEmployer" '0)
  Fields.Add "fldItemTag" '1)
  Fields.Add "fldItemDesc" '2)
  Fields.Add "fldDeptNum" '3)
  Fields.Add "fldOrigCost" '4)
  Fields.Add "fldThisYear" '5)
  Fields.Add "fldLife" '6)
  Fields.Add "fldLifeLeft" '7)
  Fields.Add "fldThisDepr" '8)
  Fields.Add "fldDepTotal" '9)
  Fields.Add "fldBookTotal" '10)
  Fields.Add "fldAcquireDate" '11)
  Fields.Add "fldDisposalDate" '12)
  End Sub
Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  Dim x As Integer
  If VBA.eof(HFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #HFile, sLine
  arr = Split(sLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("fldEmployer").Value = arr(0)
  Fields("fldItemTag").Value = arr(1)
  Fields("fldItemDesc").Value = arr(2)
  Fields("fldDeptNum").Value = arr(3)
  Fields("fldOrigCost").Value = arr(4)
  Fields("fldThisYear").Value = arr(5)
  Fields("fldLife").Value = arr(6)
  Fields("fldLifeLeft").Value = arr(7)
  Fields("fldThisDepr").Value = arr(8)
  Fields("fldDepTotal").Value = arr(9)
  Fields("fldBookTotal").Value = arr(10)
  Fields("fldAcquireDate").Value = arr(11)
  If arr(12) <> "" Then
    Fields("fldDisposalDate").Value = "*D*"
  Else
    Fields("fldDisposalDate").Value = ""
  End If
End Sub
Private Sub ActiveReport_ReportEnd()
  Unload frmFALoadReport
  If HFile <> 0 Then
    Close #HFile
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
'    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Label11.Visible = False
  Me.fldTimeDate.Text = Now
  Me.Zoom = -1
End Sub


Private Sub ReportFooter_Format()
  PageHeader.Height = 1100
  Set SubReport1 = New arSubItemHistRpt
  Label11.Visible = True
  Label16.Visible = False
  Label17.Visible = False
  Label20.Visible = False
  Label21.Visible = False
  Label35.Visible = False
  Line1.Visible = False
  
End Sub
