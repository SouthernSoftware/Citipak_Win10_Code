VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arRetRptSC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retirement Report for South Carolina"
   ClientHeight    =   8868
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arRetRptSC.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arRetRptSC.dsx":08CA
End
Attribute VB_Name = "arRetRptSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Private HFile As Integer
  Dim LWageTot As Double
  Dim LRetTot As Double
  Dim LMatchTot As Double
  Dim GWageTot As Double
  Dim GRetTot As Double
  Dim GMatchTot As Double
  Dim DetailCnt As Integer
  Dim cnt As Integer
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
      MsgBox "File - RetRptSC.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - RetRptSC.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close
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
    MsgBox "File - RetRptSC.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - RetRptSC.txt, created in the Citipak Directory.", vbOKOnly
  End If
End Sub
Private Sub ExportReport(X As Integer)
  Dim oEXL As ActiveReportsExcelExport.ARExportExcel
  Dim oTXT As ActiveReportsTextExport.ARExportText
  Dim outfile As String
  If Right$(StartPath, 1) = ":" Then
    outfile = StartPath
  Else
    outfile = StartPath & "\"
  End If
  Select Case X
    Case 1   '"Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & "RetRptSC.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "RetRptSC.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  HFile = FreeFile
  Open StartPath & "\PRRPTS\SCRETIREG.RPT" For Input As #HFile
  Fields.Add "SSNdt" '0
  Fields.Add "EmpName" '1
  Fields.Add "RetWages" '2
  Fields.Add "RetDed" '3
  Fields.Add "EmprMatch" '4
  Fields.Add "gfTotal" '5
  Fields.Add "gfWageTot" '6
  Fields.Add "gfRetTot" '7
  Fields.Add "gfMtchTot" '8
  Fields.Add "ghTitle" '9
  Fields.Add "ghLawGov" '10
  Fields.Add "fldSSN" '11
  Fields.Add "fldEmpName" '12
  Fields.Add "fldRetDed" '13
  Fields.Add "fldEmpMatch" '14
  Fields.Add "fldStartEnd"
  Fields.Add "fldEmployer"
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmLoadingRpt
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True 'removes the error message
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
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
  cnt = cnt + 1
  Fields("SSNdt").Value = arr(0)
  Fields("EmpName").Value = arr(1)
  Fields("RetWages").Value = arr(2)
  Fields("RetDed").Value = arr(3)
  Fields("EmprMatch").Value = arr(4)
  Fields("gfTotal").Value = arr(5)
  Fields("gfWageTot").Value = arr(6)
  Fields("gfRetTot").Value = arr(7)
  Fields("gfMtchTot").Value = arr(8)
  Fields("ghTitle").Value = arr(9)
  Fields("ghLawGov").Value = arr(10)
  Fields("fldSSN").Value = arr(11)
  Fields("fldEmpName").Value = arr(12)
  Fields("fldRetDed").Value = arr(13)
  Fields("fldEmpMatch").Value = arr(14)
  Fields("fldStartEnd").Value = arr(15)
  Fields("fldEmployer").Value = arr(16)
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmLoadingRpt
  If HFile <> 0 Then
    Close #HFile
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Now
  Me.Zoom = -1

End Sub

Private Sub Detail_Format()
  GroupHeader1.GroupValue = Fields("ghLawGov").Value
End Sub

Private Sub GroupFooter1_Format()
  cnt = cnt
  If Fields("ghLawGov").Value = "N" Then
    txtRetDed.Visible = False
    txtEmprMatch.Visible = False
  End If
End Sub

Private Sub PageHeader_Format()
  If GroupHeader1.GroupValue = "VA" Then
    PageHeader.Height = 2000
  End If
  DetailCnt = 0

End Sub

