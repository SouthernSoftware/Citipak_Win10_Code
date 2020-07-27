VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arRetRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retirement Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12105
   Icon            =   "arRetRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   21352
   _ExtentY        =   15637
   SectionData     =   "arRetRpt.dsx":08CA
End
Attribute VB_Name = "arRetRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private hFile As Integer
Dim LWageTot As Double
Dim LRetTot As Double
Dim LMatchTot As Double
Dim GWageTot As Double
Dim GRetTot As Double
Dim GMatchTot As Double
Dim DetailCnt As Integer
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
      MsgBox "File - RetRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - RetRpt.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - RetRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - RetRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "RetRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "RetRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  hFile = FreeFile
  Open StartPath & "\PRRPTS\RETIREG.RPT" For Input As #hFile
  Fields.Add "fldSSN"
  Fields.Add "fldRetNum"
  Fields.Add "fldEmpName"
  Fields.Add "ghTitle"
  Fields.Add "gfTotal"
  Fields.Add "gfWageTot"
  Fields.Add "gfRetTot"
  Fields.Add "gfMtchTot"
  Fields.Add "ghUnitCode"
  Fields.Add "ghLawGov"
  Fields.Add "SSN"
  Fields.Add "RetNum"
  Fields.Add "EmpName"
  Fields.Add "Wages"
  Fields.Add "RetDed"
  Fields.Add "EmpMatch"
  Fields.Add "fldEmpMatch"
  Fields.Add "fldRetDed"
  Fields.Add "fldMonthYear"
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
  
  If VBA.eof(hFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #hFile, sLine
  arr = Split(sLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  DetailCnt = DetailCnt + 1
  Fields("ghLawGov").Value = arr(0)
  Fields("RetNum").Value = arr(1)
  Fields("EmpName").Value = arr(2)
  Fields("Wages").Value = arr(3)
  Fields("RetDed").Value = arr(4)
  Fields("EmpMatch").Value = arr(5)
  Fields("gfTotal").Value = arr(6)
  Fields("gfWageTot").Value = arr(7)
  Fields("gfRetTot").Value = arr(8)
  Fields("gfMtchTot").Value = arr(9)
  Fields("ghUnitCode").Value = arr(10)
  Fields("ghTitle").Value = arr(11)
  Fields("SSN").Value = arr(12)
  Fields("fldSSN").Value = arr(13)
  Fields("fldRetNum").Value = arr(14)
  Fields("fldEmpName").Value = arr(15)
  Fields("fldRetDed").Value = arr(16)
  Fields("fldEmpMatch").Value = arr(17)
  Fields("fldMonthYear").Value = arr(18)
  Fields("fldEmployer").Value = arr(19)
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmLoadingRpt
  If hFile <> 0 Then
    Close #hFile
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
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Now
  Me.Zoom = -1
End Sub

Private Sub Detail_BeforePrint()
  If DetailCnt = 30 Then
    Detail.NewPage = ddNPAfter
    DetailCnt = 0
  Else
    Detail.NewPage = ddNPNone
  End If

End Sub

Private Sub Detail_Format()
  GroupHeader1.GroupValue = Fields("ghLawGov").Value
End Sub

Private Sub PageFooter_Format()
  If DetailCnt = 0 Then
    PageFooter.Visible = False
  Else
    PageFooter.Visible = True
  End If
End Sub

Private Sub PageHeader_Format()
  DetailCnt = 0

End Sub
