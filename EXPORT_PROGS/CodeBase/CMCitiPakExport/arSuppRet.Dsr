VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSuppRet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplemental Retirement Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arSuppRet.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arSuppRet.dsx":08CA
End
Attribute VB_Name = "arSuppRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private HFile As Integer
Dim EndReport As Boolean
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
      MsgBox "File - 401KRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - 401KRpt.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - 401KRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - 401KRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "401KRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "401KRpt.txt"
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
  HFile = FreeFile
  Open StartPath & "\PRRPTS\401KG.RPT" For Input As #HFile
  Fields.Add "ghGrpHdr1" '0 'group totals on this value
  Fields.Add "fldTitle1ph"
  Fields.Add "fldBBTph"
  Fields.Add "fldEmployerph"
  Fields.Add "fldPRNumph"
  Fields.Add "fldGenLawgh"
  Fields.Add "fldEmpNamedt"
  Fields.Add "fldSSNdt"
  Fields.Add "fldPreTaxdt"
  Fields.Add "fldPostTaxdt"
  Fields.Add "fldEmpContdt"
  Fields.Add "fldDatedt"
  Fields.Add "fldGorLdt"
  Fields.Add "fldSubgf"
  Fields.Add "fldNumOfEmpsgf"
  Fields.Add "fldPreTaxSubgf"
  Fields.Add "fldPostTaxSubgf"
  Fields.Add "fldEmpContSubgf"
  Fields.Add "fldSubTotgf"
  Fields.Add "fldTotVolLawrf"
  Fields.Add "fldTotVolGenrf"
  Fields.Add "fldTotLoanLawrf"
  Fields.Add "fldTotLoanGenrf"
  Fields.Add "fldTotMatchLawrf"
  Fields.Add "fldTotMatchGenrf"
  Fields.Add "fldTotNumofEmpsrf"
  Fields.Add "fldTotGrossrf"
  Fields.Add "fldRptEnd"
  Fields.Add "fldPreTaxSubpf"
  Fields.Add "fldPostTaxSubpf"
  Fields.Add "fldEmpContSubpf"
  Fields.Add "LowDate"
  Fields.Add "EndDate"
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
  Fields("ghGrpHdr1").Value = arr(0) 'group totals on this value
  Fields("fldTitle1ph").Value = arr(1)
  Fields("fldBBTph").Value = arr(2)
  Fields("fldEmployerph").Value = " 401K Sub Plan Name: " + QPTrim$(arr(3))
  Fields("fldPRNumph").Value = arr(4)
  Fields("fldGenLawgh").Value = arr(5)
  Fields("fldEmpNamedt").Value = arr(6)
  Fields("fldSSNdt").Value = arr(7)
  Fields("fldPreTaxdt").Value = arr(8)
  Fields("fldPostTaxdt").Value = arr(9)
  Fields("fldEmpContdt").Value = arr(10)
  Fields("fldDatedt").Value = arr(11)
  Fields("fldGorLdt").Value = arr(12)
  Fields("fldSubgf").Value = arr(13)
  Fields("fldNumOfEmpsgf").Value = arr(14)
  Fields("fldPreTaxSubgf").Value = arr(15)
  Fields("fldPostTaxSubgf").Value = arr(16)
  Fields("fldEmpContSubgf").Value = arr(17)
  Fields("fldSubTotgf").Value = arr(18)
  Fields("fldTotVolLawrf").Value = arr(19)
  
  Fields("fldTotVolGenrf").Value = arr(20)
  Fields("fldTotLoanLawrf").Value = arr(21)
  Fields("fldTotLoanGenrf").Value = arr(22)
  Fields("fldTotMatchLawrf").Value = arr(23)
  Fields("fldTotMatchGenrf").Value = arr(24)
  Fields("fldTotNumofEmpsrf").Value = arr(25)
  Fields("fldTotGrossrf").Value = arr(26)
  Fields("fldRptEnd").Value = arr(27)
  Fields("LowDate").Value = arr(28)
  Fields("EndDate").Value = arr(29)
  Fields("fldPreTaxSubpf").Value = arr(15)
  Fields("fldPostTaxSubpf").Value = arr(16)
  Fields("fldEmpContSubpf").Value = arr(17)
End Sub

Private Sub ActiveReport_PageStart()
  GroupHeader1.GroupValue = Fields("ghGrpHdr1").Value
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmLoadingRpt
  If HFile <> 0 Then
    Close #HFile
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Now
  Me.Zoom = -1
  ReportHeader.Height = 0
  
End Sub
Private Sub Detail_Format()
  GroupHeader1.GroupValue = Fields("ghGrpHdr1").Value
  If GroupHeader1.GroupValue = "" Then
    GroupFooter1.NewPage = 0
  End If

End Sub

Private Sub PageHeader_Format()
  If Fields("fldRptEnd").Value = "N" Then
    PageHeader.Height = 1540
  End If
End Sub

Private Sub ReportFooter_Format()
  Label27.Visible = False
  fldTotGrossrf.Visible = False
End Sub
