VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxPersBillingRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal Tax Billed Report"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arVATaxPersBillingRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arVATaxPersBillingRpt.dsx":08CA
End
Attribute VB_Name = "arVATaxPersBillingRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private hFile As Integer
Private Temp_Class As Resize_Class

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\TAXRPTS\TXPBLRPT.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldCtyPara") '1)
  Fields.Add ("fldCylPara") '2)
  Fields.Add ("fldTSPara") '3)
  Fields.Add ("fldSplPara") '4)
  Fields.Add ("fldYear") '5)
  Fields.Add ("fldBillNum") '6)
  Fields.Add ("fldCustName") '7)
  Fields.Add ("fldPersTax") '8)
  Fields.Add ("fldCustRec") '9)
  Fields.Add ("fldTotal") '10)
  Fields.Add ("fldBillCnt") '11)
  Fields.Add ("fldTotReal") '12)
  Fields.Add ("fldTotPers") '13)
  Fields.Add ("fldGTot") '14)
  Fields.Add ("fldOverPay") '15)
  Fields.Add ("fldNet") '16)
  Fields.Add ("fldTotCredit") '17)
  Fields.Add ("fldTotOwed") '18)
  Fields.Add ("fldTotMTOwed") '19)
  Fields.Add ("fldTotMCOwed") '20)
  Fields.Add ("fldTotFEOwed") '21)
  Fields.Add ("fldTotMHOwed") '22)
  Fields.Add ("fldTotOth") '23)
  Fields.Add ("fldMTTax") '24)
  Fields.Add ("fldMCTax") '25)
  Fields.Add ("fldFETax") '26)
  Fields.Add ("fldMHTax") '27)
  Fields.Add ("fldOthTax") '28)
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
'    Unload frmLoadReport
    frmVATaxMsg.Label1.Caption = "Error Number: " & Str(Number) & " " & Description
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
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
  Fields("fldTown").Value = arr(0)
  Fields("fldCtyPara").Value = arr(1)
  Fields("fldCylPara").Value = arr(2)
  Fields("fldTSPara").Value = arr(3)
  Fields("fldSplPara").Value = arr(4)
  Fields("fldYear").Value = arr(5)
  Fields("fldBillNum").Value = arr(6)
  Fields("fldCustName").Value = arr(7)
  Fields("fldPersTax").Value = arr(8)
  Fields("fldCustRec").Value = arr(9)
  Fields("fldTotal").Value = arr(10)
  Fields("fldBillCnt").Value = arr(11)
  Fields("fldTotReal").Value = arr(12)
  Fields("fldTotPers").Value = arr(13)
  Fields("fldGTot").Value = arr(14)
  Fields("fldOverPay").Value = arr(15)
  Fields("fldNet").Value = arr(16)
  Fields("fldTotCredit").Value = arr(17)
  Fields("fldTotOwed").Value = arr(18)
  Fields("fldTotMTOwed").Value = arr(19)
  Fields("fldTotMCOwed").Value = arr(20)
  Fields("fldTotFEOwed").Value = arr(21)
  Fields("fldTotMHOwed").Value = arr(22)
  Fields("fldTotOth").Value = arr(23)
  Fields("fldMTTax").Value = arr(24)
  Fields("fldMCTax").Value = arr(25)
  Fields("fldFETax").Value = arr(26)
  Fields("fldMHTax").Value = arr(27)
  Fields("fldOthTax").Value = arr(28)
  
End Sub

Private Sub ActiveReport_Initialize()
  Me.Toolbar.Tools.Add "&Close"
  Me.Toolbar.Tools.Add "Save/&Excel"
  Me.Toolbar.Tools.Add "&Text"
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
      frmVATaxMsg.Label1.Caption = "File - TaxPBillRpt.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - TaxPBillRpt.txt, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close
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
    frmVATaxMsg.Label1.Caption = "File - TaxPBillRpt.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - TaxPBillRpt.txt, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
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
        oEXL.FileName = outfile & "TaxPBillRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TaxPBillRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmVATaxLoadReport
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
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Date
  Me.Zoom = -1
End Sub

Private Sub Detail_Format()
'  Detail.Height = 270
''  Label38.Visible = False
'  Field15.Visible = False
'  Label39.Visible = False
'  Field16.Visible = False
'  Line7.Visible = False
'  If Fields("fldOverPay").Value > 0 Then
'    Detail.Height = 585
''    Label38.Visible = True
'    Field15.Visible = True
'    Label39.Visible = True
'    Field16.Visible = True
'    Line7.Visible = True
'  End If
End Sub

Private Sub PageHeader_Format()
'  If Fields("fldSplPara").Value = "REAL" Then
'    Label35.Caption = "Real Due"
'  ElseIf Fields("fldSplPara").Value = "PERSONAL" Then
'    Label35.Caption = "Pers Due"
'  End If
End Sub

Private Sub ReportFooter_Format()
'  Field17.Visible = False
'  Field18.Visible = False
'  Label40.Visible = False
'  Label41.Visible = False
'  ReportFooter.Height = 375
'  If Fields("fldTotCredit").Value > 0 Then
'    Field17.Visible = True
'    Field18.Visible = True
'    Label40.Visible = True
'    Label41.Visible = True
'    ReportFooter.Height = 1065
'  End If
End Sub
