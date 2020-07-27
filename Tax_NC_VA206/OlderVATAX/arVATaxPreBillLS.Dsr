VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxPreBillLS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Property Tax Billing: Pre-Billing Register"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arVATaxPreBillLS.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arVATaxPreBillLS.dsx":08CA
End
Attribute VB_Name = "arVATaxPreBillLS"
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
  Open StartPath & "\TAXRPTS\TaxRealPreBill.RPT" For Input As #hFile
  Fields.Add ("fldCustRec") '0)
  Fields.Add ("fldCustName") '1)
  Fields.Add ("fldBalFor") '2)
  Fields.Add ("fldBillNum") '3)
  Fields.Add ("fldTotDue") '4)
  Fields.Add ("fldRealVal") '5)
  Fields.Add ("fldDisc") '6)
  Fields.Add ("fldTotVal") '7)
  Fields.Add ("fldLateTax") '8)
  Fields.Add ("fldNumBills") '9)
  Fields.Add ("fldTotReal") '10)
  Fields.Add ("fldTotDisc") '11)
  Fields.Add ("fldTotBills") '12)
  Fields.Add ("fldTotPast") '13)
  Fields.Add ("fldTotOwed") '14)
  Fields.Add ("fldTotLate") '15)
  Fields.Add ("fldActive") '16)
  Fields.Add ("fldRealPin") '17)
  Fields.Add ("fldTown") '18)
  Fields.Add ("fldTaxYear") '19)
  Fields.Add ("fldTownship") '20)
  Fields.Add ("fldCycleName") '21)
  Fields.Add ("fldCycleNum") '22)
  Fields.Add ("fldCountyName") '23)
  Fields.Add ("fldCountyNum") '24)
  Fields.Add ("fldOptRev1") '25)
  Fields.Add ("fldOptRev2") '26)
  Fields.Add ("fldOptRev3") '27)
  Fields.Add ("fldOptRevDesc1") '28)
  Fields.Add ("fldOptRevDesc2") '29)
  Fields.Add ("fldOptRevDesc3") '30)
  Fields.Add ("fldTotOpt1") '31)
  Fields.Add ("fldTotOpt2") '32)
  Fields.Add ("fldTotOpt3") '33)
  Fields.Add ("fldCredit") '34)
  Fields.Add ("fldGTotWCredit") '35)
  Fields.Add ("fldTotOverPay") '36)
  Fields.Add ("fldRealRate") '37)
  Fields.Add ("fldLateRate") '38)
  Fields.Add ("fldTotBldgVal") '39)
  Fields.Add ("fldMultiyear") '40)
  Fields.Add ("fldBldgVal") '41)
  Fields.Add ("fldYearHeader")
  Fields.Add ("fldTotWCredit")
  Fields.Add ("fldTotRealVal")
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
  Fields("fldCustRec").Value = arr(0)
  Fields("fldCustName").Value = arr(1)
  Fields("fldBalFor").Value = arr(2)
  Fields("fldBillNum").Value = arr(3)
  Fields("fldTotDue").Value = arr(4)
  Fields("fldRealVal").Value = arr(5)
  Fields("fldDisc").Value = arr(6)
  Fields("fldTotVal").Value = arr(7)
  Fields("fldLateTax").Value = arr(8)
  Fields("fldNumBills").Value = arr(9)
  Fields("fldTotReal").Value = arr(10)
  Fields("fldTotDisc").Value = arr(11)
  Fields("fldTotBills").Value = OldRound(CDbl(arr(12)) + CDbl(arr(15)))
  Fields("fldTotPast").Value = arr(13)
  Fields("fldTotOwed").Value = OldRound(CDbl(arr(15)) + CDbl(arr(14)) - CDbl(arr(36)))
  Fields("fldTotLate").Value = arr(15)
  Fields("fldActive").Value = arr(16)
  Fields("fldRealPin").Value = QPTrim$(arr(17))
  Fields("fldTown").Value = "Town of: " + arr(18)
  Fields("fldTaxYear").Value = arr(19)
  Fields("fldTownship").Value = arr(20)
  Fields("fldCycleName").Value = arr(21)
  Fields("fldCycleNum").Value = arr(22)
  Fields("fldCountyName").Value = arr(23)
  Fields("fldCountyNum").Value = arr(24)
  Fields("fldOptRev1").Value = arr(25)
  Fields("fldOptRev2").Value = arr(26)
  Fields("fldOptRev3").Value = arr(27)
  Fields("fldOptRevDesc1").Value = arr(28)
  Fields("fldOptRevDesc2").Value = arr(29)
  Fields("fldOptRevDesc3").Value = arr(30)
  Fields("fldTotOpt1").Value = arr(31)
  Fields("fldTotOpt2").Value = arr(32)
  Fields("fldTotOpt3").Value = arr(33)
  Fields("fldCredit").Value = arr(34)
  Fields("fldGTotWCredit").Value = OldRound(CDbl(arr(35)) + CDbl(arr(15)))
  Fields("fldTotOverPay").Value = CDbl(arr(36))
  Fields("fldRealRate").Value = arr(37)
  Fields("fldLateRate").Value = arr(38)
  Fields("fldTotBldgVal").Value = arr(39)
  Fields("fldMultiyear").Value = arr(40)
  Fields("fldBldgVal").Value = arr(41)
  
  Fields("fldYearHeader").Value = "For Tax Year: " + arr(19)
  Fields("fldTotWCredit").Value = OldRound(CDbl(arr(4)) + CDbl(arr(34)))
  Fields("fldTotRealVal").Value = OldRound(CDbl(arr(10)) + CDbl(arr(39)) - CDbl(arr(11)))
End Sub

Private Sub ActiveReport_Initialize()
  Line9.Visible = False
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
      frmVATaxMsg.Label1.Caption = "File - TaxPrebillRpt.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - TaxPrebillRpt.txt, created in the Citipak Directory."
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
    frmVATaxMsg.Label1.Caption = "File - TaxPrebillRpt.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - TaxPrebillRpt.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "TaxPrebillRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TaxPrebillRpt.txt"
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
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
End Sub
'
'Private Sub Form_Resize()
'  If Me.WindowState <> vbMinimized Then
'    Me.Visible = False
'    'Temp_Class.ResizeControls Me
'    Me.Visible = True
'    Me.SetFocus
'  End If
'End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Date
  Me.Zoom = -1
End Sub

Private Sub Detail_Format()
  Detail.Height = 270
  Line3.Visible = False
  Line10.Visible = False
  Line11.Visible = False
  Label47.Visible = False
  Field53.Visible = False
  Field54.Visible = False
  Label48.Visible = False
  If Fields("fldOptRev1").Value > 0 Then
    Detail.Height = 960
    Field41.Visible = True
    Field42.Visible = True
    Label47.Visible = True
    Line10.Visible = True
    Line11.Visible = True
    Line3.Visible = True
  Else
    Field41.Visible = False
    Field42.Visible = False
  End If
    
  If Fields("fldOptRev2").Value > 0 Then
    Detail.Height = 960
    Label47.Visible = True
    Field44.Visible = True
    Field43.Visible = True
    Line10.Visible = True
    Line11.Visible = True
    Line3.Visible = True
  Else
    Field44.Visible = False
    Field43.Visible = False
  End If
  
  If Fields("fldOptRev3").Value > 0 Then
    Detail.Height = 960
    Field46.Visible = True
    Field45.Visible = True
    Label47.Visible = True
    Line10.Visible = True
    Line11.Visible = True
    Line3.Visible = True
  Else
    Field46.Visible = False
    Field45.Visible = False
  End If
  
  If Fields("fldCredit").Value <> 0 Then
    Detail.Height = 960
    Field53.Visible = True
    Field54.Visible = True
    Label48.Visible = True
    Line3.Visible = True
  End If
End Sub

Private Sub ReportFooter_Format()
  Dim cnt As Integer
  Dim OptRev1 As Boolean
  Dim OptRev2 As Boolean
  Dim OptRev3 As Boolean
  
  OptRev1 = False
  OptRev2 = False
  OptRev3 = False
  cnt = 0
  Line9.Visible = True
  If Fields("fldTotOpt1").Value > 0 Then
    OptRev1 = True
    cnt = cnt + 1
  End If
  If Fields("fldTotOpt2").Value > 0 Then
    OptRev2 = True
    cnt = cnt + 1
  End If
  If Fields("fldTotOpt3").Value > 0 Then
    OptRev3 = True
    cnt = cnt + 1
  End If
  Line4.Visible = True
  Field47.Visible = True
  Field48.Visible = True
  Line5.Visible = True
  Line6.Visible = True
  Field49.Visible = True
  Field50.Visible = True
  Line7.Visible = True
  Line8.Visible = True
  Field51.Visible = True
  Field52.Visible = True
  Label34.Visible = False
  Label49.Visible = False
  Select Case cnt
    Case 0
      Line4.Visible = False
      Field47.Visible = False
      Field48.Visible = False
      Line5.Visible = False
      Line6.Visible = False
      Field49.Visible = False
      Field50.Visible = False
      Line7.Visible = False
      Line8.Visible = False
      Field51.Visible = False
      Field52.Visible = False
      GoTo Revs
    Case 1
      Line5.Visible = False
      Line6.Visible = False
      Field49.Visible = False
      Field50.Visible = False
      Line7.Visible = False
      Line8.Visible = False
      Field51.Visible = False
      Field52.Visible = False
    Case 2
      Line7.Visible = False
      Line8.Visible = False
      Field51.Visible = False
      Field52.Visible = False
    Case 3
      GoTo Revs
    Case Else
  End Select
  
  If OptRev1 = False Then
    If OptRev2 = True And OptRev3 = True Then
      Field47 = Fields("fldOptRevDesc2").Value
      Field48 = Using$("$###,##0.00", CDbl(Fields("fldTotOpt2").Value))
      Field49 = Fields("fldOptRevDesc3").Value
      Field50 = Using$("$###,##0.00", CDbl(Fields("fldTotOpt3").Value))
    ElseIf OptRev2 = True And OptRev3 = False Then
      Field47 = Fields("fldOptRevDesc2").Value
      Field48 = Using$("$###,##0.00", CDbl(Fields("fldTotOpt2").Value))
    ElseIf OptRev2 = False And OptRev3 = True Then
      Field47 = Fields("fldOptRevDesc3").Value
      Field48 = Using$("$###,##0.00", CDbl(Fields("fldTotOpt3").Value))
    End If
  ElseIf OptRev1 = True Then
    If OptRev2 = False And OptRev3 = True Then
      Field49 = Fields("fldOptRevDesc3").Value
      Field50 = Using$("$###,##0.00", CDbl(Fields("fldTotOpt3").Value))
    End If
  End If
  
Revs:
  Label16.Visible = False
  Label30.Visible = False
  Label23.Visible = False
  Label20.Visible = False
  Label19.Visible = False
  Label31.Visible = False
  Label27.Visible = False
  Label21.Visible = False
  Label33.Visible = False
  Field28.Visible = False
  Label59.Visible = False
End Sub
