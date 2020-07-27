VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxPreBillPers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Personal Prebilling Report"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   12150
   Icon            =   "arVATaxPreBillPers.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   21431
   _ExtentY        =   15452
   SectionData     =   "arVATaxPreBillPers.dsx":08CA
End
Attribute VB_Name = "arVATaxPreBillPers"
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
  Open StartPath & "\TAXRPTS\TaxPersPreBill.RPT" For Input As #hFile
  Fields.Add ("fldCustRec") '0)
  Fields.Add ("fldCustName") '1)
  Fields.Add ("fldBalFor") '2)
  Fields.Add ("fldBillNum") '3)
  Fields.Add ("fldTotDue") '4)
  Fields.Add ("fldTotVal") '5)
  Fields.Add ("fldLateTax") '6)
  Fields.Add ("fldNumBills") '7)
  Fields.Add ("fldTotPers") '8)
  Fields.Add ("fldTotBills") '9)
  Fields.Add ("fldTotPast") '10)
  Fields.Add ("fldTotOwed") '11)
  Fields.Add ("fldTotLate") '12)
  Fields.Add ("fldActive") '13)
  Fields.Add ("fldTown") '14)
  Fields.Add ("fldTaxYear") '15)
  Fields.Add ("fldTownship") '16)
  Fields.Add ("fldCycleName") '17)
  Fields.Add ("fldCycleNum") '18)
  Fields.Add ("fldCountyName") '19)
  Fields.Add ("fldCountyNum") '20)
  Fields.Add ("fldCredit") '21)
  Fields.Add ("fldGTotWCredit") '22)
  Fields.Add ("fldTotOverPay") '23)
  Fields.Add ("fldPersRate") '24)
  Fields.Add ("fldLateRate") '25)
  Fields.Add ("fldFERate") '26)
  Fields.Add ("fldMHRate") '27)
  Fields.Add ("fldMCRate") '28)
  Fields.Add ("fldMTRate") '29)
  Fields.Add ("fldPersTaxNet") '30)
  Fields.Add ("fldMHTax") '31)
  Fields.Add ("fldMTTax") '32)
  Fields.Add ("fldMCTax") '33)
  Fields.Add ("fldFETax") '34)
  Fields.Add ("fldPersVal") '35)
  Fields.Add ("fldMHVal") '36)
  Fields.Add ("fldMTVal") '37)
  Fields.Add ("fldMCVal") '38)
  Fields.Add ("fldFEVal") '39)
  Fields.Add ("fldTotPersTax") '40)
  Fields.Add ("fldTotMHTax") '41)
  Fields.Add ("fldTotMTTax") '42)
  Fields.Add ("fldTotMCTax") '43)
  Fields.Add ("fldTotFETax") '44)
  Fields.Add ("fldTotPersVal") '45)
  Fields.Add ("fldTotMHVal") '46)
  Fields.Add ("fldTotMTVal") '47)
  Fields.Add ("fldTotMCVal") '48)
  Fields.Add ("fldTotFEVal") '49)
  Fields.Add ("fldPPTRADisc") '50)
  Fields.Add ("fldPPTRAVal") '51)
  Fields.Add ("fldTotPPTRADisc") '52)
  Fields.Add ("fldTotPPTRAVal") '53)
  Fields.Add ("fldPersTax") '54)
  Fields.Add ("fldNetPersVal") '55)
  Fields.Add ("fldGTPersTaxNet") '56)
  Fields.Add ("fldPERC") '57)
  Fields.Add ("fldMultiYear") '58)
  Fields.Add ("fldThisPPTRADisc") '59)
  Fields.Add ("fldPPTRAYN") '60)
  Fields.Add ("fldOptTax1") '61)
  Fields.Add ("fldOptTax2") '62)
  Fields.Add ("fldOptTax3") '63)
  Fields.Add ("fldOptDesc1") '64)
  Fields.Add ("fldOptDesc2") '65)
  Fields.Add ("fldOptDesc3") '66)
  Fields.Add ("fldOptTot1") '67)
  Fields.Add ("fldOptTot2") '68)
  Fields.Add ("fldOptTot3") '69)
  Fields.Add ("fldYearHeader")
  Fields.Add ("fldTotWCredit")
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
  Fields("fldTotVal").Value = arr(5)
  Fields("fldLateTax").Value = arr(6)
  Fields("fldNumBills").Value = arr(7)
  Fields("fldTotPers").Value = arr(8)
  Fields("fldTotBills").Value = arr(9)
  Fields("fldTotPast").Value = arr(10)
  Fields("fldTotOwed").Value = arr(11)
  Fields("fldTotLate").Value = arr(12)
  Fields("fldActive").Value = arr(13)
  Fields("fldTown").Value = arr(14)
  Fields("fldTaxYear").Value = arr(15)
  Fields("fldTownship").Value = arr(16)
  Fields("fldCycleName").Value = arr(17)
  Fields("fldCycleNum").Value = arr(18)
  Fields("fldCountyName").Value = arr(19)
  Fields("fldCountyNum").Value = arr(20)
  Fields("fldCredit").Value = arr(21)
  Fields("fldGTotWCredit").Value = arr(22)
  Fields("fldTotOverPay").Value = arr(23)
  Fields("fldPersRate").Value = arr(24)
  Fields("fldLateRate").Value = arr(25)
  Fields("fldFERate").Value = arr(26)
  Fields("fldMHRate").Value = arr(27)
  Fields("fldMCRate").Value = arr(28)
  Fields("fldMTRate").Value = arr(29)
  Fields("fldPersTaxNet").Value = arr(30)
  Fields("fldMHTax").Value = arr(31)
  Fields("fldMCTax").Value = arr(32)
  Fields("fldMTTax").Value = arr(33)
  Fields("fldFETax").Value = arr(34)
  Fields("fldPersVal").Value = arr(35)
  Fields("fldMHVal").Value = arr(36)
  Fields("fldMTVal").Value = arr(37)
  Fields("fldMCVal").Value = arr(38)
  Fields("fldFEVal").Value = arr(39)
  Fields("fldTotPersTax").Value = arr(40)
  Fields("fldTotMHTax").Value = arr(41)
  Fields("fldTotMTTax").Value = arr(42)
  Fields("fldTotMCTax").Value = arr(43)
  Fields("fldTotFETax").Value = arr(44)
  Fields("fldTotPersVal").Value = arr(45)
  Fields("fldTotMHVal").Value = arr(46)
  Fields("fldTotMTVal").Value = arr(47)
  Fields("fldTotMCVal").Value = arr(48)
  Fields("fldTotFEVal").Value = arr(49)
  Fields("fldPPTRADisc").Value = arr(50)
  Fields("fldPPTRAVal").Value = arr(51)
  Fields("fldTotPPTRADisc").Value = arr(52)
  Fields("fldTotPPTRAVal").Value = arr(53)
  Fields("fldPersTax").Value = arr(54)
  Fields("fldNetPersVal").Value = arr(55)
  Fields("fldGTPersTaxNet").Value = arr(56)
  Fields("fldPERC").Value = arr(57)
  Fields("fldMultiYear").Value = arr(58)
  Fields("fldThisPPTRADisc").Value = arr(59)
  Fields("fldPPTRAYN").Value = arr(60)
  Fields("fldOptTax1").Value = arr(61)
  Fields("fldOptTax2").Value = arr(62)
  Fields("fldOptTax3").Value = arr(63)
  Fields("fldOptDesc1").Value = arr(64)
  Fields("fldOptDesc2").Value = arr(65)
  Fields("fldOptDesc3").Value = arr(66)
  Fields("fldOptTot1").Value = arr(67)
  Fields("fldOptTot2").Value = arr(68)
  Fields("fldOptTot3").Value = arr(69)
  Fields("fldYearHeader").Value = "For Tax Year: " + arr(15)
  Fields("fldTotWCredit").Value = OldRound(CDbl(arr(4)) + CDbl(arr(21)))
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
      frmVATaxMsg.Label1.Caption = "File - TaxPersPrebillRpt.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - TaxPersPrebillRpt.txt, created in the Citipak Directory."
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
    frmVATaxMsg.Label1.Caption = "File - TaxPersPrebillRpt.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - TaxPersPrebillRpt.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "TaxPersPrebillRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TaxPersPrebillRpt.txt"
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

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    'Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Date
  Me.Zoom = -1
End Sub

Private Sub Detail_Format()
  Detail.Height = 1116
  Field79.Visible = True
  Field80.Visible = True
  Label86.Visible = True
  Label87.Visible = True
  Line10.Visible = True
  Line12.Visible = True
  Line13.Visible = True
  Line11.Visible = True
  Line14.Visible = True
  Line15.Visible = True
  Label48.Visible = True
  Field53.Visible = True
  Field54.Visible = True

  If CDbl(Fields("fldPPTRADisc").Value) = 0 Then 'no PPTRA discount
    Field79.Visible = False 'so remove those fields and adjust height
    Field80.Visible = False 'to accommodate an overpayment
    Label86.Visible = False
    Label87.Visible = False
    Line10.Visible = False
    Line12.Visible = False
    Line13.Visible = False
    Line11.Visible = False
    Line14.Visible = False
    Line15.Visible = False
    Detail.Height = 810
    If CDbl(Fields("fldCredit").Value) = 0 Then 'no overpayment
      Label48.Visible = False 'so remove those fields and adjust
      Field53.Visible = False 'height accordingly
      Field54.Visible = False
      Detail.Height = 540
    End If
  ElseIf CDbl(Fields("fldThisPPTRADisc").Value) > 0 Then 'PPTRA discount does
  'exist so detail height remains at max to accommodate those fields
    If CDbl(Fields("fldCredit").Value) = 0 Then 'but no overpayment
      Label48.Visible = False 'so remove those fields
      Field53.Visible = False
      Field54.Visible = False
    End If
  End If
    
  Field42.Visible = False
  Field41.Visible = False
  Field44.Visible = False
  Field43.Visible = False
  Field46.Visible = False
  Field45.Visible = False
  If Fields("fldOptTax1").Value > 0 Or Fields("fldOptTax2").Value > 0 Or Fields("fldOptTax3").Value > 0 Then
    Detail.Height = 1116
    If Fields("fldOptTax1").Value > 0 And Fields("fldOptTax2").Value > 0 And Fields("fldOptTax3").Value > 0 Then
      Field42.Visible = True
      Field41.Visible = True
      Field44.Visible = True
      Field43.Visible = True
      Field46.Visible = True
      Field45.Visible = True
    ElseIf Fields("fldOptTax1").Value > 0 And Fields("fldOptTax2").Value = 0 And Fields("fldOptTax3").Value = 0 Then
      Field42.Visible = True
      Field41.Visible = True
    ElseIf Fields("fldOptTax1").Value > 0 And Fields("fldOptTax2").Value > 0 And Fields("fldOptTax3").Value = 0 Then
      Field42.Visible = True
      Field41.Visible = True
      Field44.Visible = True
      Field43.Visible = True
    ElseIf Fields("fldOptTax1").Value = 0 And Fields("fldOptTax2").Value > 0 And Fields("fldOptTax3").Value = 0 Then
      Field44.Visible = True
      Field43.Visible = True
    ElseIf Fields("fldOptTax1").Value = 0 And Fields("fldOptTax2").Value > 0 And Fields("fldOptTax3").Value > 0 Then
      Field44.Visible = True
      Field43.Visible = True
      Field46.Visible = True
      Field45.Visible = True
    ElseIf Fields("fldOptTax1").Value > 0 And Fields("fldOptTax2").Value = 0 And Fields("fldOptTax3").Value > 0 Then
      Field42.Visible = True
      Field41.Visible = True
      Field46.Visible = True
      Field45.Visible = True
    ElseIf Fields("fldOptTax1").Value = 0 And Fields("fldOptTax2").Value = 0 And Fields("fldOptTax3").Value > 0 Then
      Field46.Visible = True
      Field45.Visible = True
    End If
  End If
End Sub

Private Sub PageHeader_Format()
  If Fields("fldPPTRAYN").Value = "False" Then
    Label90.Visible = False
    Field89.Visible = False
  End If
End Sub

Private Sub ReportFooter_Format()
  Dim Opt1 As Boolean
  Dim Opt2 As Boolean
  Dim Opt3 As Boolean
  
  Label16.Visible = False
  Label23.Visible = False
  Label29.Visible = False
  Label77.Visible = False
  Label73.Visible = False
  Label78.Visible = False
  Label79.Visible = False
  Label80.Visible = False
  Label81.Visible = False
  Label82.Visible = False
  Label83.Visible = False
  Label84.Visible = False
  Label27.Visible = False
  Label21.Visible = False
  Label33.Visible = False
  Field28.Visible = False
  Label92.Visible = False 'opt1
  Label93.Visible = False 'opt2
  Label94.Visible = False 'opt3
  Field91.Visible = False 'opt1
  Field92.Visible = False 'opt2
  Field93.Visible = False 'opt3
  Line9.Visible = False
  Line38.Visible = False 'opt
  Line39.Visible = False 'opt
  Line40.Visible = False 'opt
  Line41.Visible = False 'opt
  PageHeader.Height = 2390
  
  Opt1 = False
  Opt2 = False
  Opt3 = False
  
  If Fields("fldOptTot1").Value > 0 Or Fields("fldOptTot2").Value > 0 Or Fields("fldOptTot3").Value > 0 Then
    If Fields("fldOptTot1").Value > 0 Then Opt1 = True
    If Fields("fldOptTot2").Value > 0 Then Opt2 = True
    If Fields("fldOptTot3").Value > 0 Then Opt3 = True
    Line38.Visible = True
    Line39.Visible = True
    Line40.Visible = True
    Line41.Visible = True
  End If
    
  If Opt1 = True And Opt2 = True And Opt3 = True Then
    Label92.Visible = True
    Label92.Caption = "Total " + Fields("fldOptDesc1").Value + ":"
    Field91.Visible = True
    Label93.Visible = True
    Label93.Caption = "Total " + Fields("fldOptDesc2").Value + ":"
    Field92.Visible = True
    Label94.Visible = True
    Label94.Caption = "Total " + Fields("fldOptDesc3").Value + ":"
    Field93.Visible = True
  ElseIf Opt1 = True And Opt2 = False And Opt3 = False Then
    Label92.Visible = True
    Label92.Caption = "Total " + Fields("fldOptDesc1").Value + ":"
    Field91.Visible = True
    Line39.Y2 = 5760
    Line40.Y2 = 5760
    Line41.Y1 = 5760
    Line41.Y2 = 5760
  ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
    Label93.Visible = True
    Label93.Caption = "Total " + Fields("fldOptDesc2").Value + ":"
    Label93.Top = 5490
    Field92.Visible = True
    Field92.Top = 5490
    Line39.Y2 = 5760
    Line40.Y2 = 5760
    Line41.Y1 = 5760
    Line41.Y2 = 5760
  ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
    Label94.Visible = True
    Label94.Caption = "Total " + Fields("fldOptDesc3").Value + ":"
    Label94.Top = 5490
    Field93.Visible = True
    Field93.Top = 5490
    Line39.Y2 = 5760
    Line40.Y2 = 5760
    Line41.Y1 = 5760
    Line41.Y2 = 5760
  ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
    Label92.Visible = True
    Label92.Caption = "Total " + Fields("fldOptDesc1").Value + ":"
    Field91.Visible = True
    Label93.Visible = True
    Label93.Caption = "Total " + Fields("fldOptDesc2").Value + ":"
    Field92.Visible = True
    Line39.Y2 = 6030
    Line40.Y2 = 6030
    Line41.Y1 = 6030
    Line41.Y2 = 6030
  ElseIf Opt1 = True And Opt2 = False And Opt3 = True Then
    Label92.Visible = True
    Label92.Caption = "Total " + Fields("fldOptDesc1").Value + ":"
    Field91.Visible = True
    Label94.Visible = True
    Label94.Caption = "Total " + Fields("fldOptDesc3").Value + ":"
    Label94.Top = 5760
    Field93.Visible = True
    Field93.Top = 5760
    Line39.Y2 = 6030
    Line40.Y2 = 6030
    Line41.Y1 = 6030
    Line41.Y2 = 6030
  ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
    Label93.Visible = True
    Label93.Caption = "Total " + Fields("fldOptDesc2").Value + ":"
    Label93.Top = 5490
    Field92.Visible = True
    Field92.Top = 5490
    Label94.Visible = True
    Label94.Caption = "Total " + Fields("fldOptDesc3").Value + ":"
    Label94.Top = 5760
    Field93.Visible = True
    Field93.Top = 5760
    Line39.Y2 = 6030
    Line40.Y2 = 6030
    Line41.Y1 = 6030
    Line41.Y2 = 6030
  End If
End Sub
