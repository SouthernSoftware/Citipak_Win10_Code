VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxLaserPersItemized 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laser Itemized"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "arVATaxLaserPersItemized.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15452
   SectionData     =   "arVATaxLaserPersItemized.dsx":08CA
End
Attribute VB_Name = "arVATaxLaserPersItemized"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ReportFile As String
  Private hFile As Integer
  'Private Temp_Class As Resize_Class

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\TaxPLsrItem.RPT" For Input As #hFile
  Fields.Add ("fldBillNum") '0)
  Fields.Add ("fldCustName") '1)
  Fields.Add ("fldCustAdd1") '2)
  Fields.Add ("fldCustAdd2") '3)
  Fields.Add ("fldCustAdd3") '4)
  Fields.Add ("fldCustPin") '5)
  Fields.Add ("fldTotVal") '6)
  Fields.Add ("fldPDesc1") '7)
  Fields.Add ("fldPersVal") '8)
  Fields.Add ("fldFEVal") '9)
  Fields.Add ("fldExVal") '10)
  Fields.Add ("fldPPTRAVal") '11)
  Fields.Add ("fldPPTRADisc") '12)
  Fields.Add ("fldNetOwed") '13)
  Fields.Add ("flddoLogo") '14)
  Fields.Add ("fldMHVal") '15)
  Fields.Add ("fldMCVal") '16)
  Fields.Add ("fldMTVal") '17)
  Fields.Add ("fldPersTaxDue") '18)
  Fields.Add ("fldPersTaxNet") '19)
  Fields.Add ("fldPerTaxRate") '20)
  Fields.Add ("fldFETaxDue") '21)
  Fields.Add ("fldFETaxRate") '22)
  Fields.Add ("fldMCTaxDue") '23)
  Fields.Add ("fldMCTaxRate") '24)
  Fields.Add ("fldMHTaxDue") '25)
  Fields.Add ("fldMHTaxRate") '26)
  Fields.Add ("fldMTTaxDue") '27)
  Fields.Add ("fldMTTaxRate") '28)
  Fields.Add ("fldOpt1TaxDue") '29)
  Fields.Add ("fldOpt2TaxDue") '30)
  Fields.Add ("fldOpt3TaxDue") '31)
  Fields.Add ("fldOpt1Desc") '32)
  Fields.Add ("fldOpt2Desc") '33)
  Fields.Add ("fldOpt3Desc") '34)
  Fields.Add ("fldBZip") '35)
  Fields.Add ("fldCustZip") '36)
  Fields.Add ("flddoLogo2") '37)
  Fields.Add ("fldHead1") '38)
  Fields.Add ("fldHead2") '39)
  Fields.Add ("fldtxtOpt1") '40)
  Fields.Add ("fldtxtOpt2") '41)
  Fields.Add ("fldtxtOpt3") '42)
  Fields.Add ("fldtxtOpt4") '43)
  Fields.Add ("fldprgf0") '44)
  Fields.Add ("fldprgf1") '45)
  Fields.Add ("fldprgf2") '46)
  Fields.Add ("fldprgf3") '47)
  Fields.Add ("fldprgf4") '48)
  Fields.Add ("fldprgf5") '49)
  Fields.Add ("fldprgf6") '50)
  Fields.Add ("fldprgf7") '51)
  Fields.Add ("fldtxtOpt5") '52)
  Fields.Add ("fldHead3") '53)
  Fields.Add ("fldHead4") '54)
  Fields.Add ("fldHead5") '55)
  Fields.Add ("fldtxtOpt6") '56)
  Fields.Add ("fldtxtOpt7") '57)
  Fields.Add ("fldVIN") '58)
  Fields.Add ("fldMakeModel") '59)
  Fields.Add ("fldthisPersVal") '60)
  Fields.Add ("fldthisMTVal") '61)
  Fields.Add ("fldthisMCVal") '62)
  Fields.Add ("fldthisFEVal") '63)
  Fields.Add ("fldthisMHVal") '64)
  Fields.Add ("fldPCnt") '65) count of how many properties this cust has
  Fields.Add ("fldPrepayAmt") '66)
  Fields.Add ("fldTotalDue") '67)
  Fields.Add ("fldVINOrDesc") '68)
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  Dim ThisPct As Double
  
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
  Fields("fldBillNum").Value = arr(0)
  Fields("fldCustName").Value = arr(1)
  Fields("fldCustAdd1").Value = arr(2)
  Fields("fldCustAdd2").Value = arr(3)
  Fields("fldCustAdd3").Value = arr(4) + " " + QPTrim$(arr(36))
  Fields("fldCustPin").Value = arr(5)
  Fields("fldTotVal").Value = arr(6)
  Fields("fldPDesc1").Value = arr(7)
  Fields("fldPersVal").Value = arr(8)
  Fields("fldFEVal").Value = arr(9)
  Fields("fldExVal").Value = arr(10)
  Fields("fldPPTRAVal").Value = arr(11)
  Fields("fldPPTRADisc").Value = arr(12)
  Fields("fldNetOwed").Value = arr(13)
  Fields("flddoLogo").Value = arr(14)
  Fields("fldMHVal").Value = arr(15)
  Fields("fldMCVal").Value = arr(16)
  Fields("fldMTVal").Value = arr(17)
  Fields("fldPersTaxDue").Value = arr(18)
  Fields("fldPersTaxNet").Value = arr(19)
  If QPTrim$(arr(20)) <> "" Then
    ThisPct = CDbl(arr(20)) / 100
    Fields("fldPerTaxRate").Value = CStr(ThisPct)
  Else
    Fields("fldPerTaxRate").Value = arr(20)
  End If
  Fields("fldFETaxDue").Value = arr(21)
  If QPTrim$(arr(22)) <> "" Then
    ThisPct = CDbl(arr(22)) / 100
    Fields("fldFETaxRate").Value = CStr(ThisPct)
  Else
    Fields("fldFETaxRate").Value = arr(22)
  End If
  Fields("fldMCTaxDue").Value = arr(23)
  If QPTrim$(arr(24)) <> "" Then
    ThisPct = CDbl(arr(24)) / 100
    Fields("fldMCTaxRate").Value = CStr(ThisPct)
  Else
    Fields("fldMCTaxRate").Value = arr(24)
  End If
  Fields("fldMHTaxDue").Value = arr(25)
  If QPTrim$(arr(26)) <> "" Then
    ThisPct = CDbl(arr(26)) / 100
    Fields("fldMHTaxRate").Value = CStr(ThisPct)
  Else
    Fields("fldMHTaxRate").Value = arr(26)
  End If
  Fields("fldMTTaxDue").Value = arr(27)
  If QPTrim$(arr(28)) <> "" Then
    ThisPct = CDbl(arr(28)) / 100
    Fields("fldMTTaxRate").Value = CStr(ThisPct)
  Else
    Fields("fldMTTaxRate").Value = arr(28)
  End If
  Fields("fldOpt1TaxDue").Value = arr(29)
  Fields("fldOpt2TaxDue").Value = arr(30)
  Fields("fldOpt3TaxDue").Value = arr(31)
  Fields("fldOpt1Desc").Value = arr(32)
  Fields("fldOpt2Desc").Value = arr(33)
  Fields("fldOpt3Desc").Value = arr(34)
  Fields("fldBZip").Value = arr(35)
  Fields("fldCustZip").Value = arr(36)
  Fields("flddoLogo2").Value = arr(37)
  Fields("fldHead1").Value = arr(38)
  Fields("fldHead2").Value = arr(39)
  Fields("fldtxtOpt1").Value = arr(40)
  Fields("fldtxtOpt2").Value = arr(41)
  Fields("fldtxtOpt3").Value = arr(42)
  Fields("fldtxtOpt4").Value = arr(43)
  Fields("fldprgf0").Value = arr(44)
  Fields("fldprgf1").Value = arr(45)
  Fields("fldprgf2").Value = arr(46)
  Fields("fldprgf3").Value = arr(47)
  Fields("fldprgf4").Value = arr(48)
  Fields("fldprgf5").Value = arr(49)
  Fields("fldprgf6").Value = arr(50)
  Fields("fldprgf7").Value = arr(51)
  Fields("fldtxtOpt5").Value = arr(52)
  Fields("fldHead3").Value = arr(53)
  Fields("fldHead4").Value = arr(54)
  Fields("fldHead5").Value = arr(55)
  Fields("fldtxtOpt6").Value = arr(56)
  Fields("fldtxtOpt7").Value = arr(57)
  Fields("fldVIN").Value = arr(58)
  Fields("fldMakeModel").Value = arr(59)
  Fields("fldthisPersVal").Value = arr(60)
  Fields("fldthisMTVal").Value = arr(61)
  Fields("fldthisMCVal").Value = arr(62)
  Fields("fldthisFEVal").Value = arr(63)
  Fields("fldthisMHVal").Value = arr(64)
  Fields("fldPCnt").Value = arr(65)
  Fields("fldPrepayAmt").Value = arr(66)
  Fields("fldTotalDue").Value = arr(67)
  If QPTrim$(arr(58)) = "" Then 'VIN number
    Label76.Caption = "Description:"
    If QPTrim$(arr(7)) <> "" Then 'PDesc
      Fields("fldVINOrDesc").Value = QPTrim$(arr(7))
    Else
      Fields("fldVINorDesc").Value = "No Description"
    End If
  Else
    Label76.Caption = "VIN Number:"
    Fields("fldVINorDesc").Value = QPTrim$(arr(58))
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
    End If
    If KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - TxLsrPItem.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - TxLsrPItem.txt, created in the Citipak Directory.", vbOKOnly
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
  Unload frmVATaxLoadingRpt
'  Me.Show 1
End Sub
Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
  If Tool = "&Close" Then
    Close #hFile
    Unload Me
  End If
  If Tool = "Save/&Excel" Then
    Screen.MousePointer = vbHourglass
    ExportReport 1
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - TxLsrPItem.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - TxLsrPItem.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "TxLsrPItem.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TxLsrPItem.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub Detail_Format()
  Dim Opt1 As Boolean
  Dim Opt2 As Boolean
  Dim Opt3 As Boolean
  Dim ThisLen As Integer
  Dim NumOfOpts As Integer
  
  Opt1 = False
  Opt2 = False
  Opt3 = False
  If QPTrim$(Fields("fldOpt1Desc").Value) <> "" Then
    Opt1 = True
    NumOfOpts = NumOfOpts + 1
  End If
  If QPTrim$(Fields("fldOpt2Desc").Value) <> "" Then
    Opt2 = True
    NumOfOpts = NumOfOpts + 1
  End If
  If QPTrim$(Fields("fldOpt3Desc").Value) <> "" Then
    Opt3 = True
    NumOfOpts = NumOfOpts + 1
  End If

  If QPTrim$(Fields("flddoLogo2").Value) = "1" Then
    If Exist("towntaxlogo.bmp") Then
      DoEvents
      Image1.Picture = LoadPicture("towntaxlogo.bmp")
      Image1.Visible = True
      
      DoEvents
    End If
  End If
  
  If NumOfOpts = 1 Then
    ThisLen = 13
  ElseIf NumOfOpts = 2 Then
    ThisLen = 12
  ElseIf NumOfOpts = 3 Then
    ThisLen = 11
  Else
    ThisLen = 14
  End If
    
  If CInt(Fields("fldPCnt").Value) = ThisLen Then
    Detail.NewPage = ddNPBefore
  ElseIf CInt(Fields("fldPCnt").Value) < ThisLen Then
    Detail.NewPage = ddNPNone
  End If
  
End Sub

Private Sub GroupHeader1_Format()
  Dim Opt1 As Boolean
  Dim Opt2 As Boolean
  Dim Opt3 As Boolean

  Opt1 = False
  Opt2 = False
  Opt3 = False
  If QPTrim$(Fields("fldOpt1Desc").Value) = "" Then
    Label88.Caption = ""
    Field218.Visible = False
  Else
    Label88.Caption = QPTrim$(Fields("fldOpt1Desc").Value)
    Opt1 = True
  End If
  If QPTrim$(Fields("fldOpt2Desc").Value) = "" Then
    Label89.Caption = ""
    Field219.Visible = False
  Else
    Label89.Caption = QPTrim$(Fields("fldOpt2Desc").Value)
    Opt2 = True
  End If
  If QPTrim$(Fields("fldOpt3Desc").Value) = "" Then
    Label90.Caption = ""
    Field220.Visible = False
  Else
    Label90.Caption = QPTrim$(Fields("fldOpt3Desc").Value)
    Opt3 = True
  End If

  If Opt1 = False And Opt2 = False And Opt3 = False Then
    Label88.Visible = False
    Field218.Visible = False
    Label89.Visible = False
    Field219.Visible = False
    Label90.Visible = False
    Field220.Visible = False
    Line35.Y1 = 4410
    Line35.Y2 = 4410
    Label87.Top = 4410
    GroupHeader1.Height = 4590
    Line38.Y1 = 4590
    Line38.Y2 = 4590
  ElseIf Opt1 = True And Opt2 = False And Opt3 = False Then
    Label88.Visible = True
    Field218.Visible = True
    Label89.Visible = False
    Field219.Visible = False
    Label90.Visible = False
    Field220.Visible = False
    Line35.Y1 = 4680
    Line35.Y2 = 4680
    Label87.Top = 4680
    GroupHeader1.Height = 4860
    Line38.Y1 = 4860
    Line38.Y2 = 4860
  ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
    Label88.Visible = False
    Field218.Visible = False
    Label89.Visible = True
    Field219.Visible = True
    Label89.Top = 4410
    Field219.Top = 4410
    Label90.Visible = False
    Field220.Visible = False
    Line35.Y1 = 4680
    Line35.Y2 = 4680
    Label87.Top = 4680
    GroupHeader1.Height = 4860
    Line38.Y1 = 4860
    Line38.Y2 = 4860
  ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
    Label88.Visible = False
    Field218.Visible = False
    Label89.Visible = False
    Field219.Visible = False
    Label90.Visible = True
    Field220.Visible = True
    Label90.Top = 4410
    Field220.Top = 4410
    Line35.Y1 = 4680
    Line35.Y2 = 4680
    Label87.Top = 4680
    GroupHeader1.Height = 4860
    Line38.Y1 = 4860
    Line38.Y2 = 4860
  ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
    Label88.Visible = True
    Field218.Visible = True
    Label89.Visible = True
    Field219.Visible = True
    Label90.Visible = False
    Field220.Visible = False
    Line35.Y1 = 4950
    Line35.Y2 = 4950
    Label87.Top = 4950
    GroupHeader1.Height = 5130
    Line38.Y1 = 5130
    Line38.Y2 = 5130
  ElseIf Opt1 = True And Opt2 = False And Opt3 = True Then
    Label88.Visible = True
    Field218.Visible = True
    Label89.Visible = False
    Field219.Visible = False
    Label90.Visible = True
    Field220.Visible = True
    Label90.Top = 4680
    Field220.Top = 4680
    Line35.Y1 = 4950
    Line35.Y2 = 4950
    Label87.Top = 4950
    GroupHeader1.Height = 5130
    Line38.Y1 = 5130
    Line38.Y2 = 5130
  ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
    Label88.Visible = False
    Field218.Visible = False
    Label89.Visible = True
    Field219.Visible = True
    Label89.Top = 4410
    Field219.Top = 4410
    Label90.Visible = True
    Field220.Visible = True
    Label90.Top = 4680
    Field220.Top = 4680
    Line35.Y1 = 4950
    Line35.Y2 = 4950
    Label87.Top = 4950
    GroupHeader1.Height = 5130
    Line38.Y1 = 5130
    Line38.Y2 = 5130
  End If
End Sub
