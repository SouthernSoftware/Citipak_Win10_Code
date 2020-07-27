VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxLaserRealItemized 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laser Itemized"
   ClientHeight    =   8712
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   12132
   Icon            =   "arVATaxLaserRealItemized.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   21400
   _ExtentY        =   15367
   SectionData     =   "arVATaxLaserRealItemized.dsx":08CA
End
Attribute VB_Name = "arVATaxLaserRealItemized"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ReportFile As String
  Private hFile As Integer
  Private Temp_Class As Resize_Class

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\TaxRLsrItem.RPT" For Input As #hFile
  Fields.Add ("fldBillNum") '0)
  Fields.Add ("fldCustName") '1)
  Fields.Add ("fldCustAdd1") '2)
  Fields.Add ("fldCustAdd2") '3)
  Fields.Add ("fldCustAdd3") '4)
  Fields.Add ("fldCustPin") '5)
  Fields.Add ("fldTotVal") '6)
  Fields.Add ("fldRDesc") '7)
  Fields.Add ("fldRealVal") '8)
  Fields.Add ("fldBldgVal") '9)
  Fields.Add ("fldExVal") '10)
  Fields.Add ("fldNetOwed") '11)
  Fields.Add ("flddoLogo") '12)
  Fields.Add ("fldRealTaxDue") '13)
  Fields.Add ("fldRealTaxNet") '14)
  Fields.Add ("fldRealTaxRate") '15)
  Fields.Add ("fldOpt1TaxDue") '16)
  Fields.Add ("fldOpt2TaxDue") '17)
  Fields.Add ("fldOpt3TaxDue") '18)
  Fields.Add ("fldOpt1Desc") '19)
  Fields.Add ("fldOpt2Desc") '20)
  Fields.Add ("fldOpt3Desc") '21)
  Fields.Add ("fldBZip") '22)
  Fields.Add ("fldCustZip") '23)
  Fields.Add ("flddoLogo2") '24)
  Fields.Add ("fldHead1") '25)
  Fields.Add ("fldHead2") '26)
  Fields.Add ("fldtxtOpt1") '27)
  Fields.Add ("fldtxtOpt2") '28)
  Fields.Add ("fldtxtOpt3") '29)
  Fields.Add ("fldprgf0") '30)
  Fields.Add ("fldprgf1") '31)
  Fields.Add ("fldprgf2") '32)
  Fields.Add ("fldprgf3") '33)
  Fields.Add ("fldprgf4") '34)
  Fields.Add ("fldHead3") '35)
  Fields.Add ("fldHead4") '36)
  Fields.Add ("fldHead5") '37)
  Fields.Add ("fldthisRealVal") '38)
  Fields.Add ("fldRCnt") '39) count of how many properties this cust has
  Fields.Add ("fldMap") '40)
  Fields.Add ("fldBlock") '41)
  Fields.Add ("fldLot") '42)
  Fields.Add ("fldRealAdd") '43)
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
  Fields("fldRDesc").Value = arr(7)
  Fields("fldRealVal").Value = arr(8)
  Fields("fldBldgVal").Value = arr(9)
  Fields("fldExVal").Value = arr(10)
  Fields("fldNetOwed").Value = arr(11)
  Fields("flddoLogo").Value = arr(12)
  Fields("fldRealTaxDue").Value = arr(13)
  Fields("fldRealTaxNet").Value = arr(14)
  If QPTrim$(arr(15)) <> "" Then
    ThisPct = CDbl(arr(15)) / 100
    Fields("fldRealTaxRate").Value = CStr(ThisPct)
  Else
    Fields("fldRealTaxRate").Value = arr(15)
  End If
  Fields("fldOpt1TaxDue").Value = arr(16)
  Fields("fldOpt2TaxDue").Value = arr(17)
  Fields("fldOpt3TaxDue").Value = arr(18)
  Fields("fldOpt1Desc").Value = arr(19)
  Fields("fldOpt2Desc").Value = arr(20)
  Fields("fldOpt3Desc").Value = arr(21)
  Fields("fldBZip").Value = arr(22)
  Fields("fldCustZip").Value = arr(23)
  Fields("flddoLogo2").Value = arr(24)
  Fields("fldHead1").Value = arr(25)
  Fields("fldHead2").Value = arr(26)
  Fields("fldtxtOpt1").Value = arr(27)
  Fields("fldtxtOpt2").Value = arr(28)
  Fields("fldtxtOpt3").Value = arr(29)
  Fields("fldprgf0").Value = arr(30)
  Fields("fldprgf1").Value = arr(31)
  Fields("fldprgf2").Value = arr(32)
  Fields("fldprgf3").Value = arr(33)
  Fields("fldprgf4").Value = arr(34)
  Fields("fldHead3").Value = arr(35)
  Fields("fldHead4").Value = arr(36)
  Fields("fldHead5").Value = arr(37)
  Fields("fldthisRealVal").Value = arr(38)
  Fields("fldRCnt").Value = arr(39)  'count of how many properties this cust has
  Fields("fldMap").Value = arr(40)
  Fields("fldBlock").Value = arr(41)
  Fields("fldLot").Value = arr(42)
  Fields("fldRealAdd").Value = arr(43)
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
      MsgBox "File - TxLsrRItem.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - TxLsrRItem.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - TxLsrRItem.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - TxLsrRItem.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "TxLsrRItem.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TxLsrRItem.txt"
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
    
  If CInt(Fields("fldRCnt").Value) = ThisLen Then
    Detail.NewPage = ddNPBefore
  ElseIf CInt(Fields("fldRCnt").Value) < ThisLen Then
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
'  If QPTrim$(Fields("fldOpt1Desc").Value) = "" Then
'    Label88.Caption = ""
'    Field218.Visible = False
'  Else
'    Label88.Caption = QPTrim$(Fields("fldOpt1Desc").Value)
'    Opt1 = True
'  End If
'  If QPTrim$(Fields("fldOpt2Desc").Value) = "" Then
'    Label89.Caption = ""
'    Field219.Visible = False
'  Else
'    Label89.Caption = QPTrim$(Fields("fldOpt2Desc").Value)
'    Opt1 = True
'  End If
'  If QPTrim$(Fields("fldOpt3Desc").Value) = "" Then
'    Label90.Caption = ""
'    Field220.Visible = False
'  Else
'    Label90.Caption = QPTrim$(Fields("fldOpt3Desc").Value)
'    Opt1 = True
'  End If
'
'  If Opt1 = False And Opt2 = False And Opt3 = False Then
'    Label88.Visible = False
'    Field218.Visible = False
'    Label89.Visible = False
'    Field219.Visible = False
'    Label90.Visible = False
'    Field220.Visible = False
'    Line35.Y1 = 4410
'    Line35.Y2 = 4410
'    Label87.Top = 4410
'    GroupHeader1.Height = 4590
'    Line38.Y1 = 4590
'    Line38.Y2 = 4590
'  ElseIf Opt1 = True And Opt2 = False And Opt3 = False Then
'    Label88.Visible = True
'    Field218.Visible = True
'    Label89.Visible = False
'    Field219.Visible = False
'    Label90.Visible = False
'    Field220.Visible = False
'    Line35.Y1 = 4680
'    Line35.Y2 = 4680
'    Label87.Top = 4680
'    GroupHeader1.Height = 4860
'    Line38.Y1 = 4860
'    Line38.Y2 = 4860
'  ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
'    Label88.Visible = False
'    Field218.Visible = False
'    Label89.Visible = True
'    Field219.Visible = True
'    Label89.Top = 4410
'    Field219.Top = 4410
'    Label90.Visible = False
'    Field220.Visible = False
'    Line35.Y1 = 4680
'    Line35.Y2 = 4680
'    Label87.Top = 4680
'    GroupHeader1.Height = 4860
'    Line38.Y1 = 4860
'    Line38.Y2 = 4860
'  ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
'    Label88.Visible = False
'    Field218.Visible = False
'    Label89.Visible = False
'    Field219.Visible = False
'    Label90.Visible = True
'    Field220.Visible = True
'    Label90.Top = 4410
'    Field220.Top = 4410
'    Line35.Y1 = 4680
'    Line35.Y2 = 4680
'    Label87.Top = 4680
'    GroupHeader1.Height = 4860
'    Line38.Y1 = 4860
'    Line38.Y2 = 4860
'  ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
'    Label88.Visible = True
'    Field218.Visible = True
'    Label89.Visible = True
'    Field219.Visible = True
'    Label90.Visible = False
'    Field220.Visible = False
'    Line35.Y1 = 4950
'    Line35.Y2 = 4950
'    Label87.Top = 4950
'    GroupHeader1.Height = 5130
'    Line38.Y1 = 5130
'    Line38.Y2 = 5130
'  ElseIf Opt1 = True And Opt2 = False And Opt3 = True Then
'    Label88.Visible = True
'    Field218.Visible = True
'    Label89.Visible = False
'    Field219.Visible = False
'    Label90.Visible = True
'    Field220.Visible = True
'    Label90.Top = 4680
'    Field220.Top = 4680
'    Line35.Y1 = 4950
'    Line35.Y2 = 4950
'    Label87.Top = 4950
'    GroupHeader1.Height = 5130
'    Line38.Y1 = 5130
'    Line38.Y2 = 5130
'  ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
'    Label88.Visible = False
'    Field218.Visible = False
'    Label89.Visible = True
'    Field219.Visible = True
'    Label89.Top = 4410
'    Field219.Top = 4410
'    Label90.Visible = True
'    Field220.Visible = True
'    Label90.Top = 4680
'    Field220.Top = 4680
'    Line35.Y1 = 4950
'    Line35.Y2 = 4950
'    Label87.Top = 4950
'    GroupHeader1.Height = 5130
'    Line38.Y1 = 5130
'    Line38.Y2 = 5130
'  End If
End Sub

