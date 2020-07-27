VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxCustTHistByProp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Customer Transaction History"
   ClientHeight    =   8736
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arVATaxCustTHistByProp.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arVATaxCustTHistByProp.dsx":08CA
End
Attribute VB_Name = "arVATaxCustTHistByProp"
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
  Open StartPath & "\TAXRPTS\CHSTPROP.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldPropType") '1)
  Fields.Add ("fldType") '2)
  Fields.Add ("fldPropPin") '3)
  Fields.Add ("fldDesc") '4)
  Fields.Add ("fldAmount") '5)
  Fields.Add ("fldDate") '6)
  Fields.Add ("fldYear") '7)
  Fields.Add ("fldBillNum") '8)
  Fields.Add ("fldOpt1Desc") '9)
  Fields.Add ("fldOpt2Desc") '10)
  Fields.Add ("fldOpt3Desc") '11)
  Fields.Add ("fldPrincPersAmt") '12)
  Fields.Add ("fldIntAmt") '13)
  Fields.Add ("fldAdvMTAmt") '14)
  Fields.Add ("fldLateListMCAmt") '15)
  Fields.Add ("fldOpt1Amt") '16)
  Fields.Add ("fldOpt2Amt") '17)
  Fields.Add ("fldOpt3Amt") '18)
  Fields.Add ("fldPrincPersAmtPd") '19)
  Fields.Add ("fldIntAmtPd") '20)
  Fields.Add ("fldAdvMTAmtPd") '21)
  Fields.Add ("fldLateListMCAmtPd") '22)
  Fields.Add ("fldOpt1AmtPd") '23)
  Fields.Add ("fldOpt2AmtPd") '24)
  Fields.Add ("fldOpt3AmtPd") '25)
  Fields.Add ("fldPrincPersAmtDif") '26)
  Fields.Add ("fldIntAmtDif") '27)
  Fields.Add ("fldAdvMTAmtDif") '28)
  Fields.Add ("fldLateListMCAmtDif") '29)
  Fields.Add ("fldOpt1AmtDif") '30)
  Fields.Add ("fldOpt2AmtDif") '31)
  Fields.Add ("fldOpt3AmtDif") '32)
  Fields.Add ("fldTypeNum") '33)
  Fields.Add ("fldBalThisBill") '34)
  Fields.Add ("fldTotBal") '35)
  Fields.Add ("fldCustRec") '36)
  Fields.Add ("fldName") '37)
  Fields.Add ("fldPenaltyBal") '38)
  Fields.Add ("fldPenaltyChrg") '39)
  Fields.Add ("fldPenaltyPd") '40)
  Fields.Add ("fldFEAmt") '41)
  Fields.Add ("fldFEAmtPd") '42)
  Fields.Add ("fldFEAmtDif") '43)
  Fields.Add ("fldMHAmt") '41)
  Fields.Add ("fldMHAmtPd") '42)
  Fields.Add ("fldMHAmtDif") '43)
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
  
  Fields("fldTown").Value = arr(0)
  Fields("fldPropType").Value = arr(1)
  Fields("fldType").Value = arr(2)
  Fields("fldPropPin").Value = arr(3)
  Fields("fldDesc").Value = arr(4)
  Fields("fldAmount").Value = arr(5)
  Fields("fldDate").Value = arr(6)
  Fields("fldYear").Value = arr(7)
  Fields("fldBillNum").Value = arr(8)
  Fields("fldOpt1Desc").Value = arr(9)
  Fields("fldOpt2Desc").Value = arr(10)
  Fields("fldOpt3Desc").Value = arr(11)
  Fields("fldPrincPersAmt").Value = arr(12)
  Fields("fldIntAmt").Value = arr(13)
  Fields("fldAdvMTAmt").Value = arr(14)
  Fields("fldLateListMCAmt").Value = arr(15)
  Fields("fldOpt1Amt").Value = arr(16)
  Fields("fldOpt2Amt").Value = arr(17)
  Fields("fldOpt3Amt").Value = arr(18)
  Fields("fldPrincPersAmtPd").Value = arr(19)
  Fields("fldIntAmtPd").Value = arr(20)
  Fields("fldAdvMTAmtPd").Value = arr(21)
  Fields("fldLateListMCAmtPd").Value = arr(22)
  Fields("fldOpt1AmtPd").Value = arr(23)
  Fields("fldOpt2AmtPd").Value = arr(24)
  Fields("fldOpt3AmtPd").Value = arr(25)
  Fields("fldPrincPersAmtDif").Value = arr(26)
  Fields("fldIntAmtDif").Value = arr(27)
  Fields("fldAdvMTAmtDif").Value = arr(28)
  Fields("fldLateListMCAmtDif").Value = arr(29)
  Fields("fldOpt1AmtDif").Value = arr(30)
  Fields("fldOpt2AmtDif").Value = arr(31)
  Fields("fldOpt3AmtDif").Value = arr(32)
  Fields("fldTypeNum").Value = arr(33)
  Fields("fldBalThisBill").Value = arr(34)
  Fields("fldTotBal").Value = arr(35)
  Fields("fldCustRec").Value = arr(36)
  Fields("fldName").Value = arr(37)
  Fields("fldPenaltyBal").Value = arr(38)
  Fields("fldPenaltyChrg").Value = arr(39)
  Fields("fldPenaltyPd").Value = arr(40)
  Fields("fldFEAmt").Value = arr(41)
  Fields("fldFEAmtPd").Value = arr(42)
  Fields("fldFEAmtDif").Value = arr(43)
  Fields("fldMHAmt").Value = arr(41)
  Fields("fldMHAmtPd").Value = arr(42)
  Fields("fldMHAmtDif").Value = arr(43)
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
      frmVATaxMsg.Label1.Caption = "File - TaxCustTHistByProp.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - TaxCustTHistByProp.txt, created in the Citipak Directory."
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
    frmVATaxMsg.Label1.Caption = "File - TaxCustTHistByProp.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - TaxCustTHistByProp.txt, created in the Citipak Directory."
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
    Case 1
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & "TaxCustTHistByProp.xls"
        oEXL.Export Me.Pages
    Case 2
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TaxCustTHistByProp.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_ReportEnd()
'  Unload frmBLLoadReport
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
  Dim Opt1 As Boolean
  Dim Opt2 As Boolean
  Dim Opt3 As Boolean
  
  
  If Fields("fldPropType").Value = "PERSONAL" Then
    Label35.Caption = "Personal"
    Label37.Caption = "Machine Tools"
    Label38.Caption = "Merchant Capital"
    Label44.Caption = "Farm Equipment"
    Label45.Caption = "Mobile Homes"
  ElseIf Fields("fldPropType").Value = "REAL" Then
    Detail.Height = 2580
    Label35.Caption = "Principal"
    Label37.Caption = "Advertising"
    Label38.Caption = "Late Listing"
    Label44.Visible = False
    Field47.Visible = False
    Field48.Visible = False
    Field49.Visible = False
    Label45.Visible = False
    Field50.Visible = False
    Field51.Visible = False
    Field52.Visible = False
    Label36.Top = 1170 'Interest
    Field20.Top = 1170
    Field27.Top = 1170
    Field34.Top = 1170
    Label43.Top = 1440 'Penalty
    Field44.Top = 1440
    Field45.Top = 1440
    Field46.Top = 1440
    Field16.Top = 1710 'Opt1
    Field23.Top = 1710
    Field30.Top = 1710
    Field37.Top = 1710
    Field17.Top = 1980 'Opt2
    Field24.Top = 1980
    Field31.Top = 1980
    Field38.Top = 1980
    Field18.Top = 2250 'Opt3
    Field25.Top = 2250
    Field32.Top = 2250
    Field39.Top = 2250
    Line4.Y1 = 2520
    Line4.Y2 = 2520
  End If
  
  If Fields("fldTypeNum") <> 1 Then 'not a bill so no need for balance
    Field33.Visible = False
    Field34.Visible = False
    Field35.Visible = False
    Field36.Visible = False
    Field37.Visible = False
    Field38.Visible = False
    Field39.Visible = False
    Field41.Visible = False
    Field46.Visible = False
    Field49.Visible = False
    Field52.Visible = False
  Else
    Field33.Visible = True
    Field34.Visible = True
    Field35.Visible = True
    Field36.Visible = True
    Field37.Visible = True
    Field38.Visible = True
    Field39.Visible = True
    Field41.Visible = True
    Field46.Visible = True
    If Fields("fldPropType").Value = "PERSONAL" Then
      Field49.Visible = True
      Field52.Visible = True
    End If
  End If
  
  Opt1 = True
  Opt2 = True
  Opt3 = True
  If QPTrim$(Fields("fldOpt1Desc")) = "" Then
    Opt1 = False
  End If
  If QPTrim$(Fields("fldOpt2Desc")) = "" Then
    Opt2 = False
  End If
  If QPTrim$(Fields("fldOpt3Desc")) = "" Then
    Opt3 = False
  End If
  
  If Opt1 = True And Opt2 = True And Opt3 = True Then Exit Sub
  
  If Fields("fldPropType").Value = "REAL" Then
    If Opt1 = False And Opt2 = False And Opt3 = False Then
      Detail.Height = 1695
      Line4.Y1 = 1695
      Line4.Y2 = 1695
      Field17.Visible = False
      Field24.Visible = False
      Field31.Visible = False
      Field38.Visible = False
      Field18.Visible = False
      Field25.Visible = False
      Field32.Visible = False
      Field39.Visible = False
      Field16.Visible = False
      Field23.Visible = False
      Field30.Visible = False
      Field37.Visible = False
      Exit Sub
    End If
  
    If Opt1 = True And Opt2 = False And Opt3 = False Then
      Detail.Height = 1980
      Line4.Y1 = 1950
      Line4.Y2 = 1950
      Field17.Visible = False
      Field24.Visible = False
      Field31.Visible = False
      Field38.Visible = False
      Field18.Visible = False
      Field25.Visible = False
      Field32.Visible = False
      Field39.Visible = False
    ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
      Detail.Height = 2295
      Line4.Y1 = 2265
      Line4.Y2 = 2265
      Field18.Visible = False
      Field25.Visible = False
      Field32.Visible = False
      Field39.Visible = False
    ElseIf Opt1 = True And Opt2 = False And Opt3 = True Then
      Detail.Height = 2295
      Line4.Y1 = 2265
      Line4.Y2 = 2265
      Field17.Visible = False
      Field24.Visible = False
      Field31.Visible = False
      Field38.Visible = False
      Field18.Top = 1980
      Field25.Top = 1980
      Field32.Top = 1980
      Field39.Top = 1980
    ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
      Detail.Height = 1980
      Line4.Y1 = 1950
      Line4.Y2 = 1950
      Field16.Visible = False
      Field23.Visible = False
      Field30.Visible = False
      Field37.Visible = False
      Field18.Visible = False
      Field25.Visible = False
      Field32.Visible = False
      Field39.Visible = False
      Field17.Top = 1710
      Field24.Top = 1710
      Field31.Top = 1710
      Field38.Top = 1710
    ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
      Detail.Height = 2295
      Line4.Y1 = 2265
      Line4.Y2 = 2265
      Field16.Visible = False
      Field23.Visible = False
      Field30.Visible = False
      Field37.Visible = False
      Field17.Top = 1710
      Field24.Top = 1710
      Field31.Top = 1710
      Field38.Top = 1710
      Field18.Top = 1980
      Field25.Top = 1980
      Field32.Top = 1980
      Field39.Top = 1980
    ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
      Detail.Height = 1980
      Line4.Y1 = 1950
      Line4.Y2 = 1950
      Field16.Visible = False
      Field23.Visible = False
      Field30.Visible = False
      Field37.Visible = False
      Field17.Visible = False
      Field24.Visible = False
      Field31.Visible = False
      Field38.Visible = False
      Field18.Top = 1710
      Field25.Top = 1710
      Field32.Top = 1710
      Field39.Top = 1710
    End If
  ElseIf Fields("fldPropType").Value = "PERSONAL" Then
    If Opt1 = False And Opt2 = False And Opt3 = False Then
      Detail.Height = 2265
      Line4.Y1 = 2235
      Line4.Y2 = 2235
      Field16.Visible = False 'Opt1
      Field23.Visible = False
      Field30.Visible = False
      Field37.Visible = False
      Field17.Visible = False 'Opt2
      Field24.Visible = False
      Field31.Visible = False
      Field38.Visible = False
      Field18.Visible = False 'Opt3
      Field25.Visible = False
      Field32.Visible = False
      Field39.Visible = False
      Exit Sub
    End If
  
    If Opt1 = True And Opt2 = False And Opt3 = False Then
      Detail.Height = 2520
      Line4.Y1 = 2490
      Line4.Y2 = 2490
      Field17.Visible = False
      Field24.Visible = False
      Field31.Visible = False
      Field38.Visible = False
      Field18.Visible = False
      Field25.Visible = False
      Field32.Visible = False
      Field39.Visible = False
    ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
      Detail.Height = 3405
      Line4.Y1 = 3375
      Line4.Y2 = 3375
      Field18.Visible = False
      Field25.Visible = False
      Field32.Visible = False
      Field39.Visible = False
    ElseIf Opt1 = True And Opt2 = False And Opt3 = True Then
      Detail.Height = 2835
      Line4.Y1 = 2805
      Line4.Y2 = 2805
      Field17.Visible = False
      Field24.Visible = False
      Field31.Visible = False
      Field38.Visible = False
      Field18.Top = 2520
      Field25.Top = 2520
      Field32.Top = 2520
      Field39.Top = 2520
    ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
      Detail.Height = 2475
      Line4.Y1 = 2445
      Line4.Y2 = 2445
      Field16.Visible = False
      Field23.Visible = False
      Field30.Visible = False
      Field37.Visible = False
      Field18.Visible = False
      Field25.Visible = False
      Field32.Visible = False
      Field39.Visible = False
      Field17.Top = 2250
      Field24.Top = 2250
      Field31.Top = 2250
      Field38.Top = 2250
    ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
      Detail.Height = 2835
      Line4.Y1 = 2805
      Line4.Y2 = 2805
      Field16.Visible = False
      Field23.Visible = False
      Field30.Visible = False
      Field37.Visible = False
      Field17.Top = 2250
      Field24.Top = 2250
      Field31.Top = 2250
      Field38.Top = 2250
      Field18.Top = 2520
      Field25.Top = 2520
      Field32.Top = 2520
      Field39.Top = 2520
    ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
      Detail.Height = 2520
      Line4.Y1 = 2490
      Line4.Y2 = 2490
      Field16.Visible = False
      Field23.Visible = False
      Field30.Visible = False
      Field37.Visible = False
      Field17.Visible = False
      Field24.Visible = False
      Field31.Visible = False
      Field38.Visible = False
      Field18.Top = 2250
      Field25.Top = 2250
      Field32.Top = 2250
      Field39.Top = 2250
    End If
  End If
  
End Sub
