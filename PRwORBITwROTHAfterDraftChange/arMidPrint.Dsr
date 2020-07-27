VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arMidPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laser Mid Print Checks"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arMidPrint.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arMidPrint.dsx":08CA
End
Attribute VB_Name = "arMidPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private HFile As Integer
Dim Reprint As Boolean
Dim DDFlag As Boolean
Private Sub ActiveReport_Initialize()
  Me.Toolbar.Tools.Add "&Close"
  Me.Toolbar.Tools.Add "Save/&Excel"
  Me.Toolbar.Tools.Add "&Text"
End Sub

Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
    DoEvents
    
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
      MsgBox "File - MidCheck.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - MidCheck.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - MidCheck.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - MidCheck.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "MidCheck.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "MidCheck.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  HFile = FreeFile
  Open StartPath & "\PRRPTS\MIDCHECK.RPT" For Input As #HFile
  Fields.Add "fldEmpName" '(0)
  Fields.Add "fldSSN" '(1)
  Fields.Add "fldEmpNo" '(2)
  Fields.Add "fldBaseRate" '(3)
  Fields.Add "fldEndDate" '(4)
  Fields.Add "fldCheckDate" '(5)
  Fields.Add "fldVacUsed" '(6)
  Fields.Add "fldVacPay" '(7)
  Fields.Add "fldVacEarned" '(8)
  Fields.Add "fldVacUsedTot" '(9)
  Fields.Add "fldVacBal" '(10)
  Fields.Add "fldDedDesc1" '(11)
  Fields.Add "fldDedTot1" '(12)
  Fields.Add "fldDedYTD1" '(13)
  Fields.Add "fldSickUsed" '(14)
  Fields.Add "fldSickPay" '(15)
  Fields.Add "fldSickEarned" '(16)
  Fields.Add "fldSickUsedTot" '(17)
  Fields.Add "fldSickBal" '(18)
  Fields.Add "fldDedDesc2" '(19)
  Fields.Add "fldDedTot2" '(20)
  Fields.Add "fldDedYTD2" '(21)
  Fields.Add "fldHolUsed" '(22)
  Fields.Add "fldHolPay" '(23)
  Fields.Add "fldDedDesc3" '(24)
  Fields.Add "fldDedTot3" '(25)
  Fields.Add "fldDedYTD3" '(26)
  Fields.Add "fldCompUsed" '(27)
  Fields.Add "fldCompPay" '(28)
  Fields.Add "fldCompEarned" '(29)
  Fields.Add "fldCompUsedTot" '(30)
  Fields.Add "fldCompBal" '(31)
  Fields.Add "fldDedDesc4" '(32)
  Fields.Add "fldDedTot4" '(33)
  Fields.Add "fldDedYTD4" '(34)
  Fields.Add "fldRegHrsWkd" '(35)
  Fields.Add "fldRegXBase" '(36)
  Fields.Add "fldDedDesc5" '(37)
  Fields.Add "fldDedTot5" '(38)
  Fields.Add "fldDedYTD5" '(39)
  Fields.Add "fldOTHrsPaid" '(40)
  Fields.Add "fldTotOTWage" '(41)
  Fields.Add "fldDedDesc6" '(42)
  Fields.Add "fldDedTot6" '(43)
  Fields.Add "fldDedYTD6" '(44)
  Fields.Add "fldGrossPay" '(45)
  Fields.Add "fldYTDGrossPay" '(46)
  Fields.Add "fldDedDesc7" '(47)
  Fields.Add "fldDedTot7" '(48)
  Fields.Add "fldDedYTD7" '(49)
  Fields.Add "fldFedTax" '(50)
  Fields.Add "fldYTDFedTax" '(51)
  Fields.Add "fldDedDesc8" '(52)
  Fields.Add "fldDedTot8" '(53)
  Fields.Add "fldDedYTD8" '(54)
  Fields.Add "fldFICA" '(55)
  Fields.Add "fldYTDFICA" '(56)
  Fields.Add "fldDedDesc9" '(57)
  Fields.Add "fldDedTot9" '(58)
  Fields.Add "fldDedYTD9" '(59)
  Fields.Add "fldRetire" '(60)
  Fields.Add "fldYTDRetire" '(61)
  Fields.Add "fldDedDesc10" '(62)
  Fields.Add "fldDedTot10" '(63)
  Fields.Add "fldDedYTD10" '(64)
  Fields.Add "fldNetPay" '(65)
  Fields.Add "fldYTDNetPay" '(66)
  Fields.Add "fldDedDesc11" '(67)
  Fields.Add "fldDedTot11" '(68)
  Fields.Add "fldDedYTD11" '(69)
  Fields.Add "fldStaTax" '(70)
  Fields.Add "fldYTDStaTax" '(71)
  Fields.Add "fldDedDesc12" '(72)
  Fields.Add "fldDedTot12" '(73)
  Fields.Add "fldDedYTD12" '(74)
  Fields.Add "fldTotAddEarn" '(75)
  Fields.Add "fldYTDTotAddEarn" '(76)
  Fields.Add "fldDDFlag" '(77)
  Fields.Add "fldCheckNum" '(78)
  Fields.Add "fldSpellNum" '(79)
  Fields.Add "fldCheckDate2" '(80)
  Fields.Add "fldNetPay2" '(81)
  Fields.Add "fldEmpName2" '(82)
  Fields.Add "fldEmpAddress" '(83)
  Fields.Add "fldEmpCityStateZip" '(84)
  Fields.Add "fldHolPerUsed" '85)
  Fields.Add "fldHolPerPay" '86)
  Fields.Add "fldHolPerAmt" '87)
  Fields.Add "fldHPBalUsedT" '88)
  Fields.Add "fldHolPerBal" '89)
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
  DDFlag = False
  Fields("fldEmpName").Value = arr(0)
  Fields("fldSSN").Value = arr(1)
  Fields("fldEmpNo").Value = arr(2)
  Fields("fldBaseRate").Value = arr(3)
  Fields("fldEndDate").Value = arr(4)
  Fields("fldCheckDate").Value = arr(5)
  Fields("fldVacUsed").Value = arr(6)
  Fields("fldVacPay").Value = arr(7)
  Fields("fldVacEarned").Value = arr(8)
  Fields("fldVacUsedTot").Value = arr(9)
  Fields("fldVacBal").Value = arr(10)
  Fields("fldDedDesc1").Value = arr(11)
  If QPTrim$(arr(11)) = "" Then
    fldDedTot1.Visible = False
    fldDedYTD1.Visible = False
    fldDedTot1b.Visible = False
    fldDedYTD1b.Visible = False
  ElseIf QPTrim$(arr(11)) <> "" Then
    fldDedTot1.Visible = True
    fldDedYTD1.Visible = True
    fldDedTot1b.Visible = True
    fldDedYTD1b.Visible = True
  End If
  Fields("fldDedTot1").Value = arr(12)
  Fields("fldDedYTD1").Value = arr(13)
  Fields("fldSickUsed").Value = arr(14)
  Fields("fldSickPay").Value = arr(15)
  Fields("fldSickEarned").Value = arr(16)
  Fields("fldSickUsedTot").Value = arr(17)
  Fields("fldSickBal").Value = arr(18)
  Fields("fldDedDesc2").Value = arr(19)
  If QPTrim$(arr(19)) = "" Then
    fldDedTot2.Visible = False
    fldDedYTD2.Visible = False
    fldDedTot2b.Visible = False
    fldDedYTD2b.Visible = False
  ElseIf QPTrim$(arr(19)) <> "" Then
    fldDedTot2.Visible = True
    fldDedYTD2.Visible = True
    fldDedTot2b.Visible = True
    fldDedYTD2b.Visible = True
  End If
  Fields("fldDedTot2").Value = arr(20)
  Fields("fldDedYTD2").Value = arr(21)
  Fields("fldHolUsed").Value = arr(22)
  Fields("fldHolPay").Value = arr(23)
  Fields("fldDedDesc3").Value = arr(24)
  If QPTrim$(arr(24)) = "" Then
    fldDedTot3.Visible = False
    fldDedYTD3.Visible = False
    fldDedTot3b.Visible = False
    fldDedYTD3b.Visible = False
  ElseIf QPTrim$(arr(24)) <> "" Then
    fldDedTot3.Visible = True
    fldDedYTD3.Visible = True
    fldDedTot3b.Visible = True
    fldDedYTD3b.Visible = True
  End If
  Fields("fldDedTot3").Value = arr(25)
  Fields("fldDedYTD3").Value = arr(26)
  Fields("fldCompUsed").Value = arr(27)
  Fields("fldCompPay").Value = arr(28)
  Fields("fldCompEarned").Value = arr(29)
  Fields("fldCompUsedTot").Value = arr(30)
  Fields("fldCompBal").Value = arr(31)
  Fields("fldDedDesc4").Value = arr(32)
  If QPTrim$(arr(32)) = "" Then
    fldDedTot4.Visible = False
    fldDedYTD4.Visible = False
    fldDedTot4b.Visible = False
    fldDedYTD4b.Visible = False
  ElseIf QPTrim$(arr(32)) <> "" Then
    fldDedTot4.Visible = True
    fldDedYTD4.Visible = True
    fldDedTot4b.Visible = True
    fldDedYTD4b.Visible = True
  End If
  Fields("fldDedTot4").Value = arr(33)
  Fields("fldDedYTD4").Value = arr(34)
  Fields("fldRegHrsWkd").Value = arr(35)
  Fields("fldRegXBase").Value = arr(36)
  Fields("fldDedDesc5").Value = arr(37)
  If QPTrim$(arr(37)) = "" Then
    fldDedTot5.Visible = False
    fldDedYTD5.Visible = False
    fldDedTot5b.Visible = False
    fldDedYTD5b.Visible = False
  ElseIf QPTrim$(arr(37)) <> "" Then
    fldDedTot5.Visible = True
    fldDedYTD5.Visible = True
    fldDedTot5b.Visible = True
    fldDedYTD5b.Visible = True
  End If
  Fields("fldDedTot5").Value = arr(38)
  Fields("fldDedYTD5").Value = arr(39)
  Fields("fldOTHrsPaid").Value = arr(40)
  Fields("fldTotOTWage").Value = arr(41)
  Fields("fldDedDesc6").Value = arr(42)
  If QPTrim$(arr(42)) = "" Then
    fldDedTot6.Visible = False
    fldDedYTD6.Visible = False
    fldDedTot6b.Visible = False
    fldDedYTD6b.Visible = False
  ElseIf QPTrim$(arr(42)) <> "" Then
    fldDedTot6.Visible = True
    fldDedYTD6.Visible = True
    fldDedTot6b.Visible = True
    fldDedYTD6b.Visible = True
  End If
  Fields("fldDedTot6").Value = arr(43)
  Fields("fldDedYTD6").Value = arr(44)
  Fields("fldGrossPay").Value = arr(45)
  Fields("fldYTDGrossPay").Value = arr(46)
  Fields("fldDedDesc7").Value = arr(47)
  If QPTrim$(arr(47)) = "" Then
    fldDedTot7.Visible = False
    fldDedYTD7.Visible = False
    fldDedTot7b.Visible = False
    fldDedYTD7b.Visible = False
  ElseIf QPTrim$(arr(47)) <> "" Then
    fldDedTot7.Visible = True
    fldDedYTD7.Visible = True
    fldDedTot7b.Visible = True
    fldDedYTD7b.Visible = True
  End If
  Fields("fldDedTot7").Value = arr(48)
  Fields("fldDedYTD7").Value = arr(49)
  Fields("fldFedTax").Value = arr(50)
  Fields("fldYTDFedTax").Value = arr(51)
  Fields("fldDedDesc8").Value = arr(52)
  If QPTrim$(arr(52)) = "" Then
    fldDedTot8.Visible = False
    fldDedYTD8.Visible = False
    fldDedTot8b.Visible = False
    fldDedYTD8b.Visible = False
  ElseIf QPTrim$(arr(52)) <> "" Then
    fldDedTot8.Visible = True
    fldDedYTD8.Visible = True
    fldDedTot8b.Visible = True
    fldDedYTD8b.Visible = True
  End If
  Fields("fldDedTot8").Value = arr(53)
  Fields("fldDedYTD8").Value = arr(54)
  Fields("fldFICA").Value = arr(55)
  Fields("fldYTDFICA").Value = arr(56)
  Fields("fldDedDesc9").Value = arr(57)
  If QPTrim$(arr(57)) = "" Then
    fldDedTot9.Visible = False
    fldDedYTD9.Visible = False
    fldDedTot9b.Visible = False
    fldDedYTD9b.Visible = False
  ElseIf QPTrim$(arr(57)) <> "" Then
    fldDedTot9.Visible = True
    fldDedYTD9.Visible = True
    fldDedTot9b.Visible = True
    fldDedYTD9b.Visible = True
  End If
  Fields("fldDedTot9").Value = arr(58)
  Fields("fldDedYTD9").Value = arr(59)
  Fields("fldRetire").Value = arr(60)
  Fields("fldYTDRetire").Value = arr(61)
  Fields("fldDedDesc10").Value = arr(62)
  If QPTrim$(arr(62)) = "" Then
    fldDedTot10.Visible = False
    fldDedYTD10.Visible = False
    fldDedTot10b.Visible = False
    fldDedYTD10b.Visible = False
  ElseIf QPTrim$(arr(62)) <> "" Then
    fldDedTot10.Visible = True
    fldDedYTD10.Visible = True
    fldDedTot10b.Visible = True
    fldDedYTD10b.Visible = True
  End If
  Fields("fldDedTot10").Value = arr(63)
  Fields("fldDedYTD10").Value = arr(64)
  Fields("fldNetPay").Value = arr(65)
  Fields("fldYTDNetPay").Value = arr(66)
  Fields("fldDedDesc11").Value = arr(67)
  If QPTrim$(arr(67)) = "" Then
    fldDedTot11.Visible = False
    fldDedYTD11.Visible = False
    fldDedTot11b.Visible = False
    fldDedYTD11b.Visible = False
  ElseIf QPTrim$(arr(67)) <> "" Then
    fldDedTot11.Visible = True
    fldDedYTD11.Visible = True
    fldDedTot11b.Visible = True
    fldDedYTD11b.Visible = True
  End If
  Fields("fldDedTot11").Value = arr(68)
  Fields("fldDedYTD11").Value = arr(69)
  Fields("fldStaTax").Value = arr(70)
  Fields("fldYTDStaTax").Value = arr(71)
  Fields("fldDedDesc12").Value = arr(72)
  If QPTrim$(arr(72)) = "" Then
    fldDedTot12.Visible = False
    fldDedYTD12.Visible = False
    fldDedTot12b.Visible = False
    fldDedYTD12b.Visible = False
  ElseIf QPTrim$(arr(72)) <> "" Then
    fldDedTot12.Visible = True
    fldDedYTD12.Visible = True
    fldDedTot12b.Visible = True
    fldDedYTD12b.Visible = True
  End If
  Fields("fldDedTot12").Value = arr(73)
  Fields("fldDedYTD12").Value = arr(74)
  Fields("fldTotAddEarn").Value = arr(75)
  Fields("fldYTDTotAddEarn").Value = arr(76)
  Fields("fldDDFlag").Value = arr(77)
  If arr(77) = -1 Then DDFlag = True
  Fields("fldCheckNum").Value = arr(78)
  Fields("fldSpellNum").Value = arr(79)
  Fields("fldCheckDate2").Value = arr(80)
  Fields("fldNetPay2").Value = arr(81)
  Fields("fldEmpName2").Value = arr(82)
  Fields("fldEmpAddress").Value = arr(83)
  Fields("fldEmpCityStateZip").Value = arr(84)
  Fields("fldHolPerUsed").Value = arr(85)
  Fields("fldHolPerPay").Value = arr(86)
  Fields("fldHolPerAmt").Value = arr(87)
  Fields("fldHPBalUsedT").Value = arr(88)
  Fields("fldHolPerBal").Value = arr(89)
  
  If Len(arr(77)) > 0 Then
    Reprint = True
  Else
    Reprint = False
  End If
End Sub

Private Sub ActiveReport_ReportEnd()
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
'    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Me.Zoom = -1

End Sub

Private Sub Detail_Format()
  If DDFlag = False Then
    Label33.Visible = False
  Else
    Label33.Visible = True
  End If
  
  If DDFlag = False Then
    Label32.Visible = False
  Else
    Label32.Visible = True
  End If
 
  If DDFlag = False Then
    Label31.Visible = False
  Else
    Label31.Visible = True
  End If
  
  If DDFlag = False Then
    Label30.Visible = False
  Else
    Label30.Visible = True
  End If
  
  If DDFlag = False Then
    Label29.Visible = False
  Else
    Label29.Visible = True
  End If
  
  If DDFlag = False Then
    Label66.Visible = False
  Else
    Label66.Visible = True
  End If
  
  If DDFlag = False Then
    Label65.Visible = False
  Else
    Label65.Visible = True
  End If
  
'  If DDFlag = False Then
'    Label64.Visible = False
'  Else
'    Label64.Visible = True
'  End If
  
  If DDFlag = False Then
    Label63.Visible = False
  Else
    Label63.Visible = True
  End If
  
  If DDFlag = False Then
    Label62.Visible = False
    Label68.Visible = False
  Else
    Label62.Visible = True
    Label68.Visible = True
  End If
  
  
  If Reprint = True Then
    Reprint = False
  ElseIf Reprint = False Then
  End If

End Sub

Private Sub ReportFooter_Format()
End Sub


