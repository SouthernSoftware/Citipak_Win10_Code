VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arDistRegAcctNum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Number Distribution Register"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arDistRegAcctNum.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arDistRegAcctNum.dsx":08CA
End
Attribute VB_Name = "arDistRegAcctNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private hFile As Integer

Dim EndReport As Boolean
Dim DedCnt As Integer
Dim TotANums As Integer
Dim TotFNUms As Integer
Dim NoEscape As Boolean
Private Sub ActiveReport_Initialize()
  Me.ToolBar.Tools.Add "&Close"
  Me.ToolBar.Tools.Add "Save/&Excel"
  Me.ToolBar.Tools.Add "&Text"
End Sub

Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  If KeyCode = vbKeyEscape And NoEscape = False Then
    NoEscape = True
    DoEvents
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
      MsgBox "File - DistRegisterbyAcctNum.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - DistRegisterbyAcctNum.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close hFile
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
    MsgBox "File - DistRegisterbyAcctNum.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - DistRegisterbyAcctNum.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "DistRegisterbyAcctNum.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "DistRegisterbyAcctNum.txt"
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
  NoEscape = False
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
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
  hFile = FreeFile
  Open StartPath & "\PRRPTS\DISTRIBUACCTNUMG.RPT" For Input As #hFile

  Fields.Add "fldEmployer" '(0)
  Fields.Add "fldDate" '(1)
  Fields.Add "fldENum" '(2)
  Fields.Add "fldEmployee" '(3)
  Fields.Add "fldBaseRate" '(4)
  Fields.Add "fldOTRate" '(5)
  Fields.Add "fldEAcctNum" '(6)
  Fields.Add "fldSalPct" '(7)

  Fields.Add "fldRegHrs" '(8)
  Fields.Add "fldOTHrs" '(9)
  Fields.Add "fldRegPay" '(10)
  Fields.Add "fldOTPay" '(11)
  Fields.Add "fldETother" '(12)
  Fields.Add "fldGrsPy" '(13)
  Fields.Add "fldSocSec" '(14)
  Fields.Add "fldMed" '(15)
  Fields.Add "fldRet" '(16)
  Fields.Add "fldLast" '(17)
  Fields.Add "fldEmpSplCnt" '(18)
  Fields.Add "fldEmpFundNum" '(19)

  Fields.Add ("fldFundDedAmt1") '(20)
  Fields.Add ("fldFundDedAmt2") '(21)
  Fields.Add ("fldFundDedAmt3") '(22)
  Fields.Add ("fldFundDedAmt4") '(23)
  Fields.Add ("fldFundDedAmt5") '(24)
  Fields.Add ("fldFundDedAmt6") '(25)
  Fields.Add ("fldFundDedAmt7") '(26)
  Fields.Add ("fldFundDedAmt8") '(27)
  Fields.Add ("fldFundDedAmt9") '(28)
  Fields.Add ("fldFundDedAmt10") '(29)
  Fields.Add ("fldFundDedAmt11") '(30)
  Fields.Add ("fldFundDedAmt12") '(31)
  Fields.Add ("fldFundDedAmt13") '(32)
  Fields.Add ("fldFundDedAmt14") '(33)
  Fields.Add ("fldFundDedAmt15") '(34)
  Fields.Add ("fldFundDedAmt16") '(35)
  Fields.Add ("fldFundDedAmt17") '(36)
  Fields.Add ("fldFundDedAmt18") '(37)
  Fields.Add ("fldFundDedAmt19") '(38)
  Fields.Add ("fldFundDedAmt20") '(39)
  Fields.Add ("fldFundDedAmt21") '(40)
  Fields.Add ("fldFundDedAmt22") '(41)
  Fields.Add ("fldFundDedAmt23") '(42)
  Fields.Add ("fldFundDedAmt24") '(43)
  Fields.Add ("fldFundDedAmt25") '(44)
  Fields.Add ("fldFundDedAmt26") '(45)
  Fields.Add ("fldFundDedAmt27") '(46)
  Fields.Add ("fldFundDedAmt28") '(47)
  Fields.Add ("fldFundDedAmt29") '(48)
  Fields.Add ("fldFundDedAmt30") '(49)
  Fields.Add ("fldFundDedAmt31") '(50)
  Fields.Add ("fldFundDedAmt32") '(51)
  Fields.Add ("fldFundDedAmt33") '(52)
  Fields.Add ("fldFundDedAmt34") '(53)
  Fields.Add ("fldFundDedAmt35") '(54)
  Fields.Add ("fldFundDedAmt36") '(55)
  Fields.Add ("fldFundDedAmt37") '(56)
  Fields.Add ("fldFundDedAmt38") '(57)
  Fields.Add ("fldFundDedAmt39") '(58)
  Fields.Add ("fldFundDedAmt40") '(59)
  Fields.Add ("fldFundDedAmt41") '(60)
  Fields.Add ("fldFundDedAmt42") '(61)
  Fields.Add ("fldFundDedAmt43") '(62)
  Fields.Add ("fldFundDedAmt44") '(63)
  Fields.Add ("fldFundDedAmt45") '(64)
  Fields.Add ("fldFundDedAmt46") '(65)
  Fields.Add ("fldFundDedAmt47") '(66)
  Fields.Add ("fldFundDedAmt48") '(67)
  Fields.Add ("fldFundDedAmt49") '(68)
  Fields.Add ("fldFundDedAmt50") '(69)
  Fields.Add ("fldFundFed") '(70)
  Fields.Add ("fldFundSta") '(71)
  Fields.Add ("fldFundMed") '(72)
  Fields.Add ("fldFundSoc") '(73)
  Fields.Add ("fldFundRet") '(74)
  Fields.Add ("fldAcctCnt1") '(75)
  Fields.Add ("fldToday") '(76)
  Fields.Add ("fldEmpSplCnt1") '(77)
  End Sub
Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  Dim x As Integer

  Static AcctCnt As Integer
  Static FundCnt As Integer
  Static EmpName As String
  If VBA.eof(hFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If

  Line Input #hFile, sLine
  arr = Split(sLine, "~")

  Fields("fldEmployer").Value = arr(0)
  Fields("fldDate").Value = arr(1)
  Fields("fldENum").Value = arr(2)

  If QPTrim$(EmpName) = "" Then EmpName = arr(3)
  Fields("fldEmployee").Value = arr(3)

  Fields("fldBaseRate").Value = arr(4)
  Fields("fldOTRate").Value = arr(5)
  Fields("fldEAcctNum").Value = arr(6)
  Fields("fldSalPct").Value = arr(7)

  Fields("fldRegHrs").Value = arr(8)
  Fields("fldOTHrs").Value = arr(9)
  Fields("fldRegPay").Value = arr(10)
  Fields("fldOTPay").Value = arr(11)
  Fields("fldETother").Value = arr(12)
  Fields("fldGrsPy").Value = arr(13)
  Fields("fldSocSec").Value = arr(14)
  Fields("fldMed").Value = arr(15)
  Fields("fldRet").Value = arr(16)
  Fields("fldLast").Value = arr(17)

  'If account numbers are multiples then the employee name
  'will stay the same...as long as they are the same then we
  'maintain a value for account numbers...when the name
  'changes we reset the accumulator...doing this to know how
  'large to set the detail height

  If arr(3) = EmpName Then
    If arr(17) <> "" Then
      AcctCnt = arr(17)
      TotANums = AcctCnt
    ElseIf arr(17) = "" Then
      AcctCnt = 0
      TotANums = AcctCnt
    End If
    If arr(18) <> "" Then
      FundCnt = arr(18)
      TotFNUms = FundCnt
    End If
  ElseIf arr(3) <> EmpName Then
    AcctCnt = 0
    TotANums = AcctCnt
    If arr(17) <> "" Then
      AcctCnt = arr(17)
      TotANums = AcctCnt
    End If
    FundCnt = 0
    TotFNUms = FundCnt
    If arr(18) <> "" Then
      FundCnt = arr(18)
      TotFNUms = FundCnt
    End If
    EmpName = arr(3)
  End If

  Fields("fldEmpSplCnt").Value = arr(18)
  Fields("fldEmpFundNum").Value = arr(19)

  If arr(19) = "" Then
    Label103.Visible = False
  Else
    Label103.Visible = True
  End If

  Fields("fldFundDedAmt1").Value = arr(20)
  Fields("fldFundDedAmt2").Value = arr(21)
  Fields("fldFundDedAmt3").Value = arr(22)
  Fields("fldFundDedAmt4").Value = arr(23)
  Fields("fldFundDedAmt5").Value = arr(24)
  Fields("fldFundDedAmt6").Value = arr(25)
  Fields("fldFundDedAmt7").Value = arr(26)
  Fields("fldFundDedAmt8").Value = arr(27)
  Fields("fldFundDedAmt9").Value = arr(28)
  Fields("fldFundDedAmt10").Value = arr(29)
  Fields("fldFundDedAmt11").Value = arr(30)
  Fields("fldFundDedAmt12").Value = arr(31)
  Fields("fldFundDedAmt13").Value = arr(32)
  Fields("fldFundDedAmt14").Value = arr(33)
  Fields("fldFundDedAmt15").Value = arr(34)
  Fields("fldFundDedAmt16").Value = arr(35)
  Fields("fldFundDedAmt17").Value = arr(36)
  Fields("fldFundDedAmt18").Value = arr(37)
  Fields("fldFundDedAmt19").Value = arr(38)
  Fields("fldFundDedAmt20").Value = arr(39)
  Fields("fldFundDedAmt21").Value = arr(40)
  Fields("fldFundDedAmt22").Value = arr(41)
  Fields("fldFundDedAmt23").Value = arr(42)
  Fields("fldFundDedAmt24").Value = arr(43)
  Fields("fldFundDedAmt25").Value = arr(44)
  Fields("fldFundDedAmt26").Value = arr(45)
  Fields("fldFundDedAmt27").Value = arr(46)
  Fields("fldFundDedAmt28").Value = arr(47)
  Fields("fldFundDedAmt29").Value = arr(48)
  Fields("fldFundDedAmt30").Value = arr(49)
  Fields("fldFundDedAmt31").Value = arr(50)
  Fields("fldFundDedAmt32").Value = arr(51)
  Fields("fldFundDedAmt33").Value = arr(52)
  Fields("fldFundDedAmt34").Value = arr(53)
  Fields("fldFundDedAmt35").Value = arr(54)
  Fields("fldFundDedAmt36").Value = arr(55)
  Fields("fldFundDedAmt37").Value = arr(56)
  Fields("fldFundDedAmt38").Value = arr(57)
  Fields("fldFundDedAmt39").Value = arr(58)
  Fields("fldFundDedAmt40").Value = arr(59)
  Fields("fldFundDedAmt41").Value = arr(60)
  Fields("fldFundDedAmt42").Value = arr(61)
  Fields("fldFundDedAmt43").Value = arr(62)
  Fields("fldFundDedAmt44").Value = arr(63)
  Fields("fldFundDedAmt45").Value = arr(64)
  Fields("fldFundDedAmt46").Value = arr(65)
  Fields("fldFundDedAmt47").Value = arr(66)
  Fields("fldFundDedAmt48").Value = arr(67)
  Fields("fldFundDedAmt49").Value = arr(68)
  Fields("fldFundDedAmt50").Value = arr(69)
  Fields("fldFundFed").Value = arr(70)
  Fields("fldFundSta").Value = arr(71)
  Fields("fldFundMed").Value = arr(72)
  Fields("fldFundSoc").Value = arr(73)
  Fields("fldFundRet").Value = arr(74)
  Fields("fldAcctCnt1").Value = arr(75)
  Fields("fldToday").Value = arr(76)
  Fields("fldEmpSplCnt").Value = arr(77)

End Sub
Private Sub ActiveReport_ReportEnd()
  Unload frmLoadingRpt
  If hFile <> 0 Then
    Close #hFile
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Dim DedRec As DedCodeRecType
  Dim DHandle As Integer
  Dim x As Integer

  Me.Zoom = -1
  OpenDedCodeFile DHandle
  DedCnt = LOF(DHandle) / Len(DedRec)
  Close DHandle
  Set SubReport1.object = New arDistAcctSummary
  ReportHeader.Height = 0
  Label47.Visible = False 'Summary
  Me.fldTimeDate.Text = Now
End Sub

Private Sub Detail_Format()
  Dim UseThisOne As Integer
  If TotANums > 0 Then
    UseThisOne = TotANums
    Detail.Height = 200
    Label103.Visible = False
    Exit Sub
  ElseIf TotFNUms > 0 Then
    GoTo UseFNums
    Label103.Visible = True
  Else
    GoTo UseFNums
    Label103.Visible = True
  End If
UseFNums:
  Select Case DedCnt
    Case 1 To 10
      Detail.Height = 600
    Case 11 To 20
      Detail.Height = 1000
    Case 21 To 30
      Detail.Height = 1400
    Case 31 To 40
      Detail.Height = 1600
    Case 41 To 50
      Detail.Height = 2000

    Case Else
      Detail.Height = 2000
  End Select
End Sub

Private Sub GroupFooter1_Format()
  Dim ctrl As Control
  Dim sec As Section
  Dim y As Integer
  Set sec = arDistRegAcctNum.Sections("GroupFooter1")
  For y = 0 To sec.Controls.Count - 1
    If y - 17 > DedCnt Then
      sec.Controls(y).Visible = False
    Else
      sec.Controls(y).Visible = True
    End If
    If sec.Controls(y).Name = "Line5" Then
      sec.Controls(y).Visible = True
    ElseIf sec.Controls(y).Name = "Label104" Then
      sec.Controls(y).Visible = True
    ElseIf sec.Controls(y).Name = "Label105" Then
      sec.Controls(y).Visible = True
    End If
  Next y
  Select Case DedCnt
  Case 0 To 10
    GroupFooter1.Height = 1650
    Line5.Y1 = 1650
    Line5.Y2 = 1650
  Case 11 To 20
    GroupFooter1.Height = 1700
    Line5.Y1 = 1700
    Line5.Y2 = 1700
  Case 21 To 30
    GroupFooter1.Height = 1975
    Line5.Y1 = 1975
    Line5.Y2 = 1975
  Case 31 To 40
    GroupFooter1.Height = 2300
    Line5.Y1 = 2300
    Line5.Y2 = 2300
  Case 41 To 50
    GroupFooter1.Height = 2500
    Line5.Y1 = 2500
    Line5.Y2 = 2500
  Case Else
    GroupFooter1.Height = 2500
    Line5.Y1 = 2500
    Line5.Y2 = 2500
  End Select

End Sub

Private Sub PageHeader_Format()

  Set SubReport2.object = New arDedDescs
'  If EndReport = True Then
'    Label47.Visible = True
'    PageHeader.Height = 1000
''    Line6.Visible = False
''    Line6.Visible = False
'  End If
  Select Case DedCnt
  Case 0 To 10
    SubReport2.Height = 275
    PageHeader.Height = 3200
    Line6.Y1 = 3200
    Line6.Y2 = 3200
  Case 11 To 20
    SubReport2.Height = 750
    PageHeader.Height = 3400
    Line6.Y1 = 3420
    Line6.Y2 = 3420
  Case 21 To 30
    SubReport2.Height = 1000
    PageHeader.Height = 4000
    Line6.Y1 = 3720
    Line6.Y2 = 3720
  Case 31 To 40
    SubReport2.Height = 1200
    PageHeader.Height = 4300
    Line6.Y1 = 4020
    Line6.Y2 = 4020
  Case 41 To 50
    SubReport2.Height = 1500
    PageHeader.Height = 4300
    Line6.Y1 = 4300
    Line6.Y2 = 4300
  Case Else
    SubReport2.Height = 1500
    PageHeader.Height = 4300
    Line6.Y1 = 4300
    Line6.Y2 = 4300
  End Select
End Sub

Private Sub ReportFooter_Format()
  EndReport = True
  PageHeader.Height = 1200
  Label47.Visible = True
  Line6.Visible = False
  Line6.Visible = False
  Label95.Visible = False
  Label96.Visible = False
  Label97.Visible = False
  Label98.Visible = False
  Label99.Visible = False
  Label102.Visible = False
   SubReport2.Visible = False
  Set SubReport3.object = New arSubFundTotals
  Set SubReport4.object = New arDedDescs
  GroupHeader1.Height = 0
  Label48.Visible = False
  Label49.Visible = False
  Label50.Visible = False
  Detail.Height = 0
End Sub

Private Sub ReportHeader_Format()
  ReportHeader.Height = 0
End Sub







