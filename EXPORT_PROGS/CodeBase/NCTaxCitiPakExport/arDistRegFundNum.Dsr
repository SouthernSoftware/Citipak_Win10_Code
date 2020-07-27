VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arDistRegFundNum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distribution Fund Register"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20553
   _ExtentY        =   15642
   SectionData     =   "arDistRegFundNum.dsx":0000
End
Attribute VB_Name = "arDistRegFundNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private HFile As Integer
Dim EndReport As Boolean
Dim DedCnt As Integer
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
      MsgBox "File - DISTRIBUFUNDNUM.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - DISTRIBUFUNDNUM.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
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
    MsgBox "File - DISTRIBUFUNDNUM.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - DISTRIBUFUNDNUM.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "DISTRIBUFUNDNUM.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "DISTRIBUFUNDNUM.txt"
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
  Open App.Path & "\PRRPTS\DISTRIBUFUNDNUM.RPT" For Input As #HFile
  
  Fields.Add ("fldEmployee") '(0)
  Fields.Add ("fldENum") '(1)
  Fields.Add ("fldDedFromTo") '(2)
  
  Fields.Add ("fldEmpFundNum") '(3)
  Fields.Add ("fldFundDedAmt1") '(4)
  Fields.Add ("fldFundDedAmt2") '(5)
  Fields.Add ("fldFundDedAmt3") '(6)
  Fields.Add ("fldFundDedAmt4") '(7)
  Fields.Add ("fldFundDedAmt5") '(8)
  Fields.Add ("fldFundDedAmt6") '(9)
  Fields.Add ("fldFundDedAmt7") '(10)
  Fields.Add ("fldFundDedAmt8") '(11)
  Fields.Add ("fldFundDedAmt9") '(12)
  Fields.Add ("fldFundDedAmt10") '(13)
  Fields.Add ("fldFundDedAmt11") '(14)
  Fields.Add ("fldFundDedAmt12") '(15)
  Fields.Add ("fldFundDedAmt13") '(16)
  Fields.Add ("fldFundDedAmt14") '(17)
  Fields.Add ("fldFundDedAmt15") '(18)
  Fields.Add ("fldFundDedAmt16") '(19)
  Fields.Add ("fldFundDedAmt17") '(20)
  Fields.Add ("fldFundDedAmt18") '(21)
  Fields.Add ("fldFundDedAmt19") '(22)
  Fields.Add ("fldFundDedAmt20") '(23)
  Fields.Add ("fldFundDedAmt21") '(24)
  Fields.Add ("fldFundDedAmt22") '(25)
  Fields.Add ("fldFundDedAmt23") '(26)
  Fields.Add ("fldFundDedAmt24") '(27)
  Fields.Add ("fldFundDedAmt25") '(28)
  Fields.Add ("fldFundDedAmt26") '(29)
  Fields.Add ("fldFundDedAmt27") '(30)
  Fields.Add ("fldFundDedAmt28") '(31)
  Fields.Add ("fldFundDedAmt29") '(32)
  Fields.Add ("fldFundDedAmt30") '(33)
  Fields.Add ("fldFundDedAmt31") '(34)
  Fields.Add ("fldFundDedAmt32") '(35)
  Fields.Add ("fldFundDedAmt33") '(36)
  Fields.Add ("fldFundDedAmt34") '(37)
  Fields.Add ("fldFundDedAmt35") '(38)
  Fields.Add ("fldFundDedAmt36") '(39)
  Fields.Add ("fldFundDedAmt37") '(40)
  Fields.Add ("fldFundDedAmt38") '(41)
  Fields.Add ("fldFundDedAmt39") '(42)
  Fields.Add ("fldFundDedAmt40") '(43)
  Fields.Add ("fldFundDedAmt41") '(44)
  Fields.Add ("fldFundDedAmt42") '(45)
  Fields.Add ("fldFundDedAmt43") '(46)
  Fields.Add ("fldFundDedAmt44") '(47)
  Fields.Add ("fldFundDedAmt45") '(48)
  Fields.Add ("fldFundDedAmt46") '(49)
  Fields.Add ("fldFundDedAmt47") '(50)
  Fields.Add ("fldFundDedAmt48") '(51)
  Fields.Add ("fldFundDedAmt49") '(52)
  Fields.Add ("fldFundDedAmt50") '(53)
  
  
  Fields.Add ("fldFundFed") '(54)
  Fields.Add ("fldFundSta") '(55)
  Fields.Add ("fldFundMed") '(56)
  Fields.Add ("fldFundSoc") '(57)
  Fields.Add ("fldFundRet") '(58)
  Fields.Add ("fldEmployer") '(59)
  Fields.Add ("fldDate") '(60)
  Fields.Add ("AcctCnt") '(61)
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
  
  Fields("fldEmployee").Value = arr(0)
  Fields("fldENum").Value = arr(1)
  Fields("fldDedFromTo").Value = arr(2)
  
  Fields("fldEmpFundNum").Value = arr(3)
  Fields("fldFundDedAmt1").Value = arr(4)
  Fields("fldFundDedAmt2").Value = arr(5)
  Fields("fldFundDedAmt3").Value = arr(6)
  Fields("fldFundDedAmt4").Value = arr(7)
  Fields("fldFundDedAmt5").Value = arr(8)
  Fields("fldFundDedAmt6").Value = arr(9)
  Fields("fldFundDedAmt7").Value = arr(10)
  Fields("fldFundDedAmt8").Value = arr(11)
  Fields("fldFundDedAmt9").Value = arr(12)
  Fields("fldFundDedAmt10").Value = arr(13)
  Fields("fldFundDedAmt11").Value = arr(14)
  Fields("fldFundDedAmt12").Value = arr(15)
  Fields("fldFundDedAmt13").Value = arr(16)
  Fields("fldFundDedAmt14").Value = arr(17)
  Fields("fldFundDedAmt15").Value = arr(18)
  Fields("fldFundDedAmt16").Value = arr(19)
  Fields("fldFundDedAmt17").Value = arr(20)
  Fields("fldFundDedAmt18").Value = arr(21)
  Fields("fldFundDedAmt19").Value = arr(22)
  Fields("fldFundDedAmt20").Value = arr(23)
  Fields("fldFundDedAmt21").Value = arr(24)
  Fields("fldFundDedAmt22").Value = arr(25)
  Fields("fldFundDedAmt23").Value = arr(26)
  Fields("fldFundDedAmt24").Value = arr(27)
  Fields("fldFundDedAmt25").Value = arr(28)
  Fields("fldFundDedAmt26").Value = arr(29)
  Fields("fldFundDedAmt27").Value = arr(30)
  Fields("fldFundDedAmt28").Value = arr(31)
  Fields("fldFundDedAmt29").Value = arr(32)
  Fields("fldFundDedAmt30").Value = arr(33)
  Fields("fldFundDedAmt31").Value = arr(34)
  Fields("fldFundDedAmt32").Value = arr(35)
  Fields("fldFundDedAmt33").Value = arr(36)
  Fields("fldFundDedAmt34").Value = arr(37)
  Fields("fldFundDedAmt35").Value = arr(38)
  Fields("fldFundDedAmt36").Value = arr(39)
  Fields("fldFundDedAmt37").Value = arr(40)
  Fields("fldFundDedAmt38").Value = arr(41)
  Fields("fldFundDedAmt39").Value = arr(42)
  Fields("fldFundDedAmt40").Value = arr(43)
  Fields("fldFundDedAmt41").Value = arr(44)
  Fields("fldFundDedAmt42").Value = arr(45)
  Fields("fldFundDedAmt43").Value = arr(46)
  Fields("fldFundDedAmt44").Value = arr(47)
  Fields("fldFundDedAmt45").Value = arr(48)
  Fields("fldFundDedAmt46").Value = arr(49)
  Fields("fldFundDedAmt47").Value = arr(50)
  Fields("fldFundDedAmt48").Value = arr(51)
  Fields("fldFundDedAmt49").Value = arr(52)
  Fields("fldFundDedAmt50").Value = arr(53)
  
  Fields("fldFundFed").Value = arr(54)
  Fields("fldFundSta").Value = arr(55)
  Fields("fldFundMed").Value = arr(56)
  Fields("fldFundSoc").Value = arr(57)
  Fields("fldFundRet").Value = arr(58)
  Fields("fldEmployer").Value = arr(59)
  Fields("fldDate").Value = arr(60)
  Fields("AcctCnt").Value = arr(61)
End Sub
Private Sub ActiveReport_ReportEnd()
  Unload frmLoadingRpt
  If HFile <> 0 Then
    Close #HFile
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
  
  ReportHeader.Height = 0
  Label47.Visible = False 'Summary
  Me.fldTimeDate.Text = Now
  Select Case DedCnt
  Case 0 To 10
    PageHeader.Height = 2400
    Detail.Height = 575
    GroupFooter1.Height = 650
    Line6.Y1 = 700
    Line6.Y2 = 700
    Line8.Y1 = 2350
    Line8.Y2 = 2350
  Case 11 To 20
    PageHeader.Height = 2650
    Detail.Height = 700
    GroupFooter1.Height = 1100
    Line6.Y1 = 1050
    Line6.Y2 = 1050
    Line8.Y1 = 2550
    Line8.Y2 = 2550
  Case 21 To 30
    PageHeader.Height = 2650
    Detail.Height = 900
    GroupFooter1.Height = 1200
    Line6.Y1 = 1250
    Line6.Y2 = 1250
    Line8.Y1 = 2775
    Line8.Y2 = 2775
  Case 31 To 40
    PageHeader.Height = 2850
    Detail.Height = 1200
    GroupFooter1.Height = 1300
    Line6.Y1 = 1500
    Line6.Y2 = 1500
    Line8.Y1 = 3100
    Line8.Y2 = 3100
  Case 41 To 50
    PageHeader.Height = 2850
  Case Else
    PageHeader.Height = 2850
  End Select
End Sub

Private Sub Detail_Format()
  If Fields("AcctCnt").Value <= CInt("1") Then
    GroupFooter1.Visible = False
  Else
    GroupFooter1.Visible = True
  End If
End Sub

Private Sub GroupFooter1_Format()
  Dim ctrl As Control
  Dim sec As Section
  Dim Y As Integer
  Set sec = arDistRegFundNum.Sections("GroupFooter1")
  For Y = 0 To sec.Controls.Count - 1
    If Y - 5 > DedCnt Then
      sec.Controls(Y).Visible = False
    Else
      sec.Controls(Y).Visible = True
    End If
    If sec.Controls(Y).Name = "Line6" Then
      sec.Controls(Y).Visible = True
    End If
  Next Y

End Sub
Private Sub PageHeader_Format()
  If EndReport = True Then
    Label47.Visible = True
  End If
  Set SubReport1.object = New arDedDescs
  Select Case DedCnt
  Case 0 To 10
    SubReport1.Height = 275
  Case 11 To 20
    SubReport1.Height = 500
  Case 21 To 30
    SubReport1.Height = 800
  Case 31 To 40
    SubReport1.Height = 1250
  Case 41 To 50
    SubReport1.Height = 1250
  Case Else
    SubReport1.Height = 1250
  End Select
  
End Sub

Private Sub ReportFooter_Format()
  EndReport = True
  Set SubReport2.object = New arSubFundTotals
  GroupHeader1.Height = 0
End Sub

Private Sub ReportHeader_Format()
  ReportHeader.Height = 0
End Sub

