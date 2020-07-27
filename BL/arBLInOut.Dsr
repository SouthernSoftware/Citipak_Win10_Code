VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arBLInOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Inside/Outside City Limits Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   ControlBox      =   0   'False
   Icon            =   "arBLInOut.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arBLInOut.dsx":08CA
End
Attribute VB_Name = "arBLInOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsBLTextBoxOverrider
Private Temp_Class As Resize_Class
Private hFile As Integer

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\BLRPTS\ARIORPT.RPT" For Input As #hFile
  Fields.Add ("fld0") '0)
  Fields.Add ("fld1") '1)
  Fields.Add ("fld2") '2)
  Fields.Add ("fld3") '3)
  Fields.Add ("fld4") '4)
  Fields.Add ("fld5") '5)
  Fields.Add ("fld6") '6)
  Fields.Add ("fld7") '7)
  Fields.Add ("fld8") '8)
  Fields.Add ("fld9") '9)
  Fields.Add ("fld10") '10)
  Fields.Add ("fld11") '11)
  Fields.Add ("fld12") '12)
  Fields.Add ("fld13") '13)
  Fields.Add ("fld14") '14)
  Fields.Add ("fld15") '15)
  Fields.Add ("fld16") '16)
  Fields.Add ("fld17") '17)
  Fields.Add ("fld18") '18)
  Fields.Add ("fld19") '19)
  Fields.Add ("fld20") '20)
  Fields.Add ("fld21") '21)
  Fields.Add ("fld22") '22)
  Fields.Add ("fld23") '23)
  Fields.Add ("fld24") '24)
  Fields.Add ("fld25") '25)
  Fields.Add ("fld26") '26)
  Fields.Add ("fld27") '27)
  Fields.Add ("fld28") '28)
  Fields.Add ("fld29") '29)
  Fields.Add ("fld30") '30)
  Fields.Add ("fld31") '31)
  Fields.Add ("fld32") '32)
  Fields.Add ("fld33") '33)
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmBLLoadReport
    frmBLMessageBoxJr.Label1.Caption = "Error Number: " & Str(Number) & " " & Description
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
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
  Fields("fld0").Value = arr(0) 'CustCnt1
  Fields("fld1").Value = arr(1) 'CustCnt2
  Fields("fld2").Value = arr(2) 'CustCnt3
  Fields("fld3").Value = arr(3) 'GTotBal1
  Fields("fld4").Value = arr(4) 'GTotBal2
  Fields("fld5").Value = arr(5) 'GTotBal3
  Fields("fld6").Value = arr(6) 'GTotFees1
  Fields("fld7").Value = arr(7) 'GTotFees2
  Fields("fld8").Value = arr(8) 'GTotFees3
  Fields("fld9").Value = arr(9) 'CustRecNum
  Fields("fld10").Value = arr(10) 'BillName
  Fields("fld11").Value = arr(11) 'Lic Num
  Fields("fld12").Value = arr(12) 'Ex Date
  Fields("fld13").Value = arr(13) 'Cust Acct Bal
  Fields("fld14").Value = arr(14) 'Cust Fees
  Fields("fld15").Value = arr(15) 'TownName
  Fields("fld16").Value = arr(16) 'WhereFlag
  Fields("fld17").Value = arr(17) 'Footer Title
  Fields("fld18").Value = arr(18) 'AveBal(WhereFlag)
  Fields("fld19").Value = arr(19) 'AveBal1
  Fields("fld20").Value = arr(20) 'AveBal2
  Fields("fld21").Value = arr(21) 'AveBal3
  Fields("fld22").Value = arr(22) 'AveFee(WhereFlag)
  Fields("fld23").Value = arr(23) 'AveFee1
  Fields("fld24").Value = arr(24) 'AveFee2
  Fields("fld25").Value = arr(25) 'AveFee3
  Fields("fld26").Value = arr(26) 'GrandBalAve
  Fields("fld27").Value = arr(27) 'GrandFeeAve
  Fields("fld28").Value = arr(28) 'GrandCustCnt
  Fields("fld29").Value = arr(29) 'GrandBal
  Fields("fld30").Value = arr(30) 'GrandFees
  Fields("fld31").Value = arr(31) 'CustCnt(WhereFlag)
  Fields("fld32").Value = arr(32) 'Code Desc
  Fields("fld33").Value = arr(33) 'Iss Fee
  If Val(arr(33)) > 0 Then
    Label36.Visible = True
    Label36.Caption = "Current fees include a " + QPTrim$(Using$("$#,##0.00", CDbl(arr(33)))) + " issuance fee."
  Else
    Label36.Visible = False
  End If
End Sub

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
      frmBLMessageBoxJr.Label1.Caption = "File - BLInOutRpt.xls, created in the Citipak Directory."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmBLMessageBoxJr.Label1.Caption = "File - BLInOutRpt.txt, created in the Citipak Directory."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
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
    frmBLMessageBoxJr.Label1.Caption = "File - BLInOutRpt.xls, created in the Citipak Directory."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmBLMessageBoxJr.Label1.Caption = "File - BLInOutRpt.txt, created in the Citipak Directory."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
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
        oEXL.FileName = outfile & "BLInOutRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "BLInOutRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmBLLoadReport
  If hFile <> 0 Then
    Close #hFile
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Date
  Me.Zoom = -1
  Label17.Visible = False
End Sub

Private Sub ReportFooter_Format()
  Label17.Visible = True
  Label16.Visible = False
  Label18.Visible = False
  Label19.Visible = False
  Label20.Visible = False
  Label21.Visible = False
  Label22.Visible = False
  Line1.Visible = False
End Sub

Private Sub ReportFooter_BeforePrint()
  Dim ThisPct As Double
  Dim ThisTop As Integer
  Dim ThisTotal As Double
  Dim AllHeight As Integer
  Dim AllTop As Integer
  
  AllHeight = Inside.Height 'establish max height
  AllTop = Inside.Top 'establish max top
  
'  Field14.DataValue = 200 'test for "%" label placements
'  Field15.DataValue = 100 'test for "%" label placements
'  Field16.DataValue = 2700 'test for "%" label placements
'  ThisTotal = Field14.DataValue + Field15.DataValue + Field16.DataValue 'test for "%" label placements
  ThisTotal = Field17.DataValue 'total used to figure percentages
  
  If ThisTotal = 0 Then
    Label43.Caption = "NA"
    Label44.Caption = "NA"
    Label45.Caption = "NA"
    GoTo OuttaHere
  End If
  
  If Field14.DataValue < 0 Then 'in case the balance is negative
    Label43.Caption = "< 0%" 'can't drop the bar below the line because it would
    Inside.Height = 0 'interfere with graphics below the line
    Label43.Top = (AllHeight + AllTop) - 410
    Label43.ForeColor = &H0&
  Else
    ThisPct = Field14.DataValue / ThisTotal 'figure first bar's % value
    Label43.Caption = CStr(OldRound(ThisPct * 100)) + "%" 'assign % value to label for first bar
    Inside.Height = AllHeight * ThisPct 'now figure height of first bar
    
    ThisTop = AllTop - Inside.Height 'this value is how much to lower the first bar from the established top level
    Inside.Top = AllHeight + ThisTop 'this drops the top lower (adding pushes the top down) to match the first bar's percentage
    Label43.Top = Inside.Top + 50 'now position the "%" label just below the top of the first bar
    If ((AllHeight + AllTop) - Label43.Top) <= 360 Then 'if the first bar's percentage would be so small that it
    'would not produce a bar tall enough to hold the "%" label then position the "%" label above the bar and change the font color
      Label43.Top = Inside.Top - 410 'the label is 360 tall so add 50 to allow room between the "%" label and the
      'top of the bar
      Label43.ForeColor = &H0&
    End If
  End If

  If Field15.DataValue < 0 Then
    Label44.Caption = "< 0%"
    Outside.Height = 0
    Label44.Top = (AllHeight + AllTop) - 410
    Label44.ForeColor = &H0&
  Else
    ThisPct = Field15.DataValue / ThisTotal
    Label44.Caption = CStr(OldRound(ThisPct * 100)) + "%"
    Outside.Height = AllHeight * ThisPct
    ThisTop = AllTop - Outside.Height
    Outside.Top = AllHeight + ThisTop
    Label44.Top = Outside.Top + 50
    If ((AllHeight + AllTop) - Label44.Top) <= 360 Then
      Label44.Top = Outside.Top - 410
      Label44.ForeColor = &H0&
    End If
  End If

  If Field16.DataValue < 0 Then
    Label45.Caption = "< 0%"
    Unknown.Height = 0
    Label45.Top = (AllHeight + AllTop) - 410
    Label45.ForeColor = &H0&
  Else
    ThisPct = Field16.DataValue / ThisTotal
    Label45.Caption = CStr(OldRound(ThisPct * 100)) + "%"
    Unknown.Height = AllHeight * ThisPct
    ThisTop = AllTop - Unknown.Height
    Unknown.Top = AllHeight + ThisTop
    Label45.Top = Unknown.Top + 50
    If ((AllHeight + AllTop) - Label45.Top) <= 360 Then
      Label45.Top = Unknown.Top - 410
      Label45.ForeColor = &H0&
    End If
  End If

  '---------------------------------------------------
'  Field18.DataValue = 200 'test for "%" label placements
'  Field19.DataValue = 100 'test for "%" label placements
'  Field20.DataValue = 2700 'test for "%" label placements
'  ThisTotal = Field18.DataValue + Field19.DataValue + Field20.DataValue 'test for "%" label placements
OuttaHere:
  
  ThisTotal = Field21.DataValue
  If ThisTotal = 0 Then
    Label54.Caption = "NA"
    Label55.Caption = "NA"
    Label56.Caption = "NA"
    GoTo OuttaHere2
  End If
  
  ThisPct = Field18.DataValue / ThisTotal
  Label54.Caption = CStr(OldRound(ThisPct * 100)) + "%"
  Inside2.Height = AllHeight * ThisPct
  ThisTop = AllTop - Inside2.Height
  Inside2.Top = AllHeight + ThisTop
  Label54.Top = Inside2.Top + 50
  If ((AllHeight + AllTop) - Label54.Top) <= 360 Then
    Label54.Top = Inside2.Top - 410
    Label54.ForeColor = &H0&
  End If
  
  ThisPct = Field19.DataValue / ThisTotal
  Label55.Caption = CStr(OldRound(ThisPct * 100)) + "%"
  Outside2.Height = AllHeight * ThisPct
  ThisTop = AllTop - Outside2.Height
  Outside2.Top = AllHeight + ThisTop
  Label55.Top = Outside2.Top + 50
  If ((AllHeight + AllTop) - Label55.Top) <= 360 Then
    Label55.Top = Outside2.Top - 410
    Label55.ForeColor = &H0&
  End If

  ThisPct = Field20.DataValue / ThisTotal
  Label56.Caption = CStr(OldRound(ThisPct * 100)) + "%"
  Unknown2.Height = AllHeight * ThisPct
  ThisTop = AllTop - Unknown2.Height
  Unknown2.Top = AllHeight + ThisTop
  Label56.Top = Unknown2.Top + 50
  If ((AllHeight + AllTop) - Label56.Top) <= 360 Then
    Label56.Top = Unknown2.Top - 410
    Label56.ForeColor = &H0&
  End If

'----------------------------------------------------------
OuttaHere2:
  AllHeight = Inside3.Height 'establish max height
  AllTop = Inside3.Top 'establish max top
 
  ThisTotal = Field31.DataValue
  
  If ThisTotal = 0 Then
    Label67.Caption = "NA"
    Label68.Caption = "NA"
    Label69.Caption = "NA"
    Exit Sub
  End If

  ThisPct = Field22.DataValue / ThisTotal
  Label67.Caption = CStr(OldRound(ThisPct * 100)) + "%"
  Inside3.Height = AllHeight * ThisPct
  ThisTop = AllTop - Inside3.Height
  Inside3.Top = AllHeight + ThisTop
  Label67.Top = Inside3.Top + 50
  If ((AllHeight + AllTop) - Label67.Top) <= 360 Then
    Label67.Top = Inside3.Top - 410
    Label67.ForeColor = &H0&
  End If

  ThisPct = Field23.DataValue / ThisTotal
  Label68.Caption = CStr(OldRound(ThisPct * 100)) + "%"
  Outside3.Height = AllHeight * ThisPct
  ThisTop = AllTop - Outside3.Height
  Outside3.Top = AllHeight + ThisTop
  Label68.Top = Outside3.Top + 50
  If ((AllHeight + AllTop) - Label68.Top) <= 360 Then
    Label68.Top = Outside3.Top - 410
    Label68.ForeColor = &H0&
  End If

  ThisPct = Field24.DataValue / ThisTotal
  Label69.Caption = CStr(OldRound(ThisPct * 100)) + "%"
  Unknown3.Height = AllHeight * ThisPct
  ThisTop = AllTop - Unknown3.Height
  Unknown3.Top = AllHeight + ThisTop
  Label69.Top = Unknown3.Top + 50
  If ((AllHeight + AllTop) - Label69.Top) <= 360 Then
    Label69.Top = Unknown3.Top - 410
    Label69.ForeColor = &H0&
  End If

End Sub

