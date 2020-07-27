VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arBLTrnsJrnlCatAnalysis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Journal - Catagory Analysis"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15452
   SectionData     =   "arBLTrnsJrnlCatAnalysis.dsx":0000
End
Attribute VB_Name = "arBLTrnsJrnlCatAnalysis"
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
  Open StartPath & "\BLRPTS\ARCLSANL.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldBDate") '1)
  Fields.Add ("fldEDate") '2)
  Fields.Add ("fldHeader") '3)
  Fields.Add ("fldCodeDesc") '4)
  Fields.Add ("fldCodeNum") '5)
  Fields.Add ("fldCatTot") '6)
  Fields.Add ("fldTypeAmt1") '7)
  Fields.Add ("fldTypeAmt2") '8)
  Fields.Add ("fldTypeAmt3") '9)
  Fields.Add ("fldTypeAmt4") '10)
  Fields.Add ("fldTypeAmt5") '11)
  Fields.Add ("fldTypeAmt6") '12)
  Fields.Add ("fldChrgTot") '13)
  Fields.Add ("fldPayTot") '14)
  Fields.Add ("fldPenTot") '15)
  Fields.Add ("fldAdPayDn") '16)
  Fields.Add ("fldAdBillDn") '17)
  Fields.Add ("fldAdBillUp") '18)
  Fields.Add ("fldIssChrg") '19)
  Fields.Add ("fldIssPay") '20)
  Fields.Add ("fldPenPay") '21)
  Fields.Add ("fldLicChrg")
  Fields.Add ("fldLicPay")
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
  Fields("fldTown").Value = arr(0)
  Fields("fldBDate").Value = arr(1)
  Fields("fldEDate").Value = arr(2)
  Fields("fldHeader").Value = arr(3)
  Fields("fldCodeDesc").Value = arr(4)
  Fields("fldCodeNum").Value = arr(5)
  Fields("fldCatTot").Value = arr(6)
  Fields("fldTypeAmt1").Value = arr(7)
  Fields("fldTypeAmt2").Value = arr(8)
  Fields("fldTypeAmt3").Value = arr(9)
  Fields("fldTypeAmt4").Value = arr(10)
  Fields("fldTypeAmt5").Value = arr(11)
  Fields("fldTypeAmt6").Value = arr(12)
  Fields("fldChrgTot").Value = OldRound(CDbl(arr(13)) + CDbl(arr(15)))
  Fields("fldPayTot").Value = arr(14)
  Fields("fldPenTot").Value = arr(15)
  Fields("fldAdPayDn").Value = arr(16)
  Fields("fldAdBillDn").Value = arr(17)
  Fields("fldAdBillUp").Value = arr(18)
  Fields("fldIssChrg").Value = arr(19)
  Fields("fldIssPay").Value = arr(20)
  Fields("fldPenPay").Value = arr(21)
  Fields("fldLicChrg").Value = OldRound(CDbl(arr(13)) - CDbl(arr(19)))
  Fields("fldLicPay").Value = OldRound(CDbl(arr(14)) - CDbl(arr(20)) - CDbl(arr(21)))
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
      frmBLMessageBoxJr.Label1.Caption = "File - BLTJrnlAnlys.xls, created in the Citipak Directory."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmBLMessageBoxJr.Label1.Caption = "File - BLTJrnlAnlys.txt, created in the Citipak Directory."
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
    frmBLMessageBoxJr.Label1.Caption = "File - BLTJrnlAnlys.xls, created in the Citipak Directory."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmBLMessageBoxJr.Label1.Caption = "File - BLTJrnlAnlys.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "BLTJrnlAnlys.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "BLTJrnlAnlys.txt"
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
    DoEvents
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Date
  Me.Zoom = -1
End Sub

Private Sub ReportFooter_Format()
  Line1.Y1 = 1170
  Line1.Y2 = 1170
  Label40.Visible = False
  Label18.Visible = False
  PageHeader.Height = 1362
End Sub



