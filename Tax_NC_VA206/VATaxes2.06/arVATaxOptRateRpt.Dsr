VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxOptRateRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Optional Revenue Rate Report"
   ClientHeight    =   8736
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arVATaxOptRateRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arVATaxOptRateRpt.dsx":08CA
End
Attribute VB_Name = "arVATaxOptRateRpt"
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
  Open StartPath & "\TAXRPTS\TXOPRATE.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldDesc") '1)
  Fields.Add ("fldFlatAmt") '2)
  Fields.Add ("fldOptRevNum") '3)
  Fields.Add ("fldType") '4)
  Fields.Add ("fldFrom1") '5)
  Fields.Add ("fldTo1") '6)
  Fields.Add ("fldTaxF1") '7)
  Fields.Add ("fldTaxP1") ' 8)
  Fields.Add ("fldFrom2") ' 9)
  Fields.Add ("fldTo2") '10)
  Fields.Add ("fldTaxF2") '11)
  Fields.Add ("fldTaxP2") '12)
  Fields.Add ("fldFrom3") '13)
  Fields.Add ("fldTo3") '14)
  Fields.Add ("fldTaxF3") '15)
  Fields.Add ("fldTaxP3") '16)
  Fields.Add ("fldFrom4") '17)
  Fields.Add ("fldTo4") '18)
  Fields.Add ("fldTaxF4") '19)
  Fields.Add ("fldTaxP4") '20)
  Fields.Add ("fldFrom5") '21)
  Fields.Add ("fldTo5") '22)
  Fields.Add ("fldTaxF5") ' 23)
  Fields.Add ("fldTaxP5") '24)
  Fields.Add ("fldFrom6") '25)
  Fields.Add ("fldTo6") '26)
  Fields.Add ("fldTaxF6") '27)
  Fields.Add ("fldTaxP6") '28)
  Fields.Add ("fldFrom7") '29)
  Fields.Add ("fldTo7") '30)
  Fields.Add ("fldTaxF7") ' 31)
  Fields.Add ("fldTaxP7") '32)
  Fields.Add ("fldFrom8") '33)
  Fields.Add ("fldTo8") '34)
  Fields.Add ("fldTaxF8") '35)
  Fields.Add ("fldTaxP8") '36)
  Fields.Add ("fldFrom9") '37)
  Fields.Add ("fldTo9") '38)
  Fields.Add ("fldTaxF9") '39)
  Fields.Add ("fldTaxP9") '40)
  Fields.Add ("fldFrom10") ' 41)
  Fields.Add ("fldTo10") '42)
  Fields.Add ("fldTaxF10") '43)
  Fields.Add ("fldTaxP10") '44)
  Fields.Add ("fldRevType") '45)
  
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
  Fields("fldTown").Value = arr(0)
  Fields("fldDesc").Value = arr(1)
  Fields("fldFlatAmt").Value = arr(2)
  Fields("fldOptRevNum").Value = arr(3)
  If arr(4) = "F" Then
    Fields("fldType").Value = "Flat Rate"
  ElseIf arr(4) = "S" Then
    Fields("fldType").Value = "Step Flat"
  ElseIf arr(4) = "P" Then
    Fields("fldType").Value = "Step Pct"
  End If
  Fields("fldFrom1").Value = arr(5)
  Fields("fldTo1").Value = arr(6)
  Fields("fldTaxF1").Value = arr(7)
  Fields("fldTaxP1").Value = CDbl(arr(8)) / 100
  Fields("fldFrom2").Value = arr(9)
  Fields("fldTo2").Value = arr(10)
  Fields("fldTaxF2").Value = arr(11)
  Fields("fldTaxP2").Value = CDbl(arr(12)) / 100
  Fields("fldFrom3").Value = arr(13)
  Fields("fldTo3").Value = arr(14)
  Fields("fldTaxF3").Value = arr(15)
  Fields("fldTaxP3").Value = CDbl(arr(16)) / 100
  Fields("fldFrom4").Value = arr(17)
  Fields("fldTo4").Value = arr(18)
  Fields("fldTaxF4").Value = arr(19)
  Fields("fldTaxP4").Value = CDbl(arr(20)) / 100
  Fields("fldFrom5").Value = arr(21)
  Fields("fldTo5").Value = arr(22)
  Fields("fldTaxF5").Value = arr(23)
  Fields("fldTaxP5").Value = CDbl(arr(24)) / 100
  Fields("fldFrom6").Value = arr(25)
  Fields("fldTo6").Value = arr(26)
  Fields("fldTaxF6").Value = arr(27)
  Fields("fldTaxP6").Value = CDbl(arr(28)) / 100
  Fields("fldFrom7").Value = arr(29)
  Fields("fldTo7").Value = arr(30)
  Fields("fldTaxF7").Value = arr(31)
  Fields("fldTaxP7").Value = CDbl(arr(32)) / 100
  Fields("fldFrom8").Value = arr(33)
  Fields("fldTo8").Value = arr(34)
  Fields("fldTaxF8").Value = arr(35)
  Fields("fldTaxP8").Value = CDbl(arr(36)) / 100
  Fields("fldFrom9").Value = arr(37)
  Fields("fldTo9").Value = arr(38)
  Fields("fldTaxF9").Value = arr(39)
  Fields("fldTaxP9").Value = CDbl(arr(40)) / 100
  Fields("fldFrom10").Value = arr(41)
  Fields("fldTo10").Value = arr(42)
  Fields("fldTaxF10").Value = arr(43)
  Fields("fldTaxP10").Value = CDbl(arr(44)) / 100
  Fields("fldRevType").Value = arr(45)
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
      frmVATaxMsg.Label1.Caption = "File - TaxOptRevRpt.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - TaxOptRevRpt.txt, created in the Citipak Directory."
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
    frmVATaxMsg.Label1.Caption = "File - TaxOptRevRpt.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - TaxOptRevRpt.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "TaxOptRevRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TaxOptRevRpt.txt"
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
  If Fields("fldType").Value = "Flat Rate" Then
    Detail.Height = 270
    Label30.Visible = False
    Label31.Visible = False
    Label32.Visible = False
    Label33.Visible = False
  Else
    Detail.Height = 3720
    Label30.Visible = True
    Label31.Visible = True
    Label32.Visible = True
    Label33.Visible = True
  End If
End Sub

Private Sub GroupHeader1_Format()
  If Fields("fldRevType").Value = "R" Then
    Label34.Caption = "REAL REVENUE"
  ElseIf Fields("fldRevType").Value = "P" Then
    Label34.Caption = "PERSONAL REVENUE"
  End If
End Sub
