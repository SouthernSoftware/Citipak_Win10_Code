VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxDiscsAtPayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Discounts Awarded at Bill Payment"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "arVATaxDiscsAtPayment.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15452
   SectionData     =   "arVATaxDiscsAtPayment.dsx":08CA
End
Attribute VB_Name = "arVATaxDiscsAtPayment"
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
  Open StartPath & "\TAXRPTS\DISCATPAY.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldTDate") '1)
  Fields.Add ("fldCustName") '2)
  Fields.Add ("fldCustPin") '3)
  Fields.Add ("fldOptFlag") '4)
  Fields.Add ("fldDiscAmt") '5)
  Fields.Add ("fldTAmt") '6)
  Fields.Add ("fldTaxYear") '7)
  Fields.Add ("fldTownship") '8)
  Fields.Add ("fldTotAmt") '9)
  Fields.Add ("fldTotDiscs") '10)
  Fields.Add ("fldBegDate") '11)
  Fields.Add ("fldEndDate") '12)
  Fields.Add ("fldTaxYr") '13)
  Fields.Add ("fldRange")
End Sub
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmVATaxLoadReport
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
  Fields("fldTDate").Value = arr(1)
  Fields("fldCustName").Value = arr(2)
  Fields("fldCustPin").Value = arr(3)
  Fields("fldOptFlag").Value = arr(4)
  Fields("fldDiscAmt").Value = arr(5)
  Fields("fldTAmt").Value = arr(6)
  Fields("fldTaxYear").Value = arr(7)
  Fields("fldTownship").Value = arr(8)
  Fields("fldTotAmt").Value = arr(9)
  Fields("fldTotDiscs").Value = arr(10)
  Fields("fldBegDate").Value = arr(11)
  Fields("fldEndDate").Value = arr(12)
  Fields("fldTaxYr").Value = arr(13)
  If CInt(arr(4)) = 1 Then
    Fields("fldRange").Value = "From: " + arr(11) + " To: " + arr(12)
  Else
    Fields("fldRange").Value = "Tax Year: " + arr(13)
  End If
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
      frmVATaxMsg.Label1.Caption = "File - DiscGiven.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - DiscGiven.txt, created in the Citipak Directory."
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
    frmVATaxMsg.Label1.Caption = "File - DiscGiven.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - DiscGiven.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "DiscGiven.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "DiscGiven.txt"
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

Private Sub ReportFooter_Format()
  Set SubReport1 = New arVASubDiscsAtPayment
End Sub