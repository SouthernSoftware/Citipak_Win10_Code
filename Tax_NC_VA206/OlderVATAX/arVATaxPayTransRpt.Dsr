VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxPayTransRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Transaction Report"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arVATaxPayTransRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arVATaxPayTransRpt.dsx":08CA
End
Attribute VB_Name = "arVATaxPayTransRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private hFile As Integer
  'Private Temp_Class As Resize_Class
  Dim Completed As Boolean
  Dim ROPAmt As Boolean
  Dim POPAmt As Boolean

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\TAXRPTS\TaxEdPay.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldPayDate") '1)
  Fields.Add ("fldCustNum") '2)
  Fields.Add ("fldCustName") '3)
  Fields.Add ("fldCash") '4)
  Fields.Add ("fldCheck") '5)
  Fields.Add ("fldCharge") '6)
  Fields.Add ("fldDisc") '7)
  Fields.Add ("fldTotPaid") '8)
  Fields.Add ("fldChange") '9)
  Fields.Add ("fldOper") '10)
  Fields.Add ("fldCheckCnt") '11)
  Fields.Add ("fldOPAmt") '12)
  Fields.Add ("fldType") '13)
  Fields.Add ("fldTCnt") '14)
  Fields.Add ("fldTotCash") '15)
  Fields.Add ("fldTotChrg") '16)
  Fields.Add ("fldTotChk") '17)
  Fields.Add ("fldTotChng") '18)
  Fields.Add ("fldGTotChk") '19)
  Fields.Add ("fldGTotCnt") '20)
  Fields.Add ("fldGTotDisc") '21)
  Fields.Add ("fldGTotPaid") '22)
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
    Completed = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #hFile, sLine
  arr = Split(sLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("fldTown").Value = arr(0)
  Fields("fldPayDate").Value = arr(1)
  Fields("fldCustNum").Value = arr(2)
  Fields("fldCustName").Value = arr(3)
  Fields("fldCash").Value = arr(4)
  Fields("fldCheck").Value = arr(5)
  Fields("fldCharge").Value = arr(6)
  Fields("fldDisc").Value = arr(7)
  Fields("fldTotPaid").Value = arr(8)
  Fields("fldChange").Value = arr(9)
  Fields("fldOper").Value = arr(10)
  Fields("fldCheckCnt").Value = arr(11)
  Fields("fldOPAmt").Value = arr(12)
  Fields("fldType").Value = arr(13)
  Fields("fldTCnt").Value = arr(14)
  Fields("fldTotCash").Value = arr(15)
  Fields("fldTotChrg").Value = arr(16)
  Fields("fldTotChk").Value = arr(17)
  Fields("fldTotChng").Value = arr(18)
  Fields("fldGTotChk").Value = arr(19)
  Fields("fldGTotCnt").Value = arr(20)
  Fields("fldGTotDisc").Value = arr(21)
  Fields("fldGTotPaid").Value = arr(22)
  If CDbl(arr(12)) > 0 And Fields("fldType").Value = "R" Then ROPAmt = True
  If CDbl(arr(12)) > 0 And Fields("fldType").Value = "P" Then POPAmt = True
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
      frmVATaxMsg.Label1.Caption = "File - TaxPayTransRpt.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - TaxPayTransRpt.txt, created in the Citipak Directory."
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
    frmVATaxMsg.Label1.Caption = "File - TaxPayTransRpt.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - TaxPayTransRpt.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "TaxPayTransRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TaxPayTransRpt.txt"
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
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  Completed = False
  ROPAmt = False
  POPAmt = False
End Sub

'Private Sub Form_Resize()
'  If Me.WindowState <> vbMinimized Then
'    Me.Visible = False
'    'Temp_Class.ResizeControls Me
'    Me.Visible = True
'    Me.SetFocus
'  End If
'End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Date
  Me.Zoom = -1
  Label30.Visible = False
  Label31.Visible = False
End Sub

Private Sub GroupFooter1_Format()
  Set SubReport1.object = New arVASubTaxPayEditJrnl
End Sub

Private Sub GroupFooter3_Format()
  Static ThisType As Integer
  Line8.Visible = True
  ThisType = ThisType + 1
  
  If ROPAmt = False And ThisType = 1 Then
    GroupFooter2.Height = 1875
    SubReport4.Visible = False
'    SubReport4.Top = 1820
    Label37.Top = 1000
    Line8.Y1 = 1270
    Line8.Y2 = 1270
'    If ThisType = 1 Then
      Label37.Caption = "NO REAL OVERPAYMENT ACTIVITY"
'    Else
'      Label37.Caption = "NO PERSONAL OVERPAYMENT ACTIVITY"
'    End If
  Else
    SubReport4.Visible = True
    Set SubReport4.object = New arVASubTaxPayEditJrnl3
  End If
  
  If POPAmt = False And ThisType <> 1 Then
    GroupFooter2.Height = 1875
    SubReport4.Visible = False
'    SubReport4.Top = 1820
    Label37.Top = 1000
    Line8.Y1 = 1270
    Line8.Y2 = 1270
'    If ThisType = 1 Then
'      Label37.Caption = "NO REAL OVERPAYMENT ACTIVITY"
'    Else
      Label37.Caption = "NO PERSONAL OVERPAYMENT ACTIVITY"
'    End If
  Else
    SubReport4.Visible = True
    Set SubReport4.object = New arVASubTaxPayEditJrnl3
  End If

End Sub

Private Sub GroupHeader3_Format()
  Label30.Visible = True
  If Fields("fldType").Value = "R" Then
    Label33.Caption = "REAL TRANSACTIONS"
  ElseIf Fields("fldType").Value = "P" Then
    Label33.Caption = "PERSONAL TRANSACTIONS"
  End If

End Sub

Private Sub ReportFooter_Format()
  Set SubReport2.object = New arVASubTaxPayEditJrnl2
  Label31.Visible = True
  Label17.Visible = False
  Label16.Visible = False
  Label23.Visible = False
  Label19.Visible = False
  Label20.Visible = False
  Label21.Visible = False
  Label22.Visible = False
  Label24.Visible = False
  Label27.Visible = False
End Sub
