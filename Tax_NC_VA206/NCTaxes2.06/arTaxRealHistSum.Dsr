VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arTaxRealHistSum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Real History Summary"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arTaxRealHistSum.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arTaxRealHistSum.dsx":08CA
End
Attribute VB_Name = "arTaxRealHistSum"
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
  Open StartPath & "\TAXRPTS\REALHISTSUM.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldOwner") '1)
  Fields.Add ("fldTDate") '2)
  Fields.Add ("fldThisPin") '3)
  Fields.Add ("fldGTType") '4)
  Fields.Add ("fldBegDate") '5)
  Fields.Add ("fldEndDate") '6)
  Fields.Add ("fldTaxYear") '7)
  Fields.Add ("fldAmount") '8)
  Fields.Add ("fldTDesc") '9)
  Fields.Add ("fldThisTType") '10)
  Fields.Add ("fldAddr") '11)
  Fields.Add ("fldCustRec") '12)
  Fields.Add ("fldBillCustRec") '13)
  Fields.Add ("fldBillNum") '14)
  Fields.Add ("fldBill2Owner") '15)
  Fields.Add ("fldTCnt") '16)
  Fields.Add ("fldBillBal") '17)
  Fields.Add ("fldDisc") '18)
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
'    Unload frmLoadReport
    frmTaxMsg.Label1.Caption = "Error Number: " & Str(Number) & " " & Description
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
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
  Fields("fldOwner").Value = arr(1)
  Fields("fldTDate").Value = arr(2)
  Fields("fldThisPin").Value = arr(3)
  Fields("fldGTType").Value = arr(4)
  Fields("fldBegDate").Value = arr(5)
  Fields("fldEndDate").Value = arr(6)
  Fields("fldTaxYear").Value = arr(7)
  Fields("fldAmount").Value = arr(8)
  Fields("fldTDesc").Value = arr(9)
  Fields("fldThisTType").Value = arr(10)
  If arr(10) <> "Billing" Then
    Label43.Visible = False
    Field44.Visible = False
    Field42.Visible = False
    Field40.Visible = False
    Field41.Visible = True
    Field39.Visible = True
  Else
    Label43.Visible = True
    Field44.Visible = True
    Field42.Visible = True
    Field40.Visible = True
    Field41.Visible = False
    Field39.Visible = False
  End If
  
  Fields("fldAddr").Value = arr(11)
  If QPTrim$(arr(11)) = "" Then
    Fields("fldAddr").Value = "Not Saved"
  Else
    Fields("fldAddr").Value = arr(11)
  End If
  Fields("fldCustRec").Value = arr(12)
  Fields("fldBillCustRec").Value = arr(13)
  Fields("fldBillNum").Value = arr(14)
  Fields("fldBill2Owner").Value = arr(15)
  Fields("fldTCnt").Value = arr(16)
  Fields("fldBillBal").Value = arr(17)
  If arr(10) <> "Payment" Then
    Field45.Visible = False
  ElseIf arr(10) = "Payment" Then
    Field45.Visible = True
   Fields("fldDisc").Value = arr(18)
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
      frmTaxMsg.Label1.Caption = "File - TaxRealHistDet.xls, created in the Citipak Directory."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmTaxMsg.Label1.Caption = "File - TaxRealHistDet.txt, created in the Citipak Directory."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
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
    frmTaxMsg.Label1.Caption = "File - TaxRealHistDet.xls, created in the Citipak Directory."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmTaxMsg.Label1.Caption = "File - TaxRealHistDet.txt, created in the Citipak Directory."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
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
        oEXL.FileName = outfile & "TaxRealHistDet.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TaxRealHistDet.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmTaxLoadReport
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
End Sub

