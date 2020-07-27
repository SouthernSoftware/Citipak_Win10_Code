VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arTaxDiscountRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Discount Analysis Report"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arTaxDiscountRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arTaxDiscountRpt.dsx":08CA
End
Attribute VB_Name = "arTaxDiscountRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private hFile As Integer
Private Temp_Class As Resize_Class
Dim Completed As Boolean

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\TAXRPTS\DISCOUNT.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldCustAcct") '1)
  Fields.Add ("fldCustName") '2)
  Fields.Add ("fldSource") '3)
  Fields.Add ("fldThisBothDsc") '4)
  Fields.Add ("fldTotBothDsc") '5)
  Fields.Add ("fldThisLoss") '6)
  Fields.Add ("fldTotLoss") '7)
  Fields.Add ("fldThisSnrDsc") '8)
  Fields.Add ("fldTotSnrDsc") '9)
  Fields.Add ("fldThisSnrLoss") '10)
  Fields.Add ("fldTotSnrLoss") '11)
  Fields.Add ("fldThisOthDsc") '12)
  Fields.Add ("fldTotOthDsc") '13)
  Fields.Add ("fldThisOthLoss") '14)
  Fields.Add ("fldTotOthLoss") '15)
  Fields.Add ("fldCustCnt") '16)
  Fields.Add ("fldThisTownship") '17)
  Fields.Add ("fldCustTownship") '18)
  Fields.Add ("fldAddress") '19)
  Fields.Add ("fldTaxRate") '20)
  Fields.Add ("fldGOpt") '21)
  Fields.Add ("fldOptDesc") '22)
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
  Fields("fldCustAcct").Value = arr(1)
  Fields("fldCustName").Value = arr(2)
  Fields("fldSource").Value = arr(3)
  Fields("fldThisBothDsc").Value = arr(4)
  Fields("fldTotBothDsc").Value = arr(5)
  Fields("fldThisLoss").Value = arr(6)
  Fields("fldTotLoss").Value = arr(7)
  Fields("fldThisSnrDsc").Value = arr(8)
  Fields("fldTotSnrDsc").Value = arr(9)
  Fields("fldThisSnrLoss").Value = arr(10)
  Fields("fldTotSnrLoss").Value = arr(11)
  Fields("fldThisOthDsc").Value = arr(12)
  Fields("fldTotOthDsc").Value = arr(13)
  Fields("fldThisOthLoss").Value = arr(14)
  Fields("fldTotOthLoss").Value = arr(15)
  Fields("fldCustCnt").Value = arr(16)
  Fields("fldThisTownship").Value = arr(17)
  Fields("fldCustTownship").Value = arr(18)
  Fields("fldAddress").Value = arr(19)
  Fields("fldTaxRate").Value = arr(20)
  Fields("fldGOpt").Value = arr(21) + ":"
  Fields("fldOptDesc").Value = arr(22)
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
      frmTaxMsg.Label1.Caption = "File - TaxDiscountRpt.xls, created in the Citipak Directory."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmTaxMsg.Label1.Caption = "File - TaxDiscountRpt.txt, created in the Citipak Directory."
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
    frmTaxMsg.Label1.Caption = "File - TaxDiscountRpt.xls, created in the Citipak Directory."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmTaxMsg.Label1.Caption = "File - TaxDiscountRpt.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "TaxDiscountRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TaxDiscountRpt.txt"
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
  If QPTrim$(Fields("fldOptDesc").Value) <> "" Then
    Detail.Height = 810
    Line9.Y1 = 810
    Line9.Y2 = 810
    Field60.Visible = True
    Field61.Visible = True
    Field46.Top = 540
    Field47.Top = 540
    Field48.Top = 540
    Field49.Top = 540
    Field50.Top = 540
    Field51.Top = 540
  Else
    Detail.Height = 540
    Line9.Y1 = 540
    Line9.Y2 = 540
    Field60.Visible = False
    Field61.Visible = False
    Field46.Top = 270
    Field47.Top = 270
    Field48.Top = 270
    Field49.Top = 270
    Field50.Top = 270
    Field51.Top = 270
  End If

End Sub
