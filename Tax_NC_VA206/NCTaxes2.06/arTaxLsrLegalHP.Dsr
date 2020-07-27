VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arTaxLsrLegalHP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laser Legal Tax Bill"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arTaxLsrLegalHP.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15452
   SectionData     =   "arTaxLsrLegalHP.dsx":08CA
End
Attribute VB_Name = "arTaxLsrLegalHP"
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
  Open StartPath & "\TAXRPTS\TXLSRLEGALHP.RPT" For Input As #hFile
  Fields.Add ("fldTaxYear") '0)
  Fields.Add ("fldRcptNum") '1)
  Fields.Add ("fldAcctNum") '2)
  Fields.Add ("fldParcel") '3)
  Fields.Add ("fldPropName") '4)
  Fields.Add ("fldPropDesc") '5)
  Fields.Add ("fldLotAcre") '6)
  Fields.Add ("fldRealVal") '7)
  Fields.Add ("fldPersVal") '8)
  Fields.Add ("fldExemp") '9)
  Fields.Add ("fldNetTaxVal") '10)
  Fields.Add ("fldTaxRate") '11)
  Fields.Add ("fldNetTax") '12)
  Fields.Add ("fldLateList") '13)
  Fields.Add ("fldTotDue") '14)
  Fields.Add ("fldTownName") '15)
  Fields.Add ("fldTownAdd") '16)
  Fields.Add ("fldTownCSZ") '17)
  Fields.Add ("fldCustName") '18)
  Fields.Add ("fldCustAdd1") '19)
  Fields.Add ("fldCustAdd2") '20)
  Fields.Add ("fldCustCSZ") '21)
End Sub
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
'    Unload frmTaxLoadReport
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
  Fields("fldTaxYear").Value = arr(0)
  Fields("fldRcptNum").Value = arr(1)
  Fields("fldAcctNum").Value = arr(2)
  Fields("fldParcel").Value = arr(3)
  Fields("fldPropName").Value = arr(4)
  Fields("fldPropDesc").Value = arr(5)
  Fields("fldLotAcre").Value = arr(6)
  Fields("fldRealVal").Value = arr(7)
  Fields("fldPersVal").Value = arr(8)
  Fields("fldExemp").Value = arr(9)
  Fields("fldNetTaxVal").Value = arr(10)
  Fields("fldTaxRate").Value = arr(11)
  Fields("fldNetTax").Value = arr(12)
  Fields("fldLateList").Value = arr(13)
  Fields("fldTotDue").Value = arr(14)
  Fields("fldTownName").Value = arr(15)
  Fields("fldTownAdd").Value = arr(16)
  Fields("fldTownCSZ").Value = arr(17)
  Fields("fldCustName").Value = arr(18)
  Fields("fldCustAdd1").Value = arr(19)
  Fields("fldCustAdd2").Value = arr(20)
  Fields("fldCustCSZ").Value = arr(21)
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
      frmTaxMsg.Label1.Caption = "File - TaxBillLegal.xls, created in the Citipak Directory."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmTaxMsg.Label1.Caption = "File - TaxBillLegal.txt, created in the Citipak Directory."
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
    frmTaxMsg.Label1.Caption = "File - TaxBillLegal.xls, created in the Citipak Directory."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmTaxMsg.Label1.Caption = "File - TaxBillLegal.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "TaxBillLegal.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TaxBillLegal.txt"
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
  Me.Zoom = -1
End Sub



