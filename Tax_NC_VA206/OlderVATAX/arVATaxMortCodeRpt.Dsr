VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxMortCodeRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Mortgage Code Report"
   ClientHeight    =   8736
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arVATaxMortCodeRpt.dsx":0000
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arVATaxMortCodeRpt.dsx":08CA
End
Attribute VB_Name = "arVATaxMortCodeRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private hFile As Integer
  'Private Temp_Class As Resize_Class
Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\TAXRPTS\MORTCODE.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldCustName") '1)
  Fields.Add ("fldCustAcct") '2)
  Fields.Add ("fldAdd1") '3)
  Fields.Add ("fldAdd2") '4)
  Fields.Add ("fldGMortCode") '5)
  Fields.Add ("fldBankName") '6)
  Fields.Add ("fldCSZ") '7)
  Fields.Add ("fldMortCode") '8)
  Fields.Add ("fldPropPin") '9)
  Fields.Add ("fldPropVal") '10)
  Fields.Add ("fldTownship") '11)
  Fields.Add ("fldTotCnt") '12)
  Fields.Add ("fldGOpt") '13)
  Fields.Add ("fldOptDesc") '14)
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
  Fields("fldCustName").Value = arr(1)
  Fields("fldCustAcct").Value = arr(2)
  Fields("fldAdd1").Value = arr(3)
  Fields("fldAdd2").Value = arr(4)
  Fields("fldGMortCode").Value = arr(5)
  Fields("fldBankName").Value = arr(6)
  Fields("fldCSZ").Value = arr(7)
  Fields("fldMortCode").Value = arr(8)
  Fields("fldPropPin").Value = arr(9)
  Fields("fldPropVal").Value = arr(10)
  Fields("fldTownship").Value = arr(11)
  Fields("fldTotCnt").Value = arr(12)
  Fields("fldGOpt").Value = arr(13) + ":"
  Fields("fldOptDesc").Value = arr(14)
  If QPTrim$(arr(14)) <> "" Then
    Field53.Visible = True
    Field54.Visible = True
  Else
    Field53.Visible = False
    Field54.Visible = False
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
      frmVATaxMsg.Label1.Caption = "File - MortCodeRpt.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - MortCodeRpt.txt, created in the Citipak Directory."
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
    frmVATaxMsg.Label1.Caption = "File - MortCodeRpt.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - MortCodeRpt.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "MortCodeRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "MortCodeRpt.txt"
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
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    'Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Date
  Me.Zoom = -1
End Sub

Private Sub ReportFooter_Format()
  Set SubReport1.object = New arVATaxSubMortCode

End Sub
