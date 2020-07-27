VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arTCResultsRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Results Report"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "arTCResultsRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15452
   SectionData     =   "arTCResultsRpt.dsx":08CA
End
Attribute VB_Name = "arTCResultsRpt"
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
  Open App.Path & "\TCRPTS\CSTLSTSM.RPT" For Input As #hFile
  Fields.Add ("fldPinNum") '0)
  Fields.Add ("fldCustName") '1)
  Fields.Add ("fldCntyAcctS") '2)
  Fields.Add ("fldCntyAcctN") '3)
  Fields.Add ("fldRealVal") '4)
  Fields.Add ("fldRXOth") '5)
  Fields.Add ("fldRXSnr") '6)
  Fields.Add ("fldPersVal") '7)
  Fields.Add ("fldMTVal") '8)
  Fields.Add ("fldMCVal") '9)
  Fields.Add ("fldFarmVal") '10)
  Fields.Add ("fldMHVal") '11)
  Fields.Add ("fldPXOth") '12)
  Fields.Add ("fldPXSnr") '13)
  Fields.Add ("fldTRealVal") '14)
  Fields.Add ("fldTRXOth") '15)
  Fields.Add ("fldTRXSnr") '16)
  Fields.Add ("fldTPersVal") '17)
  Fields.Add ("fldTMCVal") '18)
  Fields.Add ("fldTMHVal") '19)
  Fields.Add ("fldTMTVal") '20)
  Fields.Add ("fldTFarmVal") '21)
  Fields.Add ("fldTPXOth") '22)
  Fields.Add ("fldTPXSnr") '23)
  Fields.Add ("fldGTPersVal") '24)
  Fields.Add ("fldRecCnt") '25)
  Fields.Add ("fldTBldgVal") '26)
  Fields.Add ("fldPPinNum") '27)
  Fields.Add ("fldBldgVal") '28
  Fields.Add ("fldCountyPin")
End Sub
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    frmTCMsg.Label1.Caption = "Error Number: " & Str(Number) & " " & Description
    frmTCMsg.Label1.Top = 900
    frmTCMsg.Show vbModal
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
  If QPTrim$(arr(0)) <> "" And QPTrim$(arr(0)) <> "0" Then
    Fields("fldPinNum").Value = arr(0)
  Else
    Fields("fldPinNum").Value = arr(27)
  End If
  Fields("fldCustName").Value = arr(1)
  Fields("fldCntyAcctS").Value = arr(2)
  Fields("fldCntyAcctN").Value = arr(3)
  Fields("fldRealVal").Value = arr(4)
  Fields("fldRXOth").Value = arr(5)
  Fields("fldRXSnr").Value = arr(6)
  Fields("fldPersVal").Value = arr(7)
  Fields("fldMTVal").Value = arr(8)
  Fields("fldMCVal").Value = arr(9)
  Fields("fldFarmVal").Value = arr(10)
  Fields("fldMHVal").Value = arr(11)
  Fields("fldPXOth").Value = arr(12)
  Fields("fldPXSnr").Value = arr(13)
  Fields("fldTRealVal").Value = arr(14)
  Fields("fldTRXOth").Value = arr(15)
  Fields("fldTRXSnr").Value = arr(16)
  Fields("fldTPersVal").Value = arr(17)
  Fields("fldTMCVal").Value = arr(18)
  Fields("fldTMHVal").Value = arr(19)
  Fields("fldTMTVal").Value = arr(20)
  Fields("fldTFarmVal").Value = arr(21)
  Fields("fldTPXOth").Value = arr(22)
  Fields("fldTPXSnr").Value = arr(23)
  Fields("fldGTPersVal").Value = arr(24)
  Fields("fldRecCnt").Value = arr(25)
  Fields("fldTBldgVal").Value = arr(26)
  Fields("fldPPinNum").Value = arr(27)
  Fields("fldBldgVal").Value = arr(28)
  If QPTrim$(arr(2)) <> "" Then
    Fields("fldCountyPin") = arr(2)
  Else
    Fields("fldCountyPin") = arr(3)
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
      frmTCMsg.Label1.Caption = "File - ResultsRpt.xls, created in the Citipak Directory."
      frmTCMsg.Label1.Top = 900
      frmTCMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmTCMsg.Label1.Caption = "File - ResultsRpt.txt, created in the Citipak Directory."
      frmTCMsg.Label1.Top = 900
      frmTCMsg.Show vbModal
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
    frmTCMsg.Label1.Caption = "File - ResultsRpt.xls, created in the Citipak Directory."
    frmTCMsg.Label1.Top = 900
    frmTCMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmTCMsg.Label1.Caption = "File - ResultsRpt.txt, created in the Citipak Directory."
    frmTCMsg.Label1.Top = 900
    frmTCMsg.Show vbModal
  End If
End Sub
Private Sub ExportReport(x As Integer)
  Dim oEXL As ActiveReportsExcelExport.ARExportExcel
  Dim oTXT As ActiveReportsTextExport.ARExportText
  Dim outfile As String
  If Right$(StartPath, 1) = ":" Then
    outfile = App.Path
  Else
    outfile = App.Path & "\"
  End If
  Select Case x
    Case 1   '"Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & "ResultsRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "ResultsRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_ReportEnd()
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


Private Sub GroupFooter1_Format()
  Set SubReport1 = New arTCErrorsRpt

End Sub
