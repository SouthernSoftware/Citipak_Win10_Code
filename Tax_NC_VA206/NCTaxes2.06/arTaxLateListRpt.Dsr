VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arTaxLateListRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Late List Report"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arTaxLateListRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arTaxLateListRpt.dsx":08CA
End
Attribute VB_Name = "arTaxLateListRpt"
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
  Open StartPath & "\TAXRPTS\LATELIST.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldCustAcct") '1)
  Fields.Add ("fldCustName") '2)
  Fields.Add ("fldPIN") '3)
  Fields.Add ("fldPropVal") '4)
  Fields.Add ("fldPropType") '5)
  Fields.Add ("fldTotLLCnt") '6)
  Fields.Add ("fldTotRealLLCnt") '7)
  Fields.Add ("fldTotPersLLCnt") '8)
  Fields.Add ("fldTotVal") '9)
  Fields.Add ("fldTotRealVal") '10)
  Fields.Add ("fldTotPersVal") '11)
  Fields.Add ("fldCustTownship") '12)
  Fields.Add ("fldTownship") '13)
  Fields.Add ("fldCustCnt") '14)
  Fields.Add ("fldGOpt") '15)
  Fields.Add ("fldOptDesc") '16)
  Fields.Add ("fldDesc") '17)
  Fields.Add ("fldRptOpt") '18)
End Sub
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmTaxLoadReport
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
  Fields("fldCustAcct").Value = arr(1)
  Fields("fldCustName").Value = arr(2)
  Fields("fldPIN").Value = arr(3)
  Fields("fldPropVal").Value = arr(4)
  Fields("fldPropType").Value = arr(5)
  Fields("fldTotLLCnt").Value = arr(6)
  Fields("fldTotRealLLCnt").Value = arr(7)
  Fields("fldTotPersLLCnt").Value = arr(8)
  Fields("fldTotVal").Value = arr(9)
  Fields("fldTotRealVal").Value = arr(10)
  Fields("fldTotPersVal").Value = arr(11)
  Fields("fldCustTownship").Value = arr(12)
  Fields("fldTownship").Value = arr(13)
  Fields("fldCustCnt").Value = arr(14)
  Fields("fldGOpt").Value = arr(15) + ":"
  Fields("fldOptDesc").Value = arr(16)
  Fields("fldDesc").Value = arr(17)
  Fields("fldRptOpt").Value = arr(18)
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
      frmTaxMsg.Label1.Caption = "File - LateListRpt.xls, created in the Citipak Directory."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmTaxMsg.Label1.Caption = "File - LateListRpt.txt, created in the Citipak Directory."
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
    frmTaxMsg.Label1.Caption = "File - LateListRpt.xls, created in the Citipak Directory."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmTaxMsg.Label1.Caption = "File - LateListRpt.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "LateListRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "LateListRpt.txt"
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


Private Sub GroupHeader1_Format()
  If Fields("fldOptDesc").Value <> "" Then
    GroupHeader1.Height = 540
    Line11.Y1 = 540
    Line11.Y2 = 540
    Field54.Visible = True
    Field55.Visible = True
  Else
    GroupHeader1.Height = 270
    Line11.Y1 = 270
    Line11.Y2 = 270
    Field54.Visible = False
    Field55.Visible = False
  End If

End Sub

Private Sub PageHeader_Format()
  If QPTrim$(Fields("fldRptOpt").Value) = "1" Then
    Label76.Caption = "Property Address"
  Else
    Label76.Caption = "First Line of Notes"
  End If
End Sub

