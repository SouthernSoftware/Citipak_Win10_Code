VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arTaxValSum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Valuation Summary"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arTaxValSum.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arTaxValSum.dsx":08CA
End
Attribute VB_Name = "arTaxValSum"
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
  Open StartPath & "\TAXRPTS\TXVALLSTSUM.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldCustAcct") '1)
  Fields.Add ("fldCustName") '2)
  Fields.Add ("fldLtr") '3)
  Fields.Add ("fldRealVal") '4)
  Fields.Add ("fldPersVal") '5)
  Fields.Add ("fldDscntVal") '6)
  Fields.Add ("fldNet") '7)
  Fields.Add ("fldAlphaCnt") '8)
  Fields.Add ("fldAlphaRealVal") '9)
  Fields.Add ("fldAlphaPersVal") '10)
  Fields.Add ("fldAlphaDscntVal") '11)
  Fields.Add ("fldAlphaNet") '12)
  Fields.Add ("fldGTotCnt") '13)
  Fields.Add ("fldGRealVal") '14)
  Fields.Add ("fldGPersVal") '15)
  Fields.Add ("fldGDscntVal") '16)
  Fields.Add ("fldGNet") '17)
  Fields.Add ("fldThisPin") '18)
  Fields.Add ("fldCustCnt") '19)
  Fields.Add ("fldAlphaCustCnt") '20)
  Fields.Add ("fldAdd") '21)
  Fields.Add ("fldInactiveYN") '22)
  Fields.Add ("fldGOpt") '23)
  Fields.Add ("fldOptDesc") '24)
  Fields.Add ("fldPropType") '25)
  Fields.Add ("fldPrintOrder") '26)
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
  Fields("fldCustAcct").Value = arr(1)
  Fields("fldCustName").Value = arr(2)
  Fields("fldLtr").Value = arr(3)
  Fields("fldRealVal").Value = arr(4)
  Fields("fldPersVal").Value = arr(5)
  Fields("fldDscntVal").Value = arr(6)
  Fields("fldNet").Value = arr(7)
  Fields("fldAlphaCnt").Value = arr(8)
  Fields("fldAlphaRealVal").Value = arr(9)
  Fields("fldAlphaPersVal").Value = arr(10)
  Fields("fldAlphaDscntVal").Value = arr(11)
  Fields("fldAlphaNet").Value = arr(12)
  Fields("fldGTotCnt").Value = arr(13)
  Fields("fldGRealVal").Value = arr(14)
  Fields("fldGPersVal").Value = arr(15)
  Fields("fldGDscntVal").Value = arr(16)
  Fields("fldGNet").Value = arr(17)
  Fields("fldThisPin").Value = arr(18)
  Fields("fldCustCnt").Value = arr(19)
  Fields("fldAlphaCustCnt").Value = arr(20)
  If QPTrim$(arr(21)) = "" Then
    Fields("fldAdd").Value = "Use All Addresses"
  Else
    Fields("fldAdd").Value = arr(21)
  End If
  Fields("fldInactiveYN").Value = arr(22)
  If QPTrim$(arr(22)) = "B" Then
    Label43.Caption = "Active And Inactive"
  ElseIf QPTrim$(arr(22)) = "A" Then
    Label43.Caption = "Active Only"
  ElseIf QPTrim$(arr(22)) = "I" Then
    Label43.Caption = "Inactive Only"
  End If
  Fields("fldGOpt").Value = arr(23) + ":"
  Fields("fldOptDesc").Value = arr(24)
  If QPTrim$(arr(24)) <> "" Then
    Field28.Visible = True
    Field29.Visible = True
    Detail.Height = 540
  Else
    Field28.Visible = False
    Field29.Visible = False
    Detail.Height = 270
  End If
  Fields("fldPropType").Value = arr(25)
  Fields("fldPrintOrder").Value = arr(26)
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
      frmTaxMsg.Label1.Caption = "File - TaxValSum.xls, created in the Citipak Directory."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmTaxMsg.Label1.Caption = "File - TaxValSum.txt, created in the Citipak Directory."
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
    frmTaxMsg.Label1.Caption = "File - TaxValSum.xls, created in the Citipak Directory."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmTaxMsg.Label1.Caption = "File - TaxValSum.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "TaxValSum.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TaxValSum.txt"
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

Private Sub GroupFooter1_Format()
  If Fields("fldPrintOrder").Value = "N" Then
    Label32.Visible = False
    Field11.Visible = False
    Label34.Visible = False
    Line5.Visible = False
    Label41.Visible = False
    Field12.Visible = False
    Field13.Visible = False
    Field14.Visible = False
    Field15.Visible = False
    Field16.Visible = False
  End If
End Sub
