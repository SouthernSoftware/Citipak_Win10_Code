VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxTransJournal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Transaction Journal"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arVATaxTransJournal.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arVATaxTransJournal.dsx":08CA
End
Attribute VB_Name = "arVATaxTransJournal"
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
  Open StartPath & "\TAXRPTS\TAXJRNL.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldCustName") '1)
  Fields.Add ("fldCustNum") '2)
  Fields.Add ("fldActive") '3)
  Fields.Add ("fldTransDate") '4)
  Fields.Add ("fldBillType") '5)
  Fields.Add ("fldTransType") '6)
  Fields.Add ("fldBegDate") '7)
  Fields.Add ("fldEndDate") '8)
  Fields.Add ("fldTaxYear") '9)
  Fields.Add ("fldAmount") '10)
  Fields.Add ("fldTCnt") '11)
  Fields.Add ("fldTotAmt") '12)
  Fields.Add ("fldPrePdAmt") '13)
  Fields.Add ("fldBillNum") '14)
  Fields.Add ("fldDesc") '15)
  Fields.Add ("fldThisTransType") '16)
  Fields.Add ("fldOperNum") '17)
  Fields.Add ("fldGOpt") '18)
  Fields.Add ("fldOptDesc") '19)
  Fields.Add ("fldThisOperNum") '20)
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
  
  Fields("fldTown").Value = arr(0)
  Fields("fldCustName").Value = arr(1)
  Fields("fldCustNum").Value = arr(2)
  Fields("fldActive").Value = arr(3)
  Fields("fldTransDate").Value = arr(4)
  Fields("fldBillType").Value = arr(5)
  Fields("fldTransType").Value = arr(6)
  Fields("fldBegDate").Value = arr(7)
  Fields("fldEndDate").Value = arr(8)
  Fields("fldTaxYear").Value = arr(9)
  Fields("fldAmount").Value = arr(10)
  Fields("fldTCnt").Value = arr(11)
  Fields("fldTotAmt").Value = arr(12)
  Fields("fldPrePdAmt").Value = arr(13)
  Fields("fldBillNum").Value = arr(14)
  Fields("fldDesc").Value = arr(15)
  Fields("fldThisTransType").Value = arr(16)
  Fields("fldOperNum").Value = arr(17)
  Fields("fldGOpt").Value = arr(18) + ":"
  Fields("fldOptDesc").Value = arr(19)
  Fields("fldThisOperNum").Value = arr(20)
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
      frmVATaxMsg.Label1.Caption = "File - TaxTransJrnlRpt.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - TaxTransJrnlRpt.txt, created in the Citipak Directory."
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
    frmVATaxMsg.Label1.Caption = "File - TaxTransJrnlRpt.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - TaxTransJrnlRpt.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "TaxTransJrnlRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TaxTransJrnlRpt.txt"
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
  Label31.Visible = False
End Sub

Private Sub GroupHeader1_Format()
  If QPTrim$(Fields("fldOptDesc").Value) <> "" Then
    GroupHeader1.Height = 540
    Line2.Y1 = 540
    Line2.Y2 = 540
    Field17.Visible = True
    Field18.Visible = True
  Else
    GroupHeader1.Height = 270
    Line2.Y1 = 270
    Line2.Y2 = 270
    Field17.Visible = False
    Field18.Visible = False
  End If
End Sub

Private Sub ReportFooter_Format()
  If Fields("fldBillType").Value = "Real Only" Then
    Set SubReport1.object = New arVASubTransJrnl
    Set SubReport2.object = New arSub2TransJrnl
  ElseIf Fields("fldBillType").Value = "Personal Only" Then
    Set SubReport1.object = New arVASubPersTransJrnl
    Set SubReport2.object = New arVASub2PersTransJrnl
  End If
  Label31.Visible = True
  Label19.Visible = False
  Label20.Visible = False
  Label24.Visible = False
  Label25.Visible = False
  Label26.Visible = False
  Label27.Visible = False
  Label29.Visible = False
  Label36.Visible = False
End Sub

