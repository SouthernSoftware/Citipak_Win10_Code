VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxAdjRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjustment Report"
   ClientHeight    =   9744
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arVATaxAdjRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   17198
   SectionData     =   "arVATaxAdjRpt.dsx":08CA
End
Attribute VB_Name = "arVATaxAdjRpt"
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
  Open UBPath & "\TAXRPTS\TXADJRPT.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldTransAmt") '1)
  Fields.Add ("fldName") '2)
  Fields.Add ("fldGCustNum") '3)
  Fields.Add ("fldAddress") '4)
  Fields.Add ("fldPrincHead") '5)
  Fields.Add ("fldIntHead") '6)
  Fields.Add ("fldAdvColHead") '7)
  Fields.Add ("fldLateListHead") '8)
  Fields.Add ("fldOpt1Head") '9)
  Fields.Add ("fldOpt2Head") '10)
  Fields.Add ("fldOpt3Head") '11)
  Fields.Add ("fldPrincAdj") '12)
  Fields.Add ("fldIntAdj") '13)
  Fields.Add ("fldAdvColAdj") '14)
  Fields.Add ("fldLateListAdj") '15)
  Fields.Add ("fldOpt1Adj") '16)
  Fields.Add ("fldOpt2Adj") '17)
  Fields.Add ("fldOpt3Adj") '18)
  Fields.Add ("fldOldBal") '19)
  Fields.Add ("fldTotAdj") '20)
  Fields.Add ("fldNewBal") '21)
  Fields.Add ("fldType") '22)
  Fields.Add ("fldPrincBal") '23)
  Fields.Add ("fldIntBal") '24)
  Fields.Add ("fldAdvColBal") '25)
  Fields.Add ("fldLateListBal") '26)
  Fields.Add ("fldOpt1Bal") '27)
  Fields.Add ("fldOpt2Bal") '28)
  Fields.Add ("fldOpt3Bal") '29)
  Fields.Add ("fldTotOldBal") '30)
  Fields.Add ("fldBillNum") '31)
  Fields.Add ("fldTotBal") '32)
  Fields.Add ("fldPrepay") '33)
  Fields.Add ("fldPenHead") '34)
  Fields.Add ("fldPenAdj") '35)
  Fields.Add ("fldPenBal") '36)
  Fields.Add ("fldNotes") '37)
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
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("fldTown").Value = arr(0)
  Fields("fldTransAmt").Value = arr(1)
  Fields("fldName").Value = arr(2)
  Fields("fldGCustNum").Value = arr(3)
  Fields("fldAddress").Value = arr(4)
  Fields("fldPrincHead").Value = arr(5)
  Fields("fldIntHead").Value = arr(6)
  Fields("fldAdvColHead").Value = arr(7)
  Fields("fldLateListHead").Value = arr(8)
  Fields("fldOpt1Head").Value = arr(9)
  Fields("fldOpt2Head").Value = arr(10)
  Fields("fldOpt3Head").Value = arr(11)
  Fields("fldPrincAdj").Value = arr(12)
  Fields("fldIntAdj").Value = arr(13)
  Fields("fldAdvColAdj").Value = arr(14)
  Fields("fldLateListAdj").Value = arr(15)
  Fields("fldOpt1Adj").Value = arr(16)
  Fields("fldOpt2Adj").Value = arr(17)
  Fields("fldOpt3Adj").Value = arr(18)
  Fields("fldOldBal").Value = arr(19)
  Fields("fldTotAdj").Value = arr(20)
  Fields("fldNewBal").Value = arr(21)
  Fields("fldType").Value = arr(22)
  Fields("fldPrincBal").Value = arr(23)
  Fields("fldIntBal").Value = arr(24)
  Fields("fldAdvColBal").Value = arr(25)
  Fields("fldLateListBal").Value = arr(26)
  Fields("fldOpt1Bal").Value = arr(27)
  Fields("fldOpt2Bal").Value = arr(28)
  Fields("fldOpt3Bal").Value = arr(29)
  Fields("fldTotOldBal").Value = arr(30)
  Fields("fldBillNum").Value = arr(31)
  Fields("fldTotBal").Value = arr(32)
  Fields("fldPrepay").Value = arr(33)
  Fields("fldPenHead").Value = arr(34)
  Fields("fldPenAdj").Value = arr(35)
  Fields("fldPenBal").Value = arr(36)
  Fields("fldNotes").Value = arr(37)
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
      frmVATaxMsg.Label1.Caption = "File - TaxAdjRpt.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - TaxAdjRpt.txt, created in the Citipak Directory."
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
    frmVATaxMsg.Label1.Caption = "File - TaxAdjRpt.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - TaxAdjRpt.txt, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
End Sub
Private Sub ExportReport(x As Integer)
  Dim oEXL As ActiveReportsExcelExport.ARExportExcel
  Dim oTXT As ActiveReportsTextExport.ARExportText
  Dim outfile As String
  If Right$(UBPath, 1) = ":" Then
    outfile = UBPath
  Else
    outfile = UBPath & "\"
  End If
  Select Case x
    Case 1   '"Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & "TaxAdjRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TaxAdjRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmLoadingRpt
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
   ''' Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Date
  Me.Zoom = -1
End Sub



