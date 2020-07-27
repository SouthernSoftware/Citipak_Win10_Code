VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arBLCustListRpt 
   Caption         =   "ActiveReport1"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   ControlBox      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20505
   _ExtentY        =   11271
   SectionData     =   "arBLCustListRpt.dsx":0000
End
Attribute VB_Name = "arBLCustListRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsBLTextBoxOverrider
Private Temp_Class As Resize_Class
Private hFile As Integer

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\BLRPTS\ARDetCus.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldCustNum") '1)
  Fields.Add ("fldCActive") '2)
  Fields.Add ("fldProrate") '3)
  Fields.Add ("fldLicense") '4)
  Fields.Add ("fldValidThru") '5)
  Fields.Add ("fldBillName") '6)
  Fields.Add ("fldCategories") '7)
  Fields.Add ("fldAdd1") '8)
  Fields.Add ("fldWrkPhn") '9)
  Fields.Add ("fldAdd2") '10)
  Fields.Add ("fldCity") '11)
  Fields.Add ("fldState") '12)
  Fields.Add ("fldZip") '13)
  Fields.Add ("fldCustFee") '14)
  Fields.Add ("fldCatDesc") '15)
  Fields.Add ("fldIssFee") '16)
'  Fields.Add ("fldDone") ' 17)
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmBLLoadReport
    frmBLMessageBoxJr.Label1.Caption = "Error Number: " & Str(Number) & " " & Description
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
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
  Fields("fldCustNum").Value = arr(1)
  Fields("fldCActive").Value = arr(2)
  Fields("fldProrate").Value = arr(3)
  Fields("fldLicense").Value = arr(4)
  Fields("fldValidThru").Value = arr(5)
  Fields("fldBillName").Value = arr(6)
  Fields("fldCategories").Value = arr(7)
  Fields("fldAdd1").Value = arr(8)
  Fields("fldWrkPhn").Value = arr(9)
  Fields("fldAdd2").Value = arr(10)
  Fields("fldCity").Value = arr(11) + ", " + arr(12) + " " + arr(13)
  If QPTrim$(arr(14)) <> "" Then
    Fields("fldCustFee").Value = arr(14)
    Label24.Visible = True
  Else
    Label24.Visible = False
  End If
  Fields("fldCatDesc").Value = arr(15)
  Fields("fldIssFee").Value = arr(16)
  If Val(arr(16)) > 0 Then
    Label26.Visible = True
    Label26.Caption = "All fees include a " + QPTrim$(Using$("$#,##0.00", CDbl(arr(16)))) + " issuance fee."
  Else
    Label26.Visible = False
  End If
'  Fields("fldDone").Value = arr(17)
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
      frmBLMessageBoxJr.Label1.Caption = "File - BLCustListRpt.xls, created in the Citipak Directory."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmBLMessageBoxJr.Label1.Caption = "File - BLCustListRpt.txt, created in the Citipak Directory."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  If Field4.DataValue = "ALL" Then
    Unload arBLSubCustList
  End If
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
    frmBLMessageBoxJr.Label1.Caption = "File - BLCustListRpt.xls, created in the Citipak Directory."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmBLMessageBoxJr.Label1.Caption = "File - BLCustListRpt.txt, created in the Citipak Directory."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
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
        oEXL.FileName = outfile & "BLCustListRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "BLCustListRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmBLLoadReport
  If hFile <> 0 Then
    Close #hFile
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Label27.Visible = False
'  Label28.Visible = False
'  Field5.Visible = False
  Me.fldTimeDate.Text = Date
  Me.Zoom = -1
End Sub

Private Sub Detail_Format()
'  If Fields("fldDone").Value = "DONE" Then
'    Label28.Visible = True
'    Field5.Visible = True
'    Detail.Height = 1485
'  Else
'    Detail.Height = 1125
'  End If

End Sub

Private Sub ReportFooter_Format()
'  If Field4.DataValue = "ALL" Then
'    Label27.Visible = True
'    Set SubReport1.object = New arBLSubCustList
'  End If

End Sub
