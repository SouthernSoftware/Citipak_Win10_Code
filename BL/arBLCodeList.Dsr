VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arBLCodeList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Category Code List"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   ControlBox      =   0   'False
   Icon            =   "arBLCodeList.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arBLCodeList.dsx":08CA
End
Attribute VB_Name = "arBLCodeList"
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
  Open StartPath & "\BLRPTS\MNCODLST.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldCatCode") '1)
  Fields.Add ("fldDesc") '2)
  Fields.Add ("fldRevGLNum") '3)
  Fields.Add ("fldGLAcct") '4)
  Fields.Add ("fldCashAcct") '5)
  Fields.Add ("fldFee") '6)
  Fields.Add ("fldBAmt1") '7)
  Fields.Add ("fldRecpt1") '8)
  Fields.Add ("fldPct1") '9)
  Fields.Add ("fldMax1") '10)
  Fields.Add ("fldBAmt2") '11)
  Fields.Add ("fldRecpt2") '12)
  Fields.Add ("fldPct2") '13)
  Fields.Add ("fldMax2") '14)
  Fields.Add ("fldBAmt3") '15)
  Fields.Add ("fldRecpt3") '16)
  Fields.Add ("fldPct3") '17)
  Fields.Add ("fldMax3") '18)
  Fields.Add ("fldBAmt4") '19)
  Fields.Add ("fldRecpt4") '20)
  Fields.Add ("fldPct4") '21)
  Fields.Add ("fldMax4") '22)
  Fields.Add ("fldBAmt5") '23)
  Fields.Add ("fldRecpt5") '24)
  Fields.Add ("fldPct5") '25)
  Fields.Add ("fldMax5") '26)
  Fields.Add ("fldBAmt6") '27)
  Fields.Add ("fldRecpt6") '28)
  Fields.Add ("fldPct6") '29)
  Fields.Add ("fldMax6") '30)
  Fields.Add ("fldType") '31)
  
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
  Fields("fldCatCode").Value = arr(1)
  Fields("fldDesc").Value = arr(2)
  Fields("fldRevGLNum").Value = arr(3)
  Fields("fldGLAcct").Value = arr(4)
  Fields("fldCashAcct").Value = arr(5)
  Fields("fldFee").Value = arr(6)
'  Fields("fldCType").Value = arr(7)
  Fields("fldBAmt1").Value = arr(7)
  Fields("fldRecpt1").Value = arr(8)
  If QPTrim$(arr(9)) <> "" Then
    Fields("fldPct1").Value = arr(9) / 100
  Else
    Fields("fldPct1").Value = arr(9)
  End If
  Fields("fldMax1").Value = arr(10)
  Fields("fldBAmt2").Value = arr(11)
  Fields("fldRecpt2").Value = arr(12)
  If QPTrim$(arr(14)) <> "" Then
    Fields("fldPct2").Value = arr(13) / 100
  Else
    Fields("fldPct2").Value = arr(13)
  End If
  Fields("fldMax2").Value = arr(14)
  Fields("fldBAmt3").Value = arr(15)
  Fields("fldRecpt3").Value = arr(16)
  If QPTrim$(arr(18)) <> "" Then
    Fields("fldPct3").Value = arr(17) / 100
  Else
    Fields("fldPct3").Value = arr(17)
  End If
  Fields("fldMax3").Value = arr(18)
  Fields("fldBAmt4").Value = arr(19)
  Fields("fldRecpt4").Value = arr(20)
  If QPTrim$(arr(22)) <> "" Then
    Fields("fldPct4").Value = arr(21) / 100
  Else
    Fields("fldPct4").Value = arr(21)
  End If
  Fields("fldMax4").Value = arr(22)
  Fields("fldBAmt5").Value = arr(23)
  Fields("fldRecpt5").Value = arr(24)
  If QPTrim$(arr(25)) <> "" Then
    Fields("fldPct5").Value = arr(25) / 100
  Else
    Fields("fldPct5").Value = arr(25)
  End If
  Fields("fldMax5").Value = arr(26)
  Fields("fldBAmt6").Value = arr(27)
  Fields("fldRecpt6").Value = arr(28)
  If QPTrim$(arr(29)) <> "" Then
    Fields("fldPct6").Value = arr(29) / 100
  Else
    Fields("fldPct6").Value = arr(29)
  End If
  Fields("fldMax6").Value = arr(30)
  Fields("fldType").Value = arr(31)
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
      frmBLMessageBoxJr.Label1.Caption = "File - BLCatListRpt.xls, created in the Citipak Directory."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmBLMessageBoxJr.Label1.Caption = "File - BLCatListRpt.txt, created in the Citipak Directory."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
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
    frmBLMessageBoxJr.Label1.Caption = "File - BLCatListRpt.xls, created in the Citipak Directory."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmBLMessageBoxJr.Label1.Caption = "File - BLCatListRpt.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "BLCatListRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "BLCatListRpt.txt"
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
  Me.fldTimeDate.Text = Date
  Me.Zoom = -1
End Sub

Private Sub Detail_Format()
  If Fields("fldBAmt1").Value <> "" Then
    Detail.Height = 2616
    Label24.Visible = True
    Label25.Visible = True
    Label26.Visible = True
    Label27.Visible = True
    Line2.Visible = True
    Line3.Visible = False
  Else
    Detail.Height = 372
    Label24.Visible = False
    Label25.Visible = False
    Label26.Visible = False
    Label27.Visible = False
    Line2.Visible = False
    Line3.Visible = True
  End If
End Sub

Private Sub ReportFooter_BeforePrint()
  Line2.LineStyle = ddLSSolid

End Sub

