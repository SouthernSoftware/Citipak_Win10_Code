VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arBLAdjRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Balance Adjustment Report"
   ClientHeight    =   8868
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11640
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20505
   _ExtentY        =   15637
   SectionData     =   "arBLAdjRpt.dsx":0000
End
Attribute VB_Name = "arBLAdjRpt"
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
  Open StartPath & "\BLRPTS\ARADJRPT.RPT" For Input As #hFile
  Fields.Add ("fld0") '0)
  Fields.Add ("fld1") '1)
  Fields.Add ("fld2") '2)
  Fields.Add ("fld3") '3)
  Fields.Add ("fld4") '4)
  Fields.Add ("fld5") '5)
  Fields.Add ("fld6") '6)
  Fields.Add ("fld7") '7)
  Fields.Add ("fld8") '8)
  Fields.Add ("fld9") '9)
  Fields.Add ("fld10") '10)
  Fields.Add ("fld11") '11)
  Fields.Add ("fld12") '12)
  Fields.Add ("fld13") '13)
  Fields.Add ("fld14") '14)
  Fields.Add ("fld15") '15)
  Fields.Add ("fld16") '16)
  Fields.Add ("fld17") '17)
  Fields.Add ("fld18") '18)
  Fields.Add ("fld19") '19)
  Fields.Add ("fld20") '20)
  Fields.Add ("fld21") '21)
  Fields.Add ("fld22") '22)
  Fields.Add ("fld23") '23)
  Fields.Add ("fld24") '24)
  Fields.Add ("fld25") '25)
  Fields.Add ("fld26") '26)
  Fields.Add ("fld27") '27)
  Fields.Add ("fld28") '28)
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
  Fields("fld0").Value = arr(0)
  Fields("fld1").Value = arr(1)
  Fields("fld2").Value = arr(2)
  Fields("fld3").Value = arr(3)
  Fields("fld4").Value = arr(4)
  Fields("fld5").Value = arr(5)
  Fields("fld6").Value = arr(6)
  Fields("fld7").Value = arr(7)
  Fields("fld8").Value = arr(8)
  Fields("fld9").Value = arr(9)
  Fields("fld10").Value = arr(10)
  Fields("fld11").Value = arr(11)
  Fields("fld12").Value = arr(12)
  Fields("fld13").Value = arr(13)
  Fields("fld14").Value = arr(14)
  Fields("fld15").Value = arr(15)
  Fields("fld16").Value = arr(16)
  Fields("fld17").Value = arr(17)
  Fields("fld18").Value = arr(18)
  Fields("fld19").Value = arr(19)
  Fields("fld20").Value = arr(20)
  Fields("fld21").Value = arr(21)
  Fields("fld22").Value = arr(22)
  Fields("fld23").Value = arr(23)
  Fields("fld24").Value = arr(24)
  Fields("fld25").Value = arr(25)
  Fields("fld26").Value = arr(26)
  Fields("fld27").Value = arr(27)
  Fields("fld28").Value = arr(28)
  If QPTrim$(arr(11)) = "" Then 'cat 1
    Field9.Visible = False
    Field18.Visible = False
    Field19.Visible = False
  Else
    Field9.Visible = True
    Field18.Visible = True
    Field19.Visible = True
  End If
  If QPTrim$(arr(14)) = "" Then 'cat 2
    Field10.Visible = False
    Field20.Visible = False
    Field21.Visible = False
  Else
    Field10.Visible = True
    Field20.Visible = True
    Field21.Visible = True
  End If
  If QPTrim$(arr(17)) = "" Then 'cat 3
    Field11.Visible = False
    Field22.Visible = False
    Field23.Visible = False
  Else
    Field11.Visible = True
    Field22.Visible = True
    Field23.Visible = True
  End If
  If QPTrim$(arr(20)) = "" Then 'cat 4
    Field12.Visible = False
    Field24.Visible = False
    Field25.Visible = False
  Else
    Field12.Visible = True
    Field24.Visible = True
    Field25.Visible = True
  End If
  If QPTrim$(arr(23)) = "" Then 'cat 5
    Field13.Visible = False
    Field26.Visible = False
    Field27.Visible = False
  Else
    Field13.Visible = True
    Field26.Visible = True
    Field27.Visible = True
  End If
  If Val(arr(7)) = 0 And Val(arr(8)) = 0 Then 'penalty
    Label26.Visible = False
    Field16.Visible = False
    Field17.Visible = False
  Else
    Label26.Visible = True
    Field16.Visible = True
    Field17.Visible = True
  End If
  If Val(arr(9)) = 0 And Val(arr(10)) = 0 Then 'issuance
    Label25.Visible = False
    Field14.Visible = False
    Field15.Visible = False
  Else
    Label25.Visible = True
    Field14.Visible = True
    Field15.Visible = True
  End If
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
      frmBLMessageBoxJr.Label1.Caption = "File - BLAdjRpt.xls, created in the Citipak Directory."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmBLMessageBoxJr.Label1.Caption = "File - BLAdjRpt.txt, created in the Citipak Directory."
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
    frmBLMessageBoxJr.Label1.Caption = "File - BLAdjRpt.xls, created in the Citipak Directory."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmBLMessageBoxJr.Label1.Caption = "File - BLAdjRpt.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "BLAdjRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "BLAdjRpt.txt"
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
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Unload frmBLLoadReport
  Me.fldTimeDate.Text = Date
  Me.Zoom = -1
End Sub











