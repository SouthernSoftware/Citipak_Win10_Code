VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arEmpMessage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Message"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arEmpMessage.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arEmpMessage.dsx":08CA
End
Attribute VB_Name = "arEmpMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private HFile As Integer

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
      MsgBox "File - EmpMsgRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - EmpMsgRpt.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
  If Tool = "&Close" Then
    Unload Me
  End If
  If Tool = "Save/&Excel" Then
    Screen.MousePointer = vbHourglass
    ExportReport 1
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - EmpMsgRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - EmpMsgRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "EmpMsgRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "EmpMsgRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  HFile = FreeFile
  Open StartPath & "\PRRPTS\EMPMSG.RPT" For Input As #HFile
  Fields.Add "fld1" '0
  Fields.Add "fld2" '1
  Fields.Add "fld3" '2
  Fields.Add "fld4" '3
  Fields.Add "fld5" '4
  Fields.Add "fld6" '5
  Fields.Add "fld7" '6
  Fields.Add "fld8" '7
  Fields.Add "fld9" '8
  Fields.Add "fld10" '9
  Fields.Add "fld11" '10
  Fields.Add "fld12" '11
  Fields.Add "fld13" '12
  Fields.Add "fld14" '13
  Fields.Add "fld15" '14
  Fields.Add "fld16" '15
  Fields.Add "fld17" '16
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmLoadingRpt
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True 'removes the error message
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  
  If VBA.eof(HFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #HFile, sLine
  arr = Split(sLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("fld1").Value = arr(0)
  Fields("fld2").Value = arr(1)
  Fields("fld3").Value = arr(2)
  Fields("fld4").Value = arr(3)
  Fields("fld5").Value = arr(4)
  Fields("fld6").Value = arr(5)
  Fields("fld7").Value = arr(6)
  Fields("fld8").Value = arr(7)
  Fields("fld9").Value = arr(8)
  Fields("fld10").Value = arr(9)
  Fields("fld11").Value = arr(10)
  Fields("fld12").Value = arr(11)
  Fields("fld13").Value = arr(12)
  Fields("fld14").Value = arr(13)
  Fields("fld15").Value = arr(14)
  Fields("fld16").Value = arr(15)
  Fields("fld17").Value = arr(16)
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmLoadingRpt
  If HFile <> 0 Then
    Close #HFile
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
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Me.Zoom = -1
End Sub
