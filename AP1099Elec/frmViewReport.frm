VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form frmViewReport 
   BackColor       =   &H008A775B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Report"
   ClientHeight    =   8916
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   12216
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8916
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewer21 
      Height          =   372
      Left            =   -360
      TabIndex        =   0
      Top             =   -24
      Width           =   12132
      _ExtentX        =   21400
      _ExtentY        =   656
      SectionData     =   "frmViewReport.frx":0000
   End
End
Attribute VB_Name = "frmViewReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim rpt As ActiveReport

Public ReportNM As String
Private hFile As Integer
'Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
'  KillFile ReportFile$
'End Sub



Private Sub ARViewer21_ToolbarClick(ByVal Tool As DDActiveReportsViewer2Ctl.DDTool)
  If Tool = "&Close" Then
    Unload Me
  End If
  If Tool = "&Save Report" Then
  End If

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      
      Unload Me
      KeyCode = 0
    Case Else:
  End Select
End Sub
Public Sub sendrptname(fromrpt As ActiveReport)
  rpt = fromrpt
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  ARViewer21.Toolbar.Tools.Add "&Close"
  ARViewer21.Toolbar.Tools.Add "&Save Report"
  'MakeWindowTopMost hwnd, True
 
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Sub ExportReport(x As Integer)
  Dim oEXL As ActiveReportsExcelExport.ARExportExcel
  Dim oTXT As ActiveReportsTextExport.ARExportText
  Select Case x
    Case 1   '"Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = App.Path & "\EXLExport.xls"
        oEXL.Export rpt.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = App.Path & "\TXTExport.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export rpt.Pages
  End Select
End Sub


