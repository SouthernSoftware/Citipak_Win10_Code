VERSION 5.00
Begin VB.Form frmRptSavOpt 
   BackColor       =   &H008A775B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save As...."
   ClientHeight    =   780
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   1620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   1620
   Begin VB.CommandButton cmdText 
      Caption         =   "&Text"
      Height          =   276
      Left            =   264
      TabIndex        =   1
      Top             =   432
      Width           =   1116
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "&Excel"
      Height          =   276
      Left            =   264
      TabIndex        =   0
      Top             =   72
      Width           =   1116
   End
End
Attribute VB_Name = "frmRptSavOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim rpt As ActiveReport
Private Sub cmdExcel_Click()
  ExportReport 1
  Unload Me
End Sub

Private Sub cmdText_Click()
  ExportReport 2
  Unload Me
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
  MakeWindowTopMost hwnd, True
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


