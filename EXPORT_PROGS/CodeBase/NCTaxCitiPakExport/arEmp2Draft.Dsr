VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arEmp2Draft 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employees To Draft"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   Icon            =   "arEmp2Draft.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20553
   _ExtentY        =   15642
   SectionData     =   "arEmp2Draft.dsx":08CA
End
Attribute VB_Name = "arEmp2Draft"
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
      MsgBox "File - Emp2Draft.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - Emp2Draft.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - Emp2Draft.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - Emp2Draft.txt, created in the Citipak Directory.", vbOKOnly
  End If
End Sub
Private Sub ExportReport(X As Integer)
  Dim oEXL As ActiveReportsExcelExport.ARExportExcel
  Dim oTXT As ActiveReportsTextExport.ARExportText
  Dim outfile As String
  If Right$(StartPath, 1) = ":" Then
    outfile = StartPath
  Else
    outfile = StartPath & "\"
  End If
  Select Case X
    Case 1   '"Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & "Emp2Draft.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "Emp2Draft.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  Dim X As Integer
  HFile = FreeFile
  Open StartPath & "\PRRPTS\PPDFG.RPT" For Input As #HFile
  Fields.Add "fldEmployer" '(0)
  Fields.Add "fldDate" '(1)
  Fields.Add "fldEmpNum" '(2)
  Fields.Add "fldEmployee" '(3)
  Fields.Add "fldDraftCode" '(4)
  Fields.Add "fldBankAcctNum" '(5)
  Fields.Add "fldDraftAmt" '(6)
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
  Fields("fldEmployer").Value = arr(0)
  Fields("fldDate").Value = arr(1)
  Fields("fldEmpNum").Value = arr(2)
  Fields("fldEmployee").Value = arr(3)
  Fields("fldDraftCode").Value = arr(4)
  Fields("fldBankAcctNum").Value = arr(5)
  Fields("fldDraftAmt").Value = arr(6)
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
  Me.fldTimeDate.Text = Now
  Me.Zoom = -1

End Sub

Private Sub ReportFooter_Format()
  Label64.Visible = False
  Label95.Visible = False
  Label96.Visible = False
  Label97.Visible = False
  Line1.Visible = False
  Line2.Visible = False
End Sub
