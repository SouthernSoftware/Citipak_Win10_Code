VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arW2Report 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "W2 Report"
   ClientHeight    =   8850
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11640
   Icon            =   "arW2Report.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15610
   SectionData     =   "arW2Report.dsx":08CA
End
Attribute VB_Name = "arW2Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private hFile As Integer
Dim Reprint As Boolean
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
      MsgBox "File - W2Report.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - W2Report.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close '5/28/2004
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
    MsgBox "File - W2Report.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - W2Report.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "W2Report.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "W2Report.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\PRRPTS\W2REPORTG.RPT" For Input As #hFile
  Fields.Add "fldEmployer" '(0)
  Fields.Add "fldDate" '(1)
  Fields.Add "fldEmployee" '(2)
  Fields.Add "fldAdvEIC" '(3)
  Fields.Add "fldBox10" '(4)
  Fields.Add "fldBox11" '(5)
  Fields.Add "fldBenefitBox" '(6)
  Fields.Add "fldFedGrs" '(7)
  Fields.Add "fldStateGrs" '(8)
  Fields.Add "fldSocGrs" '(9)
  Fields.Add "fldMedGrs" '(10)
  Fields.Add "fldBox12a" '(11)
  Fields.Add "fldBox12b" '(12)
  Fields.Add "fldFedTax" '(13)
  Fields.Add "fldStateTax" '(14)
  Fields.Add "fldSocTax" '(15)
  Fields.Add "fldMedTax" '(16)
  Fields.Add "fldBox14a" '(17)
  Fields.Add "fldBox14b" '(18)
  Fields.Add "fldBox12c" '(19)
  Fields.Add "fldBox12d" '(20)
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
'    Unload frmLoadingRpt
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
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
  Fields("fldEmployer").Value = arr(0)
  Fields("fldDate").Value = arr(1)
  Fields("fldEmployee").Value = arr(2)
  Fields("fldAdvEIC").Value = arr(3)
  Fields("fldBox10").Value = arr(4)
  Fields("fldBox11").Value = arr(5)
  Fields("fldBenefitBox").Value = arr(6)
  Fields("fldFedGrs").Value = arr(7)
  Fields("fldStateGrs").Value = arr(8)
  Fields("fldSocGrs").Value = arr(9)
  Fields("fldMedGrs").Value = arr(10)
  Fields("fldBox12a").Value = arr(11)
  Fields("fldBox12b").Value = arr(12)
  Fields("fldFedTax").Value = arr(13)
  Fields("fldStateTax").Value = arr(14)
  Fields("fldSocTax").Value = arr(15)
  Fields("fldMedTax").Value = arr(16)
  Fields("fldBox14a").Value = arr(17)
  Fields("fldBox14b").Value = arr(18)
  Fields("fldBox12c").Value = arr(19)
  Fields("fldBox12d").Value = arr(20)
  
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
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Now
  Me.Zoom = -1
  ReportHeader.Height = 0
  Label47.Visible = False

End Sub

Private Sub ReportFooter_Format()
  PageHeader.Height = 1600
  Label47.Visible = True
  Label64.Visible = False
  Label97.Visible = False
  Label99.Visible = False
  Label100.Visible = False
  Label95.Visible = False
  Label102.Visible = False
  Label103.Visible = False
  Label104.Visible = False
  Label105.Visible = False
  Label98.Visible = False
  Label106.Visible = False
  Label107.Visible = False
  Label108.Visible = False
  Label109.Visible = False
  Label110.Visible = False
  Label111.Visible = False
  Label129.Visible = False
  Label130.Visible = False
  Line1.Visible = False
End Sub
