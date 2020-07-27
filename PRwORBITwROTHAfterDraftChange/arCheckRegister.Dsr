VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arCheckRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Register"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arCheckRegister.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arCheckRegister.dsx":08CA
End
Attribute VB_Name = "arCheckRegister"
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
      MsgBox "File - CheckRegister.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - CheckRegister.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - CheckRegister.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - CheckRegister.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "CheckRegister.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "CheckRegister.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\PRRPTS\CHECKREGG.RPT" For Input As #hFile
  Fields.Add "fldEmployer" '(0)
  Fields.Add "fldDate" '(1)
  Fields.Add "fldCheckNum" '(2)
  Fields.Add "fldEmployee" '(3)
  Fields.Add "fldDraft" '(4)
  Fields.Add "fldCheckAmt" '(5)
  Fields.Add "fldReprint" '(6)
  Fields.Add "fldNumOfChecks" '(7)
  Fields.Add "fldAmtOfChecks" '(8)
  Fields.Add "fldNumOfDrafts" '(9)
  Fields.Add "fldAmtOfDrafts" '(10)
  Fields.Add "fldTotalAmt" '(11)
  Fields.Add "fldTotalWritten" '(12)
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
  Fields("fldCheckNum").Value = arr(2)
  Fields("fldEmployee").Value = arr(3)
  Fields("fldDraft").Value = arr(4)
  Fields("fldCheckAmt").Value = arr(5)
  Fields("fldReprint").Value = arr(6)
  Fields("fldNumOfChecks").Value = arr(7)
  Fields("fldAmtOfChecks").Value = arr(8)
  Fields("fldNumOfDrafts").Value = arr(9)
  Fields("fldAmtOfDrafts").Value = arr(10)
  Fields("fldTotalAmt").Value = arr(11)
  Fields("fldTotalWritten").Value = arr(12)
  
  If Len(arr(6)) > 0 Then
    Reprint = True
  Else
    Reprint = False
  End If
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

Private Sub Detail_Format()
  If Reprint = True Then
    fldReprint.Visible = True
    fldCheckNum.Visible = False
    fldEmployee.Visible = False
    fldDraft.Visible = False
    fldCheckAmt.Visible = False
    Reprint = False
  ElseIf Reprint = False Then
    fldReprint.Visible = False
    fldCheckNum.Visible = True
    fldEmployee.Visible = True
    fldCheckAmt.Visible = True
    fldDraft.Visible = True
  End If

End Sub

Private Sub ReportFooter_Format()
  If Fields("fldNumOfDrafts").Value = 0 Then
    Label100.Visible = False
    Label101.Visible = False
    fldNumOfDrafts.Visible = False
    fldAmtOfDrafts.Visible = False
  End If
  Label47.Visible = True
  Label64.Visible = False
  Label95.Visible = False
  Label97.Visible = False
End Sub

