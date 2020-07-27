VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arGLRegisterNS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "G/L Interface Non-Split"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arGLRegisterNS.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arGLRegisterNS.dsx":08CA
End
Attribute VB_Name = "arGLRegisterNS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private hFile As Integer
Dim NumOfAccts As Integer
Dim NoEscape As Boolean
Private Sub ActiveReport_Initialize()
  Me.ToolBar.Tools.Add "&Close"
  Me.ToolBar.Tools.Add "Save/&Excel"
  Me.ToolBar.Tools.Add "&Text"
End Sub

Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  If KeyCode = vbKeyEscape And NoEscape = False Then
    NoEscape = True
    DoEvents
    Unload Me
    DoEvents
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
      MsgBox "File - GLRegisterNS.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - GLRegisterNS.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
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
    MsgBox "File - GLRegisterNS.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - GLRegisterNS.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "GLRegisterNS.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "GLRegisterNS.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\PRRPTS\PRGLIFNSG.RPT" For Input As #hFile
  
  Fields.Add "fldAcctNum" '(0)
  Fields.Add "fldNoAcctNum" '(1)
  Fields.Add "fldDesc" '(2)
  Fields.Add "fldDebit" '(3)
  Fields.Add "fldCredit" '(4)
  Fields.Add "fldTotalDebit" '(5)
  Fields.Add "fldTotalCredit" '(6)
  Fields.Add "fldError" '(7)
1  Fields.Add "fldEmployer" '(8)
  Fields.Add "fldDate" '(9)
  Fields.Add "fldNumOfAccts" '(10)

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
  Fields("fldAcctNum").Value = arr(0)
  Fields("fldNoAcctNum").Value = arr(1)
  Fields("fldDesc").Value = arr(2)
  Fields("fldDebit").Value = arr(3)
  Fields("fldCredit").Value = arr(4)
  Fields("fldTotalDebit").Value = arr(5)
  Fields("fldTotalCredit").Value = arr(6)
  Fields("fldError").Value = arr(7)
  Fields("fldEmployer").Value = arr(8)
  Fields("fldDate").Value = arr(9)
  Fields("fldNumOfAccts").Value = arr(10)
  If arr(10) > 0 Then NumOfAccts = arr(10)
  
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
  NoEscape = False

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

Private Sub GroupFooter1_Format()
  If NumOfAccts = 0 Then
    fldError.Visible = True
    fldTotalDebit.Visible = False
    fldTotalCredit.Visible = False
  Else
    fldError.Visible = False
    fldTotalDebit.Visible = True
    fldTotalCredit.Visible = True
  End If
    
End Sub

Private Sub ReportFooter_Format()
  Set SubReport1.object = New arSubGLFundTotals
  PageHeader.Height = 1620
  Label47.Visible = True
  Label64.Visible = False
  Label95.Visible = False
  Label96.Visible = False
  Label97.Visible = False

End Sub

