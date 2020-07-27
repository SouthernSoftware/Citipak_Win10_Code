VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arFAAddDelTagOnly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Additions/Deletions"
   ClientHeight    =   8868
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arFAAddDelTagOnly.dsx":0000
   MaxButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arFAAddDelTagOnly.dsx":08CA
End
Attribute VB_Name = "arFAAddDelTagOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsFATextBoxOverRider
Private Temp_Class As Resize_Class
Private HFile As Integer
Dim NumOfDs As Integer
Dim NumOfAs As Integer
Dim PurchA As Double
Dim PurchD As Double
Dim TotDeprA As Double
Dim TotDeprD As Double
Dim BookValA As Double
Dim BookValD As Double

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
      MsgBox "File - FAAddDelTagOnly.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - FAAddDelTagOnly.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
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
    MsgBox "File - FAAddDelTagOnly.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - FAAddDelTagOnly.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "FAAddDelTagOnly.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "FAAddDelTagOnly.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  HFile = FreeFile
  Open StartPath & "\FARPTS\FAADDDELTAG.RPT" For Input As #HFile
  Fields.Add "fldEmployer" '0
  Fields.Add "fldSDate" '1
  Fields.Add "fldEDate" '2
  Fields.Add "fldAddOrDel" '3
  Fields.Add "fldTagNum" '4
  Fields.Add "fldItemDesc" '5
  Fields.Add "fldDeptNum" '6
  Fields.Add "fldAssetLife" '7
  Fields.Add "fldOrigCost" '8
  Fields.Add "fldTotDepr" '9
  Fields.Add "fldBookVal" '10
  Fields.Add "fldNoDep" '11
  Fields.Add "fldLifeLeft" '12
  Fields.Add "fldLifeData"
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
  Fields("fldSDate").Value = arr(1)
  Fields("fldEDate").Value = arr(2)
  Fields("fldAddOrDel").Value = arr(3)
  Fields("fldTagNum").Value = arr(4)
  Fields("fldItemDesc").Value = arr(5)
  Fields("fldDeptNum").Value = arr(6)
  Fields("fldAssetLife").Value = arr(7)
  Fields("fldOrigCost").Value = arr(8)
  Fields("fldTotDepr").Value = arr(9)
  Fields("fldBookVal").Value = arr(10)
  Fields("fldNoDep").Value = arr(11)
  Fields("fldLifeLeft").Value = arr(12)
  Fields("fldLifeData").Value = arr(7) + "/" + arr(12)
End Sub

Private Sub ActiveReport_ReportEnd()
'  Unload frmLoadingRpt
  If HFile <> 0 Then
    Close #HFile
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
'    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Now
  Me.Zoom = -1
  Label11.Visible = False
  Label34.Visible = False
  Label35.Visible = False
  Label36.Visible = False
  Line5.Visible = False
End Sub

Private Sub PageHeader_Format()
  If Label11.Caption = "By Department" Then
    Label37.Visible = False
  End If

End Sub

Private Sub ReportFooter_Format()
  PageHeader.Height = 1400
  Set SubReport2.object = New arFASubAddDelTagGrand
  Line5.Visible = False
  Label11.Caption = "Summary"
  Label11.Visible = True
  Label14.Visible = False
  Label5.Visible = False
  Label3.Visible = False
  Label2.Visible = False
  Label16.Visible = False
  Label17.Visible = False
  Label18.Visible = False
  Label19.Visible = False
  Label34.Visible = False
  Label35.Visible = False
  Label36.Visible = False
  Label37.Visible = False
  Line1.Visible = False
End Sub

Private Sub GroupFooter1_Format()
  PageHeader.Height = 1400
  Set SubReport1.object = New arFASubAddDelTagDepts
  Line5.Visible = True
  Line1.Visible = False
  Label34.Visible = True
  Label35.Visible = True
  Label36.Visible = True
  Label11.Visible = True
  Label11.Caption = "By Department"
  Label14.Visible = False
  Label5.Visible = False
  Label3.Visible = False
  Label2.Visible = False
  Label16.Visible = False
  Label17.Visible = False
  Label18.Visible = False
  Label19.Visible = False
End Sub
