VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arFADisposedOFTagOnly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets Disposed Of Report"
   ClientHeight    =   8868
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   10620
   Icon            =   "arFADisposedOFTagOnly.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   18733
   _ExtentY        =   15637
   SectionData     =   "arFADisposedOFTagOnly.dsx":08CA
End
Attribute VB_Name = "arFADisposedOFTagOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsFATextBoxOverRider
Private Temp_Class As Resize_Class
Private HFile As Integer
Dim GTCnt As Integer

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
      MsgBox "File - FADisposedOfTagOnly.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - FADisposedOfTagOnly.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - FADisposedOfTagOnly.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - FADisposedOfTagOnly.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "FADisposedOfTagOnly.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "FADisposedOfTagOnly.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  HFile = FreeFile
  Open StartPath & "\FARPTS\FAADDDELOPTTAG.RPT" For Input As #HFile
  Fields.Add "fldDept" '0
  Fields.Add "fldRptType" '1
  Fields.Add "fldEmployer" '2
  Fields.Add "fldTagNum" '3
  Fields.Add "fldItemDesc" '4
  Fields.Add "fldDeptNum" '5
  Fields.Add "fldMethod" '6
  Fields.Add "fldOrigCost" '7
  Fields.Add "fldTotDepr" '8
  Fields.Add "fldBookVal" '9
  Fields.Add "fldDispDate" '10
  Fields.Add "fldDispPrice" '11
  Fields.Add "fldStart" '12)
  Fields.Add "fldEnd" '13)
  Fields.Add "fldLifeLeft" '14)
  Fields.Add "fldLife" '15)
  Fields.Add "fldGTCnt"
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmFALoadReport
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
  Fields("fldDept").Value = arr(0)
  Fields("fldRptType").Value = arr(1)
  Fields("fldEmployer").Value = arr(2)
  Fields("fldTagNum").Value = arr(3)
  GTCnt = GTCnt + 1
  Fields("fldItemDesc").Value = arr(4)
  Fields("fldDeptNum").Value = arr(5)
  Fields("fldMethod").Value = arr(6)
  Fields("fldOrigCost").Value = arr(7)
  Fields("fldTotDepr").Value = arr(8)
  Fields("fldBookVal").Value = arr(9)
  Fields("fldDispDate").Value = arr(10)
  Fields("fldDispPrice").Value = arr(11)
  Fields("fldStart").Value = arr(12)
  Fields("fldEnd").Value = arr(13)
  Fields("fldLifeLeft").Value = arr(14)
  Fields("fldLife").Value = arr(15)
  Fields("fldGTCnt").Value = GTCnt
  If QPTrim$(arr(1)) <> "D" Then
    Fields("fldMethod").Value = arr(15) + "/" + arr(14)
  End If
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmFALoadReport
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
End Sub

Private Sub Detail_Format()
  If Fields("fldRptType").Value <> "D" Then
    fldMethod.Left = 5500
    fldOrigCost.Left = 6500
    fldBookVal.Left = 8200
  End If

End Sub

Private Sub PageHeader_Format()
  If Fields("fldRptType").Value = "D" Then
    Label1.Caption = "Fixed Assets Disposed Of"
  Else
    Label1.Caption = "Fixed Assets Acquired"
    Label19.Caption = "Life/Left"
    Label21.Visible = False
    Label35.Visible = False
    Label23.Caption = "Date Acquired"
    Label19.Left = 5500
    Label20.Left = 6490
    Label22.Left = 8240
  End If
End Sub

Private Sub ReportFooter_Format()
  Set SubReport1.object = New arFASubDisposedOf
  PageHeader.Height = 1400
  Line1.Visible = False
  Label11.Visible = True
  Label11.Caption = "Summary"
  Label16.Visible = False
  Label17.Visible = False
  Label18.Visible = False
  Label19.Visible = False
  Label20.Visible = False
  Label21.Visible = False
  Label22.Visible = False
  Label23.Visible = False
  If Fields("fldRptType").Value <> "D" Then
    Field6.Visible = False
    Line3.X2 = 9800
    Line4.X2 = 9800
    Label34.Left = 8300
    Field7.Left = 7950
  End If
End Sub

Private Sub GroupFooter1_Format()
  PageHeader.Height = 1700
  Set SubReport1.object = New arFASubDisposedOf
  Line1.Visible = True
  Label11.Visible = True
  Label11.Caption = "Department Totals"
  Label16.Visible = False
  Label17.Alignment = ddTXLeft
  Label17.Caption = "Department     # Items"
  Label18.Visible = False
  Label19.Visible = False
  Label20.Visible = True
  Label21.Visible = True
  Label22.Visible = True
  Label23.Visible = False
End Sub

