VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arFANewAddDelRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Additions/Deletions"
   ClientHeight    =   8868
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arFANewAddDelRpt.dsx":0000
   MaxButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arFANewAddDelRpt.dsx":08CA
End
Attribute VB_Name = "arFANewAddDelRpt"
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
      MsgBox "File - FANewAddDelRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - FANewAddDelRpt.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - FANewAddDelRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - FANewAddDelRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "FANewAddDelRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "FANewAddDelRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  HFile = FreeFile
  Open StartPath & "\FARPTS\FAADDNEW.RPT" For Input As #HFile
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
  Fields.Add "fldDOrigCost" '11
  Fields.Add "fldDYDep" '12
  Fields.Add "fldDBookTotal" '13
  Fields.Add "fldAOrigCost" '14
  Fields.Add "fldAYDep" '15
  Fields.Add "fldABookTotal" '16
  Fields.Add "fldGTDOrigCost" '17
  Fields.Add "fldGTDYDep" '18
  Fields.Add "fldGTDBookTotal" '19
  Fields.Add "fldGTAOrigCost" '20
  Fields.Add "fldGTAYDep" '21
  Fields.Add "fldGTABookTotal" '22
  Fields.Add "fldACnt" '23
  Fields.Add "fldDCnt" '24
  Fields.Add "fldNoDep" '25
  Fields.Add "fldTACnt" '26
  Fields.Add "fldTDCnt" '27
  Fields.Add "fldDeptDesc" '28
  Fields.Add "fldLifeLeft" '29
  Fields.Add "fldLifeData"
  Fields.Add "fldTotDsplPrice"
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
  Fields("fldDOrigCost").Value = arr(11)
  Fields("fldDYDep").Value = arr(12)
  Fields("fldDBookTotal").Value = arr(13)
  Fields("fldAOrigCost").Value = arr(14)
  Fields("fldAYDep").Value = arr(15)
  Fields("fldABookTotal").Value = arr(16)
  Fields("fldGTDOrigCost").Value = arr(17)
  Fields("fldGTDYDep").Value = arr(18)
  Fields("fldGTDBookTotal").Value = arr(19)
  Fields("fldGTAOrigCost").Value = arr(20)
  Fields("fldGTAYDep").Value = arr(21)
  Fields("fldGTABookTotal").Value = arr(22)
  Fields("fldACnt").Value = arr(23)
  Fields("fldDCnt").Value = arr(24)
  Fields("fldNoDep").Value = arr(25)
  Fields("fldTACnt").Value = arr(26)
  Fields("fldTDCnt").Value = arr(27)
  Fields("fldDeptDesc").Value = arr(28)
  Fields("fldLifeLeft").Value = arr(29)
  Fields("fldLifeData").Value = arr(7) + "/" + arr(29)
  Fields("fldTotDsplPrice").Value = arr(30)
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
End Sub

Private Sub ReportFooter_Format()
  PageHeader.Height = 1800
  Label11.Visible = True
  Label14.Visible = False
  Label5.Visible = False
  Label3.Visible = False
  Label2.Visible = False
  Label16.Visible = False
  Label17.Visible = False
  Label18.Visible = False
  Label19.Visible = False
  Label35.Visible = False
  Line1.Visible = False
  If Fields("fldTotDsplPrice").Value = 0 Then
    Label36.Visible = False
    Field6.Visible = False
  End If
End Sub
