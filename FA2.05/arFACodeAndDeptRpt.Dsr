VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arFACodeAndDeptRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets Codes by Department"
   ClientHeight    =   8868
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arFACodeAndDeptRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arFACodeAndDeptRpt.dsx":08CA
End
Attribute VB_Name = "arFACodeAndDeptRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsFATextBoxOverRider
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
      MsgBox "File - FADeptAndFundRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - FADeptAndFundFundRpt.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - FADeptAndFundFundRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - FADeptAndFundFundRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "FADeptAndFundFundRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "FADeptAndFundFundRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  HFile = FreeFile
  Open StartPath & "\FARPTS\FADPTCDE.RPT" For Input As #HFile
  Fields.Add ("fldEmployer") '0)
  Fields.Add ("fldTagNum") '1)
  Fields.Add ("fldItemDesc") '2)
  Fields.Add ("fldItemCode") '3)
  Fields.Add ("fldAssetLife") '4)
  Fields.Add ("fldOrigCost") '5)
  Fields.Add ("fldDepr") '6)
  Fields.Add ("fldStar") '7)
  Fields.Add ("fldBookVal") '8)
  Fields.Add ("fldCode") '9)
  Fields.Add ("fldCodeDesc1") '10)
  Fields.Add ("fldThisCodeNum") '11)
  Fields.Add ("fldCodePurchPr") '12)
  Fields.Add ("fldCodeDpr") '13)
  Fields.Add ("fldCodeBookVal") '14)
  Fields.Add ("fldCodeCnt") '15)
  Fields.Add ("fldGTPurchPr") '16)
  Fields.Add ("fldGTDpr") '17)
  Fields.Add ("fldGTBookVal") '18)
  Fields.Add ("fldTotalCnt") '19)
  Fields.Add ("fldDEPYN") '20)
  Fields.Add ("fldLifeLeft") '21)
  Fields.Add ("fldDispDate") '22)
  Fields.Add ("fldShowDisp") '23)
  Fields.Add ("fldDeptNum") '24)
  Fields.Add ("fldDeptDesc") '25)
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
  Fields("fldEmployer").Value = arr(0)
  Fields("fldTagNum").Value = arr(1)
  Fields("fldItemDesc").Value = arr(2)
  Fields("fldItemCode").Value = arr(3)
  Fields("fldAssetLife").Value = arr(4)
  Fields("fldOrigCost").Value = arr(5)
  Fields("fldDepr").Value = arr(6)
  Fields("fldStar").Value = arr(7)
  Fields("fldBookVal").Value = arr(8)
  Fields("fldCode").Value = arr(9)
  Fields("fldCodeDesc1").Value = arr(10)
  Fields("fldThisCodeNum").Value = arr(11)
  Fields("fldCodePurchPr").Value = arr(12)
  Fields("fldCodeDpr").Value = arr(13)
  Fields("fldCodeBookVal").Value = arr(14)
  Fields("fldCodeCnt").Value = arr(15)
  Fields("fldGTPurchPr").Value = arr(16)
  Fields("fldGTDpr").Value = arr(17)
  Fields("fldGTBookVal").Value = arr(18)
  Fields("fldTotalCnt").Value = arr(19)
  Fields("fldDEPYN").Value = arr(20)
  Fields("fldLifeLeft").Value = arr(21)
  Fields("fldDispDate").Value = arr(22)
  Fields("fldShowDisp").Value = arr(23)
  Fields("fldDeptNum").Value = arr(24)
  Fields("fldDeptDesc").Value = arr(25)
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
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Now
  Me.Zoom = -1
  PageHeader.Height = 1300
  Label70.Visible = False
  Line11.Visible = False
  Label40.Visible = False
'  Label11.Visible = False
  Label64.Visible = False
  Label65.Visible = False
  Label66.Visible = False
  Label67.Visible = False
  Label68.Visible = False
  Label69.Visible = False
  If Fields("fldShowDisp").Value = "N" Then
    Line1.X2 = 10530
    Line4.X2 = 10530
    Line6.X2 = 10530
    Line7.X2 = 10530
  Else
    Label40.Visible = True
    Label41.Visible = True
    Line1.X2 = 11790
    Line4.X2 = 11790
    Line6.X2 = 11790
    Line7.X2 = 11790
  End If
End Sub

Private Sub ReportFooter_Format()
  If Fields("fldCode").Value <> "ALL" Then
    Set SubReport1.object = New arFASubCode
    PageHeader.Height = 1200
    Exit Sub
  Else
    Label70.Visible = True
    PageHeader.Height = 1600
    Set SubReport2.object = New arFASub2Code
    SubReport1.Visible = False
    SubReport2.Top = 300
    Line10.Visible = False
    Label52.Visible = False
    Label53.Visible = False
    Label54.Visible = False
    Label55.Visible = False
    Label56.Visible = False
    Label57.Visible = False
  End If
  Label11.Visible = True
  Label16.Visible = False
  Label17.Visible = False
  Label18.Visible = False
  Label19.Visible = False
  Label20.Visible = False
  Label21.Visible = False
  Label22.Visible = False

End Sub






