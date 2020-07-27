VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ar941 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Form 941 Assitance"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "ar941.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "ar941.dsx":08CA
End
Attribute VB_Name = "ar941"
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
    DoEvents
  End If
  If Shift = 4 Then
    If KeyCode = vbKeyC Then
      Unload Me
      DoEvents
    ElseIf KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - 941FormsRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - 941FormsRpt.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close '5/28/2004
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
  If Tool.Caption = "&Close" Then
    Unload Me
    DoEvents
  End If
  If Tool = "Save/&Excel" Then
    Screen.MousePointer = vbHourglass
    ExportReport 1
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - 941FormsRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - 941FormsRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "941FormsRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "941FormsRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub
Private Sub ActiveReport_DataInitialize()
  HFile = FreeFile
  Open StartPath & "\PRRPTS\941FORMS.RPT" For Input As #HFile
  Fields.Add "EmpName" '(0)
  Fields.Add "TradeName" '(1)
  Fields.Add "Add1" '(2)
  Fields.Add "Add2" '(3)
  Fields.Add "City" '(4)
  Fields.Add "State" '(5)
  Fields.Add "Zip" '(6)
  Fields.Add "EmpIDNum" '(7)
  Fields.Add "QtrEnd" '(8)
  Fields.Add "NumOfEmps" '(9)
  Fields.Add "Signer" '(10)
  Fields.Add "Title" '(11)
  Fields.Add "Box2" '(12)
  Fields.Add "Box3" '(13)
  Fields.Add "Box4" '(14)
  Fields.Add "Box5" '(15)
  Fields.Add "Box6a" '(16)
  Fields.Add "Box6aPct" '(17)
  Fields.Add "Box6b" '(18)
  Fields.Add "Box6c" '(19)
  Fields.Add "Box6cPct" '(20)
  Fields.Add "Box6d" '(21)
  Fields.Add "Box7a" '(22)
  Fields.Add "Box7aPct" '(23)
  Fields.Add "Box7b" '(24)
  Fields.Add "Box8" '(25)
  Fields.Add "SickPay" '(26)
  Fields.Add "FracCents" '(27)
  Fields.Add "Other" '(28)
  Fields.Add "Box9" '(29)
  Fields.Add "Box10" '(30)
  Fields.Add "Box11" '(31)
  Fields.Add "Box12" '(32)
  Fields.Add "Box13" '(33)
  Fields.Add "Box14" '(34)
  Fields.Add "Box15" '(35)
  Fields.Add "Box16" '(36)
  Fields.Add "Box17a" '(37)
  Fields.Add "Box17b" '(38)
  Fields.Add "Box17c" '(39)
  Fields.Add "Box17d" '(40)
  Fields.Add "CSZ"
  End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  Unload frmLoadingRpt
  CancelDisplay = True 'removes the error message
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  
  ' We reached the end of the file we exit leaving the
  ' eof parameter as True (default except on first call) that will
  ' tell AR that we are done feeding data
  ' otherwise we have to set the eof parameter to False so that
  ' AR continues fetching data, until we're done
  ' if the report had a data control, the value of the parameter
  ' will be ignored, AR will always follow the data control's recordset
  ' EOF property
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
  Fields("EmpName").Value = arr(0)
  Fields("TradeName").Value = arr(1)
  Fields("Add1").Value = arr(2)
  Fields("Add2").Value = arr(3)
  Fields("City").Value = arr(4)
  Fields("State").Value = arr(5)
  Fields("Zip").Value = arr(6)
  Fields("CSZ").Value = QPTrim$(arr(4)) + ", " + QPTrim$(arr(5)) + "  " + QPTrim$(arr(6))
  Fields("EmpIDNum").Value = arr(7)
  Fields("QtrEnd").Value = arr(8)
  Fields("NumOfEmps").Value = arr(9)
  Fields("Signer").Value = arr(10)
  Fields("Title").Value = arr(11)
  Fields("Box2").Value = arr(12)
  Fields("Box3").Value = arr(13)
  Fields("Box4").Value = arr(14)
  Fields("Box5").Value = arr(15)
  Fields("Box6a").Value = arr(16)
  Fields("Box6aPct").Value = arr(17)
  Fields("Box6b").Value = arr(18)
  Fields("Box6c").Value = arr(19)
  Fields("Box6cPct").Value = arr(20)
  Fields("Box6d").Value = arr(21)
  Fields("Box7a").Value = arr(22)
  Fields("Box7aPct").Value = arr(23)
  Fields("Box7b").Value = arr(24)
  Fields("Box8").Value = arr(25)
  Fields("SickPay").Value = arr(26)
  Fields("FracCents").Value = arr(27)
  Fields("Other").Value = arr(28)
  Fields("Box9").Value = arr(29)
  Fields("Box10").Value = arr(30)
  Fields("Box11").Value = arr(31)
  Fields("Box12").Value = arr(32)
  Fields("Box13").Value = arr(33)
  Fields("Box14").Value = arr(34)
  Fields("Box15").Value = arr(35)
  Fields("Box16").Value = arr(36)
  Fields("Box17a").Value = arr(37)
  Fields("Box17b").Value = arr(38)
  Fields("Box17c").Value = arr(39)
  Fields("Box17d").Value = arr(40)
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


