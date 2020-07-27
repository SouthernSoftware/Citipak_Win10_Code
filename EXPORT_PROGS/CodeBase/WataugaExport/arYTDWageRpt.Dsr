VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arYTDWageRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Year To Date Wage Report"
   ClientHeight    =   8868
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arYTDWageRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arYTDWageRpt.dsx":08CA
End
Attribute VB_Name = "arYTDWageRpt"
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
      MsgBox "File - YTDWageRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - YTDWageRpt.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - YTDWageRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - YTDWageRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "YTDWageRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "YTDWageRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  Dim X As Integer
  HFile = FreeFile
  Open StartPath & "\PRRPTS\YTDWAGEG.RPT" For Input As #HFile
  Fields.Add "fldMonth" '0
  Fields.Add "fldFundCnt" '1
  Fields.Add "fldTotalFlag" '2
  Fields.Add "fldEmployer" '3
  Fields.Add "fldYear" '4
  Fields.Add "fldToday" '5
  Fields.Add "fldFundNo" '6
  Fields.Add "fldRegHrs" '7
  Fields.Add "fldOTHrs" '8
  Fields.Add "fldRegWgs" '9
  Fields.Add "fldOTWgs" '10
  Fields.Add "fldFundT" '11
  Fields.Add "fldRegHrsTot" '12
  Fields.Add "fldOTHrsTot" '13
  Fields.Add "fldRegWgTot" '14
  Fields.Add "fldOTWgsTot" '15
  Fields.Add "fldLbl1" '16
  Fields.Add "fldLbl2" '17
  Fields.Add "fldTrip" '18
  Fields.Add "fldTrip2"
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
  Fields("fldMonth").Value = arr(0)
  Fields("fldFundCnt").Value = arr(1)
  Fields("fldTotalFlag").Value = arr(2)
  Fields("fldEmployer").Value = arr(3)
  Fields("fldYear").Value = arr(4)
  Fields("fldToday").Value = arr(5)
  Fields("fldFundNo").Value = arr(6)
  Fields("fldRegHrs").Value = arr(7)
  Fields("fldOTHrs").Value = arr(8)
  Fields("fldRegWgs").Value = arr(9)
  Fields("fldOTWgs").Value = arr(10)
  Fields("fldFundT").Value = arr(11)
  Fields("fldRegHrsTot").Value = arr(12)
  Fields("fldOTHrsTot").Value = arr(13)
  Fields("fldRegWgTot").Value = arr(14)
  Fields("fldOTWgsTot").Value = arr(15)
  Fields("fldLbl1").Value = arr(16)
  Fields("fldLbl2").Value = arr(17)
  Fields("fldTrip").Value = arr(18)
  Fields("fldTrip2").Value = arr(19)
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
  Set SubReport1.object = New arSubYTDTotals
  fldLbl1.Visible = False
  Fields("fldLbl2").Value = "Fund Number"
  Label1.Caption = "YTD Wage Distribution Fund Summary"
End Sub
