VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arLvBftSplitRpt 
   Caption         =   "Leave Benefit Report"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11610
   Icon            =   "arLvBftSplitRpt.dsx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20479
   _ExtentY        =   14975
   SectionData     =   "arLvBftSplitRpt.dsx":08CA
End
Attribute VB_Name = "arLvBftSplitRpt"
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
      MsgBox "File - LeaveBenefitRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - LeaveBenefitRpt.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - LeaveBenefitRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - LeaveBenefitRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "LeaveBenefitRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "LeaveBenefitRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  HFile = FreeFile
  Open StartPath & "\PRRPTS\BENSPLIT.RPT" For Input As #HFile
  Fields.Add "fldEmployer" '0
  Fields.Add "fldDate" '1
  Fields.Add "fldEmpNo" '2
  Fields.Add "fldEmpName" '3
  Fields.Add "fldVacB" '4
  Fields.Add "fldSLB" '5
  Fields.Add "fldCTB" '6
  Fields.Add "fldPerB" '7
  Fields.Add "fldHolB" '8
  Fields.Add "fldLvTbl" '9
  Fields.Add "fldTotEmps" '10
  Fields.Add "fldVBTot" '11
  Fields.Add "fldSLBTot" '12
  Fields.Add "fldCTBTot" '13
  Fields.Add "fldPerBTot" '14
  Fields.Add "fldHolBTot" '15
  Fields.Add "fldPayYN" '16
  Fields.Add "fldTotPay" '17
  Fields.Add "fldVPay" '18
  Fields.Add "fldSPay" '19)
  Fields.Add "fldCPay" '20)
  Fields.Add "fldPPay" '21)
  Fields.Add "fldHPay" '22)
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
  Fields("fldEmpNo").Value = arr(2)
  Fields("fldEmpName").Value = arr(3)
  Fields("fldVacB").Value = arr(4)
  Fields("fldSLB").Value = arr(5)
  Fields("fldCTB").Value = arr(6)
  Fields("fldPerB").Value = arr(7)
  Fields("fldHolB").Value = arr(8)
  Fields("fldLvTbl").Value = arr(9)
  Fields("fldTotEmps").Value = arr(10)
  Fields("fldVBTot").Value = arr(11)
  Fields("fldSLBTot").Value = arr(12)
  Fields("fldCTBTot").Value = arr(13)
  Fields("fldPerBTot").Value = arr(14)
  Fields("fldHolBTot").Value = arr(15)
  Fields("fldPayYN").Value = arr(16)
  Fields("fldTotPay").Value = arr(17)
'  Label11.Visible = True
'  Field1.Visible = True
'  If Fields("fldPayYN").Value = False Then
'    Label11.Visible = False
'    Field1.Visible = False
'  End If
  Fields("fldVPay").Value = arr(18)
  Fields("fldSPay").Value = arr(19)
  Fields("fldCPay").Value = arr(20)
  Fields("fldPPay").Value = arr(21)
  Fields("fldHPay").Value = arr(22)
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

Private Sub Detail_Format()
  GroupHeader1.GroupValue = Fields("fldEmployer").Value
  
End Sub

Private Sub GroupFooter1_Format()
'  If Fields("fldPayYN").Value = True Then
'    Field2.Visible = True
'    Line3.X2 = 11250
'  Else
'    Field2.Visible = False
'  End If
End Sub

Private Sub GroupHeader1_Format()
'  If Fields("fldPayYN").Value = True Then
'    Line1.X1 = 0
'    Line1.X2 = 11250
'    Line2.X1 = 0
'    Line2.X2 = 11250
'  End If

End Sub

