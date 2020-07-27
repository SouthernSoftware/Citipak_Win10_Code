VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arFAMasterRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Asset Listing by Department"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arFAMasterRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arFAMasterRpt.dsx":08CA
End
Attribute VB_Name = "arFAMasterRpt"
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
      MsgBox "File - FAMasterRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - FAMasterRpt.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - FAMasterRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - FAMasterRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "FAMasterRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "FAMasterRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  HFile = FreeFile
  Open StartPath & "\FARPTS\FAMASTER.RPT" For Input As #HFile
  Fields.Add ("fldItemTag") '0)
  Fields.Add ("fldSerialNumb") '1)
  Fields.Add ("fldIDesc1") '2)
  Fields.Add ("fldMfg") '3)
  Fields.Add ("fldDsplYN") '4)
  Fields.Add ("fldContact") '5)
  Fields.Add ("fldLocation") '6)
  Fields.Add ("fldIDept") '7)
  Fields.Add ("fldOrgCost") '8)
  Fields.Add ("fldAcqDate") '9)
  Fields.Add ("fldILife") '10)
  Fields.Add ("fldDpr2Date") '11)
  Fields.Add ("fldIStatus") '12)
  Fields.Add ("fldBookTot") '13)
  Fields.Add ("fldAssCode") '14)
  Fields.Add ("fldThisDept") '15)
  Fields.Add ("fldThisDesc") '16)
  Fields.Add ("fldEmployer") '17)
  Fields.Add ("fldDitemCnt") '18)
  Fields.Add ("fldDOrigCost") '19)
  Fields.Add ("fldDDepTotal") '20)
  Fields.Add ("fldDBookTot") '21)
  Fields.Add ("fldOneDept") '22)
  Fields.Add ("fldLifeLeft") '23)
  Fields.Add ("fldDEPYN") '24)
  Fields.Add ("fldDisposalDate") '25)
  Fields.Add ("fldDeprYN") '26)
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
  Fields("fldItemTag").Value = arr(0)
  Fields("fldSerialNumb").Value = arr(1)
  Fields("fldIDesc1").Value = arr(2)
  Fields("fldMfg").Value = arr(3)
  Fields("fldDsplYN").Value = arr(4)
  Fields("fldContact").Value = arr(5)
  Fields("fldLocation").Value = arr(6)
  Fields("fldIDept").Value = arr(7)
  Fields("fldOrgCost").Value = arr(8)
  Fields("fldAcqDate").Value = arr(9)
  Fields("fldILife").Value = arr(10)
  Fields("fldDpr2Date").Value = arr(11)
  Fields("fldIStatus").Value = arr(12)
  If QPTrim$(arr(12)) = "A" Then
    Fields("fldIStatus").Value = "Active"
  Else
    Fields("fldIStatus").Value = "Inactive"
  End If
  Fields("fldBookTot").Value = arr(13)
  Fields("fldAssCode").Value = arr(14)
  Fields("fldThisDept").Value = arr(15)
  Fields("fldThisDesc").Value = arr(16)
  Fields("fldEmployer").Value = arr(17)
  Fields("fldDitemCnt").Value = arr(18)
  Fields("fldDOrigCost").Value = arr(19)
  Fields("fldDDepTotal").Value = arr(20)
  Fields("fldDBookTot").Value = arr(21)
  Fields("fldOneDept").Value = arr(22)
  Fields("fldLifeLeft").Value = arr(10) + "/" + arr(23)
  Fields("fldDEPYN").Value = arr(24)
  Fields("fldDisposalDate").Value = arr(25)
  Fields("fldDeprYN").Value = arr(26)
'  If QPTrim$(arr(24)) <> "" Then
'    Label49.Visible = True
'  Else
'    Label49.Visible = False
'  End If
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
  If Fields("fldDsplYN").Value = "Y" Then
    Label49.Visible = True
  Else
    Label49.Visible = False
  End If
  Me.fldTimeDate.Text = Now
  Me.Zoom = -1
  Label11.Visible = False
End Sub

Private Sub ReportFooter_Format()
  If Fields(22).Value = 1 Or Fields(22).Value = 0 Then
    ReportFooter.Visible = False
  End If
  Line2.Visible = True
  Line1.Visible = False
  Label50.Visible = False
  Set SubReport2.object = New arFASubMasterDept
  Set SubReport1.object = New arFASubMaster
  Label11.Caption = "Summary"
  Label11.Visible = True
  Label64.Visible = False
End Sub

Private Sub Detail_Format()
  Line2.Visible = True
End Sub

Private Sub GroupHeader2_Format()
  Line2.Visible = False
End Sub
