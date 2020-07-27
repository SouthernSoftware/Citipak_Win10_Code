VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arORBIT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ORBIT Report"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   Icon            =   "arORBIT.dsx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15584
   SectionData     =   "arORBIT.dsx":08CA
End
Attribute VB_Name = "arORBIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private hFile As Integer
Dim PL21() As String
Dim PL22() As String
Dim PL23() As String
Dim PL24() As String
Dim PL1() As String
Dim PL2() As String
Dim PL3() As String
Dim PL4() As String
Dim GStart As Integer

Private Sub ActiveReport_Initialize()
  Me.ToolBar.Tools.Add "&Close"
  Me.ToolBar.Tools.Add "Save/&Excel"
  Me.ToolBar.Tools.Add "&Text"
End Sub
Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
    DoEvents
    If frmORBITPost.Visible = False Then
      frmORBITMenu.Show
    End If
  End If
  If Shift = 4 Then
    If KeyCode = vbKeyC Then
      Unload Me
      DoEvents
      If frmORBITPost.Visible = False Then
        frmORBITMenu.Show
      End If
    ElseIf KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - ORBITRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - ORBITRpt.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
  If Tool.Caption = "&Close" Then
    Unload Me
    DoEvents
    If frmORBITPost.Visible = False Then
     frmORBITMenu.Show
    End If
  End If
  If Tool = "Save/&Excel" Then
    Screen.MousePointer = vbHourglass
    ExportReport 1
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - ORBITRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - ORBITRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "ORBITRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "ORBITRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub
Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\PRRPTS\ORBIT.RPT" For Input As #hFile
  Fields.Add "Employer" '(0)
  Fields.Add "Agency" '(1)
  Fields.Add "FileDate" '(2)
  Fields.Add "FormatVrs" '(3)
  Fields.Add "HdrStartDate" '(4)
  Fields.Add "HdrEndDate" '(5)
  Fields.Add "RptPrd" '(6)
  Fields.Add "LastName" '(7)
  Fields.Add "FirstName" '(8)
  Fields.Add "MiddleName" '(9)
  Fields.Add "Suffix" '(10)
  Fields.Add "EmployeeContr" '(11)
  Fields.Add "EmployerContr" '(12)
  Fields.Add "JobClass" '(13)
  Fields.Add "MemberID" '(14)
  Fields.Add "OTPay" '(15)
  Fields.Add "DtlStartDate" '(16)
  Fields.Add "DtlEndDate" '(17)
  Fields.Add "PlanCode" '(18)
  Fields.Add "Salary" '(19)
  Fields.Add "EmpNum" '(20)
  For x = 1 To 28
    Fields.Add ("PL1" & CStr(x))
    Fields.Add ("PL2" & CStr(x))
    Fields.Add ("PL3" & CStr(x))
    Fields.Add ("PL4" & CStr(x))
  Next x
  Fields.Add "TJCPay"
  Fields.Add "TJCCnt"
  Fields.Add "TJCEmpCont"
  Fields.Add "TJCCityCont"
  Fields.Add "TPCPay"
  Fields.Add "TPCCnt"
  Fields.Add "TPCEmpCont"
  Fields.Add "TPCCityCont"
  Fields.Add "Name"
  Fields.Add "PayType"
  Fields.Add "ADjustment"
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
  Dim x As Integer
  Dim Test As Integer
  Dim ArrCnt As Integer
  Static Start As Integer
  If VBA.eof(hFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Start = Start + 1
  GStart = Start
  If Start > 1 Then
    ReDim PL21(1 To 68) As String
    ReDim PL22(1 To 68) As String
    ReDim PL23(1 To 68) As String
    ReDim PL24(1 To 68) As String
    For x = 1 To 28
      PL21(x) = PL1(x)
      PL22(x) = PL2(x)
      PL23(x) = PL3(x)
      PL24(x) = PL4(x)
    Next x
  End If
  Line Input #hFile, sLine
  arr = Split(sLine, "~")
  Fields("Employer").Value = arr(0)
  Fields("Agency").Value = arr(1)
  Fields("FileDate").Value = MakeRegDate(CInt(arr(2)))
  Fields("FormatVrs").Value = arr(3)
  Fields("HdrStartDate").Value = MakeRegDate(CInt(arr(4)))
  Fields("HdrEndDate").Value = MakeRegDate(CInt(arr(5)))
  Fields("RptPrd").Value = FormatThisPayPd(arr(6), 1)
  Fields("LastName").Value = arr(7)
  Fields("FirstName").Value = arr(8)
  Fields("MiddleName").Value = arr(9)
  Fields("Suffix").Value = arr(10)
  Fields("EmployeeContr").Value = arr(11)
  Fields("EmployerContr").Value = arr(12)
  Fields("JobClass").Value = arr(13)
  Fields("MemberID").Value = arr(14)
  Fields("OTPay").Value = arr(15)
  Fields("DtlStartDate").Value = FormatThisPayPd(arr(16), 2)
  Fields("DtlEndDate").Value = FormatThisPayPd(arr(17), 2)
  Fields("PlanCode").Value = arr(18)
  Fields("Salary").Value = arr(19)
  Fields("EmpNum").Value = arr(20)
  ReDim PL1(1 To 28) As String
  ReDim PL2(1 To 28) As String
  ReDim PL3(1 To 28) As String
  ReDim PL4(1 To 28) As String
  ArrCnt = 0
  For x = 1 To 28
    ArrCnt = ArrCnt + 1
    PL1(x) = arr(ArrCnt + 20)
    ArrCnt = ArrCnt + 1
    PL2(x) = arr(ArrCnt + 20)
    ArrCnt = ArrCnt + 1
    PL3(x) = arr(ArrCnt + 20)
    ArrCnt = ArrCnt + 1
    PL4(x) = arr(ArrCnt + 20)
    Fields("PL1" & CStr(x)).Value = PL1(x)
    Fields("PL2" & CStr(x)).Value = PL2(x)
    Fields("PL3" & CStr(x)).Value = PL3(x)
    Fields("PL4" & CStr(x)).Value = PL4(x)
  Next x
  Fields("TJCPay").Value = arr(133)
  Fields("TJCCnt").Value = arr(134)
  Fields("TJCEmpCont").Value = arr(135)
  Fields("TJCCityCont").Value = arr(136)
  Fields("TPCPay").Value = arr(137)
  Fields("TPCCnt").Value = arr(138)
  Fields("TPCEmpCont").Value = arr(139)
  Fields("TPCCityCont").Value = arr(140)
  Fields("PayType").Value = arr(141)
  Fields("Adjustment").Value = arr(142)
  
  If QPTrim$(arr(10)) <> "" Then
    Fields("Name").Value = QPTrim$(arr(7)) & ", " & QPTrim$(arr(8)) & " " & QPTrim$(arr(9)) & ", " & QPTrim$(arr(10))
  Else
    Fields("Name").Value = QPTrim$(arr(7)) & ", " & QPTrim$(arr(8)) & " " & QPTrim$(arr(9))
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
  lblSummary.Visible = False
  Me.fldTimeDate.Text = Now
  Me.Zoom = -1
End Sub

Private Function FormatThisPayPd(ByRef ThisDate As String, ByVal Vers As Integer) As String
  Dim ch As String
  Dim DateLen As Integer
  Dim x As Integer
  Dim ThisDay As String
  Dim ThisMonth As String
  Dim ThisYear As String
  
  
  ThisMonth = Mid(ThisDate, 5, 2)
  ThisDay = Mid(ThisDate, 7, 2)
  ThisYear = Mid(ThisDate, 1, 4)
  If Vers = 2 Then
    ThisDate = ThisMonth & "/" & ThisDay & "/" & ThisYear
  ElseIf Vers = 1 Then
    ThisDate = ThisMonth & "/" & ThisYear
  End If
  FormatThisPayPd = ThisDate
  
End Function

Private Sub GroupFooter1_Format()
  Dim x As Integer
  GroupFooter1.Height = 1050
  Field16.Visible = False
  Field44.Visible = False
  Field45.Visible = False
  Field46.Visible = False
  
  Field17.Visible = False
  Field50.Visible = False
  Field51.Visible = False
  Field52.Visible = False
  
  Field18.Visible = False
  Field47.Visible = False
  Field48.Visible = False
  Field49.Visible = False

  Field19.Visible = False
  Field56.Visible = False
  Field57.Visible = False
  Field58.Visible = False

  Field20.Visible = False
  Field59.Visible = False
  Field60.Visible = True
  Field61.Visible = True
  
  Field21.Visible = False
  Field62.Visible = False
  Field63.Visible = False
  Field64.Visible = False
  
  Field22.Visible = False
  Field65.Visible = False
  Field66.Visible = False
  Field67.Visible = False
  
  Field23.Visible = False
  Field68.Visible = False
  Field69.Visible = False
  Field70.Visible = False
  
  Field24.Visible = False
  Field71.Visible = False
  Field72.Visible = False
  Field73.Visible = False
  
  Field25.Visible = False
  Field74.Visible = False
  Field75.Visible = False
  Field76.Visible = False
  
  Field26.Visible = False
  Field77.Visible = False
  Field78.Visible = False
  Field79.Visible = False
  
  Field27.Visible = False
  Field80.Visible = False
  Field81.Visible = False
  Field82.Visible = False
  
  Field28.Visible = False
  Field83.Visible = False
  Field84.Visible = False
  Field85.Visible = False
  
  Field29.Visible = False
  Field86.Visible = False
  Field87.Visible = False
  Field88.Visible = False
  
  Field30.Visible = False
  Field89.Visible = False
  Field90.Visible = False
  Field91.Visible = False
  
  Field31.Visible = False
  Field92.Visible = False
  Field93.Visible = False
  Field94.Visible = False
  
  Field32.Visible = False
  Field95.Visible = False
  Field96.Visible = False
  Field97.Visible = False
  
  Field33.Visible = False
  Field98.Visible = False
  Field99.Visible = False
  Field100.Visible = False
  
  Field34.Visible = False
  Field101.Visible = False
  Field102.Visible = False
  Field103.Visible = False
  
  Field35.Visible = False
  Field104.Visible = False
  Field105.Visible = False
  Field106.Visible = False
  
  Field36.Visible = False
  Field107.Visible = False
  Field108.Visible = False
  Field109.Visible = False
  
  Field37.Visible = False
  Field110.Visible = False
  Field111.Visible = False
  Field112.Visible = False
  
  Field38.Visible = False
  Field113.Visible = False
  Field114.Visible = False
  Field115.Visible = False
  
  Field39.Visible = False
  Field116.Visible = False
  Field117.Visible = False
  Field118.Visible = False
  
  Field40.Visible = False
  Field119.Visible = False
  Field120.Visible = False
  Field121.Visible = False
  
  Field41.Visible = False
  Field122.Visible = False
  Field123.Visible = False
  Field124.Visible = False
  
  Field42.Visible = False
  Field125.Visible = False
  Field126.Visible = False
  Field127.Visible = False
  
  Field43.Visible = False
  Field128.Visible = False
  Field129.Visible = False
  Field130.Visible = False
  If GStart <= 1 Then Exit Sub
  For x = 1 To 28
    If QPTrim$(PL21(x)) <> "" Then
      Select Case x
        Case 1
          GroupFooter1.Height = GroupFooter1.Height + 270
          Field16.Visible = True
          Field44.Visible = True
          Field45.Visible = True
          Field46.Visible = True
        Case 2
          Field17.Top = GroupFooter1.Height
          Field50.Top = GroupFooter1.Height
          Field51.Top = GroupFooter1.Height
          Field52.Top = GroupFooter1.Height
          Field17.Visible = True
          Field50.Visible = True
          Field51.Visible = True
          Field52.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 3
          Field18.Top = GroupFooter1.Height
          Field47.Top = GroupFooter1.Height
          Field48.Top = GroupFooter1.Height
          Field49.Top = GroupFooter1.Height
          Field18.Visible = True
          Field47.Visible = True
          Field48.Visible = True
          Field49.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 4
          Field19.Top = GroupFooter1.Height
          Field56.Top = GroupFooter1.Height
          Field57.Top = GroupFooter1.Height
          Field58.Top = GroupFooter1.Height
          Field19.Visible = True
          Field56.Visible = True
          Field57.Visible = True
          Field58.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 5
          Field20.Top = GroupFooter1.Height
          Field59.Top = GroupFooter1.Height
          Field60.Top = GroupFooter1.Height
          Field51.Top = GroupFooter1.Height
          Field20.Visible = True
          Field59.Visible = True
          Field60.Visible = True
          Field61.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 6
          Field21.Top = GroupFooter1.Height
          Field62.Top = GroupFooter1.Height
          Field63.Top = GroupFooter1.Height
          Field64.Top = GroupFooter1.Height
          Field21.Visible = True
          Field62.Visible = True
          Field63.Visible = True
          Field64.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 7
          Field22.Top = GroupFooter1.Height
          Field65.Top = GroupFooter1.Height
          Field66.Top = GroupFooter1.Height
          Field67.Top = GroupFooter1.Height
          Field22.Visible = True
          Field65.Visible = True
          Field66.Visible = True
          Field67.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 8
          Field23.Top = GroupFooter1.Height
          Field68.Top = GroupFooter1.Height
          Field69.Top = GroupFooter1.Height
          Field70.Top = GroupFooter1.Height
          Field23.Visible = True
          Field68.Visible = True
          Field69.Visible = True
          Field70.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 9
          Field24.Top = GroupFooter1.Height
          Field71.Top = GroupFooter1.Height
          Field72.Top = GroupFooter1.Height
          Field73.Top = GroupFooter1.Height
          Field24.Visible = True
          Field71.Visible = True
          Field72.Visible = True
          Field73.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 10
          Field25.Top = GroupFooter1.Height
          Field74.Top = GroupFooter1.Height
          Field75.Top = GroupFooter1.Height
          Field76.Top = GroupFooter1.Height
          Field25.Visible = True
          Field74.Visible = True
          Field75.Visible = True
          Field76.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 11
          Field26.Top = GroupFooter1.Height
          Field77.Top = GroupFooter1.Height
          Field78.Top = GroupFooter1.Height
          Field79.Top = GroupFooter1.Height
          Field26.Visible = True
          Field77.Visible = True
          Field78.Visible = True
          Field79.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 12
          Field27.Top = GroupFooter1.Height
          Field80.Top = GroupFooter1.Height
          Field81.Top = GroupFooter1.Height
          Field82.Top = GroupFooter1.Height
          Field27.Visible = True
          Field80.Visible = True
          Field81.Visible = True
          Field82.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 13
          Field28.Top = GroupFooter1.Height
          Field83.Top = GroupFooter1.Height
          Field84.Top = GroupFooter1.Height
          Field85.Top = GroupFooter1.Height
          Field28.Visible = True
          Field83.Visible = True
          Field84.Visible = True
          Field85.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 14
          Field29.Top = GroupFooter1.Height
          Field86.Top = GroupFooter1.Height
          Field87.Top = GroupFooter1.Height
          Field88.Top = GroupFooter1.Height
          Field29.Visible = True
          Field86.Visible = True
          Field87.Visible = True
          Field88.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 15
          Field30.Top = GroupFooter1.Height
          Field89.Top = GroupFooter1.Height
          Field90.Top = GroupFooter1.Height
          Field91.Top = GroupFooter1.Height
          Field30.Visible = True
          Field89.Visible = True
          Field90.Visible = True
          Field91.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 16
          Field31.Top = GroupFooter1.Height
          Field92.Top = GroupFooter1.Height
          Field93.Top = GroupFooter1.Height
          Field94.Top = GroupFooter1.Height
          Field31.Visible = True
          Field92.Visible = True
          Field93.Visible = True
          Field94.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 17
          Field32.Top = GroupFooter1.Height
          Field95.Top = GroupFooter1.Height
          Field96.Top = GroupFooter1.Height
          Field97.Top = GroupFooter1.Height
          Field32.Visible = True
          Field95.Visible = True
          Field96.Visible = True
          Field97.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 18
          Field33.Top = GroupFooter1.Height
          Field98.Top = GroupFooter1.Height
          Field99.Top = GroupFooter1.Height
          Field100.Top = GroupFooter1.Height
          Field33.Visible = True
          Field98.Visible = True
          Field99.Visible = True
          Field100.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 19
          Field34.Top = GroupFooter1.Height
          Field101.Top = GroupFooter1.Height
          Field102.Top = GroupFooter1.Height
          Field103.Top = GroupFooter1.Height
          Field34.Visible = True
          Field101.Visible = True
          Field102.Visible = True
          Field103.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 20
          Field35.Top = GroupFooter1.Height
          Field104.Top = GroupFooter1.Height
          Field105.Top = GroupFooter1.Height
          Field106.Top = GroupFooter1.Height
          Field35.Visible = True
          Field104.Visible = True
          Field105.Visible = True
          Field106.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 21
          Field36.Top = GroupFooter1.Height
          Field107.Top = GroupFooter1.Height
          Field108.Top = GroupFooter1.Height
          Field109.Top = GroupFooter1.Height
          Field36.Visible = True
          Field107.Visible = True
          Field108.Visible = True
          Field109.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 22
          Field37.Top = GroupFooter1.Height
          Field110.Top = GroupFooter1.Height
          Field111.Top = GroupFooter1.Height
          Field112.Top = GroupFooter1.Height
          Field37.Visible = True
          Field110.Visible = True
          Field111.Visible = True
          Field112.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 23
          Field38.Top = GroupFooter1.Height
          Field113.Top = GroupFooter1.Height
          Field114.Top = GroupFooter1.Height
          Field115.Top = GroupFooter1.Height
          Field38.Visible = True
          Field113.Visible = True
          Field114.Visible = True
          Field115.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 24
          Field39.Top = GroupFooter1.Height
          Field116.Top = GroupFooter1.Height
          Field117.Top = GroupFooter1.Height
          Field118.Top = GroupFooter1.Height
          Field39.Visible = True
          Field116.Visible = True
          Field117.Visible = True
          Field118.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 25
          Field40.Top = GroupFooter1.Height
          Field119.Top = GroupFooter1.Height
          Field120.Top = GroupFooter1.Height
          Field121.Top = GroupFooter1.Height
          Field40.Visible = True
          Field119.Visible = True
          Field120.Visible = True
          Field121.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 26
          Field41.Top = GroupFooter1.Height
          Field122.Top = GroupFooter1.Height
          Field123.Top = GroupFooter1.Height
          Field124.Top = GroupFooter1.Height
          Field41.Visible = True
          Field122.Visible = True
          Field123.Visible = True
          Field124.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 27
          Field42.Top = GroupFooter1.Height
          Field125.Top = GroupFooter1.Height
          Field126.Top = GroupFooter1.Height
          Field127.Top = GroupFooter1.Height
          Field42.Visible = True
          Field125.Visible = True
          Field126.Visible = True
          Field127.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case 28
          Field43.Top = GroupFooter1.Height
          Field128.Top = GroupFooter1.Height
          Field129.Top = GroupFooter1.Height
          Field130.Top = GroupFooter1.Height
          Field43.Visible = True
          Field128.Visible = True
          Field129.Visible = True
          Field130.Visible = True
          GroupFooter1.Height = GroupFooter1.Height + 270
        Case Else
      End Select
    End If
  Next x
  GroupFooter1.Height = GroupFooter1.Height + 450
  ReDim PL1(1 To 28) As String
  ReDim PL2(1 To 28) As String
  ReDim PL3(1 To 28) As String
  ReDim PL4(1 To 28) As String
End Sub

Private Sub ReportFooter_Format()
  lblSummary.Visible = True
  Set SubReport1.object = New arORBITSub
  Set SubReport2.object = New arORBITSub2
End Sub
