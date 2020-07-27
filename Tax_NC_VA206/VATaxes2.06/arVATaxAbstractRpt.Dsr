VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxAbstractRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Abstract Report"
   ClientHeight    =   8736
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arVATaxAbstractRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arVATaxAbstractRpt.dsx":08CA
End
Attribute VB_Name = "arVATaxAbstractRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private hFile As Integer
  Private Temp_Class As Resize_Class
Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\TAXRPTS\ABSTLIST.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldCustName") '1)
  Fields.Add ("fldCustAcct") '2)
  Fields.Add ("fldCustAdd1") '3)
  Fields.Add ("fldCustAdd2") '4)
  Fields.Add ("fldCSZ") '5)
  Fields.Add ("fldPropType") '6)
  Fields.Add ("fldRealPin") '7)
  Fields.Add ("fldRealVal") '8)
  Fields.Add ("fldRealAdd") '9)
  Fields.Add ("fldRealMBL") '10)
  Fields.Add ("fldRealDesc1") '11)
  Fields.Add ("fldRealDesc2") '12)
  Fields.Add ("fldRealDesc3") '13)
  Fields.Add ("fldPersPVal") '14)
  Fields.Add ("fldPersCVal") '15)
  Fields.Add ("fldPersMHVal") '16)
  Fields.Add ("fldPersMTVal") '17)
  Fields.Add ("fldPersMCVal") '18)
  Fields.Add ("fldPersDesc1") '19)
  Fields.Add ("fldPersDesc2") '20)
  Fields.Add ("fldPersDesc3") '21)
  Fields.Add ("fldPersDesc4") '22)
  Fields.Add ("fldPersDesc5") '23)
End Sub
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmVATaxLoadReport
    frmVATaxMsg.Label1.Caption = "Error Number: " & Str(Number) & " " & Description
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
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
  Fields("fldTown").Value = arr(0)
  Fields("fldCustName").Value = arr(1)
  Fields("fldCustAcct").Value = arr(2)
  Fields("fldCustAdd1").Value = arr(3)
  Fields("fldCustAdd2").Value = arr(4)
  Fields("fldCSZ").Value = arr(5)
  Fields("fldPropType").Value = arr(6)
  Fields("fldRealPin").Value = arr(7)
  Fields("fldRealVal").Value = arr(8)
  Fields("fldRealAdd").Value = arr(9)
  Fields("fldRealMBL").Value = arr(10)
  Fields("fldRealDesc1").Value = arr(11)
  Fields("fldRealDesc2").Value = arr(12)
  Fields("fldRealDesc3").Value = arr(13)
  Fields("fldPersPVal").Value = arr(14)
  Fields("fldPersCVal").Value = arr(15)
  Fields("fldPersMHVal").Value = arr(16)
  Fields("fldPersMTVal").Value = arr(17)
  Fields("fldPersMCVal").Value = arr(18)
  Fields("fldPersDesc1").Value = arr(19)
  Fields("fldPersDesc2").Value = arr(20)
  Fields("fldPersDesc3").Value = arr(21)
  Fields("fldPersDesc4").Value = arr(22)
  Fields("fldPersDesc5").Value = arr(23)
  If QPTrim$(arr(11)) = "" And QPTrim$(arr(13)) = "" And QPTrim$(arr(13)) = "" Then
    Fields("fldRealDesc1").Value = "NO DESCRIPTION AVAILABLE"
  End If
End Sub

Private Sub ActiveReport_Initialize()
  Me.Toolbar.Tools.Add "&Close"
  Me.Toolbar.Tools.Add "Save/&Excel"
  Me.Toolbar.Tools.Add "&Text"
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
      frmVATaxMsg.Label1.Caption = "File - AbstractRpt.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - AbstractRpt.txt, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close
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
    frmVATaxMsg.Label1.Caption = "File - AbstractRpt.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - AbstractRpt.txt, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
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
        oEXL.FileName = outfile & "AbstractRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "AbstractRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmVATaxLoadReport
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
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Me.fldTimeDate.Text = Date
  Me.Zoom = -1
End Sub

Private Sub Detail_Format()
  Dim Real1 As Boolean
  Dim Real2 As Boolean
  Dim Real3 As Boolean
  Dim Pers1 As Boolean
  Dim Pers2 As Boolean
  Dim Pers3 As Boolean
  Dim Pers4 As Boolean
  Dim Pers5 As Boolean
   
  Real1 = False
  Real2 = False
  Real3 = False
  Pers1 = False
  Pers2 = False
  Pers3 = False
  Pers4 = False
  Pers5 = False
  
  If QPTrim$(Fields("fldRealDesc1").Value) <> "" Then
    Real1 = True
  End If
  
  If QPTrim$(Fields("fldRealDesc2").Value) <> "" Then
    Real2 = True
  End If
  
  If QPTrim$(Fields("fldRealDesc3").Value) <> "" Then
    Real3 = True
  End If
  
  If QPTrim$(Fields("fldPersDesc1").Value) <> "" Then
    Pers1 = True
  End If
  
  If QPTrim$(Fields("fldPersDesc2").Value) <> "" Then
    Pers2 = True
  End If
  
  If QPTrim$(Fields("fldPersDesc3").Value) <> "" Then
    Pers3 = True
  End If
  
  If QPTrim$(Fields("fldPersDesc4").Value) <> "" Then
    Pers4 = True
  End If
  
  If QPTrim$(Fields("fldPersDesc5").Value) <> "" Then
    Pers5 = True
  End If
  
  Label80.Visible = True
  Field7.Visible = True
  Label81.Visible = True
  Field8.Visible = True
  Label82.Visible = True
  Field9.Visible = True
  Label83.Visible = True
  Field10.Visible = True
  Label84.Visible = True
  Field11.Visible = True
  Field12.Visible = True
  Field13.Visible = True
  Label85.Visible = True
  Field14.Visible = True
  Label91.Visible = True
  Label86.Visible = True
  Field15.Visible = True
  Field20.Visible = True
  Label87.Visible = True
  Field16.Visible = True
  Field21.Visible = True
  Label88.Visible = True
  Field17.Visible = True
  Field22.Visible = True
  Label89.Visible = True
  Field18.Visible = True
  Field23.Visible = True
  Label90.Visible = True
  Field19.Visible = True
  Field24.Visible = True
  
  If Fields("fldPropType").Value = "REAL" Then
    Label85.Visible = False
    Field14.Visible = False
    Label91.Visible = False
    Label86.Visible = False
    Field15.Visible = False
    Field20.Visible = False
    Label87.Visible = False
    Field16.Visible = False
    Field21.Visible = False
    Label88.Visible = False
    Field17.Visible = False
    Field22.Visible = False
    Label89.Visible = False
    Field18.Visible = False
    Field23.Visible = False
    Label90.Visible = False
    Field19.Visible = False
    Field24.Visible = False
    Detail.Height = 1395
    If Real1 = False And Real2 = False And Real3 = False Then
      Fields("fldRealDesc1").Value = "NO DESCRIPTIONS AVAILABLE"
      Field12.Visible = False
      Field13.Visible = False
      Detail.Height = 810
    ElseIf Real1 = False And Real2 = True And Real3 = True Then
      Field11.Visible = False
      Field12.Top = 540
      Field13.Top = 810
      Detail.Height = 1080
    ElseIf Real1 = False And Real2 = False And Real3 = True Then
      Field11.Visible = False
      Field12.Visible = False
      Field13.Top = 540
      Detail.Height = 810
    ElseIf Real1 = True And Real2 = True And Real3 = False Then
      Field13.Visible = False
      Detail.Height = 1080
    ElseIf Real1 = True And Real2 = False And Real3 = True Then
      Field12.Visible = False
      Field13.Top = 810
      Detail.Height = 1080
    ElseIf Real1 = False And Real2 = True And Real3 = False Then
      Field11.Visible = False
      Field13.Visible = False
      Field12.Top = 540
      Detail.Height = 810
    ElseIf Real1 = False And Real2 = False And Real3 = True Then
      Field11.Visible = False
      Field12.Visible = False
      Field13.Top = 540
      Detail.Height = 810
    ElseIf Real1 = True And Real2 = False And Real3 = False Then
      Field12.Visible = False
      Field13.Visible = False
      Detail.Height = 810
    End If
  ElseIf Fields("fldPropType").Value = "PERSONAL" Then
    Label80.Visible = False
    Field7.Visible = False
    Label81.Visible = False
    Field8.Visible = False
    Label82.Visible = False
    Field9.Visible = False
    Label83.Visible = False
    Field10.Visible = False
    Label84.Visible = False
    Field11.Visible = False
    Field12.Visible = False
    Field13.Visible = False
    Label85.Top = 0
    Field14.Top = 0
    Label91.Top = 0
    Label86.Top = 270
    Field15.Top = 270
    Field20.Top = 270
    Label87.Top = 540
    Field16.Top = 540
    Field21.Top = 540
    Label88.Top = 810
    Field17.Top = 810
    Field22.Top = 810
    Label89.Top = 1080
    Field18.Top = 1080
    Field23.Top = 1080
    Label90.Top = 1350
    Field19.Top = 1350
    Field24.Top = 1350
    Detail.Height = 1620
  End If
    
    
End Sub
