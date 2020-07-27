VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arLvBnfts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leave Benefits Earned"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arLvBnfts.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arLvBnfts.dsx":08CA
End
Attribute VB_Name = "arLvBnfts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private HFile As Integer
Dim table$
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
      MsgBox "File - AccrueBenefitRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - AccrueBenefitRpt.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - AccrueBenefitRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - AccrueBenefitRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "AccrueBenefitRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "AccrueBenefitRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  HFile = FreeFile
  Open StartPath & "\PRRPTS\ACCRUALG.RPT" For Input As #HFile
  Fields.Add "fldEmployer" '(0)
  Fields.Add "fldDate" '(1)
  Fields.Add "fldEmpNum" '(2)
  Fields.Add "fldEmployee" '(3)
  Fields.Add "fldTable" '(4)
  Fields.Add "fldYears" '(5)
  Fields.Add "fldBenePct" '(6)
  Fields.Add "fldVac" '(7)
  Fields.Add "fldSick" '(8)
  Fields.Add "fldStarS" '(9)
  Fields.Add "fldStarV" '(10)
  Fields.Add "fldHol" '(7)
  Fields.Add "fldPer" '(8)
  Fields.Add "fldStarH" '(9)
  Fields.Add "fldStarP" '(10)
  
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
  Fields("fldEmpNum").Value = arr(2)
  Fields("fldEmployee").Value = arr(3)
  Fields("fldTable").Value = arr(4)
  Fields("fldYears").Value = arr(5)
  Fields("fldBenePct").Value = arr(6)
  Fields("fldVac").Value = arr(7)
  Fields("fldSick").Value = arr(8)
  Fields("fldStarS").Value = arr(9)
  Fields("fldStarV").Value = arr(10)
  Fields("fldHol").Value = arr(11)
  Fields("fldPer").Value = arr(12)
  Fields("fldStarH").Value = arr(13)
  Fields("fldStarP").Value = arr(14)
  table = arr(4)
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
  Label31.Visible = False
  Label32.Visible = False
  Label33.Visible = False
End Sub

Private Sub Detail_Format()
  If table = -1 Then
'    lblError.Caption = "Invalid hire date."
    lblError.Visible = True
    fldTable.Visible = False
    fldYears.Visible = False
    fldBenePct.Visible = False
  Else
    lblError.Visible = False
    fldTable.Visible = True
    fldYears.Visible = True
    fldBenePct.Visible = True
  End If
    
End Sub

Private Sub ReportFooter_Format()
  Label31.Visible = True
  Label32.Visible = True
  Label33.Visible = True
  Label4.Visible = False
  Label5.Visible = False
  Label6.Visible = False
  Label7.Visible = False
  Label8.Visible = False
  Label29.Visible = False

End Sub
