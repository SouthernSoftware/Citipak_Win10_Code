VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arTaxRealClassRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Real Property Classification Report"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "arTaxRealClassRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15452
   SectionData     =   "arTaxRealClassRpt.dsx":08CA
End
Attribute VB_Name = "arTaxRealClassRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private hFile As Integer
  Private Temp_Class As Resize_Class
  Dim RepFooter As Boolean
Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\TAXRPTS\REALCLAS.RPT" For Input As #hFile
  Fields.Add ("fldTown") '(0)
  Fields.Add ("fldCustName") '(1)
  Fields.Add ("fldCustAcct") '(2)
  Fields.Add ("fldRptClassType") '(3)
  Fields.Add ("fldPropVal") '(4)
  Fields.Add ("fldPropDisc") '(5)
  Fields.Add ("fldPropPin") '(6)
  Fields.Add ("fldDesc") '(7)
  Fields.Add ("fldPropNet") '(8)
  Fields.Add ("fldRptDesc") '(9)
  Fields.Add ("fldClassType") '(10)
  Fields.Add ("fldRealTownship") '(11)
  Fields.Add ("fldTownship") '(12)
  Fields.Add ("fldGOpt") '(13)
  Fields.Add ("fldOptDesc") '(14)
  Fields.Add ("fldThisClass") '(15)
  Fields.Add ("fldTownshipYN") '16)
  Fields.Add ("fldInactiveFlag") '17)
End Sub
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then '.Value = arr(ignore the no printer warning
    Unload frmTaxLoadReport
    frmTaxMsg.Label1.Caption = "Error Number: " & Str(Number) & " " & Description
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Unload Me
  End If
  CancelDisplay = True '.Value = arr(removes the error message
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
' .Value = arr( Here we set the values of the fields that we defines as unbound
' .Value = arr( or user defined.
  Fields("fldTown").Value = arr(0)
  Fields("fldCustName").Value = arr(1)
  Fields("fldCustAcct").Value = arr(2)
  Fields("fldRptClassType").Value = arr(3)
  Fields("fldPropVal").Value = arr(4)
  Fields("fldPropDisc").Value = arr(5)
  Fields("fldPropPin").Value = arr(6)
  Fields("fldDesc").Value = arr(7)
  Fields("fldPropNet").Value = arr(8)
  Fields("fldRptDesc").Value = arr(9)
  Fields("fldClassType").Value = arr(10)
  Fields("fldRealTownship").Value = arr(11)
  Fields("fldTownship").Value = arr(12)
  Fields("fldGOpt").Value = arr(13)
  Fields("fldOptDesc").Value = arr(14)
  Fields("fldThisClass").Value = arr(15)
  Fields("fldTownshipYN").Value = arr(16)
  Fields("fldInactiveFlag").Value = arr(17)
End Sub

Private Sub ActiveReport_Initialize()
  Me.ToolBar.Tools.Add "&Close"
  Me.ToolBar.Tools.Add "Save/&Excel"
  Me.ToolBar.Tools.Add "&Text"
  RepFooter = False
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
      frmTaxMsg.Label1.Caption = "File - RealClass.xls, created in the Citipak Directory."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmTaxMsg.Label1.Caption = "File - RealClass.txt, created in the Citipak Directory."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
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
    frmTaxMsg.Label1.Caption = "File - RealClass.xls, created in the Citipak Directory."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmTaxMsg.Label1.Caption = "File - RealClass.txt, created in the Citipak Directory."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
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
    Case 1 ' .Value' = arr'("Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & "RealClass.xls"
        oEXL.Export Me.Pages
    Case 2 ' .Value = arr'("Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "RealClass.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmTaxLoadingRpt
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
  If QPTrim$(Fields("fldCustAcct").Value) = "" Then
    Detail.Visible = False
  Else
    Detail.Visible = True
  End If
End Sub

Private Sub GroupHeader1_Format()
  If Fields("fldOptDesc").Value <> "" Then
    GroupHeader1.Height = 540
    Line11.Y1 = 540
    Line11.Y2 = 540
    Field54.Visible = True
    Field55.Visible = True
  Else
    GroupHeader1.Height = 270
    Line11.Y1 = 270
    Line11.Y2 = 270
    Field54.Visible = False
    Field55.Visible = False
  End If
  If QPTrim$(Fields("fldCustAcct").Value) = "" Then
    GroupHeader1.Visible = False
    Line13.Visible = False
  Else
    GroupHeader1.Visible = True
    Line13.Visible = True
  End If

End Sub

Private Sub PageHeader_Format()
 If Fields("fldInactiveFlag").Value = True Then
   Label102.Visible = True
 Else
   Label102.Visible = False
 End If
 If RepFooter = True Then
  PageHeader.Height = 1500
  Line7.Y1 = 1480
  Line7.Y2 = 1480
 End If
 If QPTrim$(Fields("fldRptDesc").Value) = "Address of Property" Then
   Label76.Caption = "Address of Property"
 Else
   Label76.Caption = "First Line of Notes"
 End If
End Sub

Private Sub ReportFooter_Format()
  If Fields("fldTownshipYN").Value = True Then
    Set SubReport1 = New arSub1RealClassRpt
  Else
    Label90.Visible = False
    Label82.Visible = False
    Label87.Visible = False
    Label88.Visible = False
    Label89.Visible = False
    Label96.Visible = False
    Line14.Visible = False
  End If
  Set SubReport2 = New arSub2RealClassRpt

  RepFooter = True
  Label102.Visible = False
  Label62.Visible = False
  Label63.Visible = False
  Label73.Visible = False
  Label80.Visible = False
  Label76.Visible = False
  Label74.Visible = False
  Label78.Visible = False
  Label79.Visible = False
  Label75.Visible = False
End Sub

