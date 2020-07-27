VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arFAAssetsByFundRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets By Fund Report"
   ClientHeight    =   8868
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12000
   Icon            =   "arFAAssetsByFundRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   21140
   _ExtentY        =   15637
   SectionData     =   "arFAAssetsByFundRpt.dsx":08CA
End
Attribute VB_Name = "arFAAssetsByFundRpt"
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
      MsgBox "File - FAAssetsByFundRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - FAAssetsByFundRpt.txt, created in the Citipak Directory.", vbOKOnly
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
    MsgBox "File - FAAssetsByFundRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - FAAssetsByFundRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "FAAssetsByFundRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "FAAssetsByFundRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  HFile = FreeFile
  Open StartPath & "\FARPTS\FABYFUND.RPT" For Input As #HFile
  Fields.Add ("fldEmployer") '0)
  Fields.Add ("fldTagNum") '1)
  Fields.Add ("fldItemDesc") '2)
  Fields.Add ("fldItemFund") '3)
  Fields.Add ("fldAssetLife") '4)
  Fields.Add ("fldOrigCost") '5)
  Fields.Add ("fldDepr") '6)
  Fields.Add ("fldStar") '7)
  Fields.Add ("fldBookVal") '8)
  Fields.Add ("fldFund") '9)
  Fields.Add ("fldFundDesc1") '10)
  Fields.Add ("fldThisFundNum") '11)
  Fields.Add ("fldFundPurchPr") '12)
  Fields.Add ("fldFundDpr") '13)
  Fields.Add ("fldFundBookVal") '14)
  Fields.Add ("fldFundCnt") '15)
  Fields.Add ("fldGTPurchPr") '16)
  Fields.Add ("fldGTDpr") '17)
  Fields.Add ("fldGTBookVal") '18)
  Fields.Add ("fldTotalCnt") '19)
  Fields.Add ("fldDEPYN") '20)
  Fields.Add ("fldLifeLeft") '21)
  Fields.Add ("fldDispDate") '22)
  Fields.Add ("fldShowDisp") '23)
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
  Fields("fldItemFund").Value = arr(3)
  Fields("fldAssetLife").Value = arr(4)
  Fields("fldOrigCost").Value = arr(5)
  Fields("fldDepr").Value = arr(6)
  Fields("fldStar").Value = arr(7)
  Fields("fldBookVal").Value = arr(8)
  Fields("fldFund").Value = arr(9)
  Fields("fldFundDesc1").Value = arr(10)
  Fields("fldThisFundNum").Value = arr(11)
  Fields("fldFundPurchPr").Value = arr(12)
  Fields("fldFundDpr").Value = arr(13)
  Fields("fldFundBookVal").Value = arr(14)
  Fields("fldFundCnt").Value = arr(15)
  Fields("fldGTPurchPr").Value = arr(16)
  Fields("fldGTDpr").Value = arr(17)
  Fields("fldGTBookVal").Value = arr(18)
  Fields("fldTotalCnt").Value = arr(19)
  Fields("fldDEPYN").Value = arr(20)
  Fields("fldLifeLeft").Value = arr(21)
  Fields("fldDispDate").Value = arr(22)
  Fields("fldShowDisp").Value = arr(23)
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
  Label11.Visible = False
  If Fields("fldShowDisp").Value = "N" Then
    Label40.Visible = False
    Line1.X2 = 10890
    Line5.X2 = 10890
    Line4.X2 = 10890
  End If
End Sub

Private Sub GroupFooter1_Format()
'  If Fields("fldFund").Value = "ALL" Then
'    Fund.DataValue = "ALL"
'    Field6.Visible = False
'  End If
End Sub

Private Sub GroupHeader1_Format()
'  If Fields("fldFund").Value = "ALL" Then
'    fldFundDesc1.Visible = False
'    fldFund.DataValue = "All"
'  End If
End Sub

Private Sub ReportFooter_Format()
  If Fields("fldFund").Value <> "ALL" Then
    Me.ReportFooter.Visible = False
    Exit Sub
  End If
  Set SubReport1.object = New arFASubTagByFund
  PageHeader.Height = 1000
  Label40.Visible = False
  Label41.Visible = False
  Label11.Caption = "Summary"
  Line1.Visible = False
  Label11.Visible = True
  Label16.Visible = False
  Label17.Visible = False
  Label18.Visible = False
  Label19.Visible = False
  Label20.Visible = False
  Label21.Visible = False
  Label22.Visible = False
  Label24.Visible = False
  
End Sub




