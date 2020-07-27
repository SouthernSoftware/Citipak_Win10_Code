VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptBillPayTax 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Bill/Payment  Report"
   ClientHeight    =   4380
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   8952
   Icon            =   "ARptBillPayTax.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   15790
   _ExtentY        =   7726
   SectionData     =   "ARptBillPayTax.dsx":08CA
End
Attribute VB_Name = "ARptBillPayTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Public SubFile As String
Private hFile As Integer
Dim rpt As ActiveReport
Dim cnt As Integer
Dim headers(1 To 34) As String
Dim Det As Boolean
Dim Rev As Integer

Public Sub GetName(RName As String, SName As String, Detail As Boolean, RevSource As Integer)
  ReportFile$ = RName$
  SubFile$ = SName$
  Det = Detail
  Rev = RevSource
End Sub

Private Sub ActiveReport_DataInitialize()
    headers(1) = "Date"
    headers(2) = "AcctNo"
    headers(3) = "CustName"
    headers(4) = "TaxAmt"
    headers(5) = "Rev1N"
    headers(6) = "Rev1A"
    headers(7) = "Rev2N"
    headers(8) = "Rev2A"
    headers(9) = "Rev3N"
    headers(10) = "Rev3A"
    headers(11) = "Rev4N"
    headers(12) = "Rev4A"
    headers(13) = "Rev5N"
    headers(14) = "Rev5A"
    headers(15) = "Rev6N"
    headers(16) = "Rev6A"
    headers(17) = "Rev7N"
    headers(18) = "Rev7A"
    headers(19) = "Rev8N"
    headers(20) = "Rev8A"
    headers(21) = "Rev9N"
    headers(22) = "Rev9A"
    headers(23) = "Rev10N"
    headers(24) = "Rev10A"
    headers(25) = "Rev11N"
    headers(26) = "Rev11A"
    headers(27) = "Rev12N"
    headers(28) = "Rev12A"
    headers(29) = "Rev13N"
    headers(30) = "Rev13A"
    headers(31) = "Rev14N"
    headers(32) = "Rev14A"
    headers(33) = "Rev15N"
    headers(34) = "Rev15A"

    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    For cnt = 1 To 34
      Fields.Add headers(cnt)
    Next

End Sub
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmLoadingRpt
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sLine As String
Dim arr() As String
'On Error GoTo ERRORSTUFF
'    ' We reached the end of the file we exit leaving the
'    ' eof parameter as True (default except on first call) that will
'    ' tell AR that we are done feeding data
'    ' otherwise we have to set the eof parameter to False so that
'    ' AR continues fetching data, until we're done
'    ' if the report had a data control, the value of the parameter
'    ' will be ignored, AR will always follow the data control's recordset
'    ' EOF property
    If VBA.eof(hFile) Then
        eof = True
        Exit Sub
    Else
        eof = False
    End If
    frmLoadingRpt.ShowHowMuch
    Line Input #hFile, sLine
    arr = Split(sLine, "~")
'    ' Here we set the values of the fields that we defines as unbound
'    ' or user defined.
    For cnt = 1 To 34
      Fields(headers(cnt)) = arr(cnt - 1)
    Next
'If something wrong in file give message instead of crashing
Exit Sub
ERRORSTUFF:
      Unload frmLoadingRpt
'  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "ARptVendHist", "Fetch Data", Erl)
'    Case emrExitProc:
'      Resume Proc_Exit
'    Case emrResume:
'      Resume
'    Case emrResumeNext:
'      Resume Next
'    Case Else
'      '--- Technically, this should never happen.
'      Resume Proc_Exit
'  End Select
   MsgBox "Err.Number, Err.Description, Err.Source", vbOKOnly, "Error"
   GoSub Proc_Exit
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    Unload Me
End Sub

Public Sub startrpt()
  Me.Run
End Sub

Private Sub ActiveReport_Initialize()
  Me.Toolbar.Tools.Add "&Close"
  Me.Toolbar.Tools.Add "Save/&Excel"
  Me.Toolbar.Tools.Add "/&Text"

End Sub

Private Sub ActiveReport_ReportStart()
  If Det = True Then
    If Rev < 5 Then
      Me.Detail.Height = 270
    ElseIf Rev > 4 And Rev < 9 Then
      Me.Detail.Height = 540
    ElseIf Rev > 8 And Rev < 13 Then
      Me.Detail.Height = 810
    Else
      Me.Detail.Height = 1332
    End If
  Else
    Me.Detail.Height = 0
  End If
'  If Rev > 0 Then
'    Me.Label14.Visible = False
'    Me.Label15.Visible = False
'    txtcur.Visible = False
'    txtHead.Visible = True
'    Me.Label18.Visible = True
'    Me.Debit.Left = 7110
'    Me.txtTotCur.Left = 7110
'    Me.txtTotAcctBal.Visible = False
'    Me.txtTotPast.Visible = False
'  End If
  Me.Label18 = Me.txtRptParm1
  Me.Label16 = Me.txtRptParm2
End Sub

Private Sub PageHeader_Format()
If Me.pageNumber = 1 Then
  Label15.Visible = True
  Shape1.Visible = True
  txtRptParm1.Visible = True
  txtRptParm2.Visible = True
  Me.PageHeader.Height = 1536
Else
  Label15.Visible = False
  Shape1.Visible = False
  txtRptParm1.Visible = False
  txtRptParm2.Visible = False
  Me.PageHeader.Height = 1156
  Label1.Top = 630
  Label2.Top = 630
  Label3.Top = 630
  Label6.Top = 630
 ' Line1.Y1 = 990
 ' Line1.Y2 = 990

End If
End Sub

'Private Sub GroupFooter1_Format()
'  If Det > 0 Then
'    If Det < 5 Then
'      Me.GroupFooter1.Height = 270
'    ElseIf Det > 4 And Det < 9 Then
'      Me.GroupFooter1.Height = 540
'    ElseIf Det > 8 And Det < 13 Then
'      Me.GroupFooter1.Height = 810
'    Else
'      Me.GroupFooter1.Height = 1080
'    End If
'  Else
'    Me.GroupFooter1.Visible = False
'  End If
'End Sub

Private Sub ReportFooter_Format()
'  If Det > 0 Then
    Set Me.SubReport1.object = New ARSubTot
'  End If
End Sub

Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
    KeyCode = 0
  End If
  If Shift = 4 Then
    If KeyCode = vbKeyC Then
      Unload Me
    End If
    If KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - UBTax.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - UBTax.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If
End Sub

'Private Sub Detail_AfterPrint()
'  Me.GroupHeader1.Visible = False
'End Sub
'
'
'Private Sub PageHeader_AfterPrint()
'  GHP = False
'End Sub

'Private Sub GroupHeader1_BeforePrint()
'  if
'End Sub

'Private Sub PageHeader_Format()
'  If Me.Fields("LineTyp").Value = "D" Then
'    Me.PageHeader.Height = 1752
'    Me.GroupHeader1.Visible = False
'  End If
'  If Fields("LineTyp").Value = "TE" Then
'    Me.PageHeader.Height = 1068
'  End If
'  If Fields("LineTyp").Value = "T" Then
'    Me.PageHeader.Height = 1380
'  End If
'  If Me.Fields("LineTyp").Value = "H" Then
'    If Not Me.Fields("Date").Value = "Account" Then
'       Me.PageHeader.Height = 1068
'    Else
'      Me.PageHeader.Height = 1380
'    End If
'  End If
'End Sub
'
'Private Sub Detail_Format()
'  If Me.Fields("Code").Value <> "" Then
'  If Me.Fields("Code").Value = 4 Then
'    totOpnPO = totOpnPO + Me.Fields("Credit").Value
'  End If
'  If Me.Fields("Code").Value = -4 Then
'    totClsPO = totClsPO + Me.Fields("Credit").Value
'  End If
'  If Me.Fields("Code").Value = 1 Then
'    totInv = totInv + Me.Fields("Credit").Value
'      If Left$(Me.Fields("Status").Value, 4) = "  Pd" Then
'        totPInv = totPInv + Me.Fields("Credit").Value
'      End If
'  End If
'  If Me.Fields("Code").Value = -1 Then
'    'totInv = totInv + Me.Fields("Credit").Value
'    totVInv = totVInv + Me.Fields("Credit").Value
'  End If
'  If Me.Fields("Code").Value = 3 Then
'    totCks = totCks + Me.Fields("Debit").Value
'  End If
'  If Me.Fields("Code").Value = -3 Then
'    totVCks = totVCks + Me.Fields("Debit").Value
'  End If
'  End If
'End Sub

'Private Sub GroupFooter1_AfterPrint()
'  totInv = 0
'  totOpnPO = 0
'  totPInv = 0
'  totVInv = 0
'  totClsPO = 0
'  totCks = 0
'  totVCks = 0
'End Sub

'Private Sub GroupFooter1_BeforePrint()
'Dim tmpPO As Double, Bal As Double
'Bal = 0
'tmpPO = 0
'  'Me.txtTotInv = totInv
'  Me.txtTotOpnInv.DataValue = Round(totInv - totPInv)
'  'Me.txtTotPO = totOpnPO
'  If totOpnPO > 0 Then
'    tmpPO = Round(totOpnPO)
'    If tmpPO > 0 Then
'      Me.txtTotOpnPO.DataValue = tmpPO
'    Else
'      Me.txtTotOpnPO.DataValue = 0
'    End If
'  Else
'    Me.txtTotOpnPO.DataValue = 0
'  End If
'  Me.txtTotCks.DataValue = totCks
'  Me.txtTotInv.DataValue = totInv
'  Bal = Round(totInv - totCks)
'  Me.txtTotBAlance.DataValue = Bal
'
'End Sub

'Private Sub GroupHeader1_AfterPrint()
'  Me.GroupHeader1.Visible = False
'  Me.GroupHeader1.GrpKeepTogether = ddGrpNone
'  GHP = False
'End Sub
'
'Private Sub GroupHeader1_Format()
'  If Det = 0 Then
'    Me.GroupHeader1.Visible = False
'  End If
'End Sub
'  If Fields("Linetyp").Value = "H" Then
'    Me.GroupHeader1.Visible = False
'    Me.GroupHeader1.GrpKeepTogether = ddGrpNone
'    'Me.Line1.Visible = False
'  End If
'  If Fields("LineTyp").Value = "D" Then
'    Me.GroupHeader1.Visible = True
'    Me.GroupHeader1.GrpKeepTogether = ddGrpFirstDetail
'    GHP = True
'  End If
''  If Fields("linetyp").Value = "T" Then
''    Me.GroupHeader1.Visible = False
''  End If
'  If Fields("linetyp").Value = "TE" Then
'     Me.GroupHeader1.Visible = False
'     Me.GroupFooter1.Visible = True
'     Me.GroupHeader1.GrpKeepTogether = ddGrpNone
'     'Me.Line1.Visible = True
'  End If
'End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
  KillFile ReportFile$
  KillFile SubFile$
End Sub
Private Sub ActiveReport_ReportEnd()
    If hFile <> 0 Then
        Close #hFile
    End If
  Unload frmLoadingRpt
  Me.Show 1
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
    MsgBox "File - UBTax.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - UBTax.txt, created in the Citipak Directory.", vbOKOnly
  End If
End Sub
Private Sub ExportReport(x As Integer)
  Dim oEXL As ActiveReportsExcelExport.ARExportExcel
  Dim oTXT As ActiveReportsTextExport.ARExportText
  Dim outfile As String
  If Right$(UBPath, 1) = ":" Then
    outfile = UBPath
  Else
    outfile = UBPath & "\"
  End If
  Select Case x
    Case 1   '"Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & "UBTax.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "UBTax.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub


'Private Sub GroupHeader1_Format()
'  If Me.Fields("Code").Value = "" Then
'    Me.GroupFooter1.Visible = False
'    Me.Detail.Visible = False
'    Me.txtNoTrans.Visible = True
'    Me.GroupHeader1.GrpKeepTogether = ddGrpNone
'  Else
'    Me.txtNoTrans.Visible = False
'    Me.GroupFooter1.Visible = True
'    Me.Detail.Visible = True
'    Me.GroupHeader1.GrpKeepTogether = ddGrpFirstDetail
'  End If
'
'End Sub

