VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptFnctBudVAct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Budget vs Actual"
   ClientHeight    =   7104
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   12216
   Icon            =   "ARptFnctBudVAct.dsx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   21548
   _ExtentY        =   12531
   SectionData     =   "ARptFnctBudVAct.dsx":08CA
End
Attribute VB_Name = "ARptFnctBudVAct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Private hFile As Integer
Dim rpt As ActiveReport
Dim headers(1 To 10) As String
Dim cnt As Integer
Dim FRevTotM As Double, FRevTotY As Double
Dim FExpTotM As Double, FExpTotY As Double
Public rptnum As Integer
Public detopt As Integer
Public deptpage As Boolean  'use this for subtotal revs
Public overunder As Boolean

Public Sub GetName(RName As String)
  ReportFile$ = RName$
End Sub

Private Sub ActiveReport_DataInitialize()
    headers(1) = "Fund"
    headers(2) = "Typ"
    headers(3) = "Dept"
    headers(4) = "DeptName"
    headers(5) = "AcctDesc"
    headers(6) = "Budget"
    headers(7) = "MTD/Enc"
    headers(8) = "YTD"
    headers(9) = "Variance"
    headers(10) = "Pct"

    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    For cnt = 1 To 10
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

'
Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sLine As String
Dim arr() As String
'
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

    Line Input #hFile, sLine
    arr = Split(sLine, "~")

'    ' Here we set the values of the fields that we defines as unbound
'    ' or user defined.
    For cnt = 1 To 10
      Fields(headers(cnt)) = arr(cnt - 1)
    Next
'    ("Fund").Value = arr(0)
'    Fields("Dept").Value = arr(1)
'    Fields("DeptName").Value = arr(2)
'    Fields("AcctDesc").Value = arr(3)
'    Fields("Budget").Value = Val(arr(4))
'    Fields("MTD/Enc").Value = arr(5)
'    Fields("YTD").Value = Val(arr(6))
'    Fields("Variance").Value = Val(arr(7))
'    Fields("Pct").Value = arr(8)

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
    End If
    If KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - FBudVAct.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - FBudVAct.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
  KillFile ReportFile$
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
    MsgBox "File - FBudVAct.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - FBudVAct.txt, created in the Citipak Directory.", vbOKOnly
  End If
End Sub

Public Sub startrpt()
  Me.Run
End Sub
Private Sub ActiveReport_Initialize()
  Me.Toolbar.Tools.Add "&Close"
  Me.Toolbar.Tools.Add "Save/&Excel"
  Me.Toolbar.Tools.Add "/&Text"
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
        oEXL.FileName = outfile & "FBudVAct.xls"
        oEXL.Export Me.Pages
        
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "FBudVAct.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub


Private Sub Detail_Format()
 Const cstrProcName As String = "Detail Format"
 On Error GoTo ERRORSTUFF
  If detopt <> 1 Then
    Detail.Visible = False
  Else
    Detail.Visible = True

  End If
  If Fields("Typ").Value = "R" Then
    'do this so will print revenues on a seperate page
    'Me.GroupHeader3.NewPage = ddNPAfter
    Me.Label9.Caption = "Total Revenues for Function"
    FRevTotM = FRevTotM + Fields("MTD/Enc").Value
    FRevTotY = FRevTotY + Fields("YTD").Value
  Else
  'but not for expenses unless selected dept on sep page
    Me.Label9.Caption = "Total Expenses for Function"
    FExpTotM = FExpTotM + Fields("MTD/Enc").Value
    FExpTotY = FExpTotY + Fields("YTD").Value
  End If
  Exit Sub
''If something wrong in file give message instead of crashing
ERRORSTUFF:
      Unload frmLoadingRpt
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "ARptBudVAct", cstrProcName, Erl)
    Case emrExitProc:
      Resume Proc_Exit
    Case emrResume:
      Resume
    Case emrResumeNext:
      Resume Next
    Case Else
      '--- Technically, this should never happen.
      Resume Proc_Exit
  End Select

Proc_Exit:
  '--- Cleanup code goes here...
    Close
    Unload Me
End Sub


Private Sub GroupHeader1_Format()
  Me.Labelfnc.Caption = "Function - " + Me.Fields(0).Value + " - " + Me.Fields(3).Value
End Sub

'Private Sub GroupFooter3_Format()
'  If detopt = 1 Then
'    GroupFooter3.Height = 400
'  End If
'
'End Sub

Private Sub GroupHeader2_Format()
  'If detopt <> 1 Then
    If Fields("typ").Value = "R" Then
      If Me.deptpage = True Then
        GroupHeader2.NewPage = ddNPBefore
      End If
      Label10.Caption = "Revenues - Function " + Me.Fields(0).Value
      Label10.Visible = True
    Else
      If Me.deptpage = True Then
        GroupHeader2.NewPage = ddNPBefore
      End If
      Label10.Caption = "Expenditures - Function " + Me.Fields(0).Value
      Label10.Visible = True
    End If
End Sub


Private Sub GroupFooter1_AfterPrint()
  FRevTotM = 0
  FRevTotY = 0
  FExpTotM = 0
  FExpTotY = 0
End Sub

Private Sub GroupFooter2_BeforePrint()
  Dim TotTyp2 As Double
  If rptnum = 2 Then
    TotTyp2 = MTDTyp.DataValue + YTDTyp.DataValue
    PctTyp.Text = GetPct$(TotTyp2, BgtTyp)
  Else
    PctTyp.Text = GetPct$(YTDTyp, BgtTyp)
  End If
End Sub


'Private Sub GroupFooter3_BeforePrint()
'  Dim Tot2 As Double
'  If rptnum = 2 Then
'    Tot2 = MTDEncDept.DataValue + YTDDept.DataValue
'    PctDept.Text = GetPct$(Tot2, BgtDept)
'  Else
'    PctDept.Text = GetPct$(YTDDept, BgtDept)
'  End If
'
'End Sub
'
Private Sub GroupFooter1_BeforePrint()
  Dim MTDBal As Double, YTDBal As Double
  If overunder = True Then
    MTDBal = Round(FRevTotM - FExpTotM)
    YTDBal = Round(FRevTotY - FExpTotY)
  Else
    MTDBal = 0
    YTDBal = 0
    Me.Label1.Visible = False
    Me.Field6.Visible = False
    Me.FundTot1.Visible = False
    Me.FundTot2.Visible = False
    Me.Shape1.Visible = False
  End If
    FundTot1.DataValue = MTDBal
    FundTot2.DataValue = YTDBal

'  If rptnum = 2 Then
'        MTDBal# = Round#(FundRevMTD# - FundExpMTD#)
'      End If
'      BGTBal# = Round#(FundRevBgt# - FundExpBgt#)
'      YTDBal# = Round#(FundRevYTD# - FundExpYTD#)
'      EncBal# = Round#(FundEncYTD#)

'      Case 1, 3
'        Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(MTDBal#))
'        Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(YTDBal#))
'        '--Reset MTD Variables
'        FundRevMTD# = 0
'        FundExpMTD# = 0
'        DeptMTDSum# = 0
'      Case 2
'        Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(YTDBal#))
'      End Select

End Sub
'Private Sub GroupHeader3_Format()
'  If Me.deptpage = True Then 'And Me.Fields("Typ") <> "R" Then
'  'only do new page if new dept and selected seperate page option
'    Me.GroupHeader3.NewPage = ddNPAfter
'  End If
'End Sub
