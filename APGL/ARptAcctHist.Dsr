VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptAcctHist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account History"
   ClientHeight    =   6885
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   11295
   Icon            =   "ARptAcctHist.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   19923
   _ExtentY        =   12144
   SectionData     =   "ARptAcctHist.dsx":08CA
End
Attribute VB_Name = "ARptAcctHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Private hFile As Integer
Dim rpt As ActiveReport
Dim AC As String
Dim GHP As Boolean
Public Sub GetName(RName As String)
  ReportFile$ = RName$
End Sub
Private Sub ActiveReport_DataInitialize()
   
    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    Fields.Add "LineTyp"
    Fields.Add "Date"
    Fields.Add "Desc"
    Fields.Add "Ref"
    Fields.Add "Debit"
    Fields.Add "Credit"
    Fields.Add "BalSrc"
    
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
On Error GoTo ERRORSTUFF
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
    Fields("LineTyp").Value = arr(0)
    Fields("Date").Value = arr(1)
    Fields("Desc").Value = arr(2)
    Fields("Ref").Value = arr(3)
    Fields("Debit").Value = arr(4)
    Fields("Credit").Value = arr(5)
    Fields("BalSrc").Value = arr(6)
'If something wrong in file give message instead of crashing
Exit Sub
ERRORSTUFF:
      Unload frmLoadingRpt
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "ARptBudVAct", "Fetch Data", Erl)
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

Public Sub startrpt()
  Me.Run
End Sub

Private Sub ActiveReport_Initialize()
  Me.Toolbar.Tools.Add "&Close"
  Me.Toolbar.Tools.Add "Save/&Excel"
  Me.Toolbar.Tools.Add "/&Text"
  

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
      MsgBox "File - AcctHist.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - AcctHist.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If

End Sub

Private Sub Detail_AfterPrint()
  Me.GroupHeader1.Visible = False
End Sub


Private Sub PageHeader_AfterPrint()
  GHP = False
End Sub

'Private Sub GroupHeader1_BeforePrint()
'  if
'End Sub

Private Sub PageHeader_Format()
  If Me.Fields("LineTyp").Value = "D" Then
    Me.PageHeader.Height = 1752
    Me.GroupHeader1.Visible = False
  End If
  If Fields("LineTyp").Value = "TE" Then
    Me.PageHeader.Height = 1068
  End If
  If Fields("LineTyp").Value = "T" Then
    Me.PageHeader.Height = 1380
  End If
  If Me.Fields("LineTyp").Value = "H" Then
    If Not Me.Fields("Date").Value = "Account" Then
       Me.PageHeader.Height = 1068
    Else
      Me.PageHeader.Height = 1380
    End If
  End If
End Sub

Private Sub Detail_Format()
  If Me.Fields("Date").Value = "Account" Then
    AC$ = Me.Fields("Desc")
  End If
  If Me.Fields("LineTyp").Value = "H" Then
    If Not Me.Fields("Date").Value = "Account" Then
      Me.Label8 = "Account " & AC$ & " -Continued"
      Me.Label8.Visible = True
    End If
  End If
  If Me.Fields("LineTyp").Value = "D" Then
    Me.Label8 = "Account " & AC$ & " -Continued"
    Me.Label8.Visible = True
    If GHP = False Then
    Me.Label9.Visible = True
    Me.Label10.Visible = True
    Me.Label11.Visible = True
    Me.Label12.Visible = True
    Me.Label13.Visible = True
    Me.Label14.Visible = True
    Me.Line9.Visible = True
    Me.Line10.Visible = True
    Me.Line11.Visible = True
    Me.Line12.Visible = True
    Me.Line13.Visible = True
    Me.Line14.Visible = True
    End If
  End If
  If Fields("LineTyp").Value = "TE" Then
    Me.Label8.Visible = False
  End If
  If Fields("LineTyp").Value = "T" Then
    Me.Label9.Visible = False
    Me.Label10.Visible = False
    Me.Label11.Visible = False
    Me.Label12.Visible = False
    Me.Label13.Visible = False
    Me.Label14.Visible = False
    Me.Line9.Visible = False
    Me.Line10.Visible = False
    Me.Line11.Visible = False
    Me.Line12.Visible = False
    Me.Line13.Visible = False
    Me.Line14.Visible = False
    
  End If

End Sub

Private Sub GroupHeader1_AfterPrint()
  Me.GroupHeader1.Visible = False
  Me.GroupHeader1.GrpKeepTogether = ddGrpNone
  GHP = False
End Sub

Private Sub GroupHeader1_Format()
  If Fields("Linetyp").Value = "H" Then
    Me.GroupHeader1.Visible = False
    Me.GroupHeader1.GrpKeepTogether = ddGrpNone
    'Me.Line1.Visible = False
  End If
  If Fields("LineTyp").Value = "D" Then
    Me.GroupHeader1.Visible = True
    Me.GroupHeader1.GrpKeepTogether = ddGrpFirstDetail
    GHP = True
  End If
'  If Fields("linetyp").Value = "T" Then
'    Me.GroupHeader1.Visible = False
'  End If
  If Fields("linetyp").Value = "TE" Then
     Me.GroupHeader1.Visible = False
     Me.GroupFooter1.Visible = True
     Me.GroupHeader1.GrpKeepTogether = ddGrpNone
     'Me.Line1.Visible = True
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
    MsgBox "File - AcctHist.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - AcctHist.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "AcctHist.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "AcctHist.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub




