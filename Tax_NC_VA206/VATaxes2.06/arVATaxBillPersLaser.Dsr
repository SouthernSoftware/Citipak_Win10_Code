VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxBillPersLaser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal Laser Tax Bill"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "arVATaxBillPersLaser.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15452
   SectionData     =   "arVATaxBillPersLaser.dsx":08CA
End
Attribute VB_Name = "arVATaxBillPersLaser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Private hFile As Integer
Dim rpt As ActiveReport
Dim tempTot As Integer
Dim cnt As Integer, dcnt As Integer
Dim headers(1 To 37) As String

Public Sub GetName(RName As String)
  ReportFile$ = RName$
End Sub
Private Sub ActiveReport_DataInitialize()
    headers(1) = "BillNum" '0)
    headers(2) = "CustName" '1)
    headers(3) = "Addr1" '2)
    headers(4) = "Addr2" '3)
    headers(5) = "City" '4)
    headers(6) = "TaxID" '5)
    headers(7) = "TotVal" '6)
    headers(8) = "PropDesc" '7)
    headers(9) = "PersVal" '8)
    headers(10) = "FEVal" '9)
    headers(11) = "ExemptVal" '10)
    headers(12) = "PPTRAVal" '11)
    headers(13) = "PPTRADisc" '12)
    headers(14) = "TaxesDue" '13)
    headers(15) = "logo" '14)
    headers(16) = "MHVal" '15)
    headers(17) = "MCVal" '16)
    headers(18) = "MTVal" '17)
    headers(19) = "PersTaxDue" '18)
    headers(20) = "PersTaxNet" '19)
    headers(21) = "PersTaxRate" '20)
    headers(22) = "FETaxDue" '21)
    headers(23) = "FETaxRate" '22)
    headers(24) = "MCTaxDue" '23)
    headers(25) = "MCTaxRate" '24)
    headers(26) = "MHTaxDue" '25)
    headers(27) = "MHTaxRate" '26)
    headers(28) = "MTTaxDue" '27)
    headers(29) = "MTTaxRate" '28)
    headers(30) = "Opt1TaxDue" '29)
    headers(31) = "Opt2TaxDue" '30)
    headers(32) = "Opt3TaxDue" '31)
    headers(33) = "Opt1Desc" '32)
    headers(34) = "Opt2Desc" '33)
    headers(35) = "Opt3Desc" '34)
    headers(36) = "BZip" '35)
    headers(37) = "CZip" '36)
    hFile = FreeFile
    Open ReportFile$ For Input As #hFile

'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    For cnt = 1 To 37
      Fields.Add headers(cnt)
    Next
    Fields.Add ("fldLogo") '37)
    Fields.Add ("fldOverPay") '38)
    Fields.Add ("fldPriorTax") '39)
    Fields.Add ("fldPrintPrTax") '40)
End Sub
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmVATaxLoadingRpt
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

  Dim sLine As String
  Dim arr() As String
  Dim ThisOpt As String * 16
'  On Error GoTo ERRORSTUFF
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
    For cnt = 1 To 37
      Fields(headers(cnt)) = arr(cnt - 1)
    Next
    If Len(QPTrim$(arr(35))) <> 0 Then
      Barcode1.Visible = True
    Else
      Barcode1.Visible = False
    End If
    Fields("fldLogo").Value = arr(37)
    Field177.Visible = False
    If QPTrim$(arr(38)) <> "0" Then
      Field177.Visible = True
      Fields("fldOverPay").Value = "** Applied Credit Of " + QPTrim$(Using$("$##,##0.00", CDbl(arr(38)))) + " **"
    End If
    Fields(headers(5)).Value = Fields(headers(5)).Value + " " + Fields(headers(37)).Value
    Fields("fldPriorTax").Value = arr(39)
    If arr(40) = True Then
      arr(39) = arr(39)
      Label62.Caption = "PRIOR BALANCE"
      Field178.Visible = True
    Else
      Label62.Caption = ""
      Field178.Visible = False
    End If

    'If something wrong in file give message instead of crashing
  Exit Sub
  
ERRORSTUFF:
    Unload frmVATaxLoadingRpt
    Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "arVATaxBillPersLaser", "Fetch Data", Erl)
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
   MsgBox "Err.Number, Err.Description, Err.Source", vbOKOnly, "Error"
   GoSub Proc_Exit
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    Unload Me
End Sub

Public Sub startrpt()
  Me.Run True
End Sub

Private Sub ActiveReport_Initialize()
  Me.Toolbar.Tools.Add "&Close"
  Me.Toolbar.Tools.Add "Save/&Excel"
  Me.Toolbar.Tools.Add "/&Text"
  dcnt = 0
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
      MsgBox "File - TxPersBill1.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - TxPersBill1.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
 ' KillFile ReportFile$
End Sub
Private Sub ActiveReport_ReportEnd()
    If hFile <> 0 Then
        Close #hFile
    End If
  Unload frmVATaxLoadingRpt
'  Me.Show 1
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
    MsgBox "File - TxPersBill1.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - TxPersBill1.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "TxPersBill1.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TxPersBill1.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub Detail_Format()
  Label57.Caption = ""
  Field174.Visible = False
  Label58.Caption = ""
  Field175.Visible = False
  Label59.Caption = ""
  Field176.Visible = False
  If Fields(headers(33)).Value > 0 Then
    Label57.Caption = QPTrim$(Fields(headers(33)).Value)
    Field174.Visible = True
  End If
  If Fields(headers(34)).Value > 0 Then
    Label58.Caption = QPTrim$(Fields(headers(34)).Value)
    Field175.Visible = True
  End If
  If Fields(headers(35)).Value > 0 Then
    Label59.Caption = QPTrim$(Fields(headers(35)).Value)
    Field176.Visible = True
  End If
'  If QPTrim$(Fields("fldLogo").Value) = "1" Then
'    If Exist("towntaxlogo.bmp") Then
'      DoEvents
'      Image1.Picture = LoadPicture("towntaxlogo.bmp")
'      DoEvents
'    End If
'  End If
End Sub

