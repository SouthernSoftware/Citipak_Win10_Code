VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptVATempTaxBill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laser Bill "
   ClientHeight    =   6735
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   11460
   Icon            =   "ARptVATempTaxBill.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20214
   _ExtentY        =   11880
   SectionData     =   "ARptVATempTaxBill.dsx":08CA
End
Attribute VB_Name = "ARptVATempTaxBill"
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
Dim headers(1 To 16) As String

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
    headers(7) = "ParcelID" '6)
    headers(8) = "PropDesc" '7)
    headers(9) = "RealVal" '8)
    headers(10) = "BldgVal" '9)
    headers(11) = "ExemptVal" '10)
    headers(12) = "TaxableVal" '11)
    headers(13) = "RatePer100" '12)
    headers(14) = "TaxesDue" '13)
    headers(15) = "BZip" '14)
    headers(16) = "CZip" '15)
    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    For cnt = 1 To 16
      Fields.Add headers(cnt)
    Next
    Fields.Add ("fldLogo") '16)
    Fields.Add ("fldOpt1Due") '17)
    Fields.Add ("fldOpt2Due") '18)
    Fields.Add ("fldOpt3Due") '19)
    Fields.Add ("fldOpt1Desc") '20)
    Fields.Add ("fldOpt2Desc") '21)
    Fields.Add ("fldOpt3Desc") '22)
    Fields.Add ("fldOverPay") '23)
    Fields.Add ("fldLateList") '24)
    Fields.Add ("fldPriorYrTax") '25)
    Fields.Add ("fldPrintPrYr") '26
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
    Line Input #hFile, sLine
    arr = Split(sLine, "~")
'    ' Here we set the values of the fields that we defines as unbound
'    ' or user defined.
    For cnt = 1 To 16 '5
      Fields(headers(cnt)) = arr(cnt - 1)
    Next cnt
    
    If Len(QPTrim$(arr(14))) <> 0 Then
      Barcode1.Visible = True
    Else
      Barcode1.Visible = False
    End If
    
    Fields("fldLogo").Value = arr(15)
    Fields(headers(5)).Value = Fields(headers(5)).Value + " " + Fields(headers(16)).Value
    Fields("fldOpt1Due").Value = arr(17)
    Fields("fldOpt2Due").Value = arr(18)
    Fields("fldOpt3Due").Value = arr(19)
    Fields("fldOpt1Desc").Value = arr(20)
    Fields("fldOpt2Desc").Value = arr(21)
    Fields("fldOpt3Desc").Value = arr(22)
    If QPTrim$(arr(20)) = "" Then
      Field157.Visible = False
      Label42.Caption = ""
    Else
      Field157.Visible = True
      Label42.Caption = QPTrim$(arr(20))
    End If
    If QPTrim$(arr(21)) = "" Then
      Field158.Visible = False
      Label43.Caption = ""
    Else
      Field158.Visible = True
      Label43.Caption = QPTrim$(arr(21))
    End If
    If QPTrim$(arr(22)) = "" Then
      Field159.Visible = False
      Label44.Caption = ""
    Else
      Field159.Visible = True
      Label44.Caption = QPTrim$(arr(22))
    End If
    Field160.Visible = False
    Fields("fldOverPay").Value = arr(23)
    If QPTrim$(arr(23)) <> "0" Then
      Field160.Visible = True
      Fields("fldOverPay").Value = "** Applied Credit of " + QPTrim$(Using$("$##,##0.00", CDbl(arr(23)))) + " **"
    End If
    Fields("fldLateList").Value = arr(24)
    Fields("fldPriorYrTax").Value = arr(25)
    Fields("fldPrintPrYr").Value = arr(26)
    
    If arr(26) = True Then
      Label49.Caption = "Prior Balance"
      Field163.Visible = True
    Else
      Label49.Caption = ""
      Field163.Visible = False
    End If
    
    
'If something wrong in file give message instead of crashing
Exit Sub
ERRORSTUFF:
      Unload frmVATaxLoadingRpt
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
      MsgBox "File - TxBill1.xls, created in the Citipak Directory.", vbOKOnly
    End If
    If KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - TxBill1.txt, created in the Citipak Directory.", vbOKOnly
    End If
  End If
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
  Close #hFile
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
    Close
    Unload Me
  End If
  If Tool = "Save/&Excel" Then
    Screen.MousePointer = vbHourglass
    ExportReport 1
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - TxBill1.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "/&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - TxBill1.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "TxBill1.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TxBill1.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub Detail_Format()
  Label46.Visible = False
  Field161.Visible = False
  If QPTrim$(Fields("fldLateList").Value) <> "" Then
    If CDbl(Fields("fldLateList").Value) > 0 Then
      Label46.Visible = True
      Field161.Visible = True
    End If
  End If
'  If QPTrim$(Fields("fldLogo").Value) = "1" Then
'    If Exist("towntaxlogo.bmp") Then
'      DoEvents
'      Image1.Picture = LoadPicture("towntaxlogo.bmp")
'      DoEvents
'    End If
'  End If
End Sub
