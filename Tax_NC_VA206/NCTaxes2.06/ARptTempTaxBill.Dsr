VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARptTempTaxBill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laser Bill "
   ClientHeight    =   9465
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   11460
   Icon            =   "ARptTempTaxBill.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20214
   _ExtentY        =   16695
   SectionData     =   "ARptTempTaxBill.dsx":08CA
End
Attribute VB_Name = "ARptTempTaxBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportFile As String
Private hFile As Integer
Dim rpt As ActiveReport
Dim tempTot As Integer
Dim Cnt As Integer, dcnt As Integer
Dim headers(1 To 26) As String

Public Sub GetName(RName As String)
  ReportFile$ = RName$
End Sub
Private Sub ActiveReport_DataInitialize()
    headers(1) = "BillNum" '0
    headers(2) = "CustName" '1
    headers(3) = "Addr1" '2
    headers(4) = "Addr2" '3
    headers(5) = "City" '4
    headers(6) = "TaxID" '5
    headers(7) = "ParcelID" '6
    headers(8) = "PropDesc" '7
    headers(9) = "RealVal" '8
    headers(10) = "PersVal" '9
    headers(11) = "ExemptVal" '10
    headers(12) = "TaxableVal" '11
    headers(13) = "RatePer100" '12
    headers(14) = "TaxesDue" '13
    headers(15) = "BarCode" '14
    headers(16) = "fldLogo" '15
    headers(17) = "Opt1Val" '16
    headers(18) = "Opt2Val" '17
    headers(19) = "Opt3Val" '18
    headers(20) = "Opt1Desc" '19
    headers(21) = "Opt2Desc" '20
    headers(22) = "Opt3Desc" '21
    headers(23) = "CZip" '22
    headers(24) = "LateList" '23
    headers(25) = "Prepay" '24
    headers(26) = "PrintPriorYN" '25
    hFile = FreeFile
    Open ReportFile$ For Input As #hFile
'
'    ' This sets up the fields used in data binding
'    'Fields.Add "ProductID"
    For Cnt = 1 To 26
      Fields.Add headers(Cnt)
    Next
'    Fields.Add ("fldLogo") '1f)
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
  
'  'on error goto ERRORSTUFF
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
    For Cnt = 1 To 26
      Fields(headers(Cnt)) = arr(Cnt - 1)
    Next
    If QPTrim$(arr(14)) = "" Then
      Barcode1.Visible = False
    Else
      Barcode1.Visible = True
    End If
    
    If arr(19) <> "" Then
      Label42.Caption = arr(19)
    Else
      Label42.Caption = "TAX OPT #1"
    End If
    
    If arr(20) <> "" Then
      Label44.Caption = arr(20)
    Else
      Label44.Caption = "TAX OPT #2"
    End If
    
    If arr(21) <> "" Then
      Label45.Caption = arr(21)
    Else
      Label45.Caption = "TAX OPT #3"
    End If
    Fields(headers(5)).Value = Fields(headers(5)).Value + " " + Fields(headers(23)).Value

    If Val(CDbl(arr(24))) < 0 Or arr(25) = True Then
      Label47.Caption = "Prior Balance"
      Field162.Visible = True
    Else
      Label47.Caption = ""
      Field162.Visible = False
    End If
      
Exit Sub

ERRORSTUFF:
   Unload frmLoadingRpt
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
  Me.ToolBar.Tools.Add "&Close"
  Me.ToolBar.Tools.Add "Save/&Excel"
  Me.ToolBar.Tools.Add "/&Text"
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
  Unload frmLoadingRpt
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
'  If QPTrim$(Fields("fldLogo").Value) = "1" Then
'    If Exist("towntaxlogo.bmp") Then
'      DoEvents
'      Image1.Picture = LoadPicture("towntaxlogo.bmp")
'      DoEvents
'    End If
'  End If
End Sub
