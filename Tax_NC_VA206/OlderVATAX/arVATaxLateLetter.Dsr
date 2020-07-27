VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxLateLetter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Late Letter"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arVATaxLateLetter.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arVATaxLateLetter.dsx":08CA
End
Attribute VB_Name = "arVATaxLateLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private hFile As Integer
Private Temp_Class As Resize_Class
Dim Completed As Boolean

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\TAXRPTS\LATENOTICE.RPT" For Input As #hFile
  Fields.Add ("fldAdd1") '0)
  Fields.Add ("fldAdd2") '1)
  Fields.Add ("fldAdvBal") '2)
  Fields.Add ("fldAdvDate") '3)
  Fields.Add ("fldCity") '4)
  Fields.Add ("fldCustName") '5)
  Fields.Add ("fldIntBal") '6)
  Fields.Add ("fldLateListBal") '7)
  Fields.Add ("fldSeqNum") '8)
  Fields.Add ("fldOpt1Bal") '9)
  Fields.Add ("fldOpt2Bal") '10)
  Fields.Add ("fldOpt3Bal") '11)
  Fields.Add ("fldPayDate") '12)
  Fields.Add ("fldPersExemp") '13)
  Fields.Add ("fldPersVal") '14)
  Fields.Add ("fldPrincBal") '15)
  Fields.Add ("fldRealExemp") '16)
  Fields.Add ("fldRealValue") '17)
  Fields.Add ("fldState") '18)
  Fields.Add ("fldTaxYear") '19)
  Fields.Add ("fldTownname") '20)
  Fields.Add ("fldZip") '21)
  Fields.Add ("fldTownAdd1") '22)
  Fields.Add ("fldTownAdd2") '23)
  Fields.Add ("fldTownCSZ") '24)
  Fields.Add ("fldLtrDate") '25)
  Fields.Add ("fldHead1") '26)
  Fields.Add ("fldHead2") '27)
  Fields.Add ("fldHead3") '28)
  Fields.Add ("fldHead4") '29)
  Fields.Add ("fldHead5") '30)
  Fields.Add ("fldBody1") '31)
  Fields.Add ("fldBody2") '32)
  Fields.Add ("fldBody3") '33)
  Fields.Add ("fldBody4") '34)
  Fields.Add ("fldBody5") '35)
  Fields.Add ("fldBody6") '36)
  Fields.Add ("fldBody7") '37)
  Fields.Add ("fldBody8") '38)
  Fields.Add ("fldBody9") '39)
  Fields.Add ("fldBody10") '40)
  Fields.Add ("fldBody11") '41)
  Fields.Add ("fldBody12") '42)
  Fields.Add ("fldBody13") '43)
  Fields.Add ("fldBody14") '44)
  Fields.Add ("fldBody15") '45)
  Fields.Add ("fldBody16") '46)
  Fields.Add ("fldBody17") '47)
  Fields.Add ("fldBody18") '48)
  Fields.Add ("fldBody19") '49)
  Fields.Add ("fldBody20") '50)
  Fields.Add ("fldTotBal") '51)
  Fields.Add ("fldCurrBal") '52)
  Fields.Add ("fldPrevBal") '53)
  Fields.Add ("fldCustAcct") '54)
  Fields.Add ("fldCurrTaxYear") '55)
  Fields.Add ("fldCSZ")
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
'    Unload frmLoadReport
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
    Completed = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #hFile, sLine
  arr = Split(sLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("fldAdd1").Value = arr(0)
  Fields("fldAdd2").Value = arr(1)
  Fields("fldAdvBal").Value = arr(2)
  Fields("fldAdvDate").Value = arr(3)
  Fields("fldCity").Value = arr(4)
  Fields("fldCustName").Value = arr(5)
  Fields("fldIntBal").Value = arr(6)
  Fields("fldLateListBal").Value = arr(7)
  Fields("fldSeqNum").Value = arr(8)
  Fields("fldOpt1Bal").Value = arr(9)
  Fields("fldOpt2Bal").Value = arr(10)
  Fields("fldOpt3Bal").Value = arr(11)
  Fields("fldPayDate").Value = arr(12)
  Fields("fldPersExemp").Value = arr(13)
  Fields("fldPersVal").Value = arr(14)
  Fields("fldPrincBal").Value = arr(15)
  Fields("fldRealExemp").Value = arr(16)
  Fields("fldRealValue").Value = arr(17)
  Fields("fldState").Value = arr(18)
  Fields("fldTaxYear").Value = arr(19)
  Fields("fldTownname").Value = arr(20)
  Fields("fldZip").Value = arr(21)
  Fields("fldTownAdd1").Value = arr(22)
  Fields("fldTownAdd2").Value = arr(23)
  Fields("fldTownCSZ").Value = arr(24)
  Fields("fldLtrDate").Value = arr(25)
  Fields("fldHead1").Value = QPTrim$(arr(26))
  Fields("fldHead2").Value = QPTrim$(arr(27))
  Fields("fldHead3").Value = QPTrim$(arr(28))
  Fields("fldHead4").Value = QPTrim$(arr(29))
  Fields("fldHead5").Value = QPTrim$(arr(30))
  Fields("fldBody1").Value = arr(31)
  Fields("fldBody2").Value = arr(32)
  Fields("fldBody3").Value = arr(33)
  Fields("fldBody4").Value = arr(34)
  Fields("fldBody5").Value = arr(35)
  Fields("fldBody6").Value = arr(36)
  Fields("fldBody7").Value = arr(37)
  Fields("fldBody8").Value = arr(38)
  Fields("fldBody9").Value = arr(39)
  Fields("fldBody10").Value = arr(40)
  Fields("fldBody11").Value = arr(41)
  Fields("fldBody12").Value = arr(42)
  Fields("fldBody13").Value = arr(43)
  Fields("fldBody14").Value = arr(44)
  Fields("fldBody15").Value = arr(45)
  Fields("fldBody16").Value = arr(46)
  Fields("fldBody17").Value = arr(47)
  Fields("fldBody18").Value = arr(48)
  Fields("fldBody19").Value = arr(49)
  Fields("fldBody20").Value = arr(50)
  Fields("fldTotBal").Value = arr(51)
  Fields("fldCurrBal").Value = arr(52)
  Fields("fldPrevBal").Value = arr(53)
  Fields("fldCustAcct").Value = arr(54)
  Fields("fldCurrTaxYear").Value = arr(55)
  Fields("fldCSZ").Value = QPTrim$(arr(4)) + ", " + QPTrim$(arr(18)) + "  " + QPTrim$(arr(21))
  If arr(19) = "ALL" Then
    Label3.Caption = "All Years"
  ElseIf CInt(arr(55)) > CInt(arr(19)) Then
    Label3.Caption = "Other Tax Years:"
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
      frmVATaxMsg.Label1.Caption = "File - TaxLateLtr.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - TaxLateLtr.txt, created in the Citipak Directory."
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
    frmVATaxMsg.Label1.Caption = "File - TaxLateLtr.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - TaxLateLtr.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "TaxLateLtr.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TaxLateLtr.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_ReportEnd()
'  Unload frmBLLoadReport
  If hFile <> 0 Then
    Close #hFile
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    'Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
'  Me.fldTimeDate.Text = Date
'  Me.Zoom = -1
End Sub

Private Sub Detail_Format()
 If Exist("towntaxlogo.bmp") Then
  DoEvents
  Image1.Picture = LoadPicture("towntaxlogo.bmp")
  Image1.Visible = True
  DoEvents
 End If

End Sub
