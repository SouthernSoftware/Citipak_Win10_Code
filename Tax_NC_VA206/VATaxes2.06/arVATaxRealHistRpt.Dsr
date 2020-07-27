VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxRealHistRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Real Estate History Report"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arVATaxRealHistRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arVATaxRealHistRpt.dsx":08CA
End
Attribute VB_Name = "arVATaxRealHistRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private hFile As Integer
Private Temp_Class As Resize_Class

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\TAXRPTS\REALHIST.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldOwner") '1)
  Fields.Add ("fldTDate") '2)
  Fields.Add ("fldThisPin") '3)
  Fields.Add ("fldGTType") '4)
  Fields.Add ("fldBegDate") '5)
  Fields.Add ("fldEndDate") '6)
  Fields.Add ("fldTaxYear") '7)
  Fields.Add ("fldAmount") '8)
  Fields.Add ("fldTDesc") '9)
  Fields.Add ("fldThisTType") '10)
  Fields.Add ("fldPrinc") '11)
  Fields.Add ("fldPrincPd") '12)
  Fields.Add ("fldPrincDif") '13)
  Fields.Add ("fldInt") '14)
  Fields.Add ("fldIntPd") '15)
  Fields.Add ("fldIntDif") '16)
  Fields.Add ("fldAdv") '17)
  Fields.Add ("fldAdvPd") '18)
  Fields.Add ("fldAdvDif") '19)
  Fields.Add ("fldLateList") '20)
  Fields.Add ("fldLateListPd") '21)
  Fields.Add ("fldLateListDif") '22)
  Fields.Add ("fldOpt1") '23)
  Fields.Add ("fldOpt1Pd") '24)
  Fields.Add ("fldOpt1Dif") '25)
  Fields.Add ("fldOpt2") '26)
  Fields.Add ("fldOpt2Pd") '27)
  Fields.Add ("fldOpt2Dif") '28)
  Fields.Add ("fldOpt3") '29)
  Fields.Add ("fldOpt3Pd") '30)
  Fields.Add ("fldOpt3Dif") '31)
  Fields.Add ("fldOpt1Desc") '32)
  Fields.Add ("fldOpt2Desc") '33)
  Fields.Add ("fldOpt3Desc") '34)
  Fields.Add ("fldAddr") '35)
  Fields.Add ("fldCustRec") '36)
  Fields.Add ("fldBillCustrec") '37
  Fields.Add ("fldBillNum") '38)
  Fields.Add ("fldBill2Owner") '39)
  Fields.Add ("fldTCnt") '40)
  Fields.Add ("fldBillBal") '41)
  Fields.Add ("fldDisc") '42)
  Fields.Add ("fldPenDif") '43)
  Fields.Add ("fldPenalty") '44)
  Fields.Add ("fldPenPd") '45)
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
    Exit Sub
  Else
    eof = False
  End If
  Line Input #hFile, sLine
  arr = Split(sLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("fldTown").Value = arr(0)
  Fields("fldOwner").Value = arr(1)
  Fields("fldTDate").Value = arr(2)
  Fields("fldThisPin").Value = arr(3)
  Fields("fldGTType").Value = arr(4)
  Fields("fldBegDate").Value = arr(5)
  Fields("fldEndDate").Value = arr(6)
  Fields("fldTaxYear").Value = arr(7)
  Fields("fldAmount").Value = arr(8)
  Fields("fldTDesc").Value = arr(9)
  Fields("fldThisTType").Value = arr(10)
  Fields("fldPrinc").Value = arr(11)
  Fields("fldPrincPd").Value = arr(12)
  Fields("fldPrincDif").Value = arr(13)
  Fields("fldInt").Value = arr(14)
  Fields("fldIntPd").Value = arr(15)
  Fields("fldIntDif").Value = arr(16)
  Fields("fldAdv").Value = arr(17)
  Fields("fldAdvPd").Value = arr(18)
  Fields("fldAdvDif").Value = arr(19)
  Fields("fldLateList").Value = arr(20)
  Fields("fldLateListPd").Value = arr(21)
  Fields("fldLateListDif").Value = arr(22)
  Fields("fldOpt1").Value = arr(23)
  Fields("fldOpt1Pd").Value = arr(24)
  Fields("fldOpt1Dif").Value = arr(25)
  Fields("fldOpt2").Value = arr(26)
  Fields("fldOpt2Pd").Value = arr(27)
  Fields("fldOpt2Dif").Value = arr(28)
  Fields("fldOpt3").Value = arr(29)
  Fields("fldOpt3Pd").Value = arr(30)
  Fields("fldOpt3Dif").Value = arr(31)
  Fields("fldOpt1Desc").Value = arr(32)
  Fields("fldOpt2Desc").Value = arr(33)
  Fields("fldOpt3Desc").Value = arr(34)
  If QPTrim$(arr(35)) = "" Then
    Fields("fldAddr").Value = "Not Saved"
  Else
    Fields("fldAddr").Value = arr(35)
  End If
  Fields("fldCustRec").Value = arr(36)
  Fields("fldBillCustRec").Value = arr(37)
  Fields("fldBillNum").Value = arr(38)
  Fields("fldBill2Owner").Value = arr(39)
  Fields("fldTCnt").Value = arr(40)
  Fields("fldBillBal").Value = arr(41)
  Fields("fldDisc").Value = arr(42)
  Fields("fldPenDif").Value = arr(43)
  Fields("fldPenalty").Value = arr(44)
  Fields("fldPenPd").Value = arr(45)
  If arr(10) <> "Billing" Then
    Field17.Visible = False
    Field23.Visible = False
    Field29.Visible = False
    Field26.Visible = False
    Field36.Visible = False
    Field37.Visible = False
    Field38.Visible = False
    Label35.Visible = False
    Field41.Visible = True
    Field39.Visible = True
    Field42.Visible = False
    Field40.Visible = False
    Field44.Visible = False
    Label43.Visible = False
    Field45.Visible = True
  Else
    Field17.Visible = True
    Field23.Visible = True
    Field29.Visible = True
    Field26.Visible = True
    Field36.Visible = True
    Field37.Visible = True
    Field38.Visible = True
    Label35.Visible = True
    Field41.Visible = False
    Field39.Visible = False
    Field42.Visible = True
    Field40.Visible = True
    Field44.Visible = True
    Label43.Visible = True
    Field45.Visible = False
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
      frmVATaxMsg.Label1.Caption = "File - TaxRealHistDet.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - TaxRealHistDet.txt, created in the Citipak Directory."
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
    frmVATaxMsg.Label1.Caption = "File - TaxRealHistDet.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - TaxRealHistDet.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "TaxRealHistDet.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TaxRealHistDet.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmVATaxLoadReport
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
'  Label31.Visible = False
End Sub

Private Sub ReportFooter_Format()
'  Label31.Visible = True
'  Label19.Visible = False
'  Label20.Visible = False
'  Label24.Visible = False
'  Label25.Visible = False
'  Label26.Visible = False
'  Label27.Visible = False
'  Label29.Visible = False
'  Label32.Visible = False
'  Label33.Visible = False
'  Label34.Visible = False
'  Label35.Visible = False
End Sub

Private Sub Detail_Format()
  Dim Opt1 As Boolean
  Dim Opt2 As Boolean
  Dim Opt3 As Boolean

  If QPTrim$(Fields("fldOpt1Desc").Value) = "" Then
    Opt1 = False
  Else
    Opt1 = True
  End If

  If QPTrim$(Fields("fldOpt2Desc").Value) = "" Then
    Opt2 = False
  Else
    Opt2 = True
  End If

  If QPTrim$(Fields("fldOpt3Desc").Value) = "" Then
    Opt3 = False
  Else
    Opt3 = True
  End If

'  Field18.Visible = True
'  Field30.Visible = True
'  Field33.Visible = True
'  Field36.Visible = True
'  Field19.Visible = True
'  Field31.Visible = True
'  Field34.Visible = True
'  Field37.Visible = True
'  Field20.Visible = True
'  Field32.Visible = True
'  Field35.Visible = True
'  Field38.Visible = True
'
  If Opt1 = True And Opt2 = True And Opt3 = True Then Exit Sub
  If Opt1 = False And Opt2 = False And Opt3 = False Then
    Detail.Height = 1990
    Line5.Y1 = 1890
    Line5.Y2 = 1890
    Field18.Visible = False
    Field30.Visible = False
    Field33.Visible = False
    Field36.Visible = False
    Field19.Visible = False
    Field31.Visible = False
    Field34.Visible = False
    Field37.Visible = False
    Field20.Visible = False
    Field32.Visible = False
    Field35.Visible = False
    Field38.Visible = False
    Label40.Top = 1620
    Label41.Top = 1620
    Field41.Top = 1620
    Field42.Top = 1620
    Field44.Top = 1620
    Label43.Top = 1620
    Exit Sub
  End If
  If Opt1 = True And Opt2 = False And Opt3 = False Then
    Detail.Height = 2370
    Line5.Y1 = 2160
    Line5.Y2 = 2160
    Field19.Visible = False
    Field31.Visible = False
    Field34.Visible = False
    Field37.Visible = False
    Field20.Visible = False
    Field32.Visible = False
    Field35.Visible = False
    Field38.Visible = False
    Label40.Top = 1890
    Label41.Top = 1890
    Field41.Top = 1890
    Field42.Top = 1890
    Field39.Top = 1890
    Field40.Top = 1890
    Field44.Top = 1890
    Label43.Top = 1890
  ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
    Detail.Height = 2625
    Line5.Y1 = 2430
    Line5.Y2 = 2430
    Field20.Visible = False
    Field32.Visible = False
    Field35.Visible = False
    Field38.Visible = False
    Label40.Top = 2160
    Label41.Top = 2160
    Field41.Top = 2160
    Field42.Top = 2160
    Field44.Top = 2160
    Label43.Top = 2160
  ElseIf Opt1 = True And Opt2 = False And Opt3 = True Then
    Detail.Height = 2625
    Line5.Y1 = 2430
    Line5.Y2 = 2430
    Field19.Visible = False
    Field31.Visible = False
    Field34.Visible = False
    Field37.Visible = False
    Field20.Top = 1890
    Field25.Top = 1890
    Field32.Top = 1890
    Field38.Top = 1890
    Label40.Top = 2160
    Label41.Top = 2160
    Field41.Top = 2160
    Field42.Top = 2160
    Field44.Top = 2160
    Label43.Top = 2160
  ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
    Detail.Height = 2370
    Line5.Y1 = 2160
    Line5.Y2 = 2160
    Field18.Visible = False
    Field30.Visible = False
    Field33.Visible = False
    Field36.Visible = False
    Field20.Visible = False
    Field32.Visible = False
    Field35.Visible = False
    Field38.Visible = False
    Field19.Top = 1620
    Field24.Top = 1620
    Field31.Top = 1620
    Field38.Top = 1620
    Label40.Top = 1890
    Label41.Top = 1890
    Field41.Top = 1890
    Field42.Top = 1890
    Field44.Top = 1890
    Label43.Top = 1890
  ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
    Detail.Height = 2625
    Line5.Y1 = 2430
    Line5.Y2 = 2430
    Field18.Visible = False
    Field30.Visible = False
    Field33.Visible = False
    Field36.Visible = False
    Field19.Top = 1620
    Field31.Top = 1620
    Field34.Top = 1620
    Field37.Top = 1620
    Field20.Top = 1890
    Field32.Top = 1890
    Field35.Top = 1890
    Field38.Top = 1890
    Label40.Top = 2160
    Label41.Top = 2160
    Field41.Top = 2160
    Field42.Top = 2160
    Field44.Top = 2160
    Label43.Top = 2160
  ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
    Detail.Height = 2370
    Line5.Y1 = 2160
    Line5.Y2 = 2160
    Field18.Visible = False
    Field30.Visible = False
    Field33.Visible = False
    Field36.Visible = False
    Field19.Visible = False
    Field31.Visible = False
    Field34.Visible = False
    Field37.Visible = False
    Field20.Top = 1620
    Field32.Top = 1620
    Field35.Top = 1620
    Field38.Top = 1620
    Label40.Top = 1890
    Label41.Top = 1890
    Field41.Top = 1890
    Field42.Top = 1890
    Field44.Top = 1890
    Label43.Top = 1890
  End If
  
End Sub

