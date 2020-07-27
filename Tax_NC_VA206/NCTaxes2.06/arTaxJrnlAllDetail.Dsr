VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arTaxJrnlAllDetail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Transaction Journal"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arTaxJrnlAllDetail.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arTaxJrnlAllDetail.dsx":08CA
End
Attribute VB_Name = "arTaxJrnlAllDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private hFile As Integer
Private Temp_Class As Resize_Class
Dim BillNum As Long
Dim RptFtr As Boolean

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  RptFtr = False
  hFile = FreeFile
  Open StartPath & "\TAXRPTS\TXJRLDT.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldCustName") '1)
  Fields.Add ("fldCustNum") '2)
  Fields.Add ("fldActive") '3)
  Fields.Add ("fldTransDate") '4)
  Fields.Add ("fldBillType") '5)
  Fields.Add ("fldTransType") '6)
  Fields.Add ("fldBegDate") '7)
  Fields.Add ("fldEndDate") '8)
  Fields.Add ("fldTaxYear") '9)
  Fields.Add ("fldAmount") '10)
  Fields.Add ("fldTCnt") '11)
  Fields.Add ("fldTotAmt") '12)
  Fields.Add ("fldPrePdAmt") '13)
  Fields.Add ("fldBillNum") '14)
  Fields.Add ("fldDesc") '15)
  Fields.Add ("fldThisTransType") '16)
  Fields.Add ("fldPrinc") '17)
  Fields.Add ("fldPrincPd") '18)
  Fields.Add ("fldPrincDif") '19)
  Fields.Add ("fldInt") '20)
  Fields.Add ("fldIntPd") '21)
  Fields.Add ("fldIntDif") '22)
  Fields.Add ("fldAdv") '23)
  Fields.Add ("fldAdvPd") '24)
  Fields.Add ("fldAdvDif") '25)
  Fields.Add ("fldLateList") '26)
  Fields.Add ("fldLateListPd") '27)
  Fields.Add ("fldLateListDif") '28)
  Fields.Add ("fldOpt1") '29)
  Fields.Add ("fldOpt1Pd") '30)
  Fields.Add ("fldOpt1Dif") '31)
  Fields.Add ("fldOpt2") '32)
  Fields.Add ("fldOpt2Pd") '33)
  Fields.Add ("fldOpt2Dif") '34)
  Fields.Add ("fldOpt3") '35)
  Fields.Add ("fldOpt3Pd") '36)
  Fields.Add ("fldOpt3Dif") '37)
  Fields.Add ("fldOpt1Desc") '38)
  Fields.Add ("fldOpt2Desc") '39)
  Fields.Add ("fldOpt3Desc") '40)
  Fields.Add ("fldBill") '41)
  Fields.Add ("fldCustBal") '42)
  Fields.Add ("fldBillBal") '43)
  Fields.Add ("fldTranType") '44)
  Fields.Add ("fldThisOperNum") '45)
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
'    Unload frmLoadReport
    frmTaxMsg.Label1.Caption = "Error Number: " & Str(Number) & " " & Description
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
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
  Fields("fldCustName").Value = arr(1)
  Fields("fldCustNum").Value = arr(2)
  Fields("fldActive").Value = arr(3)
  Fields("fldTransDate").Value = arr(4)
  Fields("fldBillType").Value = arr(5)
  Fields("fldTransType").Value = arr(6)
  Fields("fldBegDate").Value = arr(7)
  Fields("fldEndDate").Value = arr(8)
  Fields("fldTaxYear").Value = arr(9)
  Fields("fldAmount").Value = arr(10)
  Fields("fldTCnt").Value = arr(11)
  Fields("fldTotAmt").Value = arr(12)
  Fields("fldPrePdAmt").Value = arr(13)
  Fields("fldBillNum").Value = arr(14)
  Fields("fldDesc").Value = arr(15)
  Fields("fldThisTransType").Value = arr(16)
  Fields("fldPrinc").Value = arr(17)
  Fields("fldPrincPd").Value = arr(18)
  Fields("fldPrincDif").Value = arr(19)
  Fields("fldInt").Value = arr(20)
  Fields("fldIntPd").Value = arr(21)
  Fields("fldIntDif").Value = arr(22)
  Fields("fldAdv").Value = arr(23)
  Fields("fldAdvPd").Value = arr(24)
  Fields("fldAdvDif").Value = arr(25)
  Fields("fldLateList").Value = arr(26)
  Fields("fldLateListPd").Value = arr(27)
  Fields("fldLateListDif").Value = arr(28)
  Fields("fldOpt1").Value = arr(29)
  Fields("fldOpt1Pd").Value = arr(30)
  Fields("fldOpt1Dif").Value = arr(31)
  Fields("fldOpt2").Value = arr(32)
  Fields("fldOpt2Pd").Value = arr(33)
  Fields("fldOpt2Dif").Value = arr(34)
  Fields("fldOpt3").Value = arr(35)
  Fields("fldOpt3Pd").Value = arr(36)
  Fields("fldOpt3Dif").Value = arr(37)
  Fields("fldOpt1Desc").Value = arr(38)
  Fields("fldOpt2Desc").Value = arr(39)
  Fields("fldOpt3Desc").Value = arr(40)
  Fields("fldBill").Value = arr(41)
  Fields("fldCustBal").Value = arr(42)
  Fields("fldBillBal").Value = arr(43)
  Fields("fldTranType").Value = arr(44)
  Fields("fldThisOperNum").Value = arr(45)
End Sub

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
      frmTaxMsg.Label1.Caption = "File - TaxTransJrnlRptDet.xls, created in the Citipak Directory."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmTaxMsg.Label1.Caption = "File - TaxTransJrnlRptDet.txt, created in the Citipak Directory."
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
    frmTaxMsg.Label1.Caption = "File - TaxTransJrnlRptDet.xls, created in the Citipak Directory."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmTaxMsg.Label1.Caption = "File - TaxTransJrnlRptDet.txt, created in the Citipak Directory."
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
    Case 1   '"Excel"
        Set oEXL = New ActiveReportsExcelExport.ARExportExcel
        oEXL.FileName = outfile & "TaxTransJrnlRptDet.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "TaxTransJrnlRptDet.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_ReportEnd()
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
  Label31.Visible = False
End Sub

Private Sub PageHeader_Format()
  If RptFtr = True Then
    Line1.Y1 = 1800
    Line1.Y2 = 1800
    PageHeader.Height = 1860
  End If
End Sub

Private Sub ReportFooter_Format()
  RptFtr = True
  Set SubReport1.object = New arSubTaxJrnlDetAll
  Label31.Visible = True
  Label19.Visible = False
  Label20.Visible = False
  Label24.Visible = False
  Label25.Visible = False
  Label26.Visible = False
  Label27.Visible = False
  Label29.Visible = False
  Label32.Visible = False
  Label33.Visible = False
  Label34.Visible = False
  Label35.Visible = False
  Set SubReport2.object = New arSub2TaxJrnlDetAll
End Sub
Private Sub Detail_Format()
  Dim Opt1 As Boolean
  Dim Opt2 As Boolean
  Dim Opt3 As Boolean
  
  If BillNum <> CLng(Fields("fldBill").Value) Then
    Detail.Height = 0
    Field8.Visible = False
    Field9.Visible = False
    Field10.Visible = False
    Field11.Visible = False
    Field12.Visible = False
    Field13.Visible = False
    Field14.Visible = False

    Label36.Visible = False
    Field15.Visible = False
    Field16.Visible = False

    Label37.Visible = False
    Field21.Visible = False
    Field22.Visible = False

    Label38.Visible = False
    Field27.Visible = False
    Field28.Visible = False

    Label39.Visible = False
    Field24.Visible = False
    Field25.Visible = False

    Field12.Visible = False
    Field13.Visible = False
    Field14.Visible = False

    Field18.Visible = False
    Field30.Visible = False
    Field33.Visible = False

    Field19.Visible = False
    Field31.Visible = False
    Field34.Visible = False

    Field20.Visible = False
    Field32.Visible = False
    Field35.Visible = False
    Label27.Visible = False
    Line5.Visible = False
  Else
    Detail.Height = 2250
    Field8.Visible = True
    Field9.Visible = True
    Field10.Visible = True
    Field11.Visible = True
    Field12.Visible = True
    Field13.Visible = True
    Field14.Visible = True

    Label36.Visible = True
    Field15.Visible = True
    Field16.Visible = True

    Label37.Visible = True
    Field21.Visible = True
    Field22.Visible = True

    Label38.Visible = True
    Field27.Visible = True
    Field28.Visible = True

    Label39.Visible = True
    Field24.Visible = True
    Field25.Visible = True

    Field12.Visible = True
    Field13.Visible = True
    Field14.Visible = True

    Field18.Visible = True
    Field30.Visible = True
    Field33.Visible = True

    Field19.Visible = True
    Field31.Visible = True
    Field34.Visible = True

    Field20.Visible = True
    Field32.Visible = True
    Field35.Visible = True
    Label27.Visible = True
    Line5.Visible = True
  End If
  BillNum = CLng(Fields("fldBill").Value)

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

  If Opt1 = True And Opt2 = True And Opt3 = True Then Exit Sub
  If Opt1 = False And Opt2 = False And Opt3 = False Then
    Detail.Height = 1425
    Field18.Visible = False
    Field30.Visible = False
    Field33.Visible = False
    Field19.Visible = False
    Field31.Visible = False
    Field34.Visible = False
    Field20.Visible = False
    Field32.Visible = False
    Field35.Visible = False
    Exit Sub
  End If

  If Opt1 = True And Opt2 = False And Opt3 = False Then
    Detail.Height = 1695
    Field19.Visible = False
    Field31.Visible = False
    Field34.Visible = False
    Field20.Visible = False
    Field32.Visible = False
    Field35.Visible = False
  ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
    Detail.Height = 1965
    Field20.Visible = False
    Field32.Visible = False
    Field35.Visible = False
  ElseIf Opt1 = True And Opt2 = False And Opt3 = True Then
    Detail.Height = 1965
    Field19.Visible = False
    Field31.Visible = False
    Field34.Visible = False
    Field20.Top = 1440
    Field32.Top = 1440
    Field35.Top = 1440
  ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
    Detail.Height = 1695
    Field18.Visible = False
    Field30.Visible = False
    Field33.Visible = False
    Field20.Visible = False
    Field32.Visible = False
    Field35.Visible = False
    Field19.Top = 1440
    Field31.Top = 1440
    Field34.Top = 1440
  ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
    Detail.Height = 1965
    Field18.Visible = False
    Field30.Visible = False
    Field33.Visible = False
    Field19.Top = 1440
    Field31.Top = 1440
    Field34.Top = 1440
    Field20.Top = 1710
    Field32.Top = 1710
    Field35.Top = 1710
  ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
    Detail.Height = 1695
    Field18.Visible = False
    Field30.Visible = False
    Field33.Visible = False
    Field19.Visible = False
    Field31.Visible = False
    Field34.Visible = False
    Field20.Top = 1440
    Field32.Top = 1440
    Field35.Top = 1440
  End If
  
End Sub

Private Sub GroupHeader2_Format()
  Dim Opt1 As Boolean
  Dim Opt2 As Boolean
  Dim Opt3 As Boolean
  
  'commented out 1/18/07
'  If QPTrim$(Fields("fldTranType").Value) = "12" Or QPTrim$(Fields("fldTranType").Value) = "11" Then
'    Field94.Visible = False
'    Field95.Visible = False
'    Field96.Visible = False
'    Field97.Visible = False
'    Field98.Visible = False
'    Field100.Visible = False
'    GroupHeader2.Height = 0
'    Exit Sub
'  Else
'    Field94.Visible = True
'    Field95.Visible = True
'    Field96.Visible = True
'    Field97.Visible = True
'    Field98.Visible = True
'    Field100.Visible = True
'    GroupHeader2.Height = 2430
'  End If
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
  
  Field72.Visible = True
  Field81.Visible = True
  Field84.Visible = True
  Field91.Visible = True
  Field91.Visible = True
  Field73.Visible = True
  Field82.Visible = True
  Field85.Visible = True
  Field92.Visible = True
  Field74.Visible = True
  Field83.Visible = True
  Field86.Visible = True
  Field93.Visible = True
  
  If Opt1 = True And Opt2 = True And Opt3 = True Then Exit Sub
  If Opt1 = False And Opt2 = False And Opt3 = False Then
    GroupHeader2.Height = 1605
    Field72.Visible = False
    Field81.Visible = False
    Field84.Visible = False
    Field91.Visible = False
    Field73.Visible = False
    Field82.Visible = False
    Field85.Visible = False
    Field92.Visible = False
    Field74.Visible = False
    Field83.Visible = False
    Field86.Visible = False
    Field93.Visible = False
    Label45.Top = 1350
    Field102.Top = 1350
    Exit Sub
  End If
  
  If Opt1 = True And Opt2 = False And Opt3 = False Then
    GroupHeader2.Height = 1860
    Field73.Visible = False
    Field82.Visible = False
    Field85.Visible = False
    Field92.Visible = False
    Field74.Visible = False
    Field83.Visible = False
    Field86.Visible = False
    Field93.Visible = False
    Label45.Top = 1620
    Field102.Top = 1620
  ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
    GroupHeader2.Height = 2115
    Field74.Visible = False
    Field83.Visible = False
    Field86.Visible = False
    Field93.Visible = False
    Label45.Top = 1890
    Field102.Top = 1890
  ElseIf Opt1 = True And Opt2 = False And Opt3 = True Then
    GroupHeader2.Height = 2115
    Field73.Visible = False
    Field82.Visible = False
    Field85.Visible = False
    Field92.Visible = False
    Field74.Top = 1350
    Field83.Top = 1350
    Field86.Top = 1350
    Field93.Top = 1350
    Label45.Top = 1620
    Field102.Top = 1620
  ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
    GroupHeader2.Height = 1860
    Field72.Visible = False
    Field81.Visible = False
    Field84.Visible = False
    Field91.Visible = False
    Field74.Visible = False
    Field83.Visible = False
    Field86.Visible = False
    Field93.Visible = False
    Field73.Top = 1350
    Field82.Top = 1350
    Field85.Top = 1350
    Field92.Top = 1350
    Label45.Top = 1620
    Field102.Top = 1620
  ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
    GroupHeader2.Height = 1860
    Field72.Visible = False
    Field81.Visible = False
    Field84.Visible = False
    Field94.Visible = False
    Field73.Top = 1350
    Field82.Top = 1350
    Field85.Top = 1350
    Field92.Top = 1350
    Field74.Top = 1620
    Field83.Top = 1620
    Field86.Top = 1620
    Field93.Top = 1620
    Label45.Top = 1890
    Field102.Top = 1890
  ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
    GroupHeader2.Height = 1860
    Field72.Visible = False
    Field81.Visible = False
    Field84.Visible = False
    Field91.Visible = False
    Field73.Visible = False
    Field82.Visible = False
    Field85.Visible = False
    Field92.Visible = False
    Field74.Top = 1350
    Field83.Top = 1350
    Field86.Top = 1350
    Field93.Top = 1350
    Label45.Top = 1620
    Field102.Top = 1620
  End If

End Sub
