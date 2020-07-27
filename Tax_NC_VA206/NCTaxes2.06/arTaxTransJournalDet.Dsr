VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arTaxTransJournalDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Transaction Journal In Detail"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arTaxTransJournalDet.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arTaxTransJournalDet.dsx":08CA
End
Attribute VB_Name = "arTaxTransJournalDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private hFile As Integer
Private Temp_Class As Resize_Class
Dim RptHdr As Boolean

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
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
  Fields.Add ("fldGOpt") '41)
  Fields.Add ("fldOptDesc") '42)
  Fields.Add ("fldThisOperNum") '43)
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
  If arr(16) <> "Billing" Then
    Field17.Visible = False
    Field23.Visible = False
    Field29.Visible = False
    Field26.Visible = False
    Field36.Visible = False
    Field37.Visible = False
    Field38.Visible = False
    Label35.Visible = False
  Else
    Field17.Visible = True
    Field23.Visible = True
    Field29.Visible = True
    Field26.Visible = True
    Field36.Visible = True
    Field37.Visible = True
    Field38.Visible = True
    Label35.Visible = True
  End If
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
  Fields("fldGOpt").Value = arr(41) + ":"
  Fields("fldOptDesc").Value = arr(42)
  Fields("fldThisOperNum").Value = arr(43)
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
'  Unload frmBLLoadReport
  If hFile <> 0 Then
    Close #hFile
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  RptHdr = False
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

Private Sub GroupHeader1_Format()
  If QPTrim$(Fields("fldOptDesc").Value) = "" Then
    GroupHeader1.Height = 270
    Line2.Y1 = 270
    Line2.Y2 = 270
    Field40.Visible = False
    Field41.Visible = False
  Else
    GroupHeader1.Height = 540
    Line2.Y1 = 540
    Line2.Y2 = 540
    Field40.Visible = True
    Field41.Visible = True
  End If
End Sub

Private Sub PageHeader_Format()
  If Fields("fldTransType").Value = "Billing" Then
    Label27.Visible = False
  End If
  If RptHdr = True Then
    Line1.Y1 = 1530
    Line1.Y2 = 1530
    PageHeader.Height = 1575
  End If
End Sub

Private Sub ReportFooter_Format()
  RptHdr = True
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
  Dim Bill As Boolean
  
  Bill = False
'  If Fields("fldTransType").Value = "Billing" Then
'    Bill = True
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
  
  Field18.Visible = True
  Field30.Visible = True
  Field33.Visible = True
  Field36.Visible = True
  Field19.Visible = True
  Field31.Visible = True
  Field34.Visible = True
  Field37.Visible = True
  Field20.Visible = True
  Field32.Visible = True
  Field35.Visible = True
  Field38.Visible = True
  Label37.Visible = True 'Interest
  Field21.Visible = True
  Field22.Visible = True
  Field23.Visible = True
  Label38.Visible = True 'advertising
  Field27.Visible = True
  Field28.Visible = True
  Field29.Visible = True
    
  If Bill = True Then
    Detail.Height = 1800
    Label37.Visible = False 'Interest
    Field21.Visible = False
    Field22.Visible = False
    Field23.Visible = False
    Label38.Visible = False 'advertising
    Field27.Visible = False
    Field28.Visible = False
    Field29.Visible = False
    Label39.Top = 540 'late list
    Field24.Top = 540
    Field25.Top = 540
    Field26.Top = 540
    Field18.Top = 810 'opt 1
    Field30.Top = 810
    Field33.Top = 810
    Field36.Top = 810
    Field19.Top = 1080 'opt 2
    Field31.Top = 1080
    Field34.Top = 1080
    Field37.Top = 1080
    Field20.Top = 1350 'opt 3
    Field32.Top = 1350
    Field35.Top = 1350
    Field38.Top = 1350
    Field13.Visible = False
  Else
    Field23.Visible = False 'balance fields
    Field29.Visible = False
    Field26.Visible = False
    Field36.Visible = False
    Field37.Visible = False
    Field38.Visible = False
  End If
    
  If Opt1 = True And Opt2 = True And Opt3 = True Then Exit Sub
  If Opt1 = False And Opt2 = False And Opt3 = False Then
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
    If Bill = False Then
      Detail.Height = 1515
    Else
      Detail.Height = 975
    End If
    Exit Sub
  End If
  If Opt1 = True And Opt2 = False And Opt3 = False Then
    Field19.Visible = False
    Field31.Visible = False
    Field34.Visible = False
    Field37.Visible = False
    Field20.Visible = False
    Field32.Visible = False
    Field35.Visible = False
    Field38.Visible = False
    If Bill = False Then
      Detail.Height = 1770
    Else
      Detail.Height = 1230
      Field18.Top = 810 'opt 1
      Field30.Top = 810
      Field33.Top = 810
      Field36.Top = 810
    End If
  ElseIf Opt1 = True And Opt2 = True And Opt3 = False Then
    Field20.Visible = False
    Field32.Visible = False
    Field35.Visible = False
    Field38.Visible = False
    If Bill = False Then
      Detail.Height = 2025
    Else
      Detail.Height = 1500
      Field18.Top = 810 'opt 1
      Field30.Top = 810
      Field33.Top = 810
      Field36.Top = 810
      Field19.Top = 1080 'opt 2
      Field31.Top = 1080
      Field34.Top = 1080
      Field37.Top = 1080
    End If
  ElseIf Opt1 = True And Opt2 = False And Opt3 = True Then
    Field19.Visible = False
    Field31.Visible = False
    Field34.Visible = False
    Field37.Visible = False
    If Bill = False Then
      Detail.Height = 2025
      Field20.Top = 1620
      Field35.Top = 1620
      Field32.Top = 1620
      Field38.Top = 1620
    Else
      Detail.Height = 1500
      Field18.Top = 810 'opt 1
      Field30.Top = 810
      Field33.Top = 810
      Field36.Top = 810
      Field20.Top = 1080 'opt 3
      Field32.Top = 1080
      Field35.Top = 1080
      Field38.Top = 1080
    End If
  ElseIf Opt1 = False And Opt2 = True And Opt3 = False Then
    Field18.Visible = False
    Field30.Visible = False
    Field33.Visible = False
    Field36.Visible = False
    Field20.Visible = False
    Field32.Visible = False
    Field35.Visible = False
    Field38.Visible = False
    If Bill = False Then
      Detail.Height = 1770
      Field19.Top = 1350
      Field24.Top = 1350
      Field31.Top = 1350
      Field38.Top = 1350
    Else
      Detail.Height = 1230
      Field19.Top = 810 'opt 2
      Field31.Top = 810
      Field34.Top = 810
      Field37.Top = 810
    End If
  ElseIf Opt1 = False And Opt2 = True And Opt3 = True Then
    Field18.Visible = False
    Field30.Visible = False
    Field33.Visible = False
    Field36.Visible = False
    If Bill = False Then
      Detail.Height = 2025
      Field19.Top = 1350
      Field31.Top = 1350
      Field34.Top = 1350
      Field37.Top = 1350
      Field20.Top = 1620
      Field32.Top = 1620
      Field35.Top = 1620
      Field38.Top = 1620
    Else
      Detail.Height = 1485
      Field19.Top = 810 'opt 2
      Field31.Top = 810
      Field34.Top = 810
      Field37.Top = 810
      Field20.Top = 1080 'opt 3
      Field32.Top = 1080
      Field35.Top = 1080
      Field38.Top = 1080
    End If
  ElseIf Opt1 = False And Opt2 = False And Opt3 = True Then
    Field18.Visible = False
    Field30.Visible = False
    Field33.Visible = False
    Field36.Visible = False
    Field19.Visible = False
    Field31.Visible = False
    Field34.Visible = False
    Field37.Visible = False
    If Bill = False Then
      Detail.Height = 1770
      Field20.Top = 1350
      Field32.Top = 1350
      Field35.Top = 1350
      Field38.Top = 1350
    Else
      Detail.Height = 1230
      Field20.Top = 810 'opt 3
      Field32.Top = 810
      Field35.Top = 810
      Field38.Top = 810
    End If
  End If
  
End Sub
