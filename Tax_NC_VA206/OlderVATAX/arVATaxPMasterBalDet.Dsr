VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxPMasterBalDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal Master Balance Report"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arVATaxPMasterBalDet.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   16113
   SectionData     =   "arVATaxPMasterBalDet.dsx":08CA
End
Attribute VB_Name = "arVATaxPMasterBalDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private hFile As Integer
  'Private Temp_Class As Resize_Class
Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\TAXRPTS\TXPMSTBALDET.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldCustName") '1)
  Fields.Add ("fldCustRec") '2)
  Fields.Add ("fldYear") '3)
  Fields.Add ("fldYrAmt") '4)
  Fields.Add ("fldTotEntries") '5)
  Fields.Add ("fldTotBal") '6)
  Fields.Add ("fldPersBal") '7)
  Fields.Add ("fldIntBal") '8)
  Fields.Add ("fldMTBal") '9)
  Fields.Add ("fldMCBal") '10)
  Fields.Add ("fldOpt1Bal") '11)
  Fields.Add ("fldOpt2Bal") '12)
  Fields.Add ("fldOpt3Bal") '13)
  Fields.Add ("fldGPersTot") '14)
  Fields.Add ("fldGIntTot") '15)
  Fields.Add ("fldGMTTot") '16)
  Fields.Add ("fldGMCTot") '17)
  Fields.Add ("fldGOpt1Tot") '18)
  Fields.Add ("fldGOpt2Tot") '19)
  Fields.Add ("fldGOpt3Tot") '20)
  Fields.Add ("fldOpt1Desc") '21)
  Fields.Add ("fldOpt2Desc") '22)
  Fields.Add ("fldOpt3Desc") '23)
  Fields.Add ("fldGOpt") '24)
  Fields.Add ("fldOptDesc") '25)
  Fields.Add ("fldActiveFlag") '26)
  Fields.Add ("fldPropType") '27)
  Fields.Add ("fldThisPin") '28)
  Fields.Add ("fldCustTotBal") '29)
  Fields.Add ("fldFEBal") '30)
  Fields.Add ("fldGFETot") '31)
  Fields.Add ("fldMHBal") '32)
  Fields.Add ("fldGMHTot") '33)
  Fields.Add ("fldPenBal") '34)
  Fields.Add ("fldGPenTot") '35)
  Fields.Add ("fldBillNum") '36)
  Fields.Add ("fldOP") '37)
  Fields.Add ("fldMainYear") '38)
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
  Fields("fldCustName").Value = arr(1)
  Fields("fldCustRec").Value = arr(2)
  Fields("fldYear").Value = arr(3)
  Fields("fldYrAmt").Value = arr(4)
  Fields("fldTotEntries").Value = arr(5)
  Fields("fldTotBal").Value = arr(6)
  Fields("fldPersBal").Value = arr(7)
  Fields("fldIntBal").Value = arr(8)
  Fields("fldMTBal").Value = arr(9)
  Fields("fldMCBal").Value = arr(10)
  Fields("fldOpt1Bal").Value = arr(11)
  Fields("fldOpt2Bal").Value = arr(12)
  Fields("fldOpt3Bal").Value = arr(13)
  Fields("fldGPersTot").Value = arr(14)
  Fields("fldGIntTot").Value = arr(15)
  Fields("fldGMTTot").Value = arr(16)
  Fields("fldGMCTot").Value = arr(17)
  Fields("fldGOpt1Tot").Value = arr(18)
  Fields("fldGOpt2Tot").Value = arr(19)
  Fields("fldGOpt3Tot").Value = arr(20)
  Fields("fldOpt1Desc").Value = arr(21)
  Fields("fldOpt2Desc").Value = arr(22)
  Fields("fldOpt3Desc").Value = arr(23)
  Fields("fldGOpt").Value = arr(24) + ": "
  Fields("fldOptDesc").Value = arr(25)
  If arr(26) = "B" Then
    Fields("fldActiveFlag").Value = "Active And Inactive"
  ElseIf arr(26) = "A" Then
    Fields("fldActiveFlag").Value = "Active Only"
  ElseIf arr(26) = "I" Then
    Fields("fldActiveFlag").Value = "Inactive Only"
  End If
  Fields("fldPropType").Value = arr(27)
  Fields("fldThisPin").Value = arr(28)
  Fields("fldCustTotBal").Value = arr(29)
  Fields("fldFEBal").Value = arr(30)
  Fields("fldGFETot").Value = arr(31)
  Fields("fldMHBal").Value = arr(32)
  Fields("fldGMHTot").Value = arr(33)
  Fields("fldPenBal").Value = arr(34)
  Fields("fldGPenTot").Value = arr(35)
  Fields("fldBillNum").Value = arr(36)
  Fields("fldOP").Value = arr(37)
  Fields("fldMainYear").Value = arr(38)
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
      frmVATaxMsg.Label1.Caption = "File - MasterBalDet.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - MasterBalDet.txt, created in the Citipak Directory."
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
    frmVATaxMsg.Label1.Caption = "File - MasterBalDet.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - MasterBalDet.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "MasterBalDet.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "MasterBalDet.txt"
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
  Me.fldTimeDate.Text = Date
  Me.Zoom = -1
End Sub

Private Sub GroupHeader1_Format()
  If QPTrim$(Fields("fldOptDesc").Value) <> "" Then
    Field31.Visible = True
    Field30.Visible = True
    GroupHeader1.Height = 540
    Line4.Y1 = 540
    Line4.Y2 = 540
  Else
    Field31.Visible = False
    Field30.Visible = False
    GroupHeader1.Height = 270
    Line4.Y1 = 270
    Line4.Y2 = 270
  End If

End Sub

Private Sub PageHeader_Format()
  If Fields("fldActiveFlag").Value = "Active And Inactive" Then
    Label47.Visible = True
  Else
    Label47.Visible = False
  End If
End Sub

Private Sub ReportFooter_Format()
  Dim Opt1 As Boolean
  Dim Opt2 As Boolean
  Dim Opt3 As Boolean
  
  If Fields("fldMainYear").Value <> "All" Then
    Label36.Visible = False
    Field43.Visible = False
  End If
  
  Opt1 = True
  Opt2 = True
  Opt3 = True
  
  If QPTrim$(Fields("fldOpt1Desc").Value) = "" Then
    Opt1 = False
  End If
  
  If QPTrim$(Fields("fldOpt2Desc").Value) = "" Then
    Opt2 = False
  End If
  
  If QPTrim$(Fields("fldOpt3Desc").Value) = "" Then
    Opt3 = False
  End If
  
  If Opt1 = True And Opt2 = True And Opt3 = True Then
    GoTo AllTrue
  End If
  
  If Opt1 = True And Opt2 = False And Opt3 = False Then
    Field25.Visible = False
    Field28.Visible = False
    Field26.Visible = False
    Field29.Visible = False
  End If
  
  If Opt1 = True And Opt2 = True And Opt3 = False Then
    Field26.Visible = False
    Field29.Visible = False
  End If
  
  If Opt1 = True And Opt2 = False And Opt3 = True Then
    Field25.Visible = False
    Field28.Visible = False
    Field26.Top = 1080
    Field29.Top = 1080
  End If
  
  If Opt1 = False And Opt2 = True And Opt3 = True Then
    Field24.Visible = False
    Field27.Visible = False
    Field25.Top = 810
    Field28.Top = 810
    Field26.Top = 1080
    Field29.Top = 1080
  End If
  
  If Opt1 = False And Opt2 = True And Opt3 = False Then
    Field24.Visible = False
    Field27.Visible = False
    Field26.Visible = False
    Field29.Visible = False
    Field25.Top = 810
    Field28.Top = 810
  End If
  
  If Opt1 = False And Opt2 = False And Opt3 = True Then
    Field24.Visible = False
    Field27.Visible = False
    Field25.Visible = False
    Field28.Visible = False
    Field26.Top = 810
    Field29.Top = 810
  End If
  
  If Opt1 = False And Opt2 = False And Opt3 = False Then
    Field24.Visible = False
    Field27.Visible = False
    Field25.Visible = False
    Field28.Visible = False
    Field26.Visible = False
    Field29.Visible = False
  End If
  
AllTrue:
  Line1.Visible = True
  Set SubReport1 = New arVASubTaxPMastBalDet
  Label27.Visible = False
  Label22.Visible = False
  Label50.Visible = False
End Sub

Private Sub Detail_Format()
  Dim Opt1 As Boolean
  Dim Opt2 As Boolean
  Dim Opt3 As Boolean
  
  Line1.Visible = False
  Opt1 = True
  Opt2 = True
  Opt3 = True
  
  If QPTrim$(Fields("fldOpt1Desc").Value) = "" Then
    Field10.Visible = False
    Field11.Visible = False
    Opt1 = False
  End If
  
  If QPTrim$(Fields("fldOpt2Desc").Value) = "" Then
    Field12.Visible = False
    Field13.Visible = False
    Opt2 = False
  End If
  
  If QPTrim$(Fields("fldOpt3Desc").Value) = "" Then
    Field14.Visible = False
    Field15.Visible = False
    Opt3 = False
  End If
  
  If Opt1 = False And Opt2 = False And Opt3 = False Then
    Detail.Height = 795
    Line2.Y1 = 795
    Line2.Y2 = 795
  End If
  
  
End Sub

