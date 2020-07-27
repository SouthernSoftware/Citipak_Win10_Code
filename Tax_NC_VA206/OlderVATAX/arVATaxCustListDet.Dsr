VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVATaxCustListDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Master Customer List"
   ClientHeight    =   8736
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "arVATaxCustListDet.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20532
   _ExtentY        =   15399
   SectionData     =   "arVATaxCustListDet.dsx":08CA
End
Attribute VB_Name = "arVATaxCustListDet"
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
  Open StartPath & "\TAXRPTS\CSTLSTDT.RPT" For Input As #hFile
  Fields.Add ("fldTown") '0)
  Fields.Add ("fldCustAcct") '1)
  Fields.Add ("fldCustName") '2)
  Fields.Add ("fldAdd1") '3)
  Fields.Add ("fldAdd2") '4)
  Fields.Add ("fldCSZ") '5)
  Fields.Add ("fldIntYN") '6)
  Fields.Add ("fldBankruptYN") '7)
  Fields.Add ("fldDrvrsLic") '8)
  Fields.Add ("fldHPhone") '9)
  Fields.Add ("fldWPhone") '10)
  Fields.Add ("fldLateNotYN") '11)
  Fields.Add ("fldTaxExYN") '12)
  Fields.Add ("fldPIN") '13)
  Fields.Add ("fldNote1") '14)
  Fields.Add ("fldNote2") '15)
  Fields.Add ("fldNote3") '16)
  Fields.Add ("fldNote4") '17)
  Fields.Add ("fldNote5") '18)
  Fields.Add ("fldGisPos") '19)
  Fields.Add ("fldMapBlkLot") '20)
  Fields.Add ("fldOtherX") '21)
  Fields.Add ("fldSeniorX") '22)
  Fields.Add ("fldLateListYN") '23)
  Fields.Add ("fldLien") '24)
  Fields.Add ("fldMortCode") '25)
  Fields.Add ("fldPropAdd") '26)
  Fields.Add ("fldCValue") '27)
  Fields.Add ("fldMcValue") '28)
  Fields.Add ("fldMhValue") '29)
  Fields.Add ("fldMtValue") '30)
  Fields.Add ("fldPersVal") '31)
  Fields.Add ("fldPropType") '32)
  Fields.Add ("fldCustCnt") '33)
  Fields.Add ("fldActive") '34)
  Fields.Add ("fldLienDesc") '35)
  Fields.Add ("fldTownship") '36)
  Fields.Add ("fldTSHeader") '37)
  Fields.Add ("fldOptDesc") '38)
  Fields.Add ("fldGOptDesc") '39)
  Fields.Add ("fldActiveFlag") '40)
  Fields.Add ("fldFlagType") '41)
  Fields.Add ("fldRealVal") '42)
  Fields.Add ("fldBldgVal") '43)
End Sub
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmVATaxLoadReport
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
  Fields("fldCustAcct").Value = arr(1)
  Fields("fldCustName").Value = arr(2)
  Fields("fldAdd1").Value = arr(3)
  Fields("fldAdd2").Value = arr(4)
  Fields("fldCSZ").Value = arr(5)
  Fields("fldIntYN").Value = arr(6)
  Fields("fldBankruptYN").Value = arr(7)
  Fields("fldDrvrsLic").Value = arr(8)
  If QPTrim$(arr(9)) = "(" Then arr(9) = ""
  Fields("fldHPhone").Value = arr(9)
  If QPTrim$(arr(10)) = "(" Then arr(10) = ""
  Fields("fldWPhone").Value = arr(10)
  Fields("fldLateNotYN").Value = arr(11)
  Fields("fldTaxExYN").Value = arr(12)
  Fields("fldPIN").Value = arr(13)
  Fields("fldNote1").Value = arr(14)
  Fields("fldNote2").Value = arr(15)
  Fields("fldNote3").Value = arr(16)
  Fields("fldNote4").Value = arr(17)
  Fields("fldNote5").Value = arr(18)
  Fields("fldGisPos").Value = arr(19)
  Fields("fldMapBlkLot").Value = arr(20)
  Fields("fldOtherX").Value = arr(21)
  Fields("fldSeniorX").Value = arr(22)
  Fields("fldLateListYN").Value = arr(23)
  Fields("fldLien").Value = arr(24)
  Fields("fldMortCode").Value = arr(25)
  Fields("fldPropAdd").Value = arr(26)
  Fields("fldCValue").Value = arr(27)
  Fields("fldMcValue").Value = arr(28)
  Fields("fldMhValue").Value = arr(29)
  Fields("fldMtValue").Value = arr(30)
  Fields("fldPersVal").Value = arr(31)
  Fields("fldPropType").Value = arr(32)
  Fields("fldCustCnt").Value = arr(33)
  Fields("fldActive").Value = arr(34)
  Fields("fldLienDesc").Value = arr(35)
  Fields("fldTownship").Value = arr(36)
  Fields("fldTSHeader").Value = arr(37)
  Fields("fldOptDesc").Value = arr(38)
  Fields("fldGOptDesc").Value = arr(39) + ": "
  If QPTrim$(arr(38)) = "" Then
    Field42.Visible = False
    Field43.Visible = False
  Else
    Field42.Visible = True
    Field43.Visible = True
  End If
  Fields("fldActiveFlag").Value = arr(40)
  Fields("fldFlagType").Value = arr(41)
  Fields("fldRealVal").Value = arr(42)
  Fields("fldBldgVal").Value = arr(43)
  
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
      frmVATaxMsg.Label1.Caption = "File - MasterCustDet.xls, created in the Citipak Directory."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmVATaxMsg.Label1.Caption = "File - MasterCustDet.txt, created in the Citipak Directory."
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
    frmVATaxMsg.Label1.Caption = "File - MasterCustDet.xls, created in the Citipak Directory."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmVATaxMsg.Label1.Caption = "File - MasterCustDet.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "MasterCustDet.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "MasterCustDet.txt"
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

Private Sub Detail_Format()
  Detail.Height = 1700
  If Fields("fldPropType").Value = "Real" Then
    Detail.Height = 1900
    Label54.Visible = False
    Label57.Visible = False
    Label56.Visible = False
    Label53.Visible = False
    Label52.Visible = False
    Label55.Visible = False
    Field28.Visible = False
    Field31.Visible = False
    Field30.Visible = False
    Field32.Visible = False
    Field27.Visible = False
    Field33.Visible = False
    Field26.Visible = False
    Field34.Visible = False
    Field29.Visible = False
    Field35.Visible = False
    
    Label43.Visible = True
    Label43.Top = 270
    Label42.Visible = True
    Label42.Top = 540
    Label44.Visible = True
    Label44.Top = 810
    Label45.Visible = True
    Label45.Top = 1080
    Label46.Visible = True
    Label46.Top = 1350
    Label47.Visible = True
    Label47.Top = 810
    Label48.Visible = True
    Label48.Top = 1080
    Label49.Visible = True
    Label49.Top = 1350
    Label50.Visible = True
    Label50.Top = 270
    Label58.Visible = True
    Label58.Top = 540
    Field16.Visible = True
    Field16.Top = 270
    Field17.Visible = True
    Field17.Top = 540
    Field18.Visible = True
    Field18.Top = 810
    Field19.Visible = True
    Field19.Top = 1080
    Field20.Visible = True
    Field20.Top = 1350
    Field21.Visible = True
    Field21.Top = 810
    Field22.Visible = True
    Field22.Top = 1080
    Field23.Visible = True
    Field23.Top = 1350
    Field24.Visible = True
    Field24.Top = 270
    Field36.Visible = True
    Field36.Top = 540
    Field37.Visible = True
    Field37.Top = 810
    Field38.Visible = True
    Field38.Top = 1080
    Label64.Visible = True
    Label64.Top = 1620
    Field46.Visible = True
    Field46.Top = 1620
    Label65.Visible = True
    Label65.Top = 1620
    Field47.Visible = True
    Field47.Top = 1620
    Line9.Y1 = 1890
    Line9.Y2 = 1890
  Else
    Label54.Visible = True
    Label57.Visible = True
    Label56.Visible = True
    Label53.Visible = True
    Label52.Visible = True
    Label55.Visible = True
    Field28.Visible = True
    Field31.Visible = True
    Field30.Visible = True
    Field32.Visible = True
    Field27.Visible = True
    Field33.Visible = True
    Field26.Visible = True
    Field34.Visible = True
    Field29.Visible = True
    Field35.Visible = True
    
    Label43.Visible = False
    Label42.Visible = False
    Label44.Visible = False
    Label45.Visible = False
    Label46.Visible = False
    Label47.Visible = False
    Label48.Visible = False
    Label49.Visible = False
    Label50.Visible = False
    Label58.Visible = False
    Field16.Visible = False
    Field17.Visible = False
    Field18.Visible = False
    Field19.Visible = False
    Field20.Visible = False
    Field21.Visible = False
    Field22.Visible = False
    Field23.Visible = False
    Field24.Visible = False
    Field36.Visible = False
    Field37.Visible = False
    Field38.Visible = False
    Label64.Visible = False
    Field46.Visible = False
    Label65.Visible = False
    Field47.Visible = False
    Line9.Y1 = 1620
    Line9.Y2 = 1620
  End If
End Sub

Private Sub PageHeader_Format()
  If Fields("fldFlagType").Value = "None" Then
    Field45.Visible = False
    Label63.Visible = False
  End If
  If Fields("fldActiveFlag").Value = "B" Then
    Field44.Text = "Active and Inactive"
  ElseIf Fields("fldActiveFlag").Value = "A" Then
    Field44.Text = "Active Only"
  ElseIf Fields("fldActiveFlag").Value = "I" Then
    Field44.Text = "Inactive Only"
  End If
  

End Sub
