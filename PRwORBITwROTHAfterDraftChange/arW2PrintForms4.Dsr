VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arW2PrintForms4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "W2 Print Forms"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "arW2PrintForms4.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arW2PrintForms4.dsx":08CA
End
Attribute VB_Name = "arW2PrintForms4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private hFile As Integer
Private Sub ActiveReport_Initialize()
  Me.Toolbar.Tools.Add "&Close"
  Me.Toolbar.Tools.Add "Save/&Excel"
  Me.Toolbar.Tools.Add "&Text"
End Sub
Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
    DoEvents
    frmW2FormsPrinting.Show
  End If
  If Shift = 4 Then
    If KeyCode = vbKeyC Then
      Unload Me
      DoEvents
      frmW2FormsPrinting.Show
    ElseIf KeyCode = vbKeyE Then
      Screen.MousePointer = vbHourglass
      ExportReport 1
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - W2FormsRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - W2FormsRpt.txt, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_Terminate()
  Close '5/28/2004
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
  If Tool.Caption = "&Close" Then
    Unload Me
    DoEvents
    frmW2FormsPrinting.Show
  End If
  If Tool = "Save/&Excel" Then
    Screen.MousePointer = vbHourglass
    ExportReport 1
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - W2FormsRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - W2FormsRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "W2FormsRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "W2FormsRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub
Private Sub ActiveReport_DataInitialize()
  hFile = FreeFile
  Open StartPath & "\PRRPTS\W2FRMS4.RPT" For Input As #hFile
  Fields.Add "fldCntrNum" '(0)
  Fields.Add "fldFedID" '(1)
  Fields.Add "fldFedWage" '(2)
  Fields.Add "fldFedTaxWH" '(3)
  Fields.Add "fldEmployer" '(4)
  Fields.Add "fldSOCWAGE" '(5)
  Fields.Add "fldSOCTAXWH" '(6)
  Fields.Add "fldAddr1" '(7)
  Fields.Add "fldAddr2" '(8)
  Fields.Add "fldMEDWAGES" '(9)
  Fields.Add "fldMEDTAXWH" '(10)
  Fields.Add "fldCity" '(11)
  Fields.Add "fldState" '(12)
  Fields.Add "fldZip" '(13)
  Fields.Add "fldSocTip" '(14)
  Fields.Add "fldAlocTip" '(15)
  Fields.Add "fldEmpSSN" '(16)
  Fields.Add "fldAdvEicP" '(17)
  Fields.Add "fldDepCare" '(18)
  Fields.Add "fldEmpFName" '(19)
  Fields.Add "fldEmpLName" '(20)
  Fields.Add "fldNQP" '(21)
  Fields.Add "fldBOX13TXt" '(22)
  Fields.Add "fldBox13Amt1" '(23)
  Fields.Add "fldEmpAddr1" '(24)
  Fields.Add "fldEmpAddr2" '(25)
  Fields.Add "fldBOX15A" '(26)
  Fields.Add "fldBOX15c" '(27)
  Fields.Add "fldBOX15G" '(28)
  Fields.Add "fldBOX13TX1" '(29)
  Fields.Add "fldBox13Amt1vs2" '(30)
  Fields.Add "fldEmpCity" '(31)
  Fields.Add "fldEmpState" '(32)
  Fields.Add "fldEmpZip" '(33)
  Fields.Add "fldBTxt14" '(34)
  Fields.Add "fldBox14Amt1" '(35)
  Fields.Add "fldBTxt14vs2" '(36)
  Fields.Add "fldBox14Amt1vs2" '(37)
  Fields.Add "fldState2" '(38)
  Fields.Add "fldSTAID" '(39)
  Fields.Add "fldStateWage" '(40)
  Fields.Add "fldStateTax" '(41)
  Fields.Add "fldBOX13AM2" '(42)
  Fields.Add "fldBox13TX2" '(43)
  Fields.Add "fldBOX13AM3" '(44)
  Fields.Add "fldBox13TX3" '(45)
  End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  Unload frmLoadingRpt
  CancelDisplay = True 'removes the error message
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  
  ' We reached the end of the file we exit leaving the
  ' eof parameter as True (default except on first call) that will
  ' tell AR that we are done feeding data
  ' otherwise we have to set the eof parameter to False so that
  ' AR continues fetching data, until we're done
  ' if the report had a data control, the value of the parameter
  ' will be ignored, AR will always follow the data control's recordset
  ' EOF property
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
  Fields("fldCntrNum").Value = arr(0)
  Fields("fldFedID").Value = arr(1)
  Fields("fldFedWage").Value = arr(2)
  Fields("fldFedTaxWH").Value = arr(3)
  Fields("fldEmployer").Value = arr(4)
  Fields("fldSOCWAGE").Value = arr(5)
  Fields("fldSOCTAXWH").Value = arr(6)
  Fields("fldAddr1").Value = arr(7)
  Fields("fldAddr2").Value = arr(8)
  Fields("fldMEDWAGES").Value = arr(9)
  Fields("fldMEDTAXWH").Value = arr(10)
  Fields("fldCity").Value = arr(11)
  Fields("fldState").Value = arr(12)
  Fields("fldZip").Value = arr(13)
'  If arr(14) = "0" Then arr(14) = ""
  Fields("fldSocTip").Value = arr(14)
'  If arr(15) = "0" Then arr(15) = ""
  Fields("fldAlocTip").Value = arr(15)
  Fields("fldEmpSSN").Value = arr(16)
'  If arr(17) = "0" Then arr(17) = ""
  Fields("fldAdvEicP").Value = arr(17)
'  If arr(18) = "0" Then arr(18) = ""
  Fields("fldDepCare").Value = arr(18)
'  Fields("fldEmpFName").Value = arr(19)
'  Fields("fldEmpLName").Value = arr(20)
  Fields("fldEmpFName").Value = QPTrim$(arr(19)) + " " + QPTrim$(arr(20))
'  If arr(21) = "0" Then arr(21) = ""
  Fields("fldNQP").Value = arr(21)
  Fields("fldBOX13TXt").Value = arr(22)
'  If arr(23) = "0" Then arr(23) = ""
  Fields("fldBox13Amt1").Value = arr(23)
  Fields("fldEmpAddr1").Value = QPTrim$(arr(24))
  Fields("fldEmpAddr2").Value = QPTrim$(arr(25))
  Fields("fldBOX15A").Value = arr(26)
  Fields("fldBOX15c").Value = arr(27)
  Fields("fldBOX15G").Value = arr(28)
  Fields("fldBOX13TX1").Value = arr(29)
'  If arr(30) = "0" Then arr(30) = ""
  Fields("fldBox13Amt1vs2").Value = arr(30)
  Fields("fldEmpCity").Value = QPTrim$(arr(31)) + ", " + QPTrim$(arr(32)) + " " + QPTrim$(arr(33))
'  Fields("fldEmpState").Value = arr(32)
'  Fields("fldEmpZip").Value = arr(33)
  Fields("fldBTxt14").Value = arr(34)
'  If arr(35) = "0" Then arr(35) = ""
  Fields("fldBox14Amt1").Value = arr(35)
  Fields("fldBTxt14vs2").Value = arr(36)
'  If arr(37) = "0" Then arr(37) = ""
  Fields("fldBox14Amt1vs2").Value = arr(37)
  Fields("fldState2").Value = arr(38)
  Fields("fldSTAID").Value = arr(39)
  Fields("fldStateWage").Value = arr(40)
  Fields("fldStateTax").Value = arr(41)
  Fields("fldBOX13AM2").Value = arr(42)
  Fields("fldBox13TX2").Value = arr(43)
  Fields("fldBOX13AM3").Value = arr(44)
  Fields("fldBox13TX3").Value = arr(45)
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmLoadingRpt
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
''    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  PageHeader.Height = 0
'  GroupHeader1.Height = 0
  Me.Zoom = -1
End Sub


Private Sub PageHeader_Format()
'  Static Count As Integer
'
'  Count = Count + 1
'  If Count = 2 Then
'    GroupHeader1.NewPage = True
'    Count = 0
'  End If

End Sub

