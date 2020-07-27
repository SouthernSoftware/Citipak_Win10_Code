VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arW3Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "W3 Form"
   ClientHeight    =   10995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14535
   Icon            =   "arW3Form.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   25638
   _ExtentY        =   19394
   SectionData     =   "arW3Form.dsx":08CA
End
Attribute VB_Name = "arW3Form"
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
      MsgBox "File - W3FormsRpt.xls, created in the Citipak Directory.", vbOKOnly
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      MsgBox "File - W3FormsRpt.txt, created in the Citipak Directory.", vbOKOnly
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
  End If
  If Tool = "Save/&Excel" Then
    Screen.MousePointer = vbHourglass
    ExportReport 1
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - W3FormsRpt.xls, created in the Citipak Directory.", vbOKOnly
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    MsgBox "File - W3FormsRpt.txt, created in the Citipak Directory.", vbOKOnly
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
        oEXL.FileName = outfile & "W3FormsRpt.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "W3FormsRpt.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub
Private Sub ActiveReport_DataInitialize()
  hFile = FreeFile
  Open StartPath & "\PRRPTS\W3FORMS.RPT" For Input As #hFile
  Fields.Add "ControlNum" '(0)
  Fields.Add "Payer" '(1)
  Fields.Add "ThirdParty" '(2)
  Fields.Add "NumOfForms" '(3)
  Fields.Add "EstNum" '(4)
  Fields.Add "EmpIDNum" '(5)
  Fields.Add "EmpName" '(6)
  Fields.Add "Add1" '(7)
  Fields.Add "Add2" '(8)
  Fields.Add "City" '(9)
  Fields.Add "Zip" '(10)
  Fields.Add "State" '(11)
  Fields.Add "OtherEIN" '(12)
  Fields.Add "FedWages" '(13)
  Fields.Add "FedTax" '(14)
  Fields.Add "SSWages" '(15)
  Fields.Add "SSTax" '(16)
  Fields.Add "MedWages" '(17)
  Fields.Add "MedTax" '(18)
  Fields.Add "SSTips" '(19)
  Fields.Add "AlloTips" '(20)
  'Fields.Add "AdvEIC" '(21)
  Fields.Add "DepCare" '(22)
  Fields.Add "NQP" '(23)
  Fields.Add "DefComp" '(24)
  Fields.Add "SickOnly" '(25)
  Fields.Add "SickAmt" '(26)
  Fields.Add "State2" '(27)
  Fields.Add "StateID" '(28)
  Fields.Add "StateWages" '(29)
  Fields.Add "StateTax" '(30)
  Fields.Add "LocalWages" '(31)
  Fields.Add "LocalTax" '(32)
  Fields.Add "Contact" '(33)
  Fields.Add "Phone" '(34)
  Fields.Add "Email" '(35)
  Fields.Add "Fax" '(36)
  Fields.Add "CSZ"
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
  fldA.Text = ""
  fldB.Text = ""
  fldC.Text = ""
  fldD.Text = ""
  fldE.Text = ""
  fldF.Text = ""
  fldG.Text = ""
  arr = Split(sLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("ControlNum").Value = arr(0)
  Fields("Payer").Value = arr(1)
  Select Case QPTrim$(arr(1))
    Case "941"
      fldA.Text = "X"
    Case "Military"
      fldB.Text = "X"
    Case "943"
      fldC.Text = "X"
    Case "CT-1"
      fldD.Text = "X"
    Case "Household Emp"
      fldE.Text = "X"
    Case "Medicare Govt-Emp"
      fldF.Text = "X"
    Case Else
  End Select
  Fields("ThirdParty").Value = arr(2)
  If arr(2) = "Yes" Then fldG.Text = "X"

  Fields("NumOfForms").Value = arr(3)
  Fields("EstNum").Value = arr(4)
  Fields("EmpIDNum").Value = arr(5)
  Fields("EmpName").Value = arr(6)
  Fields("Add1").Value = arr(7)
  Fields("Add2").Value = arr(8)
  Fields("City").Value = arr(9)
  Fields("Zip").Value = arr(10)
  Fields("State").Value = arr(11)
  Fields("CSZ").Value = QPTrim$(arr(9)) + ", " + QPTrim$(arr(10)) + "  " + QPTrim$(arr(11))
  Fields("OtherEIN").Value = arr(12)
  Fields("FedWages").Value = arr(13)
  Fields("FedTax").Value = arr(14)
  Fields("SSWages").Value = arr(15)
  Fields("SSTax").Value = arr(16)
  Fields("MedWages").Value = arr(17)
  Fields("MedTax").Value = arr(18)
  Fields("SSTips").Value = arr(19)
  Fields("AlloTips").Value = arr(20)
  'Fields("AdvEIC").Value = arr(21)
  Fields("DepCare").Value = arr(22)
  Fields("NQP").Value = arr(23)
  Fields("DefComp").Value = arr(24)
  Fields("SickOnly").Value = arr(25)
  Fields("SickAmt").Value = arr(26)
  Fields("State2").Value = arr(27)
  Fields("StateID").Value = arr(28)
  Fields("StateWages").Value = arr(29)
  Fields("StateTax").Value = arr(30)
  Fields("LocalWages").Value = arr(31)
  Fields("LocalTax").Value = arr(32)
  Fields("Contact").Value = arr(33)
  Fields("Phone").Value = arr(34)
  Fields("Email").Value = arr(35)
  Fields("Fax").Value = arr(36)
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
  Me.Zoom = -1
End Sub

