VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arBLAdvanceLtr3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Advance Letter 3"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   ControlBox      =   0   'False
   Icon            =   "arBLAdvanceLtr3.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   20558
   _ExtentY        =   15637
   SectionData     =   "arBLAdvanceLtr3.dsx":08CA
End
Attribute VB_Name = "arBLAdvanceLtr3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsBLTextBoxOverrider
Private Temp_Class As Resize_Class
Private hFile As Integer

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\BLRPTS\ARADVLT3.RPT" For Input As #hFile
  Fields.Add ("fld0") '0)
  Fields.Add ("fld1") '1)
  Fields.Add ("fld2") '2)
  Fields.Add ("fld3") '3)
  Fields.Add ("fld4") '4)
  Fields.Add ("fld5") '5)
  Fields.Add ("fld6") '6)
  Fields.Add ("fld7") '7)
  Fields.Add ("fld8") '8)
  Fields.Add ("fld9") '9)
  Fields.Add ("fld10") '10)
  Fields.Add ("fld11") '11)
  Fields.Add ("fld12") '12)
  Fields.Add ("fld13") '13)
  Fields.Add ("fld14") '14)
  Fields.Add ("fld15") '15)
  Fields.Add ("fld16") '16)
  Fields.Add ("fld17") '17)
  Fields.Add ("fld18") '18)
  Fields.Add ("fld19") '19)
  Fields.Add ("fld20") '20)
  Fields.Add ("fld21") '21)
  Fields.Add ("fld22") '22)
  Fields.Add ("fld23") '23)
  Fields.Add ("fld24") '24)
  Fields.Add ("fld25") '25)
  Fields.Add ("fld26") '26)
  Fields.Add ("fld27") '27)
  Fields.Add ("fld28") '28)
  Fields.Add ("fld29") '29)
  Fields.Add ("fld30") '30)
  Fields.Add ("fld31") '31)
  Fields.Add ("fld32") '32)
  Fields.Add ("fld33") '33)
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmBLLoadReport
    frmBLMessageBoxJr.Label1.Caption = "Error Number: " & Str(Number) & " " & Description
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
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
  Line8.Visible = True
  Field17.Visible = True
  Label5.Visible = True
  Field32.Visible = True
  Fields("fld0").Value = arr(0) 'cust add 2
  Fields("fld1").Value = arr(1) 'billname/cust num
  Fields("fld2").Value = arr(2) 'cust add 1
  Fields("fld3").Value = arr(3) 'csz
  Fields("fld4").Value = arr(4) 'line1 1
  Fields("fld5").Value = arr(5) 'line1 2
  Fields("fld6").Value = arr(6) 'line1 3
  Fields("fld7").Value = arr(7) 'line1 4
  Fields("fld8").Value = arr(8) 'line1 5
  Fields("fld9").Value = arr(9) 'line1 6
  Fields("fld10").Value = arr(10) 'line2 1
  Fields("fld11").Value = arr(11) 'line2 2
  Fields("fld12").Value = arr(12) 'line2 3
  Fields("fld13").Value = arr(13) 'line2 4
  Fields("fld14").Value = arr(14) 'code 1
  Fields("fld15").Value = arr(15) 'desc 1
  Fields("fld16").Value = arr(16) 'fee amt 1
  Fields("fld17").Value = arr(17) 'code 2
  Fields("fld18").Value = arr(18) 'desc 2
  Fields("fld19").Value = arr(19) 'fee amt 2
  Fields("fld20").Value = arr(20) 'code 3
  Fields("fld21").Value = arr(21) 'desc 3
  Fields("fld22").Value = arr(22) 'fee amt 3
  Fields("fld23").Value = arr(23) 'code 4
  Fields("fld24").Value = arr(24) 'desc 4
  Fields("fld25").Value = arr(25) 'fee amt 4
  Fields("fld26").Value = arr(26) 'code 5
  Fields("fld27").Value = arr(27) 'desc 2
  Fields("fld28").Value = arr(28) 'fee amt 3
  Fields("fld29").Value = arr(29) 'tot fee
  Fields("fld30").Value = arr(30) 'iss fee
  If QPTrim$(arr(30)) = "0" Then
    Field17.Visible = False
    Label5.Visible = False
  End If
  Fields("fld31").Value = arr(31) 'print date
  Fields("fld32").Value = arr(32) 'signer
  If Len(arr(33)) <= 1 Then
   Field32.Visible = False
  End If
  Fields("fld33").Value = arr(33) 'phone
  If QPTrim$(arr(32)) = "" Then Line8.Visible = False
  
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
      frmBLMessageBoxJr.Label1.Caption = "File - BLAdvLt3.xls, created in the Citipak Directory."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmBLMessageBoxJr.Label1.Caption = "File - BLAdvLt3.txt, created in the Citipak Directory."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
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
    frmBLMessageBoxJr.Label1.Caption = "File - BLAdvLt3.xls, created in the Citipak Directory."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmBLMessageBoxJr.Label1.Caption = "File - BLAdvLt3.txt, created in the Citipak Directory."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
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
        oEXL.FileName = outfile & "BLAdvLt3.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "BLAdvLt3.txt"
        oTXT.PageDelimiter = ";"
        oTXT.TextDelimiter = ","
        oTXT.Export Me.Pages
  End Select
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmBLLoadReport
  If hFile <> 0 Then
    Close #hFile
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  If Exist("townlogoadvltr3.bmp") Then
    Image1.Picture = LoadPicture("townlogoadvltr3.bmp")
  Else
    Image1.Visible = False
  End If
  Unload frmBLLoadReport
  Me.Zoom = -1
End Sub











