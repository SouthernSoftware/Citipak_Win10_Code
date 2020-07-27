VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arBLLaser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Laser Forms"
   ClientHeight    =   9720
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   13185
   ControlBox      =   0   'False
   Icon            =   "arBLLaser.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   23257
   _ExtentY        =   17145
   SectionData     =   "arBLLaser.dsx":08CA
End
Attribute VB_Name = "arBLLaser"
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
  Open StartPath & "\BLRPTS\ARLASER.RPT" For Input As #hFile
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
  Fields.Add ("fld34") '34)
  Fields.Add ("fld35") '35)
  Fields.Add ("fld36") '36)
  Fields.Add ("fld37") '37)
  
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
  
  Fields("fld0").Value = arr(0) 'heading1
  Fields("fld1").Value = arr(1) 'heading2
  Fields("fld2").Value = arr(2) 'heading3
  Fields("fld3").Value = arr(3) 'Year
  Fields("fld4").Value = arr(4) 'customer number
  Fields("fld5").Value = arr(5) 'prorate(optional)
  Fields("fld6").Value = arr(6) 'bill name
  Fields("fld7").Value = arr(7) 'License number
  Fields("fld8").Value = arr(8) 'address
  Fields("fld9").Value = arr(9) 'citystatezip
  Fields("fld10").Value = arr(10) 'issue date
  Fields("fld11").Value = arr(11) 'expire date
  Fields("fld12").Value = arr(12) 'customer name
  Fields("fld13").Value = arr(13) 'authorized by
  If QPTrim$(arr(13)) = "" Then
    Label10.Visible = False
    fld13.Visible = False
    Label26.Visible = False
    Field13.Visible = False
  Else
    Label10.Visible = True
    fld13.Visible = True
    Label26.Visible = True
    Field13.Visible = True
  End If
  
  Fields("fld14").Value = arr(14) 'cat code1
  Fields("fld15").Value = arr(15) 'cat desc1
  Fields("fld16").Value = arr(16) 'cat code2
  Fields("fld17").Value = arr(17) 'cat desc2
  Fields("fld18").Value = arr(18) 'cat code3
  Fields("fld19").Value = arr(19) 'cat desc3
  Fields("fld20").Value = arr(20) 'cat code4
  Fields("fld21").Value = arr(21) 'cat desc4
  Fields("fld22").Value = arr(22) 'cat code5
  Fields("fld23").Value = arr(23) 'cat desc5
  Fields("fld24").Value = arr(24) 'catfee1
  Fields("fld25").Value = arr(25) 'catfee2
  Fields("fld26").Value = arr(26) 'catfee3
  Fields("fld27").Value = arr(27) 'catfee4
  Fields("fld28").Value = arr(28) 'catfee5
  Fields("fld29").Value = arr(29) 'issue fee
  Fields("fld30").Value = arr(30) 'total bill amount
  Fields("fld31").Value = arr(31) 'from date
  Fields("fld32").Value = arr(32) 'total outstanding balance - this balance
  Fields("fld33").Value = arr(33) 'total outstanding balance
  Fields("fld34").Value = arr(34) 'balanceflag
  Fields("fld35").Value = arr(35) 'license year
  Fields("fld36").Value = arr(36) 'cust bill name
  Fields("fld37").Value = arr(37) 'service address
  
  fld24.Visible = True
  fld25.Visible = True
  fld26.Visible = True
  fld27.Visible = True
  fld28.Visible = True
  fld30.Visible = True
  Field5.Visible = True
  Field29.Visible = True
  Field30.Visible = True
  Field31.Visible = True
  Field32.Visible = True
  Field33.Visible = True
  Field36.Visible = True
  Field37.Visible = True
  Field38.Visible = True
  Field39.Visible = True
  Label30.Visible = True
  Label31.Visible = True
  Label32.Visible = True
  Label33.Visible = True
  
  'if print fees is 'no' then the code being
  'sent here from the originating screen leaves
  'blank the arr() values that would normally
  'contain values
  '...the code below examines the blank fields
  'and makes visible or invisible those fields
  'that affect the license form in terms of
  'fee values
  
  If Val(arr(16)) = 0 Then 'fee amt #1
    If QPTrim$(arr(15)) = "" Then 'desc #1
      fld24.Visible = False
      Field29.Visible = False
    End If
  End If
  If Val(arr(19)) = 0 Then 'fee amt #2
    If QPTrim$(arr(18)) = "" Then 'desc #2
      fld25.Visible = False
      Field30.Visible = False
    End If
  End If
  If Val(arr(22)) = 0 Then 'fee amt #3
    If QPTrim$(arr(21)) = "" Then 'desc #3
      fld26.Visible = False
      Field31.Visible = False
    End If
  End If
  If Val(arr(25)) = 0 Then 'fee amt #4
    If QPTrim$(arr(24)) = "" Then 'desc #4
      fld27.Visible = False
      Field32.Visible = False
    End If
  End If
  If Val(arr(28)) = 0 Then 'fee amt #5
    If QPTrim$(arr(27)) = "" Then 'desc #5
      fld28.Visible = False
      Field33.Visible = False
    End If
  End If
  Label2.Visible = True
  Label17.Visible = True
  
  If arr(30) = "No" Then 'Print/Don't print fees
    Label2.Visible = False
    Label17.Visible = False
    fld30.Visible = False
    Field5.Visible = False
  End If
  
  If Val(arr(29)) = 0 Then 'iss fee
    Label28.Visible = False
    Label29.Visible = False
    fld29.Visible = False
    Field35.Visible = False
  Else
    Label28.Visible = True
    Label29.Visible = True
    fld29.Visible = True
    Field35.Visible = True
  End If
  
  If Val(arr(34)) = 1 Then 'balance flag
    Label30.Visible = False
    Label31.Visible = False
    Label32.Visible = False
    Label33.Visible = False
    Field36.Visible = False
    Field37.Visible = False
    Field38.Visible = False
    Field39.Visible = False
  End If
  
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
      frmBLMessageBoxJr.Label1.Caption = "File - BLLaser.xls, created in the Citipak Directory."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      KeyCode = 0
    ElseIf KeyCode = vbKeyT Then
      Screen.MousePointer = vbHourglass
      ExportReport 2
      Screen.MousePointer = vbDefault
      DoEvents
      frmBLMessageBoxJr.Label1.Caption = "File - BLLaser.txt, created in the Citipak Directory."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      KeyCode = 0
    End If
  End If
End Sub

Private Sub ActiveReport_PrintProgress(ByVal pageNumber As Long)
  If DidPrint = 2 Then
    MainLog ("Non-posting business license number " + CStr(pageNumber) + " printed in laser format.")
    DidPrint = 1
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
    frmBLMessageBoxJr.Label1.Caption = "File - BLLaser.xls, created in the Citipak Directory."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
  End If
  If Tool = "&Text" Then
    Screen.MousePointer = vbHourglass
    ExportReport 2
    Screen.MousePointer = vbDefault
    DoEvents
    frmBLMessageBoxJr.Label1.Caption = "File - BLLaser.txt, created in the Citipak Directory."
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
        oEXL.FileName = outfile & "BLLaser.xls"
        oEXL.Export Me.Pages
    Case 2   '"Text"
        Set oTXT = New ActiveReportsTextExport.ARExportText
        oTXT.FileName = outfile & "BLLaser.txt"
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
    DoEvents
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  
  If PrintSign = True Then 'signature line
    fldSigned.Visible = True
    fldSigned2.Visible = True
    Line6.Visible = True
    Line7.Visible = True
    If Exist("townsign.bmp") Then
      Image3.Picture = LoadPicture("townsign.bmp")
      Image4.Picture = LoadPicture("townsign.bmp")
    Else
      Image3.Visible = False
      Image4.Visible = False
    End If
  Else
    Image3.Visible = False
    Image4.Visible = False
    fldSigned.Visible = False
    fldSigned2.Visible = False
    Line6.Visible = False
    Line7.Visible = False
    Label8.Left = 2340
    Label9.Left = 4230
    Label21.Left = 2340
    Label22.Left = 4230
    Label11.Left = 6480
    Label25.Left = 6480
    fld31.Left = 5070
    fld11.Left = 6940
    Field8.Left = 5070
    Field12.Left = 6940
  End If
  
  If Exist("townlogo.bmp") Then
    Image1.Picture = LoadPicture("townlogo.bmp")
    Image5.Picture = LoadPicture("townlogo.bmp")
    fld0.Left = 4510
    fld1.Left = 4240
    fld2.Left = 4240
    Field1.Left = 4510
    Field2.Left = 4240
    Field3.Left = 4240
    fld12.Left = 4870
    Field44.Left = 4870
    Field6.Left = 4870
    Field45.Left = 4870
'    Field40.Left = 4960
'    Field41.Left = 4960
  Else
    fld12.Left = 4360
    Label6.Left = 3070
    Field44.Left = 4360
    Label39.Left = 2530
    Field6.Left = 4360
    Label19.Left = 3070
    Field45.Left = 4360
    Label38.Left = 2530
    Image1.Visible = False
    Image5.Visible = False
  End If
  
  Unload frmBLLoadReport
  Me.Zoom = -1
End Sub

