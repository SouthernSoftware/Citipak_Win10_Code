VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arBLSubLicRegister 
   BorderStyle     =   0  'None
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   15875
   _ExtentY        =   7726
   SectionData     =   "arSubLicRegister.dsx":0000
End
Attribute VB_Name = "arBLSubLicRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Over As clsBLTextBoxOverrider
Private Temp_Class As Resize_Class
Private hFile As Integer

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\BLRPTS\ARSUBLICREG.RPT" For Input As #hFile
  Fields.Add ("fldCode") '0)
  Fields.Add ("fldFee") '1)
  Fields.Add ("fldCustCnt") '2)
  Fields.Add ("fldTotalFee") '3)
  Fields.Add ("fldCatDesc") '4)
  Fields.Add ("fldCatCnt") '5)
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
  Fields("fldCode").Value = arr(0)
  Fields("fldFee").Value = arr(1)
  Fields("fldCustCnt").Value = arr(2)
  Fields("fldTotalFee").Value = arr(3)
  Fields("fldCatDesc").Value = arr(4)
  Fields("fldCatCnt").Value = arr(5)

End Sub

Private Sub ActiveReport_Initialize()
'  Me.ToolBar.Tools.Add "&Close"
'  Me.ToolBar.Tools.Add "Save/&Excel"
'  Me.ToolBar.Tools.Add "&Text"
End Sub

Private Sub ActiveReport_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
'  If KeyCode = vbKeyEscape Then
'    Unload Me
'    KeyCode = 0
'  End If
'  If Shift = 4 Then
'    If KeyCode = vbKeyC Then
'      Unload Me
'      KeyCode = 0
'    ElseIf KeyCode = vbKeyE Then
'      Screen.MousePointer = vbHourglass
'      ExportReport 1
'      Screen.MousePointer = vbDefault
'      DoEvents
'      MsgBox "File - BLLicRegister.xls, created in the Citipak Directory.", vbOKOnly
'      KeyCode = 0
'    ElseIf KeyCode = vbKeyT Then
'      Screen.MousePointer = vbHourglass
'      ExportReport 2
'      Screen.MousePointer = vbDefault
'      DoEvents
'      MsgBox "File - BLLicRegister.txt, created in the Citipak Directory.", vbOKOnly
'      KeyCode = 0
'    End If
'  End If
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


