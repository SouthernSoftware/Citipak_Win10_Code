VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arBLSubCustList 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   12224
   _ExtentY        =   6694
   SectionData     =   "arBLSubCustList.dsx":0000
End
Attribute VB_Name = "arBLSubCustList"
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
  Open StartPath & "\BLRPTS\ARDetCustSub.RPT" For Input As #hFile
  Fields.Add ("fldCatCode") '0)
  Fields.Add ("fldDesc") '1)
  Fields.Add ("fldCustNum") '2)
  Fields.Add ("fldCustName") '3)
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
  Fields("fldCatCode").Value = arr(0)
  Fields("fldDesc").Value = arr(1)
  Fields("fldCustNum").Value = arr(2)
  Fields("fldCustName").Value = arr(3)

End Sub

Private Sub ActiveReport_ReportEnd()
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





