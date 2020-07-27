VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arFASubTagByCode 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   4488
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   14790
   _ExtentY        =   7911
   SectionData     =   "arFASubTagByCode.dsx":0000
End
Attribute VB_Name = "arFASubTagByCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsFATextBoxOverRider
Private Temp_Class As Resize_Class
Private HFile As Integer

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  HFile = FreeFile
  Open StartPath & "\FARPTS\FASUBCODE.RPT" For Input As #HFile
  Fields.Add ("fldCodeNum") '0)
  Fields.Add ("fldCodeDesc") '1)
  Fields.Add ("fldCodeCnt") '2)
  Fields.Add ("fldOrigCost") '3)
  Fields.Add ("fldDepr") '4)
  Fields.Add ("fldBkTot") '5)
  Fields.Add ("fldTOrigCost") '6)
  Fields.Add ("fldTDepr") '7)
  Fields.Add ("fldTBkTot") '8)
  Fields.Add ("fldEndRpt") '9)
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    Unload frmFALoadReport
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True 'removes the error message
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  
  If VBA.eof(HFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #HFile, sLine
  arr = Split(sLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("fldCodeNum").Value = arr(0)
  Fields("fldCodeDesc").Value = arr(1)
  Fields("fldCodeCnt").Value = arr(2)
  Fields("fldOrigCost").Value = arr(3)
  Fields("fldDepr").Value = arr(4)
  Fields("fldBkTot").Value = arr(5)
  Fields("fldTOrigCost").Value = arr(6)
  Fields("fldTDepr").Value = arr(7)
  Fields("fldTBkTot").Value = arr(8)
  Fields("fldEndRpt").Value = arr(9)
End Sub

Private Sub ActiveReport_ReportEnd()
  Unload frmFALoadReport
  If HFile <> 0 Then
    Close #HFile
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsFATextBoxOverRider
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






