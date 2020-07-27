VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSub2RealClassRpt 
   BorderStyle     =   0  'None
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   15901
   _ExtentY        =   10213
   SectionData     =   "arSub2RealClassRpt.dsx":0000
End
Attribute VB_Name = "arSub2RealClassRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private hFile As Integer

Private Sub ActiveReport_DataInitialize()
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\TAXRPTS\SUB2REALCLAS.RPT" For Input As #hFile
  Fields.Add ("fldClassAmt") '0)
  Fields.Add ("fldClassDisc") '1)
  Fields.Add ("fldClassCnt") '2)
  Fields.Add ("fldClassName") '3)
  Fields.Add ("fldClassNet") '4)
  Fields.Add ("fldClassAmtTot") '5)
  Fields.Add ("fldClassDiscTot") '6)
  Fields.Add ("fldClassNetTot") '7)
  Fields.Add ("fldClassCntTot") '8)
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
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
  ' Here we set the values of the fields that we define as unbound
  ' or user defined.
  Fields("fldClassAmt").Value = arr(0)
  Fields("fldClassDisc").Value = arr(1)
  Fields("fldClassCnt").Value = arr(2)
  Fields("fldClassName").Value = arr(3)
  Fields("fldClassNet").Value = arr(4)
  Fields("fldClassAmtTot").Value = arr(5)
  Fields("fldClassDiscTot").Value = arr(6)
  Fields("fldClassNetTot").Value = arr(7)
  Fields("fldClassCntTot").Value = arr(8)
End Sub

Private Sub ActiveReport_ReportEnd()
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
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub



