VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSub1RealClassRpt 
   BorderStyle     =   0  'None
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   13361
   _ExtentY        =   10213
   SectionData     =   "arSub1RealClassRpt.dsx":0000
End
Attribute VB_Name = "arSub1RealClassRpt"
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
  Open StartPath & "\TAXRPTS\SUB1REALCLAS.RPT" For Input As #hFile
  Fields.Add ("fldTSName") '0)
  Fields.Add ("fldTSAmt") '1)
  Fields.Add ("fldTSDisc") '2)
  Fields.Add ("fldTSNet") '3)
  Fields.Add ("fldClassName") '4)
  Fields.Add ("fldTSAmtTot") '5)
  Fields.Add ("fldTSDiscTot") '6)
  Fields.Add ("fldTSNetTot") '7)
  Fields.Add ("fldTSCnt") '8)
  Fields.Add ("fldTSCntTot") '9)
  Fields.Add ("fldGTSAmtTot") '10)
  Fields.Add ("fldGTSDiscTot") '11)
  Fields.Add ("fldGTSNetTot") '12)
  Fields.Add ("fldGTSCntTot") '13)
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
  Fields("fldTSName").Value = arr(0)
  Fields("fldTSAmt").Value = arr(1)
  Fields("fldTSDisc").Value = arr(2)
  Fields("fldTSNet").Value = arr(3)
  Fields("fldClassName").Value = arr(4)
  Fields("fldTSAmtTot").Value = arr(5)
  Fields("fldTSDiscTot").Value = arr(6)
  Fields("fldTSNetTot").Value = arr(7)
  Fields("fldTSCnt").Value = arr(8)
  Fields("fldTSCntTot").Value = arr(9)
  Fields("fldGTSAmtTot").Value = arr(10)
  Fields("fldGTSDiscTot").Value = arr(11)
  Fields("fldGTSNetTot").Value = arr(12)
  Fields("fldGTSCntTot").Value = arr(13)
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


