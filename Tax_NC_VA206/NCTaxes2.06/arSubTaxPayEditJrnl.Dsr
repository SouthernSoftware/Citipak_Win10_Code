VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSubTaxPayEditJrnl 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8175
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   14420
   _ExtentY        =   8281
   SectionData     =   "arSubTaxPayEditJrnl.dsx":0000
End
Attribute VB_Name = "arSubTaxPayEditJrnl"
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
  Open StartPath & "\TAXRPTS\SubEdPay1.RPT" For Input As #hFile
  Fields.Add ("fldGPrincDisc") '0)
  Fields.Add ("fldGInt") '1)
  Fields.Add ("fldGAdvCol") '2)
  Fields.Add ("fldGLateList") '3)
  Fields.Add ("fldGRev1") '4)
  Fields.Add ("fldGRev2") '5)
  Fields.Add ("fldGRev3") '6)
  Fields.Add ("fldGTot") '7)
  Fields.Add ("fldGPrinc") '8)
  Fields.Add ("fldGDisc") '9)
  Fields.Add ("fldGOverPay") '10)
  Fields.Add ("fldRevDesc1") '11)
  Fields.Add ("fldRevDesc2") '12)
  Fields.Add ("fldRevDesc3") '13)
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
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("fldGPrincDisc").Value = arr(0)
  Fields("fldGInt").Value = arr(1)
  Fields("fldGAdvCol").Value = arr(2)
  Fields("fldGLateList").Value = arr(3)
  Fields("fldGRev1").Value = arr(4)
  Fields("fldGRev2").Value = arr(5)
  Fields("fldGRev3").Value = arr(6)
  Fields("fldGTot").Value = arr(7)
  Fields("fldGPrinc").Value = arr(8)
  Fields("fldGDisc").Value = arr(9)
  Fields("fldGOverPay").Value = arr(10)
  Fields("fldRevDesc1").Value = arr(11)
  Fields("fldRevDesc2").Value = arr(12)
  Fields("fldRevDesc3").Value = arr(13)
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


