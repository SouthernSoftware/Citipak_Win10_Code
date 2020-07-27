VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVASubDiscsAtPayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Discounts Applied At Payment"
   ClientHeight    =   5784
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   9375
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   16536
   _ExtentY        =   10213
   SectionData     =   "arVASubDiscsAtPayment.dsx":0000
End
Attribute VB_Name = "arVASubDiscsAtPayment"
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
  Open StartPath & "\TAXRPTS\SUBDISCATPAY.RPT" For Input As #hFile
  Fields.Add ("fldYear") '0)
  Fields.Add ("fldYearAmt") '1)
  Fields.Add ("fldYearDisc") '2)
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    frmVATaxMsg.Label1.Caption = "Error Number: " & Str(Number) & " " & Description
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
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
  Fields("fldYear").Value = arr(0)
  Fields("fldYearAmt").Value = arr(1)
  Fields("fldYearDisc").Value = arr(2)
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
  

