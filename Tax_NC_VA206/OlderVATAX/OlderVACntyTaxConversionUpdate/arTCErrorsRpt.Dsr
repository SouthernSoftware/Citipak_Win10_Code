VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arTCErrorsRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "arTCSubErrorRpt"
   ClientHeight    =   6000
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   10560
   Icon            =   "arTCErrorsRpt.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   18627
   _ExtentY        =   10583
   SectionData     =   "arTCErrorsRpt.dsx":08CA
End
Attribute VB_Name = "arTCErrorsRpt"
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
  Open App.Path & "\TCRPTS\TCERRORS.RPT" For Input As #hFile
  Fields.Add ("fldCntyAcctNum") '0)
  Fields.Add ("fldCntyAcctStr") '1)
  Fields.Add ("fldCustName") '2)
  Fields.Add ("fldErrorType") '3)
  Fields.Add ("fldPersTot") '4)
  Fields.Add ("fldPersXTot") '5)
  Fields.Add ("fldRPinNum") '6)
  Fields.Add ("fldRealTot") '7)
  Fields.Add ("fldRealXTot") '8)
  Fields.Add ("fldPPinNum") '9)
  Fields.Add ("fldCountyPin") '10)
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    frmTCMsg.Label1.Caption = "Error Number: " & Str(Number) & " " & Description
    frmTCMsg.Label1.Top = 900
    frmTCMsg.Show vbModal
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
  Fields("fldCntyAcctNum").Value = arr(0)
  Fields("fldCntyAcctStr").Value = arr(1)
  Fields("fldCustName").Value = arr(2)
  Fields("fldErrorType").Value = arr(3)
  Fields("fldPersTot").Value = arr(4)
  Fields("fldPersXTot").Value = arr(5)
  Fields("fldRPinNum").Value = arr(6)
  Fields("fldRealTot").Value = arr(7)
  Fields("fldRealXTot").Value = arr(8)
  Fields("fldPPinNum").Value = arr(9)
  If QPTrim$(arr(1)) <> "" Then
    Fields("fldCountyPin").Value = arr(1)
  Else
    Fields("fldCountyPin").Value = arr(0)
  End If
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
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub
  


