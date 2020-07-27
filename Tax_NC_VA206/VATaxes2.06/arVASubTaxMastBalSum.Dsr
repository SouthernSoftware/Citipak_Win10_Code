VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arVASubTaxMastBalSum 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   3348
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   13335
   _ExtentY        =   5900
   SectionData     =   "arVASubTaxMastBalSum.dsx":0000
End
Attribute VB_Name = "arVASubTaxMastBalSum"
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
  Open StartPath & "\TAXRPTS\TXMSTBALSUB.RPT" For Input As #hFile
  Fields.Add ("fldYear") '0)
  Fields.Add ("fldYearBal") '1)
  Fields.Add ("fldOverage") '2)
  Fields.Add ("fldLastOne") '3)
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
'    Field4.Visible = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #hFile, sLine
  arr = Split(sLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
'  Field4.Visible = False
  Fields("fldYear").Value = arr(0)
  Fields("fldYearBal").Value = arr(1)
  Fields("fldOverage").Value = arr(2)
  Fields("fldLastOne").Value = arr(3)
'  If CInt(Fields("fldLastOne").Value) = 1 Then
'    Field4.Visible = True
'  Else
'    Field4.Visible = False
'  End If
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
  

