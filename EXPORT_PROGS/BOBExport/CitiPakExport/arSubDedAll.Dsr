VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSubDedAll 
   BorderStyle     =   0  'None
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   11959
   _ExtentY        =   7726
   SectionData     =   "arSubDedAll.dsx":0000
End
Attribute VB_Name = "arSubDedAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private HFile As Integer
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
  End If
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True 'removes the error message
End Sub
Private Sub ActiveReport_DataInitialize()
  HFile = FreeFile
  Open StartPath & "\PRRPTS\SUBDEDALL.RPT" For Input As #HFile
  
  Fields.Add "fld0" '(0)
  Fields.Add "fld1" '(1)
  Fields.Add "fld2" '(2)
  Fields.Add "fld3" '(3)
  Fields.Add "fld4"
  End Sub
Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  Dim X As Integer
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
  Fields("fld0").Value = arr(0)
  Fields("fld1").Value = arr(1)
  Fields("fld2").Value = arr(2)
  Fields("fld3").Value = arr(3)
  Fields("fld4").Value = arr(4)
  
End Sub
Private Sub ActiveReport_ReportEnd()
  If HFile <> 0 Then
    Close #HFile
  End If
End Sub

