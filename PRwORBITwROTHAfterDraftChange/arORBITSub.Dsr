VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arORBITSub 
   BorderStyle     =   0  'None
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   13653
   _ExtentY        =   9869
   SectionData     =   "arORBITSub.dsx":0000
End
Attribute VB_Name = "arORBITSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private hFile As Integer
Dim fld1() As String
Dim Fld1Cnt As Integer
Dim fld2() As Double
Dim fld3() As Double
Dim Fld4() As Double
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
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
  Dim x As Integer
  hFile = FreeFile
  Open StartPath & "\PRRPTS\ORBITSUB.RPT" For Input As #hFile
  
  Fields.Add "fld0"
  Fields.Add "fld1"
  Fields.Add "fld2"
  Fields.Add "fld3"
  Fields.Add "fld4"
  
  End Sub
Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  Dim x As Integer
  Static ArrCnt As Integer
  If VBA.eof(hFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #hFile, sLine
  arr = Split(sLine, "~")
  Fields("fld0").Value = arr(0)
  Fields("fld1").Value = arr(1)
  Fields("fld2").Value = arr(2)
  Fields("fld3").Value = arr(3)
  Fields("fld4").Value = arr(4)
End Sub
Private Sub ActiveReport_ReportEnd()
  If hFile <> 0 Then
    Close #hFile
  End If
End Sub
