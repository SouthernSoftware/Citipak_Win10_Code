VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arManDedDesc 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4380
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   6555
   Icon            =   "arManDedDesc.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   11562
   _ExtentY        =   7726
   SectionData     =   "arManDedDesc.dsx":08CA
End
Attribute VB_Name = "arManDedDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private DFile As Integer
Dim EndReport As Boolean
Dim DedCnt As Integer
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  EndReport = False
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

Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim dLine As String
  Dim arr() As String
  If VBA.eof(DFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #DFile, dLine
  arr = Split(dLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("fldDedDsc1").Value = arr(0)
  Fields("fldDedDsc2").Value = arr(1)
  Fields("fldDedDsc3").Value = arr(2)
  Fields("fldDedDsc4").Value = arr(3)
  Fields("fldDedDsc5").Value = arr(4)
  Fields("fldDedDsc6").Value = arr(5)
  Fields("fldDedDsc7").Value = arr(6)
  Fields("fldDedDsc8").Value = arr(7)
  Fields("fldDedDsc9").Value = arr(8)
  Fields("fldDedDsc10").Value = arr(9)
  Fields("fldDedDsc11").Value = arr(10)
  Fields("fldDedDsc12").Value = arr(11)
  Fields("fldDedDsc13").Value = arr(12)
  Fields("fldDedDsc14").Value = arr(13)
  Fields("fldDedDsc15").Value = arr(14)
  Fields("fldDedDsc16").Value = arr(15)
  Fields("fldDedDsc17").Value = arr(16)
  Fields("fldDedDsc18").Value = arr(17)
  Fields("fldDedDsc19").Value = arr(18)
  Fields("fldDedDsc20").Value = arr(19)
  Fields("fldDedDsc21").Value = arr(20)
  Fields("fldDedDsc22").Value = arr(21)
  Fields("fldDedDsc23").Value = arr(22)
  Fields("fldDedDsc24").Value = arr(23)
  Fields("fldDedDsc25").Value = arr(24)
  Fields("fldDedDsc26").Value = arr(25)
  Fields("fldDedDsc27").Value = arr(26)
  Fields("fldDedDsc28").Value = arr(27)
  Fields("fldDedDsc29").Value = arr(28)
  Fields("fldDedDsc30").Value = arr(29)
  Fields("fldDedDsc31").Value = arr(30)
  Fields("fldDedDsc32").Value = arr(31)
  Fields("fldDedDsc33").Value = arr(32)
  Fields("fldDedDsc34").Value = arr(33)
  Fields("fldDedDsc35").Value = arr(34)
  Fields("fldDedDsc36").Value = arr(35)
  Fields("fldDedDsc37").Value = arr(36)
  Fields("fldDedDsc38").Value = arr(37)
  Fields("fldDedDsc39").Value = arr(38)
  Fields("fldDedDsc40").Value = arr(39)
  Fields("fldDedDsc41").Value = arr(40)
  Fields("fldDedDsc42").Value = arr(41)
  Fields("fldDedDsc43").Value = arr(42)
  Fields("fldDedDsc44").Value = arr(43)
  Fields("fldDedDsc45").Value = arr(44)
  Fields("fldDedDsc46").Value = arr(45)
  Fields("fldDedDsc47").Value = arr(46)
  Fields("fldDedDsc48").Value = arr(47)
  Fields("fldDedDsc49").Value = arr(48)
  Fields("fldDedDsc50").Value = arr(49)
End Sub

Private Sub ActiveReport_Initialize()
  ToolBar.Tools.Add "Exit"
  ToolBar.Font.Size = 10
  
End Sub
Private Sub ActiveReport_DataInitialize()
  DFile = FreeFile
  Open StartPath & "\PRRPTS\MANDEDDESC.RPT" For Input As #DFile
  Fields.Add "fldDedDsc1" '(0)
  Fields.Add "fldDedDsc2" '(1)
  Fields.Add "fldDedDsc3" '(2)
  Fields.Add "fldDedDsc4" '(3)
  Fields.Add "fldDedDsc5" '(4)
  Fields.Add "fldDedDsc6" '(5)
  Fields.Add "fldDedDsc7" '(6)
  Fields.Add "fldDedDsc8" '(7)
  Fields.Add "fldDedDsc9" '(8)
  Fields.Add "fldDedDsc10" '(9)
  Fields.Add "fldDedDsc11" '(10)
  Fields.Add "fldDedDsc12" '(11)
  Fields.Add "fldDedDsc13" '(12)
  Fields.Add "fldDedDsc14" '(13)
  Fields.Add "fldDedDsc15" '(14)
  Fields.Add "fldDedDsc16" '(15)
  Fields.Add "fldDedDsc17" '(16)
  Fields.Add "fldDedDsc18" '(17)
  Fields.Add "fldDedDsc19" '(18)
  Fields.Add "fldDedDsc20" '(19)
  Fields.Add "fldDedDsc21" '(20)
  Fields.Add "fldDedDsc22" '(21)
  Fields.Add "fldDedDsc23" '(22)
  Fields.Add "fldDedDsc24" '(23)
  Fields.Add "fldDedDsc25" '(24)
  Fields.Add "fldDedDsc26" '(25)
  Fields.Add "fldDedDsc27" '(26)
  Fields.Add "fldDedDsc28" '(27)
  Fields.Add "fldDedDsc29" '(28)
  Fields.Add "fldDedDsc30" '(29)
  Fields.Add "fldDedDsc31" '(30)
  Fields.Add "fldDedDsc32" '(31)
  Fields.Add "fldDedDsc33" '(32)
  Fields.Add "fldDedDsc34" '(33)
  Fields.Add "fldDedDsc35" '(34)
  Fields.Add "fldDedDsc36" '(35)
  Fields.Add "fldDedDsc37" '(36)
  Fields.Add "fldDedDsc38" '(37)
  Fields.Add "fldDedDsc39" '(38)
  Fields.Add "fldDedDsc40" '(39)
  Fields.Add "fldDedDsc41" '(40)
  Fields.Add "fldDedDsc42" '(41)
  Fields.Add "fldDedDsc43" '(42)
  Fields.Add "fldDedDsc44" '(43)
  Fields.Add "fldDedDsc45" '(44)
  Fields.Add "fldDedDsc46" '(45)
  Fields.Add "fldDedDsc47" '(46)
  Fields.Add "fldDedDsc48" '(47)
  Fields.Add "fldDedDsc49" '(48)
  Fields.Add "fldDedDsc50" '(49)

End Sub
Private Sub ActiveReport_ReportEnd()
  If DFile <> 0 Then
    Close #DFile
  End If
End Sub

Private Sub ActiveReport_ReportStart()
  Dim DedRec As DedCodeRecType
  Dim DHandle As Integer
  Dim x As Integer
  
  OpenDedCodeFile DHandle
  DedCnt = LOF(DHandle) / Len(DedRec)
  Close DHandle
End Sub

Private Sub ActiveReport_Terminate()
  Close
End Sub

Private Sub Detail_Format()
  Dim DedRec As DedCodeRecType
  Dim DHandle As Integer

  OpenDedCodeFile DHandle
  DedCnt = LOF(DHandle) / Len(DedRec)
  Close DHandle
  Select Case DedCnt
  Case 0 To 10
    Detail.Height = 200
  Case 11 To 20
    Detail.Height = 500
  Case 21 To 30
    Detail.Height = 800
  Case 31 To 40
    Detail.Height = 1000
  Case 41 To 50
    Detail.Height = 1500
  Case Else
    Detail.Height = 1800
  End Select

End Sub

