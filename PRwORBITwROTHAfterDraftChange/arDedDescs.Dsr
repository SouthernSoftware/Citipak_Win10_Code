VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arDedDescs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deduction Descriptions"
   ClientHeight    =   4380
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8070
   Icon            =   "arDedDescs.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   14235
   _ExtentY        =   7726
   SectionData     =   "arDedDescs.dsx":08CA
End
Attribute VB_Name = "arDedDescs"
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
  Dim arrD() As String
  If VBA.eof(DFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #DFile, dLine
  arrD = Split(dLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("fldDedDsc1").Value = arrD(0)
  Fields("fldDedDsc2").Value = arrD(1)
  Fields("fldDedDsc3").Value = arrD(2)
  Fields("fldDedDsc4").Value = arrD(3)
  Fields("fldDedDsc5").Value = arrD(4)
  Fields("fldDedDsc6").Value = arrD(5)
  Fields("fldDedDsc7").Value = arrD(6)
  Fields("fldDedDsc8").Value = arrD(7)
  Fields("fldDedDsc9").Value = arrD(8)
  Fields("fldDedDsc10").Value = arrD(9)
  Fields("fldNext").Value = arrD(10)
End Sub

Private Sub ActiveReport_Initialize()
  ToolBar.Tools.Add "Exit"
  ToolBar.Font.Size = 10
  
End Sub
Private Sub ActiveReport_DataInitialize()
  DFile = FreeFile
  Open StartPath & "\PRRPTS\DEDDESC.RPT" For Input As #DFile
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
  Fields.Add "fldNext" '(10)

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

Private Sub GroupFooter1_Format()
  GroupFooter1.Height = 0

End Sub

Private Sub GroupHeader1_Format()
  GroupHeader1.Height = 0

End Sub

Private Sub PageFooter_Format()
'  PageFooter.Height = 0

End Sub

Private Sub PageHeader_Format()
'  PageHeader.Height = 0

End Sub
