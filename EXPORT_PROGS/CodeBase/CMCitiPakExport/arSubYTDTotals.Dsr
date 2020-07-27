VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSubYTDTotals 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveReport1"
   ClientHeight    =   4380
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   7056
   Icon            =   "arSubYTDTotals.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   12446
   _ExtentY        =   7726
   SectionData     =   "arSubYTDTotals.dsx":08CA
End
Attribute VB_Name = "arSubYTDTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private HFile As Integer
Private Sub ActiveReport_Initialize()
  Me.ToolBar.Tools.Add "&Close"
End Sub
Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
  If Tool = "&Close" Then
    Unload Me
  End If
End Sub
Private Sub ExportReport(X As Integer)
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
  Open StartPath & "\PRRPTS\YTDWAGETOTAL.RPT" For Input As #HFile
  
  Fields.Add "fldFundNum" '(0)
  Fields.Add "fldReg" '(1)
  Fields.Add "fldOT" '(2)
  Fields.Add "fldRegWage" '(3)
  Fields.Add "fldOTWage" '(4)
  
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
  Fields("fldFundNum").Value = arr(0)
  Fields("fldReg").Value = arr(1)
  Fields("fldOT").Value = arr(2)
  Fields("fldRegWage").Value = arr(3)
  Fields("fldOTWage").Value = arr(4)

  
End Sub
Private Sub ActiveReport_ReportEnd()
  If HFile <> 0 Then
    Close #HFile
  End If
End Sub

