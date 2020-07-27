VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arDistAcctSummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveReport1"
   ClientHeight    =   4380
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   7905
   Icon            =   "arDistAcctSummary.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   13944
   _ExtentY        =   7726
   SectionData     =   "arDistAcctSummary.dsx":08CA
End
Attribute VB_Name = "arDistAcctSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private TFile As Integer
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
    Unload frmLoadingRpt
    MsgBox "Error Number: " & Str(Number) & " " & Description, vbOKOnly, "Printer Error"
    Unload Me
  End If
  CancelDisplay = True 'removes the error message
End Sub
Private Sub ActiveReport_Initialize()
  ToolBar.Tools.Add "Exit"
  ToolBar.Font.Size = 10
  
End Sub
Private Sub ActiveReport_DataInitialize()
  TFile = FreeFile
  Open StartPath & "\PRRPTS\DISTACCTTOT.RPT" For Input As #TFile
  
  Fields.Add "fldTAcctNum" '(0)
  Fields.Add "fldTSalPct" '(1)
  Fields.Add "fldTRegHrs" '(2)
  Fields.Add "fldTOTHrs" '(3)
  Fields.Add "fldTRegPay" '(4)
  Fields.Add "fldTOTPay" '(5)
  Fields.Add "fldTother" '(6)
  Fields.Add "fldTGrsPy" '(7)
  Fields.Add "fldTSocSec" '(8)
  Fields.Add "fldTMed" '(9)
  Fields.Add "fldTRet" '(10)
  
  End Sub
Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim tLine As String
  Dim arrT() As String
  If VBA.eof(TFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #TFile, tLine
  arrT = Split(tLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("fldTAcctNum").Value = arrT(0)
  Fields("fldTSalPct").Value = arrT(1)
  Fields("fldTRegHrs").Value = arrT(2)
  Fields("fldTOTHrs").Value = arrT(3)
  Fields("fldTRegPay").Value = arrT(4)
  Fields("fldTOTPay").Value = arrT(5)
  Fields("fldTother").Value = arrT(6)
  Fields("fldTGrsPy").Value = arrT(7)
  Fields("fldTSocSec").Value = arrT(8)
  Fields("fldTMed").Value = arrT(9)
  Fields("fldTRet").Value = arrT(10)

End Sub

Private Sub ActiveReport_ReportEnd()
  If TFile <> 0 Then
    Close #TFile
  End If
End Sub
