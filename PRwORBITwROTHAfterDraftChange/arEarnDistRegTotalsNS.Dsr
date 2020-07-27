VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arEarnDistRegTotalsNS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveReport1"
   ClientHeight    =   4380
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   6255
   Icon            =   "arEarnDistRegTotalsNS.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   11033
   _ExtentY        =   7726
   SectionData     =   "arEarnDistRegTotalsNS.dsx":08CA
End
Attribute VB_Name = "arEarnDistRegTotalsNS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private TFile As Integer

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
  Open StartPath & "\PRRPTS\DISTTOTALSNS.RPT" For Input As #TFile
  
  Fields.Add "fldEAcctNum" '(0)
  Fields.Add "fldSalPct" '(1)
  Fields.Add "fldRegHrs" '(2)
  Fields.Add "fldOTHrs" '(3)
  Fields.Add "fldRegPay" '(4)
  Fields.Add "fldOTPay" '(5)
  Fields.Add "fldETother" '(6)
  Fields.Add "fldGrsPy" '(7)
  Fields.Add "fldSocSec" '(8)
  Fields.Add "fldMed" '(9)
  Fields.Add "fldRet" '(10)
  Fields.Add "fldEmployer" '(11)
  End Sub
Private Sub ActiveReport_FetchData(eof As Boolean)
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
  Fields("fldEAcctNum").Value = arrT(0)
  Fields("fldSalPct").Value = arrT(1)
  Fields("fldRegHrs").Value = arrT(2)
  Fields("fldOTHrs").Value = arrT(3)
  Fields("fldRegPay").Value = arrT(4)
  Fields("fldOTPay").Value = arrT(5)
  Fields("fldETother").Value = arrT(6)
  Fields("fldGrsPy").Value = arrT(7)
  Fields("fldSocSec").Value = arrT(8)
  Fields("fldMed").Value = arrT(9)
  Fields("fldRet").Value = arrT(10)
  Fields("fldEmployer").Value = arrT(11)
End Sub
Private Sub ActiveReport_ReportEnd()
  Unload frmLoadingRpt
  If TFile <> 0 Then
    Close #TFile
  End If
End Sub

