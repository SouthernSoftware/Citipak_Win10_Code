VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arFASubWrnty 
   Caption         =   "ActiveReport1"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   16325
   _ExtentY        =   7726
   SectionData     =   "arFASubWrnty.dsx":0000
End
Attribute VB_Name = "arFASubWrnty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsFATextBoxOverRider
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
Private Sub ExportReport(x As Integer)
End Sub

Private Sub Form_Load()
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
'    'Me.Visible = False
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
  Open StartPath & "\FARPTS\FAGTWRNTYRPT.RPT" For Input As #HFile
  
  Fields.Add "fldDeptNum" '0)
  Fields.Add "fldDeptDesc" '1)
  Fields.Add "fldDeptCnt" '2)
  Fields.Add "fldDeptPurchPr" '3)
  Fields.Add "fldDeptBookVal" '4)
  Fields.Add "fldTotalItems" '5)
  Fields.Add "fldTotPurchPr" '6)
  Fields.Add "fldTotBookVal" '7)
  Fields.Add "fldEnd" '8)
  End Sub
Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String
  Dim x As Integer
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
  Fields("fldDeptNum").Value = arr(0)
  Fields("fldDeptDesc").Value = arr(1)
  Fields("fldDeptCnt").Value = arr(2)
  Fields("fldDeptPurchPr").Value = arr(3)
  Fields("fldDeptBookVal").Value = arr(4)
  Fields("fldTotalItems").Value = arr(5)
  Fields("fldTotPurchPr").Value = arr(6)
  Fields("fldTotBookVal").Value = arr(7)
  Fields("fldEnd").Value = arr(8)
End Sub
Private Sub ActiveReport_ReportEnd()
  If HFile <> 0 Then
    Close #HFile
  End If
End Sub








