VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arFASubMaster 
   BorderStyle     =   0  'None
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   14446
   _ExtentY        =   7726
   SectionData     =   "arFASubMaster.dsx":0000
End
Attribute VB_Name = "arFASubMaster"
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
  Open StartPath & "\FARPTS\FAGTMASTER.RPT" For Input As #HFile
  
  Fields.Add "fldDeptNum" '0
  Fields.Add "fldItemCnt" '1
  Fields.Add "fldOrigCost" '2
  Fields.Add "fldDepTotal" '3
  Fields.Add "fldBookTot" '4
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
  Fields("fldItemCnt").Value = arr(1)
  Fields("fldOrigCost").Value = arr(2)
  Fields("fldDepTotal").Value = arr(3)
  Fields("fldBookTot").Value = arr(4)
End Sub
Private Sub ActiveReport_ReportEnd()
  If HFile <> 0 Then
    Close #HFile
  End If
End Sub




