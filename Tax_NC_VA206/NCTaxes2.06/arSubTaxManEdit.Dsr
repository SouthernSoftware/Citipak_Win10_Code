VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSubTaxManEdit 
   BorderStyle     =   0  'None
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   15028
   _ExtentY        =   9022
   SectionData     =   "arSubTaxManEdit.dsx":0000
End
Attribute VB_Name = "arSubTaxManEdit"
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
  Open StartPath & "\TAXRPTS\TXMANSUB1.RPT" For Input As #hFile
  Fields.Add ("fldYear") '0)
  Fields.Add ("fldPrinc") '1)
  Fields.Add ("fldInt") '2)
  Fields.Add ("fldAdvCol") '3)
  Fields.Add ("fldLateList") '4)
  Fields.Add ("fldOpt1") '5)
  Fields.Add ("fldOpt2") '6)
  Fields.Add ("fldOpt3") '7)
  Fields.Add ("fldTotal") '8)
  Fields.Add ("fldType") '9)
  Fields.Add ("fldOpt1Desc") '10)
  Fields.Add ("fldOpt2Desc") '11)
  Fields.Add ("fldOpt3Desc") '12)
End Sub

Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
  If Number <> 5007 Then 'ignore the no printer warning
    frmTaxMsg.Label1.Caption = "Error Number: " & Str(Number) & " " & Description
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Unload Me
  End If
  CancelDisplay = True 'removes the error message

End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
  Dim sLine As String
  Dim arr() As String

  If VBA.eof(hFile) Then
    eof = True
    Exit Sub
  Else
    eof = False
  End If
  Line Input #hFile, sLine
  arr = Split(sLine, "~")
  ' Here we set the values of the fields that we defines as unbound
  ' or user defined.
  Fields("fldYear").Value = arr(0)
  Fields("fldPrinc").Value = arr(1)
  Fields("fldInt").Value = arr(2)
  Fields("fldAdvCol").Value = arr(3)
  Fields("fldLateList").Value = arr(4)
  Fields("fldOpt1").Value = arr(5)
  Fields("fldOpt2").Value = arr(6)
  Fields("fldOpt3").Value = arr(7)
  Fields("fldTotal").Value = arr(8)
  If arr(9) = "C" Then
    Fields("fldType").Value = "COMBINED"
  ElseIf arr(9) = "R" Then
    Fields("fldType").Value = "REAL ONLY"
  ElseIf arr(9) = "P" Then
    Fields("fldType").Value = "PERSONAL ONLY"
  End If
  Fields("fldOpt1Desc").Value = arr(10)
  Fields("fldOpt2Desc").Value = arr(11)
  Fields("fldOpt3Desc").Value = arr(12)
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
Private Sub GroupHeader2_Format()
  If QPTrim$(Fields("fldOpt1Desc").Value) = "" Then
    Field7.Visible = False
    Field8.Visible = False
  Else
    Field7.Visible = True
    Field8.Visible = True
  End If
  
  If QPTrim$(Fields("fldOpt2Desc").Value) = "" Then
    Field9.Visible = False
    Field10.Visible = False
  Else
    Field9.Visible = True
    Field10.Visible = True
  End If
  
  If QPTrim$(Fields("fldOpt3Desc").Value) = "" Then
    Field11.Visible = False
    Field12.Visible = False
  Else
    Field11.Visible = True
    Field12.Visible = True
  End If
  
End Sub
