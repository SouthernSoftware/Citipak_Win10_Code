VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormOverRider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmForm As Form
Attribute frmForm.VB_VarHelpID = -1

Property Set SetForm(ByVal frmSetForm As Form)
  'On Error Resume Next
  Set frmForm = frmSetForm
End Property


Private Sub frmForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then
    SendKeys "%X"
  End If
End Sub
