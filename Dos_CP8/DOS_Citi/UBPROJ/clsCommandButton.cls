VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommandButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents cmdButton As CommandButton
Attribute cmdButton.VB_VarHelpID = -1

Property Set SetCommandButton(ByVal cmdCommandButton As CommandButton)
  Set cmdButton = cmdCommandButton
End Property

Private Sub cmdButton_GotFocus()
  Dim FntSize As Double
  With cmdButton
  .FontBold = True
  .FontUnderline = True
  End With
End Sub

Private Sub cmdButton_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case Else:
  End Select
  DoEvents
End Sub

Private Sub cmdButton_LostFocus()
  With cmdButton
  .FontBold = False
  .FontUnderline = False
  End With
  DoEvents
End Sub

Private Sub cmdButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    SendKeys "%X"
    DoEvents
  End If
End Sub


