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
Dim Back As Long

Property Set SetCommandButton(ByVal cmdCommandButton As CommandButton)
  On Error Resume Next
  Set cmdButton = cmdCommandButton
  Back = cmdButton.BackColor
  ''on error goto 0
End Property

Private Sub cmdButton_Click()
'  Beep
End Sub

Private Sub cmdButton_GotFocus()
  'Dim FntSize As Double, Cnt As Double
  With cmdButton
   'FntSize = .FontSize
    .FontBold = True
    .FontUnderline = True
    'Cnt = 1
   ' Do
     ' Cnt = Cnt + 0.15
     ' .FontSize = Cnt
    '  DoEvents
    'Loop Until .FontSize >= FntSize
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
End Sub

Private Sub cmdButton_LostFocus()
  With cmdButton
    .FontBold = False
    .FontUnderline = False
  End With
End Sub



