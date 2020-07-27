VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextMBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Option Explicit
'
'Private WithEvents txtMbox As MaskEdBox
'
'
'
'Property Set SetTextMBox(ByVal txtTextMBox As MaskEdBox)
'  'On Error Resume Next
'
'  Set txtMbox = txtTextMBox
'End Property


'
'Private Sub txtMBox_GotFocus()
'  With txtMbox
'    .BackColor = QBColor(4)
'    .ForeColor = vbWindowBackground
'    '.Alignment = vbLeftJustify
'    .SelStart = 0
'    .SelLength = Len(.Text)
'    'If (Not (sbStatus Is Nothing)) Then
'    '  sbStatus.Panels("ToolTips").Text = .ToolTipText
'    'End If
'  End With
'End Sub
'
'Private Sub txtMBox_KeyDown(KeyCode As Integer, Shift As Integer)
'  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      'If (txtBox.SelStart >= Len(txtBox.Text)) Then
'        SendKeys "{Tab}"
'        KeyCode = 0
'      'End If
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
'    Case Else:
'  End Select
'End Sub
'
'Private Sub txtMBox_LostFocus()
'  With txtMbox
'    .BackColor = vbWindowBackground
'    .ForeColor = vbWindowText
'    '.Alignment = vbRightJustify
'    'If (Not (sbStatus Is Nothing)) Then
'    '  sbStatus.Panels("ToolTips").Text = ""
'    'End If
'  End With
'
'End Sub

