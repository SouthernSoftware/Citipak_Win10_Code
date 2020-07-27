VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents txtbox As TextBox
Attribute txtbox.VB_VarHelpID = -1

Private sbStatus As StatusBar

'Friend Property Set StatusBar(ByRef sb As StatusBar)
'  Set sbStatus = sb
'End Property

'Property Set SetTextBox(ByVal txtTextBox As TextBox)
'  'On Error Resume Next
'
'  Set txtbox = txtTextBox
'End Property



'Private Sub txtBox_GotFocus()
'  With txtbox
'    .BackColor = QBColor(4)
'    .ForeColor = vbWindowBackground
'    .Alignment = vbLeftJustify
'    .SelStart = 0
'    .SelLength = Len(.Text)
'    'If (Not (sbStatus Is Nothing)) Then
'    '  sbStatus.Panels("ToolTips").Text = .ToolTipText
'    'End If
'  End With
'End Sub
'
'Private Sub txtBox_KeyDown(KeyCode As Integer, Shift As Integer)
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

'Private Sub txtBox_LostFocus()
'  With txtbox
'    .BackColor = vbWindowBackground
'    .ForeColor = vbWindowText
'    '.Alignment = vbRightJustify
'    'If (Not (sbStatus Is Nothing)) Then
'    '  sbStatus.Panels("ToolTips").Text = ""
'    'End If
'  End With
'
'End Sub
'



