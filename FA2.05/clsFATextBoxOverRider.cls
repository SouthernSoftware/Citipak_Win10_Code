VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFATextBoxOverRider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Col As Collection

'Dim sbMyStatus As StatusBar

Public Sub OverRide(ByRef frm As Form, Optional ByVal strTag As String = "")
  'Dim tb As clsTextBox
  Dim cb As clsFACommandButton
  'Dim tbm As clsTextMBox
  Dim ctlLoop As Control
  
  '===  Look at all the controls...
  For Each ctlLoop In frm.Controls
    '===  Add textbox to override list only if tags match...
    If (ctlLoop.Tag = strTag) Then
      '===  Right now, only works with TextBox controls...
'      If (TypeOf ctlLoop Is TextBox) Then
'        '===  Create a new custom Textbox override object...
'        Set tb = New clsTextBox
'        '===  Set the Textbox to override equal to this control...
'        Set tb.SetTextBox = ctlLoop
'
'        '===  If
'        If (Not (sbMyStatus Is Nothing)) Then
'          Set tb.StatusBar = sbMyStatus
'        End If
'
'        '===  NOTE:  By adding the tb object to the collection, it stays referenced (in use)
'        '           until the class is destroyed because the col Collection is at the module
'        '           level.
'        col.Add tb
'        ElseIf (TypeOf ctlLoop Is MaskEdBox) Then
'        '===  Create a new custom Textbox override object...
'        Set tbm = New clsTextMBox
'        '===  Set the Textbox to override equal to this control...
'        Set tbm.SetTextMBox = ctlLoop
'
'        col.Add tbm
        If (TypeOf ctlLoop Is CommandButton) Then
        '===  Create a new custom CommandButton override object...
        Set cb = New clsFACommandButton
        '===  Set the CommandButton to override equal to this control...
        Set cb.SetCommandButton = ctlLoop
        '===  NOTE:  By adding the cb object to the collection, it stays referenced (in use)
        '           until the class is destroyed because the col Collection is at the module
        '           level.
        Col.Add cb
        End If
    End If
  Next ctlLoop
End Sub
'
'Property Set StatusBar(ByRef sb As StatusBar)
'  Set sbMyStatus = sb
'End Property

Private Sub Class_Initialize()
  '===  Create the collection...
  Set Col = New Collection
End Sub




