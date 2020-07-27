VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WResize_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ControlsProportion
    WidthProportion As Single
    HeightProportion As Single
    TopProportion As Single
    LeftProportion As Single
End Type
Private ArrayOfProportion() As ControlsProportion
Private GridCols As New Collection

Public Sub InitResizeClass(CurrentForm As Form)
    Dim Counter As Integer, CTRL As Control, tmpCols() As Single, i As Integer
    On Error GoTo ErrHendel
    ReDim ArrayOfProportion(0 To CurrentForm.Controls.Count - 1)
    For Counter = 0 To CurrentForm.Controls.Count
rNext:
      Set CTRL = CurrentForm.Controls(Counter)
      With ArrayOfProportion(Counter)
        
        .WidthProportion = CTRL.Width / CurrentForm.ScaleWidth
        .HeightProportion = CTRL.Height / CurrentForm.ScaleHeight
        .TopProportion = CTRL.Top / CurrentForm.ScaleHeight
        .LeftProportion = CTRL.Left / CurrentForm.ScaleWidth
       'If TypeOf CTRL Is MSFlexGrid Then
       '  ReDim tmpCols(CTRL.Cols - 1)
       '  For i = 0 To CTRL.Cols - 1
       '    tmpCols(i) = CTRL.ColWidth(i) / CurrentForm.ScaleWidth
       '  Next
       '  GridCols.Add tmpCols, CTRL.Name
       'End If
      End With
    Next
    Exit Sub
ErrHendel:
    If Counter < CurrentForm.Controls.Count - 1 Then
        Counter = Counter + 1
        Resume rNext
    End If
End Sub
Public Sub ResizeControls(CurrentForm As Form)
    Dim Counter As Integer, CTRL As Control, i As Integer
    'Dim t As MSFlexGrid
    Dim tmpCols() As Single
    On Error GoTo ErrHendel
    For Counter = 0 To CurrentForm.Controls.Count - 1
rNext:
         Set CTRL = CurrentForm.Controls(Counter)
         CTRL.Move ArrayOfProportion(Counter).LeftProportion _
               * CurrentForm.ScaleWidth, _
               ArrayOfProportion(Counter).TopProportion * CurrentForm.ScaleHeight, _
               ArrayOfProportion(Counter).WidthProportion * CurrentForm.ScaleWidth, _
               ArrayOfProportion(Counter).HeightProportion * CurrentForm.ScaleHeight
         'If TypeOf CTRL Is MSFlexGrid Then
         '  tmpCols = GridCols(CTRL.Name)
         '  For i = 0 To CTRL.Cols - 1
         '     CTRL.ColWidth(i) = tmpCols(i) * CurrentForm.ScaleWidth
         '  Next
         'End If
    Next
   
    Exit Sub
ErrHendel:
    If Counter < CurrentForm.Controls.Count - 1 Then
        Counter = Counter + 1
        Resume rNext
    End If
End Sub

