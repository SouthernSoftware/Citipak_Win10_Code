VERSION 5.00
Begin VB.Form frmPrint1 
   Caption         =   "Form2"
   ClientHeight    =   2508
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   3744
   LinkTopic       =   "Form2"
   ScaleHeight     =   2508
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPrint1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim vWidth%, vHeight%, vTop%, vLeft%

Private Sub cmdCancel_Click()
  
  Unload frmPrint1
  
End Sub
'Private Sub fpcboPrinters_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeySpace Then
'    fpcboPrinters.ListDown = True
'  End If
'End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub cmdPrint_Click()
'  Dim DefPrinter As String, Copies As Integer
'  If fpcboPrinters.ListIndex <> -1 Then
'    fpcboPrinters.Col = 1
'    DefPrinter = fpcboPrinters.ColText
'    If txtCopies > 0 Then
'      Copies = txtCopies
'    Else
'      Copies = 1
'    End If
'    frmViewPrint.PrintWSet DefPrinter, Copies
''    If vbKeyDown = vbKeyEscape Then
''      Printer.KillDoc
''    End If
'  Else
'    MsgBox "Make A Printer Selection Or Cancel.", vbOKOnly, "Invalid Printer Selection"
'    Exit Sub
'  End If
'  Unload frmPrint
End Sub

Private Sub Form_Initialize()
  vWidth = Screen.Width * 0.5      ' Set width of form.
  vHeight = Screen.Height * 0.33  ' Set height of form.
  vLeft = (Screen.Width - vWidth) \ 2   ' Center form horizontally.
  vTop = ((Screen.Height - vHeight) \ 2) + 10  ' Center form vertically.
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
'  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  Me.Width = vWidth
  Me.Height = vHeight
  Me.Left = vLeft
  Me.Top = vTop
''  FillPrinters fpcboPrinters
'  fpcboPrinters.Col = 1
'  fpcboPrinters.SearchText = Printer.Port
'      fpcboPrinters.Action = 0
'      If fpcboPrinters.SearchIndex <> -1 Then
'        fpcboPrinters.ListIndex = fpcboPrinters.SearchIndex
'      Else
'        fpcboPrinters.ListIndex = 0
'      End If
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub FillPrinters(combo As fpCombo)
'Dim cnt As Integer
'For cnt = 0 To (Printers.Count - 1)
'  fpcboPrinters.InsertRow = Printers(cnt).DeviceName & Chr(9) & Printers(cnt).Port
'Next
End Sub

