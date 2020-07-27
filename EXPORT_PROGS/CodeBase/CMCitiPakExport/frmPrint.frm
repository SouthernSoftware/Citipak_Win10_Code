VERSION 5.00
Begin VB.Form frmPrint 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Options"
   ClientHeight    =   2430
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   6915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox fpcboPrinters 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2685
      ScaleHeight     =   405
      ScaleWidth      =   3450
      TabIndex        =   0
      Top             =   285
      Width           =   3450
   End
   Begin VB.PictureBox cmdCancel 
      Height          =   495
      Left            =   5130
      ScaleHeight     =   435
      ScaleWidth      =   1245
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1770
      Width           =   1305
   End
   Begin VB.PictureBox txtCopies 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   2688
      ScaleHeight     =   375
      ScaleWidth      =   975
      TabIndex        =   1
      Top             =   912
      Width           =   972
   End
   Begin VB.PictureBox cmdPrint 
      Height          =   495
      Left            =   5136
      ScaleHeight     =   435
      ScaleWidth      =   1245
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1056
      Width           =   1305
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   4440
      Picture         =   "frmPrint.frx":0000
      Top             =   1155
      Width           =   360
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Number Of Copies:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Left            =   336
      TabIndex        =   3
      Top             =   960
      Width           =   2124
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select A Printer:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Left            =   624
      TabIndex        =   2
      Top             =   336
      Width           =   1836
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim vWidth%, vHeight%, vTop%, vLeft%

Private Sub cmdCancel_Click()
  Unload frmPrint
  DoEvents
End Sub

Private Sub fpcboPrinters_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPrinters.ListDown = True
  End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdCancel_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%r"
      Call cmdPrint_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub cmdPrint_Click()
  Dim DefPrinter As String, Copies As Integer, DPName As String
  If fpcboPrinters.ListIndex <> -1 Then
    fpcboPrinters.Col = 0
    DPName = QPTrim$(fpcboPrinters.ColText)
    fpcboPrinters.Col = 1
    DefPrinter = fpcboPrinters.ColText
    If txtCopies > 0 Then
      Copies = txtCopies
    Else
      Copies = 1
    End If
    If InStr(1, DPName, "\\", vbTextCompare) Then
      frmViewPrint.PrintWSet DPName, Copies
    Else
      frmViewPrint.PrintWSet DefPrinter, Copies
    End If
  Else
    MsgBox "Make A Printer Selection Or Cancel.", vbOKOnly, "Invalid Printer Selection"
    Exit Sub
  End If
  Unload frmPrint
End Sub

Private Sub Form_Initialize()
  vWidth = Screen.Width * 0.5      ' Set width of form.
  vHeight = Screen.Height * 0.33  ' Set height of form.
  vLeft = (Screen.Width - vWidth) \ 2   ' Center form horizontally.
  vTop = ((Screen.Height - vHeight) \ 2) + 10  ' Center form vertically.
End Sub

Private Sub Form_Load()
  If doAlign = True Then
    txtCopies.Visible = False
    Label2.Visible = False
  End If
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Me.Width = vWidth
  Me.Height = vHeight
  Me.Left = vLeft
  Me.Top = vTop
  FillPrinters fpcboPrinters
  fpcboPrinters.Col = 1
  fpcboPrinters.SearchText = Printer.Port
  fpcboPrinters.Action = 0
  If fpcboPrinters.SearchIndex <> -1 Then
    fpcboPrinters.ListIndex = fpcboPrinters.SearchIndex
  Else
    fpcboPrinters.ListIndex = 0
  End If
  doAlign = False
End Sub

Private Sub Form_Resize()
  Temp_Class.ResizeControls Me
  DoEvents
End Sub
'Private Sub FillPrinters(combo As fpCombo)
'  Dim cnt As Integer
'  For cnt = 0 To (Printers.Count - 1)
'    fpcboPrinters.InsertRow = Printers(cnt).DeviceName & Chr(9) & Printers(cnt).Port
'  Next
'End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'  If ((UnloadMode = vbFormControlMenu)) Then
'    Call Terminate
'    End
'  End If
End Sub

