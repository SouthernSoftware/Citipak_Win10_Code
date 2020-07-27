VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmPrint2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Options"
   ClientHeight    =   2268
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   6864
   Icon            =   "frmPrint2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2268
   ScaleWidth      =   6864
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAll 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check2"
      Height          =   204
      Left            =   3168
      TabIndex        =   7
      Top             =   1488
      Width           =   204
   End
   Begin VB.CheckBox chkSelection 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      ForeColor       =   &H00000000&
      Height          =   192
      Left            =   3168
      TabIndex        =   6
      Top             =   1152
      Width           =   204
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5328
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1536
      Width           =   1212
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5328
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   768
      Width           =   1212
   End
   Begin EditLib.fpLongInteger txtCopies 
      Height          =   372
      Left            =   2880
      TabIndex        =   2
      Top             =   384
      Width           =   972
      _Version        =   196608
      _ExtentX        =   1714
      _ExtentY        =   656
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   12632256
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   1
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   8421504
      ThreeDTextOffset=   1
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   14737632
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "1"
      MaxValue        =   "100"
      MinValue        =   "1"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      BorderGrayAreaColor=   12632256
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00D0D0D0&
      BorderWidth     =   2
      Height          =   876
      Left            =   2880
      Top             =   1008
      Width           =   1836
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Selection"
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
      Left            =   3456
      TabIndex        =   8
      Top             =   1104
      Width           =   1020
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "All"
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
      Left            =   3504
      TabIndex        =   4
      Top             =   1440
      Width           =   444
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Print Range:"
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
      Height          =   444
      Left            =   1008
      TabIndex        =   5
      Top             =   1104
      Width           =   1404
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
      Top             =   432
      Width           =   2124
   End
End
Attribute VB_Name = "frmPrint2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tried to get this to work for Budget Prep Print Options But no work....
'The copyloop caused printer dialog from vaspread to display for each copy.
'Couldn't pass value of printer selected -- was not system default printer.
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim vWidth%, vHeight%, vTop%, vLeft%

Private Sub chkAll_Click()
  If chkAll.Value = vbChecked Then
    chkSelection.Value = vbUnchecked
  Else
    chkSelection.Value = vbChecked
  End If
End Sub

'Private Sub chkLandscape_Click()
'  If chkLandscape.Value = vbChecked Then
'    chkPortrait.Value = vbUnchecked
'  Else
'    chkPortrait.Value = vbChecked
'  End If
'End Sub

'Private Sub chkPortrait_Click()
'  If chkPortrait.Value = vbChecked Then
'    chkLandscape.Value = vbUnchecked
'  Else
'    chkLandscape.Value = vbChecked
'  End If
'End Sub

Private Sub chkSelection_Click()
  If chkSelection.Value = vbChecked Then
    chkAll.Value = vbUnchecked
  Else
    chkAll.Value = vbChecked
  End If
End Sub

Private Sub cmdCancel_Click()
  Unload frmPrint2
End Sub

Private Sub cmdNext_Click()
  Dim DefPrinter As String, Copies As Integer, PRange As Integer, POrient As Integer
  Dim PPort As Integer
 
  'If fpcboPrinters.ListIndex <> -1 Then
    'fpcboPrinters.Col = 1
    'DefPrinter = fpcboPrinters.ColText
   ' fpcboPrinters.Col = 2
    'PPort = fpcboPrinters.ColText
    If txtCopies > 0 Then
      Copies = txtCopies
    Else
      Copies = 1
    End If
    If chkSelection.Value = 0 Then
      PRange = 0
    Else
      PRange = 1
    End If
'    If chkLandscape.Value = 0 Then
'      POrient = 1
'    Else
'      POrient = 2
'    End If
    'frmBudPrepMaint.PrintLandscp Copies, PRange
  'Else
    'MsgBox "Make A Printer Selection Or Cancel.", vbOKOnly, "Invalid Printer Selection"
    'Exit Sub
  'End If
  Unload frmPrint2
  
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
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  Me.Width = vWidth
  Me.Height = vHeight
  Me.Left = vLeft
  Me.Top = vTop
  'FillPrinters fpcboPrinters
  chkAll.Value = 1
  chkSelection.Value = 0
  'chkLandscape.Value = 1
  'chkPortrait.Value = 0
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
'Private Sub FillPrinters(combo As fpCombo)
'Dim cnt As Integer
'For cnt = 0 To (Printers.Count - 1)
'  fpcboPrinters.InsertRow = Printers(cnt).DeviceName & Chr(9) & Printers(cnt).Port & Chr(9) & cnt
'Next
'End Sub

