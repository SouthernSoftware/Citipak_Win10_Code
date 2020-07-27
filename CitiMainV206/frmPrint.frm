VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmPrint 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Options"
   ClientHeight    =   2436
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   6912
   ForeColor       =   &H00000000&
   Icon            =   "frmPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2436
   ScaleWidth      =   6912
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpCombo fpcboPrinters 
      Height          =   384
      Left            =   2016
      TabIndex        =   0
      Top             =   288
      Width           =   4644
      _Version        =   196608
      _ExtentX        =   8191
      _ExtentY        =   677
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BackColor       =   16777215
      ForeColor       =   0
      Text            =   ""
      Columns         =   2
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   3
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   12632256
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   12632256
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   12632256
      ExtendRow       =   0
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   5
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrint.frx":08CA
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Print"
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
      Left            =   5136
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1056
      Width           =   1308
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Esc &Cancel"
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
      Left            =   5136
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1776
      Width           =   1308
   End
   Begin EditLib.fpLongInteger txtCopies 
      Height          =   372
      Left            =   2472
      TabIndex        =   1
      Top             =   912
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
      ThreeDOutsideShadowColor=   8421504
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
      ThreeDTextShadowColor=   12632256
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
      ButtonColor     =   14737632
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Image Image1 
      Height          =   276
      Left            =   4560
      Picture         =   "frmPrint.frx":0C7D
      Top             =   1152
      Width           =   288
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
      TabIndex        =   5
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
      Left            =   192
      TabIndex        =   4
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
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim vWidth%, vHeight%, vTop%, vLeft%
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'  If ((UnloadMode = vbFormControlMenu)) Then
'    Cancel = True
'  End If
'End Sub

Private Sub cmdCancel_Click()
  
  Unload frmPrint
  
End Sub
Private Sub fpcboPrinters_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboPrinters.ListDown = True
  End If
  If fpcboPrinters.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
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
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub cmdPrint_Click()
  Dim DefPrinter As String, Copies As Integer, DPName As String
  If fpcboPrinters.ListIndex <> -1 Then
    fpcboPrinters.col = 0
    DPName = QPTrim(fpcboPrinters.ColText)
    'MsgBox DPName, vbOKOnly
    fpcboPrinters.col = 1
    DefPrinter = QPTrim(fpcboPrinters.ColText)
    'MsgBox DefPrinter, vbOKOnly
    'Find Me this is to debug problem with xp
    'MsgBox "DefPrinter Port is *" + DefPrinter + "*" + " Name *" + DPName + "*"
    If txtCopies > 0 Then
      Copies = txtCopies
    Else
      Copies = 1
    End If
    'MsgBox "DPName -" + DPName, vbOKOnly
    'MsgBox "DefPrinter -" + DefPrinter, vbOKOnly
  'Added this to allow for winxp network printer port names of ne00:, etc.
  'the device name worked so use that instead, but only for network printers.
    If InStr(1, DPName, "\\", vbTextCompare) Then
      'MsgBox "DPName Sent", vbOKOnly
      frmViewPrint.PrintWSet DPName, Copies
    Else
      'MsgBox "DefPrinter Sent", vbOKOnly
      frmViewPrint.PrintWSet DefPrinter, Copies
    End If
'    If vbKeyDown = vbKeyEscape Then
'      Printer.KillDoc
'    End If
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
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
 ' GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  Me.Width = vWidth
  Me.Height = vHeight
  Me.Left = vLeft
  Me.Top = vTop
  FillPrinters fpcboPrinters
  fpcboPrinters.col = 1
  fpcboPrinters.SearchText = Printer.Port
      fpcboPrinters.Action = 0
      If fpcboPrinters.SearchIndex <> -1 Then
        fpcboPrinters.ListIndex = fpcboPrinters.SearchIndex
      Else
        fpcboPrinters.ListIndex = 0
      End If
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub FillPrinters(combo As fpCombo)
Dim cnt As Integer

For cnt = 0 To (Printers.Count - 1)
  fpcboPrinters.InsertRow = Printers(cnt).DeviceName & Chr(9) & Printers(cnt).Port
Next
End Sub
