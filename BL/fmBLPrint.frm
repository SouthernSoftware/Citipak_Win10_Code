VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLPrint 
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Options"
   ClientHeight    =   2415
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   6900
   Icon            =   "fmBLPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpCombo fpcboPrinters 
      Height          =   405
      Left            =   2745
      TabIndex        =   0
      Top             =   210
      Width           =   3450
      _Version        =   196608
      _ExtentX        =   6085
      _ExtentY        =   714
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
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
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
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
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
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
      ColDesigner     =   "fmBLPrint.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
      Height          =   492
      Left            =   5184
      TabIndex        =   4
      Top             =   960
      Width           =   1308
      _Version        =   131072
      _ExtentX        =   2307
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "fmBLPrint.frx":0D15
   End
   Begin EditLib.fpLongInteger txtCopies 
      Height          =   372
      Left            =   2742
      TabIndex        =   1
      Top             =   840
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
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
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
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
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCancel 
      Height          =   492
      Left            =   5184
      TabIndex        =   5
      Top             =   1632
      Width           =   1308
      _Version        =   131072
      _ExtentX        =   2307
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "fmBLPrint.frx":0EF2
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   4620
      Picture         =   "fmBLPrint.frx":10CE
      Top             =   1080
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
      Left            =   390
      TabIndex        =   3
      Top             =   888
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
      Left            =   678
      TabIndex        =   2
      Top             =   264
      Width           =   1836
   End
End
Attribute VB_Name = "frmBLPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsBLTextBoxOverrider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim vWidth%, vHeight%, vTop%, vLeft%

Private Sub cmdCancel_Click()
  Unload frmBLPrint
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
      SendKeys "%X"
      Call cmdCancel_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
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
      frmBLViewPrint.PrintWSet DPName, Copies
    Else
      frmBLViewPrint.PrintWSet DefPrinter, Copies
    End If
  Else
    frmBLMessageBoxJr.Label1.Caption = "Make A Printer Selection Or Cancel."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  If DidPrint = 2 Then
    MainLog ("Non-posting business licenses printed in text format.")
    DidPrint = 1
  End If
  
  Unload frmBLPrint
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
  Set Over = New clsBLTextBoxOverrider
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
Private Sub FillPrinters(combo As fpCombo)
  Dim cnt As Integer
  For cnt = 0 To (Printers.Count - 1)
    fpcboPrinters.InsertRow = Printers(cnt).DeviceName & Chr(9) & Printers(cnt).Port
  Next
End Sub




