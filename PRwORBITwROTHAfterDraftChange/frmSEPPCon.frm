VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmSEPPCon 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEPP Contribution Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmSEPPCon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8840
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5244
      Left            =   2124
      TabIndex        =   4
      Top             =   1794
      Width           =   7404
      _Version        =   196609
      _ExtentX        =   13060
      _ExtentY        =   9250
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDShadowColor=   -2147483633
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmSEPPCon.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3480
         TabIndex        =   3
         Top             =   3450
         Width           =   2370
         _Version        =   196608
         _ExtentX        =   4180
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
         Columns         =   0
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
         ScrollBarH      =   1
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
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmSEPPCon.frx":08E6
      End
      Begin EditLib.fpDateTime fptxtStart 
         Height          =   396
         Left            =   3648
         TabIndex        =   0
         Top             =   1488
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
         _ExtentY        =   698
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
         ThreeDInsideHighlightColor=   -2147483633
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
         ButtonStyle     =   2
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "11/20/2002"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "mm/dd/yyyy"
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   1
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fptxtMatchPct 
         Height          =   396
         Left            =   4080
         TabIndex        =   2
         Top             =   2832
         Width           =   1260
         _Version        =   196608
         _ExtentX        =   2222
         _ExtentY        =   698
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
         ThreeDInsideHighlightColor=   -2147483633
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   1
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 ,"
         MaxLength       =   255
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fptxtEnd 
         Height          =   396
         Left            =   3648
         TabIndex        =   1
         Top             =   2160
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
         _ExtentY        =   698
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
         ThreeDInsideHighlightColor=   -2147483633
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
         ButtonStyle     =   2
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "11/20/2002"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "mm/dd/yyyy"
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   1
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4464
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to generate the SEPP contribution report."
         Top             =   4176
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1217
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         ButtonDesigner  =   "frmSEPPCon.frx":0BDD
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1200
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   4176
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1217
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         ButtonDesigner  =   "frmSEPPCon.frx":0DBC
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Print Option:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1605
         TabIndex        =   9
         Top             =   3555
         Width           =   1500
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "End Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1776
         TabIndex        =   8
         Top             =   2256
         Width           =   1452
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   732
         Left            =   1440
         Top             =   384
         Width           =   4620
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Start Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1776
         TabIndex        =   7
         Top             =   1584
         Width           =   1452
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Matching Pct:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2208
         TabIndex        =   6
         Top             =   2976
         Width           =   1644
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "SEPP Contribution Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   444
         Left            =   1728
         TabIndex        =   5
         Top             =   576
         Width           =   4044
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   5568
      Left            =   1980
      Top             =   1650
      Width           =   7692
   End
End
Attribute VB_Name = "frmSEPPCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdEscape_Click()
   frmReportsProcessing.Show
   DoEvents
   Unload frmSEPPCon
End Sub
Private Sub cmdProcess_Click()
  If fpcomboPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
    Exit Sub
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    Call PrintText
  Else
    Exit Sub
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
      Call cmdEscape_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call LoadSEPPScreen
  Me.HelpContextID = hlpSEPP
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub LoadSEPPScreen()
   Dim Today As String * 10
'   Date$ = FormatDateTime(Date, vbShortDate)
   Today = Date '$
   fptxtStart.Text = Mid(Today, 1, 2) + "-01-" + Mid(Today, 7, 4) '8/21
   'changed from Today
   fptxtEnd.Text = Today
   fpcomboPrintOpt.AddItem "Graphical"
   fpcomboPrintOpt.AddItem "Text"
   fpcomboPrintOpt.Text = "Graphical"
End Sub

Private Sub PrintGraphics()
  Dim Dash As String * 80, MonthNum As Integer
  Dim EPct As Currency
  Dim LowDate&, HighDate&, RptName$
  Dim EmpRecSize&, TRecSize&, TransRecNum&, TEMatch As Currency
  Dim EmpIdxLNameHandle As Integer, RptTitle$
  Dim EmpIdxRec As NameSortIdxType, RecNo&
  Dim Emp2Handle As Integer, IdxRecLen As Integer
  Dim Emp2Rec As EmpData2Type, RecNum&, x&, cnt As Integer
  Dim IdxFileSize&, NumOfRecs&, MaxLines As Integer
  Dim RHandle As Integer, DHandle As Integer, THandle As Integer
  Dim UsingThisOne As Boolean, EMatchAmt As Currency, GrossAmt As Currency, RETAMT As Currency
  Dim LineCnt As Integer, FF$, GTotal As Currency, RTotal As Currency
  Dim Page As Integer, UTemp$, UnitHandle As Integer
  Dim dlm$, SSN$, Dates$
  Dim ThisCnt As Integer
  
  dlm$ = "~"
  If fptxtStart.Text = "" Then
     MsgBox "Please enter a Starting Date"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtStart.Text) = False Or Len(fptxtStart.Text) <> 10 Then
     MsgBox "Please enter a valid Starting Date (##-##-####)"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If fptxtEnd.Text = "" Then
     MsgBox "Please enter an Ending Date"
     fptxtEnd.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtEnd.Text) = False Or Len(fptxtEnd.Text) <> 10 Then
     MsgBox "Please enter a valid Ending Date (##-##-####)"
     fptxtEnd.SetFocus
     Exit Sub
  End If
  
  If fptxtMatchPct.Text = "" Then
     MsgBox "Please enter a value in the Matching Pct field."
     fptxtMatchPct.SetFocus
     Exit Sub
  End If
  
  LowDate = Date2Num(fptxtStart.Text) '8/26 reworked this
  'error check away from DOS code using OKFlag and FirstTime flag
  HighDate = Date2Num(fptxtEnd.Text)
  MonthNum = QPTrim$(Mid$(fptxtEnd.Text, 1, 2))
  EPct = Val(fptxtMatchPct.Text)
  If LowDate > HighDate Then
    MsgBox "ERROR: The ending date is before the beginning date."
    fptxtStart.SetFocus
    Exit Sub
  End If

  RptName$ = "PRRPTS\SEPPCONTG.RPT"

  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType

  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))
  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  
  FrmShowPctComp.Label1 = "SEPP Contribution Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False

  ReDim IdxBuff(1 To NumOfRecs) As Integer
  OpenEmpIdxLNameFile EmpIdxLNameHandle
  RecNum = LOF(EmpIdxLNameHandle) \ IdxRecLen
  For x = 1 To RecNum
     Get EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  RHandle = FreeFile
  On Error GoTo ErrorHandler
  Open RptName$ For Output As #1
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TRecSize
  
  ReDim FundList(1 To NumOfRecs&) As Long
  OpenEmpData2File DHandle
  For RecNo = 1 To RecNum
    UsingThisOne = False
    EMatchAmt = 0
    GrossAmt = 0
    RETAMT = 0
    Get DHandle, CLng(IdxBuff(RecNo)), Emp2Rec
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEm8
    End If
    TransRecNum& = CLng(Emp2Rec.LastTransRec)
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HighDate
         If Len(QPTrim$(Emp2Rec.EMPRETNO)) Then
          GrossAmt = OldRound#(GrossAmt + TransHRec(1).GrossPay)
          EMatchAmt = OldRound#(EMatchAmt + OldRound#(TransHRec(1).GrossPay * (EPct * 0.01)))
          TEMatch = OldRound#(TEMatch + EMatchAmt)
          UsingThisOne = True
        End If
      Case Else
      End Select

      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          GoSub PrintEmpSEPPLine
        End If
        Exit Do
      Else
        TransRecNum& = TransHRec(1).PrevTransRec
      End If
    Loop

SkipEm8:

    FrmShowPctComp.ShowPctComp RecNo, RecNum
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next

  Close THandle
  Close DHandle   'open employee data file
  
  Close
'****************************************************************
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  If ThisCnt = 0 Then
    MsgBox "There were no wages paid for the dates entered."
    Exit Sub
  End If
  
  arSeppCon.Show
  frmLoadingRpt.Show
  MainLog ("SEPP Contribution Report processed.")
Exit Sub

PrintEmpSEPPLine:
  GTotal = OldRound#(GTotal + GrossAmt)
  RTotal = OldRound#(RTotal + RETAMT)
  If QPTrim$(Emp2Rec.EmpSSN) = "UNKNOWN" Then
    SSN$ = "UNKNOWN"
  Else
    SSN$ = Left$(Emp2Rec.EmpSSN, 3) & "-" & Mid$(Emp2Rec.EmpSSN, 4, 2) & "-" & Mid$(Emp2Rec.EmpSSN, 6, 4)
  End If
  Dates$ = MonthName$(MonthNum) & " " & Mid$(fptxtEnd.Text, 7, 4)
  ThisCnt = ThisCnt + 1
  '                    0                   1
  Print #1, QPTrim$(Unit(1).UFEMPR); dlm; Dates$; dlm;
  '          2
  Print #1, SSN$; dlm;
  '                     3                                4
  Print #1, QPTrim$(Emp2Rec.EMPRETNO); dlm; QPTrim$(Emp2Rec.EmpLName) & ", " & QPTrim$(Emp2Rec.EmpFName); dlm;
  '                       5                                    6                                   7
  Print #1, Using("###,##0.00", GrossAmt); dlm; Using("##,##0.00", EMatchAmt); dlm; Using("$#,###,##0.00", GTotal); dlm;
  '                    8
  Print #1, Using("#,###,##0.00", TEMatch)
  
  Return

ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmSEPPCon.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintText()
  Dim Dash As String * 80, MonthNum As Integer
  Dim EPct As Currency
  Dim LowDate&, HighDate&, RptName$
  Dim EmpRecSize&, TRecSize&, TransRecNum&, TEMatch As Currency
  Dim EmpIdxLNameHandle As Integer, RptTitle$
  Dim EmpIdxRec As NameSortIdxType, RecNo&
  Dim Emp2Handle As Integer, IdxRecLen As Integer
  Dim Emp2Rec As EmpData2Type, RecNum&, x&, cnt As Integer
  Dim IdxFileSize&, NumOfRecs&, MaxLines As Integer
  Dim RHandle As Integer, DHandle As Integer, THandle As Integer
  Dim UsingThisOne As Boolean, EMatchAmt As Currency, GrossAmt As Currency, RETAMT As Currency
  Dim LineCnt As Integer, FF$, GTotal As Currency, RTotal As Currency
  Dim Page As Integer, UTemp$, UnitHandle As Integer
  Dim ThisCnt As Integer
  
  If fptxtStart.Text = "" Then
     MsgBox "Please enter a Starting Date"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtStart.Text) = False Or Len(fptxtStart.Text) <> 10 Then
     MsgBox "Please enter a valid Starting Date (##-##-####)"
     fptxtStart.SetFocus
     Exit Sub
  End If

  If fptxtEnd.Text = "" Then
     MsgBox "Please enter an Ending Date"
     fptxtEnd.SetFocus
     Exit Sub
  End If

  If CheckValDate(fptxtEnd.Text) = False Or Len(fptxtEnd.Text) <> 10 Then
     MsgBox "Please enter a valid Ending Date (##-##-####)"
     fptxtEnd.SetFocus
     Exit Sub
  End If
  
  If fptxtMatchPct.Text = "" Then
     MsgBox "Please enter a value in the Matching Pct field."
     fptxtMatchPct.SetFocus
     Exit Sub
  End If
  
  FF$ = Chr$(12)
  
  LowDate = Date2Num(fptxtStart.Text) '8/26 reworked this
  'error check away from DOS code using OKFlag and FirstTime flag
  HighDate = Date2Num(fptxtEnd.Text)
  MonthNum = QPTrim$(Mid$(fptxtEnd.Text, 1, 2))
  EPct = Val(fptxtMatchPct.Text)
  If LowDate > HighDate Then
    MsgBox "ERROR: The ending date is before the beginning date."
    fptxtStart.SetFocus
    Exit Sub
  End If

  RptName$ = "PRRPTS\SEPPCONT.RPT"
  Dash = String$(80, "-")

  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3

  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))
  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  
  FrmShowPctComp.Label1 = "SEPP Contribution Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False

  ReDim IdxBuff(1 To NumOfRecs) As Integer
  OpenEmpIdxLNameFile EmpIdxLNameHandle
  RecNum = LOF(EmpIdxLNameHandle) \ IdxRecLen
  For x = 1 To RecNum
     Get EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  MaxLines = 55
  RptTitle$ = "SEPP Contribution Report."
  
  RHandle = FreeFile
  Open RptName$ For Output As #1
  RPTSetupPRN 18, 1
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TRecSize
  
  ReDim FundList(1 To NumOfRecs&) As Long
  OpenEmpData2File DHandle
  GoSub SEPPRptHeader
  For RecNo = 1 To RecNum
    UsingThisOne = False
    EMatchAmt = 0
    GrossAmt = 0
    RETAMT = 0
    Get DHandle, CLng(IdxBuff(RecNo)), Emp2Rec
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEm8
    End If
    TransRecNum& = CLng(Emp2Rec.LastTransRec)
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HighDate
         If Len(QPTrim$(Emp2Rec.EMPRETNO)) Then
          GrossAmt = OldRound#(GrossAmt + TransHRec(1).GrossPay)
          EMatchAmt = OldRound#(EMatchAmt + OldRound#(TransHRec(1).GrossPay * (EPct * 0.01)))
          TEMatch = OldRound#(TEMatch + EMatchAmt)
          UsingThisOne = True
        End If
      Case Else
      End Select

      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          GoSub PrintEmpSEPPLine
          If LineCnt >= MaxLines Then
            Print #1, FF$
            GoSub SEPPRptHeader
          End If
        End If
        Exit Do
      Else
        TransRecNum& = TransHRec(1).PrevTransRec
      End If
    Loop

SkipEm8:

    FrmShowPctComp.ShowPctComp RecNo, RecNum
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  GoSub SEPPTotals

  Close THandle
  Close DHandle   'open employee data file
  RPTSetupPRN 123, RHandle '8/15...123 is the default end code
  
  Close
'****************************************************************
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  If ThisCnt = 0 Then
    MsgBox "There were no wages paid for the dates entered."
    Exit Sub
  End If
  
  ViewPrint RptName$, RptTitle$
  
  MainLog ("SEPP Contribution Report processed.")
  
  Exit Sub

PrintEmpSEPPLine:
  ThisCnt = ThisCnt + 1
  Print #1, Left$(Emp2Rec.EmpSSN, 3); "-"; Mid$(Emp2Rec.EmpSSN, 4, 2); "-"; Mid$(Emp2Rec.EmpSSN, 6, 4);
  Print #1, Tab(14); Left$(Emp2Rec.EMPRETNO, 15);
  Print #1, Tab(30); QPTrim$(Emp2Rec.EmpLName); ", "; QPTrim$(Emp2Rec.EmpFName);
  Print #1, Tab(55); Using("###,##0.00", GrossAmt); Tab(69); Using("##,##0.00", EMatchAmt)
  'PRINT #1,
  LineCnt = LineCnt + 1     'employeesprinted = employeesprinted + 1
  GTotal = OldRound#(GTotal + GrossAmt)
  RTotal = OldRound#(RTotal + RETAMT)
  Return

SEPPTotals:
  Print #1, Dash
  Print #1, Tab(28); "Totals:";
  Print #1, Tab(52); Using("$#,###,##0.00", GTotal); Tab(66); Using("#,###,##0.00", TEMatch)
  Print #1, FF$
  Return

SEPPRptHeader:
  Page = Page + 1
  RSet Pg(1) = Str$(Page)
  UTemp$ = Space$(70)
  LSet UTemp$ = QPTrim$(Unit(1).UFEMPR)
  Mid$(UTemp$, 62) = "Page:" + Pg(1)
  Print #1, UTemp$
  Print #1, "SEPP Contribution Report."
  Print #1, "Month: "; MonthName$(MonthNum); " "; Mid$(fptxtEnd.Text, 7, 4)
  'PRINT #1, "                                                   Wages Subject     Employer  "
  Print #1, "Soc Sec #    Ret #           Employee Name           Gross Wage  Contribution"
  Print #1, Dash
  'PRINT #1, ""
  LineCnt = 5
Return

End Sub

Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboPrintOpt.ListIndex = -1
  End If
  If fpcomboPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdEscape.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

