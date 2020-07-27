VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmLeaveBenefit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Accrual Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmLeaveBenefit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5910
      Left            =   2160
      TabIndex        =   4
      Top             =   1492
      Width           =   7350
      _Version        =   196609
      _ExtentX        =   12965
      _ExtentY        =   10425
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483627
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmLeaveBenefit.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3405
         TabIndex        =   3
         Top             =   4230
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
         ColDesigner     =   "frmLeaveBenefit.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbDollar 
         Height          =   405
         Left            =   1680
         TabIndex        =   2
         Top             =   2880
         Width           =   4095
         _Version        =   196608
         _ExtentX        =   7223
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
         ColDesigner     =   "frmLeaveBenefit.frx":0BDD
      End
      Begin VB.CheckBox CheckName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Sort by Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   1200
         TabIndex        =   13
         Top             =   3570
         Value           =   1  'Checked
         Width           =   1980
      End
      Begin VB.CheckBox CheckNumber 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Sort by Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   3915
         TabIndex        =   12
         Top             =   3570
         Width           =   2340
      End
      Begin EditLib.fpText fptxtFirstEmpNo 
         Height          =   396
         Left            =   4176
         TabIndex        =   0
         Top             =   1296
         Width           =   1308
         _Version        =   196608
         _ExtentX        =   2307
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
      Begin EditLib.fpText fptxtSecEmpNo 
         Height          =   396
         Left            =   4176
         TabIndex        =   1
         Top             =   1872
         Width           =   1308
         _Version        =   196608
         _ExtentX        =   2307
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
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4170
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to generate the employee accrual report."
         Top             =   4965
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
         ButtonDesigner  =   "frmLeaveBenefit.frx":0ED4
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1290
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   4965
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
         ButtonDesigner  =   "frmLeaveBenefit.frx":10B3
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   720
         Top             =   3480
         Width           =   6015
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Include Dollar Values?"
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
         Left            =   2280
         TabIndex        =   9
         Top             =   2520
         Width           =   2655
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
         Left            =   1635
         TabIndex        =   8
         Top             =   4320
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   1395
         Top             =   390
         Width           =   4620
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Last Employee No:"
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
         Left            =   1248
         TabIndex        =   7
         Top             =   1976
         Width           =   2652
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D0D0D0&
         Caption         =   "First Employee No:"
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
         Left            =   1872
         TabIndex        =   6
         Top             =   1410
         Width           =   2124
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Accrual Report"
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
         Height          =   492
         Left            =   1632
         TabIndex        =   5
         Top             =   528
         Width           =   4044
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   6210
      Left            =   1980
      Top             =   1327
      Width           =   7695
   End
End
Attribute VB_Name = "frmLeaveBenefit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim DollarFlag As Boolean
Dim SplitFlag As Boolean

Private Sub cmdEscape_Click()
   frmReportsProcessing.Show
   DoEvents
   Unload frmLeaveBenefit
End Sub

Private Sub cmdProcess_Click()
  If fpcomboPrintOpt.Text = "Graphical" Then
    If SplitFlag = False Then
      Call PrintGraphics
      Exit Sub
    Else
      Call PrintSplitGraphics
    End If
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    If SplitFlag = False Then
      Call PrintText
      Exit Sub
    Else
      Call PrintSplitText
    End If
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
  Call LoadThisForm
  Me.HelpContextID = hlpLeaveBenefit
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub fpcomboDedNo_Click()

End Sub

Private Sub PrintGraphics()
  
  Dim RecNo As Integer
  Dim MinHrs As Long, RptTitle$, x%
  Dim Image1$, Image2$, cnt As Integer
  Dim LineCnt As Integer, Emp1RecLen As Long
  Dim EmpRecSize As Long
  Dim IdxRecLen As Integer, FF$
  Dim IdxFileSize&, NumOfRecs As Long
  Dim EmpIdxNNameHandle As Integer, Page As Integer
  Dim EmpIdxLNameHandle As Integer
  Dim RHandle As Integer, DHandle As Integer, THandle As Integer
  Dim UnitHandle As Integer, EHandle1 As Integer
  Dim FirstEmp&, LastEmp&, EmpNo&, RptName$
  Dim UTemp$, MaxLines As Integer, CrLf$
  Dim Emp2Rec As EmpData2Type
  Dim VBalTotal As String * 13
  Dim VTotal As Double
  Dim SBalTotal As String * 13
  Dim STotal As Double
  Dim CBalTotal As String * 12
  Dim CTotal As Double
  Dim PBalTotal As String * 13
  Dim PTotal As Double
  Dim HBalTotal As String * 13
  Dim HTotal As Double
  Dim NumOfEmps As String * 22
  Dim NbrOfEmps As Integer
  Dim dlm$
  Dim TotHrs As Double
  Dim GTotHrs As Double
  Dim TotPay As Double
  Dim GTotPay As Double
  Dim VPay As Double
  Dim SPay As Double
  Dim CPay As Double
  Dim PPay As Double
  Dim HPay As Double
  
  dlm$ = "~"
  MinHrs = -10000
  Image1$ = "#,##0.00"
  Image2$ = "###,##0.00"
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  ReDim ENumb(1) As String * 11
  ReDim EName(1) As String * 19
  ReDim VBal(1) As String * 13
  ReDim SBal(1) As String * 13
  ReDim CBal(1) As String * 13
  ReDim PBal(1) As String * 12
  ReDim HBal(1) As String * 13
  ReDim LTbl(1) As String * 6
  
  EmpRecSize = Len(Emp2Rec)
'--------------------------------------------------------
  RptName$ = "PRRPTS\BENEACCRG.RPT"
  If CheckNumber.Value = 1 Then
    OpenEmpIdxNNameFile EmpIdxNNameHandle
    IdxRecLen = 2
    NumOfRecs = LOF(EmpIdxNNameHandle) \ IdxRecLen
    If NumOfRecs = 0 Then
      MsgBox "No records on file."
      Close EmpIdxNNameHandle
      Exit Sub
    End If
  ElseIf CheckName.Value = 1 Then
    OpenEmpIdxLNameFile EmpIdxLNameHandle
    IdxRecLen = 2
    NumOfRecs = LOF(EmpIdxLNameHandle) \ IdxRecLen
    If NumOfRecs = 0 Then
      MsgBox "No records on file."
      Close EmpIdxLNameHandle
      Exit Sub
    End If
  End If
  
  If fptxtFirstEmpNo.Text = "" Then
    MsgBox "Please enter a First Employee Number"
    fptxtFirstEmpNo.SetFocus
    Close EmpIdxLNameHandle
    Close EmpIdxNNameHandle
    Exit Sub
  End If
  
  If fptxtSecEmpNo.Text = "" Then
    MsgBox "Please enter a Second Employee Number"
    fptxtSecEmpNo.SetFocus
    Close EmpIdxLNameHandle
    Close EmpIdxNNameHandle
    Exit Sub
  End If
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    If CheckNumber.Value = 1 Then
      Get EmpIdxNNameHandle, x, IdxBuff(x)
    ElseIf CheckName.Value = 1 Then
      Get EmpIdxLNameHandle, x, IdxBuff(x)
    End If
  Next x
  Close EmpIdxLNameHandle
  Close EmpIdxNNameHandle
'-----------------------------------------------------------
  RHandle = FreeFile
  On Error GoTo ErrorHandler
  Open RptName$ For Output As RHandle
  OpenTransHistFile THandle
  OpenEmpData2File DHandle
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  ReDim Emp1Rec(1) As EmpData1Type
  Emp1RecLen = Len(Emp1Rec(1))
  OpenEmpData1File EHandle1
  Get EHandle1, IdxBuff(1), Emp1RecLen
  FirstEmp& = Val(Emp1Rec(1).EmpNo)
  Get EHandle1, IdxBuff(NumOfRecs), Emp1RecLen
  LastEmp& = Val(Emp1Rec(1).EmpNo)
  Close EHandle1
  
  FirstEmp& = Val(fptxtFirstEmpNo.Text)
  LastEmp& = Val(fptxtSecEmpNo.Text)
  If LastEmp& < FirstEmp& Then
    MsgBox "ERROR: The Last Employee Number is less than the First Employee Number"
    fptxtSecEmpNo.SetFocus
    Close
    Exit Sub
  End If
  
  FrmShowPctComp.Label1 = "Payroll Deduction Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  OpenEmpData2File DHandle
 
  For RecNo = 1 To NumOfRecs
    Get DHandle, IdxBuff(RecNo), Emp2Rec
    EmpNo& = Val(QPTrim$(Emp2Rec.EmpNo))
    If (EmpNo& < FirstEmp& Or EmpNo& > LastEmp&) Or Emp2Rec.Deleted Or Emp2Rec.EMPTDATE <> 0 Then
      GoTo SkipEmBene
    End If
    GoSub PrintEmpBalance
SkipEmBene:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload FrmShowPctComp
      GoTo DedExitRpt
    End If
  Next
  Close DHandle
  Close RHandle
  
  arLvBnftRpt.Show
  frmLoadingRpt.Show
  EnableCloseButton Me.hwnd, True
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  MainLog ("Leave Benefit Report processed.")
  Exit Sub
  
PrintEmpBalance:
  If Emp2Rec.EMPCTBAL < MinHrs Then
    Emp2Rec.EMPCTBAL = 0
  End If
  If Emp2Rec.EMPVBAL < MinHrs Then
    Emp2Rec.EMPVBAL = 0
  End If
  If Emp2Rec.EMPSLBAL < MinHrs Then
    Emp2Rec.EMPSLBAL = 0
  End If
  If Emp2Rec.PERBAL < MinHrs Then
    Emp2Rec.PERBAL = 0
  End If
  If Emp2Rec.HOLBAL < MinHrs Then
    Emp2Rec.HOLBAL = 0
  End If
  
  LSet ENumb(1) = QPTrim$(Emp2Rec.EmpNo)
  NbrOfEmps = NbrOfEmps + 1
  NumOfEmps = Using("#####", NbrOfEmps)
  TotHrs = 0
  TotPay = 0
  LSet EName(1) = QPTrim$(Emp2Rec.EmpLName) + ", " + QPTrim$(Emp2Rec.EmpFName)
  RSet VBal(1) = Using(Image2$, Emp2Rec.EMPVBAL)
  VTotal = VTotal + Emp2Rec.EMPVBAL
  VPay = FigurePayAmts(Emp2Rec.EMPVBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  TotPay = TotPay + VPay
  TotHrs = TotHrs + Emp2Rec.EMPVBAL
  VBalTotal = Using(Image2$, VTotal)
  RSet SBal(1) = Using(Image2$, Emp2Rec.EMPSLBAL)
  STotal = STotal + Emp2Rec.EMPSLBAL
  SPay = FigurePayAmts(Emp2Rec.EMPSLBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  TotPay = TotPay + SPay
  TotHrs = TotHrs + Emp2Rec.EMPSLBAL
  SBalTotal = Using(Image2$, STotal)
  RSet CBal(1) = Using(Image2$, Emp2Rec.EMPCTBAL)
  CTotal = CTotal + Emp2Rec.EMPCTBAL
  CPay = FigurePayAmts(Emp2Rec.EMPCTBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  TotPay = TotPay + CPay
  TotHrs = TotHrs + Emp2Rec.EMPCTBAL
  CBalTotal = Using(Image2$, CTotal)
  RSet PBal(1) = Using(Image2$, Emp2Rec.PERBAL)
  PTotal = PTotal + Emp2Rec.PERBAL
  PPay = FigurePayAmts(Emp2Rec.PERBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  TotPay = TotPay + PPay
  TotHrs = TotHrs + Emp2Rec.PERBAL
  PBalTotal = Using(Image2$, PTotal)
  RSet HBal(1) = Using(Image2$, Emp2Rec.HOLBAL)
  HTotal = HTotal + Emp2Rec.HOLBAL
  HPay = FigurePayAmts(Emp2Rec.HOLBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  TotPay = TotPay + HPay
  TotHrs = TotHrs + Emp2Rec.HOLBAL
  HBalTotal = Using(Image2$, HTotal)
  GTotHrs = GTotHrs + TotHrs
  GTotPay = OldRound(GTotPay + TotPay)
  
  RSet LTbl(1) = Str$(Emp2Rec.LeaveTbl)
  '                           0                   1             2             3
  Print #RHandle, QPTrim$(Unit(1).UFEMPR); dlm; Date$; dlm; ENumb(1); dlm; EName(1); dlm;
  '                 4             5             6              7            8             9
  Print #RHandle, VBal(1); dlm; SBal(1); dlm; CBal(1); dlm; PBal(1); dlm; HBal(1); dlm; LTbl(1); dlm;
  '                  10               11             12              13              14              15
  Print #RHandle, NumOfEmps; dlm; VBalTotal; dlm; SBalTotal; dlm; CBalTotal; dlm; PBalTotal; dlm; HBalTotal; dlm;
  '                   16              17
  Print #RHandle, DollarFlag; dlm; TotPay
  Return
  
ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."

DedExitRpt:
End Sub

Private Sub LoadThisForm()
   Dim EmpData1Handle As Integer, EmpIdxNNameHandle As Integer
   Dim EmpData1Rec As EmpData1Type
   Dim IdxRecPointer As Integer, NumOfRecs As Integer

   OpenEmpData1File EmpData1Handle
   OpenEmpIdxNNameFile EmpIdxNNameHandle
   NumOfRecs = LOF(EmpIdxNNameHandle) / 2
   If NumOfRecs = 0 Then
     MsgBox "No records on file"
     Close
     Exit Sub
   End If
   Get #EmpIdxNNameHandle, 1, IdxRecPointer
   Get #EmpData1Handle, IdxRecPointer, EmpData1Rec
   fptxtFirstEmpNo.Text = Val(EmpData1Rec.EmpNo)
   
   Get #EmpIdxNNameHandle, NumOfRecs, IdxRecPointer
   Get #EmpData1Handle, IdxRecPointer, EmpData1Rec
   fptxtSecEmpNo.Text = Val(EmpData1Rec.EmpNo)
  
   Close EmpIdxNNameHandle, EmpData1Handle
   fpcmbDollar.Text = "No"
   fpcmbDollar.AddItem "No"
   fpcmbDollar.AddItem "Yes. Show Totals Only."
   fpcmbDollar.AddItem "Yes. Split By Benefit."
   fpcomboPrintOpt.AddItem "Graphical"
   fpcomboPrintOpt.AddItem "Text"
   fpcomboPrintOpt.Text = "Graphical"
   DollarFlag = False
   SplitFlag = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmLeaveBenefit.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintText()
  Dim RecNo As Integer
  Dim MinHrs As Long, RptTitle$, x%
  Dim Image1$, Image2$, cnt As Integer
  Dim LineCnt As Integer, Emp1RecLen As Long
  Dim EmpRecSize As Long
  Dim IdxRecLen As Integer, FF$
  Dim IdxFileSize&, NumOfRecs As Long
  Dim EmpIdxNNameHandle As Integer, Page As Integer
  Dim EmpIdxLNameHandle As Integer
  Dim RHandle As Integer, DHandle As Integer, THandle As Integer
  Dim UnitHandle As Integer, EHandle1 As Integer
  Dim FirstEmp&, LastEmp&, EmpNo&, RptName$
  Dim UTemp$, MaxLines As Integer, CrLf$
  Dim Emp2Rec As EmpData2Type
  Dim VBalTotal As String * 13
  Dim VTotal As Double
  Dim SBalTotal As String * 13
  Dim STotal As Double
  Dim CBalTotal As String * 12
  Dim CTotal As Double
  Dim PBalTotal As String * 13
  Dim PTotal As Double
  Dim HBalTotal As String * 13
  Dim HTotal As Double
  Dim NumOfEmps As String * 22
  Dim NbrOfEmps As Integer
  Dim TotHrs As Double
  Dim GTotHrs As Double
  Dim TotPay As Double
  Dim GTotPay As Double
  Dim Dash As String * 103, DDash As String * 120
  Dim VPay As Double
  Dim SPay As Double
  Dim CPay As Double
  Dim PPay As Double
  Dim HPay As Double
  
  FF$ = Chr$(12)
  MinHrs = -10000
  Image1$ = "#,##0.00"
  Image2$ = "###,##0.00"
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  ReDim ENumb(1) As String * 11
  ReDim EName(1) As String * 19
  ReDim VBal(1) As String * 13
  ReDim SBal(1) As String * 13
  ReDim CBal(1) As String * 13
  ReDim PBal(1) As String * 12
  ReDim HBal(1) As String * 13
  ReDim LTbl(1) As String * 6
  
  MaxLines = 55
  LineCnt = 0
  Dash = String$(103, "-") ' + CrLf$
  DDash = String$(120, "-") '
  EmpRecSize = Len(Emp2Rec)
'--------------------------------------------------------
  RptName$ = "PRRPTS\BENEACCR.RPT"
'  OpenEmpIdxNNameFile EmpIdxNNameHandle
'  IdxRecLen = 2
'  NumOfRecs = LOF(EmpIdxNNameHandle) \ IdxRecLen
'  If NumOfRecs = 0 Then
'    MsgBox "No records on file."
'    Close
'    Exit Sub
'  End If
   If CheckNumber.Value = 1 Then
    OpenEmpIdxNNameFile EmpIdxNNameHandle
    IdxRecLen = 2
    NumOfRecs = LOF(EmpIdxNNameHandle) \ IdxRecLen
    If NumOfRecs = 0 Then
      MsgBox "No records on file."
      Close EmpIdxNNameHandle
      Exit Sub
    End If
  ElseIf CheckName.Value = 1 Then
    OpenEmpIdxLNameFile EmpIdxLNameHandle
    IdxRecLen = 2
    NumOfRecs = LOF(EmpIdxLNameHandle) \ IdxRecLen
    If NumOfRecs = 0 Then
      MsgBox "No records on file."
      Close EmpIdxLNameHandle
      Exit Sub
    End If
  End If
 
  If fptxtFirstEmpNo.Text = "" Then
    MsgBox "Please enter a First Employee Number"
    fptxtFirstEmpNo.SetFocus
    Close EmpIdxLNameHandle
    Close EmpIdxNNameHandle
    Exit Sub
  End If
  
  If fptxtSecEmpNo.Text = "" Then
    MsgBox "Please enter a Second Employee Number"
    fptxtSecEmpNo.SetFocus
    Close EmpIdxLNameHandle
    Close EmpIdxNNameHandle
    Exit Sub
  End If
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    If CheckNumber.Value = 1 Then
      Get EmpIdxNNameHandle, x, IdxBuff(x)
    ElseIf CheckName.Value = 1 Then
      Get EmpIdxLNameHandle, x, IdxBuff(x)
    End If
  Next x
  Close EmpIdxLNameHandle
  Close EmpIdxNNameHandle
'-----------------------------------------------------------
  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  RPTSetupPRN 8, RHandle
  OpenTransHistFile THandle
  OpenEmpData2File DHandle
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  RptTitle$ = "Benefits Accrual Report"
  
  ReDim Emp1Rec(1) As EmpData1Type
  Emp1RecLen = Len(Emp1Rec(1))
  OpenEmpData1File EHandle1
  Get EHandle1, IdxBuff(1), Emp1RecLen
  FirstEmp& = Val(Emp1Rec(1).EmpNo)
  Get EHandle1, IdxBuff(NumOfRecs), Emp1RecLen
  LastEmp& = Val(Emp1Rec(1).EmpNo)
  Close EHandle1
  
  FirstEmp& = Val(fptxtFirstEmpNo.Text)
  LastEmp& = Val(fptxtSecEmpNo.Text)
  If LastEmp& < FirstEmp& Then
    MsgBox "ERROR: The Last Employee Number is less than the First Employee Number"
    fptxtSecEmpNo.SetFocus
    Close
    Exit Sub
  End If
  
  FrmShowPctComp.Label1 = "Payroll Deduction Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  OpenEmpData2File DHandle
 
  GoSub PrintBenefitHeader
  
  For RecNo = 1 To NumOfRecs
    Get DHandle, IdxBuff(RecNo), Emp2Rec
    EmpNo& = Val(QPTrim$(Emp2Rec.EmpNo))
    If (EmpNo& < FirstEmp& Or EmpNo& > LastEmp&) Or Emp2Rec.Deleted Or Emp2Rec.EMPTDATE <> 0 Then
      GoTo SkipEmBene
    End If
    GoSub PrintEmpBalance
    If (LineCnt > MaxLines) And RecNo < NumOfRecs Then          'bottom of page?
      Print #RHandle, FF$         'yes; form feed
      GoSub PrintBenefitHeader  'write title lines
    End If
SkipEmBene:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload FrmShowPctComp
      GoTo DedExitRpt
    End If
  Next
  
  If DollarFlag = True Then
    Print #RHandle, DDash
  Else
    Print #RHandle, Dash
  End If
  
  Print #RHandle,
  If DollarFlag = True Then
    Print #RHandle, "Totals"; Tab(12); NumOfEmps; Tab(34); VBalTotal; Tab(47); SBalTotal; Tab(60); CBalTotal; Tab(73); HBalTotal; Tab(86); PBalTotal; Tab(107); Using("$##,###,##0.00", GTotPay)
  Else
    Print #RHandle, "Totals"; Tab(12); NumOfEmps; Tab(34); VBalTotal; Tab(47); SBalTotal; Tab(60); CBalTotal; Tab(73); HBalTotal; Tab(86); PBalTotal
  End If
  Print #RHandle, FF$             'yes; form feed
  RPTSetupPRN 123, RHandle '7/24
  Close DHandle
  Close RHandle
  
  ViewPrint RptName$, RptTitle$, True
  EnableCloseButton Me.hwnd, True
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  MainLog ("Leave Benefit Report processed.")
  Exit Sub
  
PrintBenefitHeader:
  Page = Page + 1
  RSet Pg(1) = Str$(Page)
  UTemp$ = Space$(71)
  LSet UTemp$ = QPTrim$(Unit(1).UFEMPR)
  Mid$(UTemp$, 62) = "Page:" + Pg(1) + CrLf$
  Print #RHandle, UTemp$
  Print #RHandle, "Benefits Accrual Report" + CrLf$
  Print #RHandle, "Report Date: " + Date$ + CrLf$
  Print #RHandle, "                                   Vacation   Sick Leave    Comp Time      Holiday     Personal   Leave" ' + CrLf$
  If DollarFlag = True Then
    Print #RHandle, "Emp #      Name                     Balance      Balance      Balance      Balance      Balance   Table            Value" '+ CrLf$
    Print #RHandle, DDash
  Else
    Print #RHandle, "Emp #      Name                     Balance      Balance      Balance      Balance      Balance   Table" ' + CrLf$
    Print #RHandle, Dash
  End If
  LineCnt = 6
  Return
  
PrintEmpBalance:
'  If Emp2Rec.EmpLName = "Clarke" Then Stop
  If Emp2Rec.EMPCTBAL < MinHrs Then
    Emp2Rec.EMPCTBAL = 0
  End If
  If Emp2Rec.EMPVBAL < MinHrs Then
    Emp2Rec.EMPVBAL = 0
  End If
  If Emp2Rec.EMPSLBAL < MinHrs Then
    Emp2Rec.EMPSLBAL = 0
  End If
  If Emp2Rec.PERBAL < MinHrs Then
    Emp2Rec.PERBAL = 0
  End If
  If Emp2Rec.HOLBAL < MinHrs Then
    Emp2Rec.HOLBAL = 0
  End If
  
  LSet ENumb(1) = QPTrim$(Emp2Rec.EmpNo)
  NbrOfEmps = NbrOfEmps + 1
  NumOfEmps = Using("#####", NbrOfEmps)
'  If QPTrim$(Emp2Rec.EMPLNAME) = "LAYTON" Then Stop
  LSet EName(1) = QPTrim$(Emp2Rec.EmpLName) + ", " + QPTrim$(Emp2Rec.EmpFName)
  TotHrs = 0
  TotPay = 0
  RSet VBal(1) = Using(Image2$, Emp2Rec.EMPVBAL)
  VTotal = VTotal + Emp2Rec.EMPVBAL
  VPay = FigurePayAmts(Emp2Rec.EMPVBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  TotPay = TotPay + VPay
  TotHrs = TotHrs + Emp2Rec.EMPVBAL
  VBalTotal = Using(Image2$, VTotal)
  
  RSet SBal(1) = Using(Image2$, Emp2Rec.EMPSLBAL)
  STotal = STotal + Emp2Rec.EMPSLBAL
  SPay = FigurePayAmts(Emp2Rec.EMPSLBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  TotPay = TotPay + SPay
  TotHrs = TotHrs + Emp2Rec.EMPSLBAL
  SBalTotal = Using(Image2$, STotal)
  
  RSet CBal(1) = Using(Image2$, Emp2Rec.EMPCTBAL)
  CTotal = CTotal + Emp2Rec.EMPCTBAL
  CPay = FigurePayAmts(Emp2Rec.EMPCTBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  TotPay = TotPay + CPay
  TotHrs = TotHrs + Emp2Rec.EMPCTBAL
  CBalTotal = Using(Image2$, CTotal)
  
  RSet PBal(1) = Using(Image2$, Emp2Rec.PERBAL)
  PTotal = PTotal + Emp2Rec.PERBAL
  PPay = FigurePayAmts(Emp2Rec.PERBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  TotPay = TotPay + PPay
  TotHrs = TotHrs + Emp2Rec.PERBAL
  PBalTotal = Using(Image2$, PTotal)
  
  RSet HBal(1) = Using(Image2$, Emp2Rec.HOLBAL)
  HTotal = HTotal + Emp2Rec.HOLBAL
  HPay = FigurePayAmts(Emp2Rec.HOLBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  TotPay = TotPay + HPay
  TotHrs = TotHrs + Emp2Rec.HOLBAL
  HBalTotal = Using(Image2$, HTotal)
  GTotHrs = GTotHrs + TotHrs
  GTotPay = OldRound(GTotPay + TotPay)
  
  RSet LTbl(1) = Str$(Emp2Rec.LeaveTbl)
  If DollarFlag = True Then
    Print #RHandle, ENumb(1) + EName(1); Tab(31); VBal(1); Tab(44); SBal(1); Tab(57); CBal(1); Tab(70); HBal(1); Tab(84); PBal(1); Tab(98); LTbl(1); Tab(110); Using("$###,##0.00", TotPay) ' + CrLf$
  Else
    Print #RHandle, ENumb(1) + EName(1); Tab(31); VBal(1); Tab(44); SBal(1); Tab(57); CBal(1); Tab(70); HBal(1); Tab(84); PBal(1); Tab(98); LTbl(1) ' + CrLf$
  End If
  LineCnt = LineCnt + 1
  Return

DedExitRpt:

End Sub

Private Sub fpcmbDollar_Change()
  If fpcmbDollar.Text = "No" Then
    DollarFlag = False
    SplitFlag = False
  Else
    DollarFlag = True
    If fpcmbDollar.Text = "Yes. Show Totals Only." Then
      SplitFlag = False
    ElseIf fpcmbDollar.Text = "Yes. Split By Benefit." Then
      SplitFlag = True
    End If
  End If
End Sub

Private Sub fpcmbDollar_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbDollar.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDollar.ListIndex = -1
  End If
  If fpcmbDollar.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcomboPrintOpt.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

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

Private Function FigurePayAmts(Hours As Double, ThisRate As Double, ThisType$, ThisFreq$) As Double
  FigurePayAmts = 0
  If ThisType = "S" Then
    Select Case UCase(ThisFreq$)
      Case "WEEKLY"
        FigurePayAmts = OldRound((ThisRate / 40) * Hours)
      Case "BI-WEEKLY"
        FigurePayAmts = OldRound((ThisRate / 80) * Hours)
      Case "SEMI-MONTHLY"
        FigurePayAmts = OldRound((ThisRate / 86.67) * Hours)
      Case "MONTHLY"
        FigurePayAmts = OldRound((ThisRate / 173.33) * Hours)
      Case "QUARTERLY"
        FigurePayAmts = OldRound((ThisRate / 520) * Hours)
      Case "SEMI-ANNUALLY"
        FigurePayAmts = OldRound((ThisRate / 1040) * Hours)
      Case "ANNUALLY"
        FigurePayAmts = OldRound((ThisRate / 2080) * Hours)
    End Select
  ElseIf ThisType = "H" Then
    FigurePayAmts = OldRound(ThisRate * Hours)
  End If
End Function

Private Sub PrintSplitGraphics()
  Dim RecNo As Integer
  Dim MinHrs As Long, RptTitle$, x%
  Dim Image1$, Image2$, cnt As Integer
  Dim LineCnt As Integer, Emp1RecLen As Long
  Dim EmpRecSize As Long
  Dim IdxRecLen As Integer, FF$
  Dim IdxFileSize&, NumOfRecs As Long
  Dim EmpIdxNNameHandle As Integer, Page As Integer
  Dim EmpIdxLNameHandle As Integer
  Dim RHandle As Integer, DHandle As Integer, THandle As Integer
  Dim UnitHandle As Integer, EHandle1 As Integer
  Dim FirstEmp&, LastEmp&, EmpNo&, RptName$
  Dim UTemp$, MaxLines As Integer, CrLf$
  Dim Emp2Rec As EmpData2Type
  Dim VBalTotal As String * 13
  Dim VTotal As Double
  Dim SBalTotal As String * 13
  Dim STotal As Double
  Dim CBalTotal As String * 12
  Dim CTotal As Double
  Dim PBalTotal As String * 13
  Dim PTotal As Double
  Dim HBalTotal As String * 13
  Dim HTotal As Double
  Dim NumOfEmps As String * 22
  Dim NbrOfEmps As Integer
  Dim dlm$
  Dim TotHrs As Double
  Dim GTotHrs As Double
  Dim TotPay As Double
  Dim GTotPay As Double
  Dim VPay As Double
  Dim SPay As Double
  Dim CPay As Double
  Dim PPay As Double
  Dim HPay As Double
  
  On Error GoTo ErrorHandler
  dlm$ = "~"
  MinHrs = -10000
  Image1$ = "#,##0.00"
  Image2$ = "###,##0.00"
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  ReDim ENumb(1) As String * 11
  ReDim EName(1) As String * 19
  ReDim VBal(1) As String * 13
  ReDim SBal(1) As String * 13
  ReDim CBal(1) As String * 13
  ReDim PBal(1) As String * 12
  ReDim HBal(1) As String * 13
  ReDim LTbl(1) As String * 6
  
  EmpRecSize = Len(Emp2Rec)
'--------------------------------------------------------
  RptName$ = "PRRPTS\BENSPLIT.RPT"
'  OpenEmpIdxNNameFile EmpIdxNNameHandle
'  IdxRecLen = 2
'  NumOfRecs = LOF(EmpIdxNNameHandle) \ IdxRecLen
'  If NumOfRecs = 0 Then
'    MsgBox "No records on file."
'    Close
'    Exit Sub
'  End If
  If CheckNumber.Value = 1 Then
    OpenEmpIdxNNameFile EmpIdxNNameHandle
    IdxRecLen = 2
    NumOfRecs = LOF(EmpIdxNNameHandle) \ IdxRecLen
    If NumOfRecs = 0 Then
      MsgBox "No records on file."
      Close EmpIdxNNameHandle
      Exit Sub
    End If
  ElseIf CheckName.Value = 1 Then
    OpenEmpIdxLNameFile EmpIdxLNameHandle
    IdxRecLen = 2
    NumOfRecs = LOF(EmpIdxLNameHandle) \ IdxRecLen
    If NumOfRecs = 0 Then
      MsgBox "No records on file."
      Close EmpIdxLNameHandle
      Exit Sub
    End If
  End If
  
  If fptxtFirstEmpNo.Text = "" Then
    MsgBox "Please enter a First Employee Number"
    fptxtFirstEmpNo.SetFocus
    Close EmpIdxLNameHandle
    Close EmpIdxNNameHandle
    Exit Sub
  End If
  
  If fptxtSecEmpNo.Text = "" Then
    MsgBox "Please enter a Second Employee Number"
    fptxtSecEmpNo.SetFocus
    Close EmpIdxLNameHandle
    Close EmpIdxNNameHandle
    Exit Sub
  End If
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    If CheckNumber.Value = 1 Then
      Get EmpIdxNNameHandle, x, IdxBuff(x)
    ElseIf CheckName.Value = 1 Then
      Get EmpIdxLNameHandle, x, IdxBuff(x)
    End If
  Next x
  Close EmpIdxLNameHandle
  Close EmpIdxNNameHandle
'-----------------------------------------------------------
  RHandle = FreeFile
  On Error GoTo ErrorHandler
  Open RptName$ For Output As RHandle
  OpenTransHistFile THandle
  OpenEmpData2File DHandle
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  ReDim Emp1Rec(1) As EmpData1Type
  Emp1RecLen = Len(Emp1Rec(1))
  OpenEmpData1File EHandle1
  Get EHandle1, IdxBuff(1), Emp1RecLen
  FirstEmp& = Val(Emp1Rec(1).EmpNo)
  Get EHandle1, IdxBuff(NumOfRecs), Emp1RecLen
  LastEmp& = Val(Emp1Rec(1).EmpNo)
  Close EHandle1
  
  FirstEmp& = Val(fptxtFirstEmpNo.Text)
  LastEmp& = Val(fptxtSecEmpNo.Text)
  If LastEmp& < FirstEmp& Then
    MsgBox "ERROR: The Last Employee Number is less than the First Employee Number"
    fptxtSecEmpNo.SetFocus
    Close
    Exit Sub
  End If
  
  FrmShowPctComp.Label1 = "Payroll Deduction Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  OpenEmpData2File DHandle
 
  For RecNo = 1 To NumOfRecs
    Get DHandle, IdxBuff(RecNo), Emp2Rec
    EmpNo& = Val(QPTrim$(Emp2Rec.EmpNo))
    If (EmpNo& < FirstEmp& Or EmpNo& > LastEmp&) Or Emp2Rec.Deleted Or Emp2Rec.EMPTDATE <> 0 Then
      GoTo SkipEmBene
    End If
    GoSub PrintEmpBalance
SkipEmBene:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload FrmShowPctComp
      GoTo DedExitRpt
    End If
  Next
  Close DHandle
  Close RHandle
  
  arLvBftSplitRpt.Show
  frmLoadingRpt.Show
  EnableCloseButton Me.hwnd, True
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  MainLog ("Leave Benefit Report processed in benefit split format.")
  
  Exit Sub
  
PrintEmpBalance:
  If Emp2Rec.EMPCTBAL < MinHrs Then
    Emp2Rec.EMPCTBAL = 0
  End If
  If Emp2Rec.EMPVBAL < MinHrs Then
    Emp2Rec.EMPVBAL = 0
  End If
  If Emp2Rec.EMPSLBAL < MinHrs Then
    Emp2Rec.EMPSLBAL = 0
  End If
  If Emp2Rec.PERBAL < MinHrs Then
    Emp2Rec.PERBAL = 0
  End If
  If Emp2Rec.HOLBAL < MinHrs Then
    Emp2Rec.HOLBAL = 0
  End If
  
  LSet ENumb(1) = QPTrim$(Emp2Rec.EmpNo)
  NbrOfEmps = NbrOfEmps + 1
  NumOfEmps = Using("#####", NbrOfEmps)
  TotHrs = 0
  TotPay = 0
  LSet EName(1) = QPTrim$(Emp2Rec.EmpLName) + ", " + QPTrim$(Emp2Rec.EmpFName)
  RSet VBal(1) = Using(Image2$, Emp2Rec.EMPVBAL)
  VTotal = VTotal + Emp2Rec.EMPVBAL
  VPay = FigurePayAmts(Emp2Rec.EMPVBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  TotPay = TotPay + VPay
  TotHrs = TotHrs + Emp2Rec.EMPVBAL
  VBalTotal = Using(Image2$, VTotal)
  RSet SBal(1) = Using(Image2$, Emp2Rec.EMPSLBAL)
  STotal = STotal + Emp2Rec.EMPSLBAL
  SPay = FigurePayAmts(Emp2Rec.EMPSLBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  TotPay = TotPay + SPay
  TotHrs = TotHrs + Emp2Rec.EMPSLBAL
  SBalTotal = Using(Image2$, STotal)
  RSet CBal(1) = Using(Image2$, Emp2Rec.EMPCTBAL)
  CTotal = CTotal + Emp2Rec.EMPCTBAL
  CPay = FigurePayAmts(Emp2Rec.EMPCTBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  TotPay = TotPay + CPay
  TotHrs = TotHrs + Emp2Rec.EMPCTBAL
  CBalTotal = Using(Image2$, CTotal)
  RSet PBal(1) = Using(Image2$, Emp2Rec.PERBAL)
  PTotal = PTotal + Emp2Rec.PERBAL
  PPay = FigurePayAmts(Emp2Rec.PERBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  TotPay = TotPay + PPay
  TotHrs = TotHrs + Emp2Rec.PERBAL
  PBalTotal = Using(Image2$, PTotal)
  RSet HBal(1) = Using(Image2$, Emp2Rec.HOLBAL)
  HTotal = HTotal + Emp2Rec.HOLBAL
  HPay = FigurePayAmts(Emp2Rec.HOLBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  TotPay = TotPay + HPay
  TotHrs = TotHrs + Emp2Rec.HOLBAL
  HBalTotal = Using(Image2$, HTotal)
  GTotHrs = GTotHrs + TotHrs
  
  GTotPay = OldRound(GTotPay + TotPay)
  
  RSet LTbl(1) = Str$(Emp2Rec.LeaveTbl)
  '                           0                   1             2             3
  Print #RHandle, QPTrim$(Unit(1).UFEMPR); dlm; Date$; dlm; ENumb(1); dlm; EName(1); dlm;
  '                 4             5             6              7            8             9
  Print #RHandle, VBal(1); dlm; SBal(1); dlm; CBal(1); dlm; PBal(1); dlm; HBal(1); dlm; LTbl(1); dlm;
  '                  10               11             12              13              14              15
  Print #RHandle, NumOfEmps; dlm; VBalTotal; dlm; SBalTotal; dlm; CBalTotal; dlm; PBalTotal; dlm; HBalTotal; dlm;
  '                   16              17         18         19         20         21         22
  Print #RHandle, DollarFlag; dlm; TotPay; dlm; VPay; dlm; SPay; dlm; CPay; dlm; PPay; dlm; HPay
  
  Return
  
ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."

DedExitRpt:

  End Sub

Private Sub PrintSplitText()
  Dim RecNo As Integer
  Dim MinHrs As Long, RptTitle$, x%
  Dim Image1$, Image2$, cnt As Integer
  Dim LineCnt As Integer, Emp1RecLen As Long
  Dim EmpRecSize As Long
  Dim IdxRecLen As Integer, FF$
  Dim IdxFileSize&, NumOfRecs As Long
  Dim EmpIdxNNameHandle As Integer, Page As Integer
  Dim EmpIdxLNameHandle As Integer
  Dim RHandle As Integer, DHandle As Integer, THandle As Integer
  Dim UnitHandle As Integer, EHandle1 As Integer
  Dim FirstEmp&, LastEmp&, EmpNo&, RptName$
  Dim UTemp$, MaxLines As Integer, CrLf$
  Dim Emp2Rec As EmpData2Type
  Dim VBalTotal As String * 13
  Dim VTotal As Double
  Dim SBalTotal As String * 13
  Dim STotal As Double
  Dim CBalTotal As String * 12
  Dim CTotal As Double
  Dim PBalTotal As String * 13
  Dim PTotal As Double
  Dim HBalTotal As String * 13
  Dim HTotal As Double
  Dim NumOfEmps As String * 22
  Dim NbrOfEmps As Integer
  Dim TotHrs As Double
  Dim GTotHrs As Double
  Dim TotPay As Double
  Dim GTotPay As Double
  Dim Dash As String * 103, DDash As String * 120
  Dim VPay As Double, GVPay As Double
  Dim SPay As Double, GSPay As Double
  Dim CPay As Double, GCPay As Double
  Dim PPay As Double, GPPay As Double
  Dim HPay As Double, GHPay As Double
  
  FF$ = Chr$(12)
  MinHrs = -10000
  GVPay = 0
  GSPay = 0
  GCPay = 0
  GPPay = 0
  GHPay = 0
  Image1$ = "#,##0.00"
  Image2$ = "###,##0.00"
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  ReDim ENumb(1) As String * 11
  ReDim EName(1) As String * 19
  ReDim VBal(1) As String * 13
  ReDim SBal(1) As String * 13
  ReDim CBal(1) As String * 13
  ReDim PBal(1) As String * 12
  ReDim HBal(1) As String * 13
  ReDim LTbl(1) As String * 6
  
  MaxLines = 55
  LineCnt = 0
  Dash = String$(103, "-") ' + CrLf$
  DDash = String$(120, "-") '
  EmpRecSize = Len(Emp2Rec)
'--------------------------------------------------------
  RptName$ = "PRRPTS\BENEACCR.RPT"
'  OpenEmpIdxNNameFile EmpIdxNNameHandle
'  IdxRecLen = 2
'  NumOfRecs = LOF(EmpIdxNNameHandle) \ IdxRecLen
'  If NumOfRecs = 0 Then
'    MsgBox "No records on file."
'    Close
'    Exit Sub
'  End If
  
  If CheckNumber.Value = 1 Then
    OpenEmpIdxNNameFile EmpIdxNNameHandle
    IdxRecLen = 2
    NumOfRecs = LOF(EmpIdxNNameHandle) \ IdxRecLen
    If NumOfRecs = 0 Then
      MsgBox "No records on file."
      Close EmpIdxNNameHandle
      Exit Sub
    End If
  ElseIf CheckName.Value = 1 Then
    OpenEmpIdxLNameFile EmpIdxLNameHandle
    IdxRecLen = 2
    NumOfRecs = LOF(EmpIdxLNameHandle) \ IdxRecLen
    If NumOfRecs = 0 Then
      MsgBox "No records on file."
      Close EmpIdxLNameHandle
      Exit Sub
    End If
  End If
  
  If fptxtFirstEmpNo.Text = "" Then
    MsgBox "Please enter a First Employee Number"
    fptxtFirstEmpNo.SetFocus
    Close EmpIdxLNameHandle
    Close EmpIdxNNameHandle
    Exit Sub
  End If
  
  If fptxtSecEmpNo.Text = "" Then
    MsgBox "Please enter a Second Employee Number"
    fptxtSecEmpNo.SetFocus
    Close EmpIdxLNameHandle
    Close EmpIdxNNameHandle
    Exit Sub
  End If
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
'  For x = 1 To NumOfRecs
'    Get EmpIdxNNameHandle, x, IdxBuff(x)
'  Next x
'  Close EmpIdxNNameHandle
  For x = 1 To NumOfRecs
    If CheckNumber.Value = 1 Then
      Get EmpIdxNNameHandle, x, IdxBuff(x)
    ElseIf CheckName.Value = 1 Then
      Get EmpIdxLNameHandle, x, IdxBuff(x)
    End If
  Next x
  Close EmpIdxLNameHandle
  Close EmpIdxNNameHandle
'-----------------------------------------------------------
  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  RPTSetupPRN 8, RHandle
  OpenTransHistFile THandle
  OpenEmpData2File DHandle
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  RptTitle$ = "Benefits Accrual Report"
  
  ReDim Emp1Rec(1) As EmpData1Type
  Emp1RecLen = Len(Emp1Rec(1))
  OpenEmpData1File EHandle1
  Get EHandle1, IdxBuff(1), Emp1RecLen
  FirstEmp& = Val(Emp1Rec(1).EmpNo)
  Get EHandle1, IdxBuff(NumOfRecs), Emp1RecLen
  LastEmp& = Val(Emp1Rec(1).EmpNo)
  Close EHandle1
  
  FirstEmp& = Val(fptxtFirstEmpNo.Text)
  LastEmp& = Val(fptxtSecEmpNo.Text)
  If LastEmp& < FirstEmp& Then
    MsgBox "ERROR: The Last Employee Number is less than the First Employee Number"
    fptxtSecEmpNo.SetFocus
    Close
    Exit Sub
  End If
  
  FrmShowPctComp.Label1 = "Payroll Deduction Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  OpenEmpData2File DHandle
 
  GoSub PrintBenefitHeader
  
  For RecNo = 1 To NumOfRecs
    Get DHandle, IdxBuff(RecNo), Emp2Rec
    EmpNo& = Val(QPTrim$(Emp2Rec.EmpNo))
    If (EmpNo& < FirstEmp& Or EmpNo& > LastEmp&) Or Emp2Rec.Deleted Or Emp2Rec.EMPTDATE <> 0 Then
      GoTo SkipEmBene
    End If
    GoSub PrintEmpBalance
    If (LineCnt > MaxLines) And RecNo < NumOfRecs Then          'bottom of page?
      Print #RHandle, FF$         'yes; form feed
      GoSub PrintBenefitHeader  'write title lines
    End If
SkipEmBene:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload FrmShowPctComp
      GoTo DedExitRpt
    End If
  Next
  
  If LineCnt >= 49 Then
    Print #RHandle, FF$
    GoSub PrintBenefitHeader
  End If
  
  Print #RHandle, DDash
  
  Print #RHandle,
  Print #RHandle, "# Employees: " + CStr(NumOfEmps)
  Print #RHandle, Tab(5); "Total Benefit Hours: "; Tab(34); VBalTotal; Tab(47); SBalTotal; Tab(60); CBalTotal; Tab(73); HBalTotal; Tab(86); PBalTotal
  Print #RHandle, Tab(5); "Total Benefit Values: "; Tab(31); Using$("$#,###,##0.00", GVPay); Tab(44); Using$("$#,###,##0.00", GSPay); Tab(59); Using$("$###,##0.00", GCPay); Tab(72); Using$("$###,##0.00", GHPay); Tab(85); Using$("$###,##0.00", GPPay); Tab(107); Using("$##,###,##0.00", GTotPay)
  
  Print #RHandle, FF$             'yes; form feed
  RPTSetupPRN 123, RHandle '7/24
  Close DHandle
  Close RHandle
  
  ViewPrint RptName$, RptTitle$, True
  EnableCloseButton Me.hwnd, True
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  MainLog ("Leave Benefit Report processed.")
  Exit Sub
  
PrintBenefitHeader:
  Page = Page + 1
  RSet Pg(1) = Str$(Page)
  UTemp$ = Space$(71)
  LSet UTemp$ = QPTrim$(Unit(1).UFEMPR)
  Mid$(UTemp$, 62) = "Page:" + Pg(1) + CrLf$
  Print #RHandle, UTemp$
  Print #RHandle, "Benefits Accrual Report" + CrLf$
  Print #RHandle, "Report Date: " + Date$ + CrLf$
  Print #RHandle, "                                   Vacation   Sick Leave    Comp Time      Holiday     Personal   Leave" ' + CrLf$
  Print #RHandle, "Emp #      Name                     Balance      Balance      Balance      Balance      Balance   Table            Value" '+ CrLf$
  Print #RHandle, DDash
  LineCnt = 6
  Return
  
PrintEmpBalance:
'  If Emp2Rec.EmpLName = "Clarke" Then Stop
  If Emp2Rec.EMPCTBAL < MinHrs Then
    Emp2Rec.EMPCTBAL = 0
  End If
  If Emp2Rec.EMPVBAL < MinHrs Then
    Emp2Rec.EMPVBAL = 0
  End If
  If Emp2Rec.EMPSLBAL < MinHrs Then
    Emp2Rec.EMPSLBAL = 0
  End If
  If Emp2Rec.PERBAL < MinHrs Then
    Emp2Rec.PERBAL = 0
  End If
  If Emp2Rec.HOLBAL < MinHrs Then
    Emp2Rec.HOLBAL = 0
  End If
  
  LSet ENumb(1) = QPTrim$(Emp2Rec.EmpNo)
  NbrOfEmps = NbrOfEmps + 1
  NumOfEmps = Using("#####", NbrOfEmps)
  
  LSet EName(1) = QPTrim$(Emp2Rec.EmpLName) + ", " + QPTrim$(Emp2Rec.EmpFName)
  TotHrs = 0
  TotPay = 0
  RSet VBal(1) = Using(Image2$, Emp2Rec.EMPVBAL)
  VTotal = VTotal + Emp2Rec.EMPVBAL
  VPay = FigurePayAmts(Emp2Rec.EMPVBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  GVPay = GVPay + VPay
  TotPay = TotPay + VPay
  TotHrs = TotHrs + Emp2Rec.EMPVBAL
  VBalTotal = Using(Image2$, VTotal)
  
  RSet SBal(1) = Using(Image2$, Emp2Rec.EMPSLBAL)
  STotal = STotal + Emp2Rec.EMPSLBAL
  SPay = FigurePayAmts(Emp2Rec.EMPSLBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  GSPay = GSPay + SPay
  TotPay = TotPay + SPay
  TotHrs = TotHrs + Emp2Rec.EMPSLBAL
  SBalTotal = Using(Image2$, STotal)
  
  RSet CBal(1) = Using(Image2$, Emp2Rec.EMPCTBAL)
  CTotal = CTotal + Emp2Rec.EMPCTBAL
  CPay = FigurePayAmts(Emp2Rec.EMPCTBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  GCPay = GCPay + CPay
  TotPay = TotPay + CPay
  TotHrs = TotHrs + Emp2Rec.EMPCTBAL
  CBalTotal = Using(Image2$, CTotal)
  
  RSet PBal(1) = Using(Image2$, Emp2Rec.PERBAL)
  PTotal = PTotal + Emp2Rec.PERBAL
  PPay = FigurePayAmts(Emp2Rec.PERBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  GPPay = GPPay + PPay
  TotPay = TotPay + PPay
  TotHrs = TotHrs + Emp2Rec.PERBAL
  PBalTotal = Using(Image2$, PTotal)
  
  RSet HBal(1) = Using(Image2$, Emp2Rec.HOLBAL)
  HTotal = HTotal + Emp2Rec.HOLBAL
  HPay = FigurePayAmts(Emp2Rec.HOLBAL, Emp2Rec.EMPPRATE, Mid(Emp2Rec.EMPPTYPE, 1, 1), QPTrim$(UCase(Emp2Rec.EMPPFREQ)))
  GHPay = GHPay + HPay
  TotPay = TotPay + HPay
  TotHrs = TotHrs + Emp2Rec.HOLBAL
  HBalTotal = Using(Image2$, HTotal)
  GTotHrs = GTotHrs + TotHrs
  GTotPay = OldRound(GTotPay + TotPay)
  
  RSet LTbl(1) = Str$(Emp2Rec.LeaveTbl)
  Print #RHandle, ENumb(1) + EName(1); Tab(31); VBal(1); Tab(44); SBal(1); Tab(57); CBal(1); Tab(70); HBal(1); Tab(84); PBal(1); Tab(98); LTbl(1)
  Print #RHandle, Tab(34); Using$("$##,##0.00", VPay); Tab(47); Using$("$##,##0.00", SPay); Tab(60); Using$("$##,##0.00", CPay); Tab(73); Using$("$##,##0.00", HPay); Tab(86); Using$("$##,##0.00", PPay); Tab(107); Using("$##,###,##0.00", TotPay)
  Print #RHandle, DDash
  Print #RHandle,
  LineCnt = LineCnt + 4
  
  Return

DedExitRpt:

End Sub

Private Sub CheckName_Click()
  If CheckName.Value = 1 Then CheckNumber.Value = 0
End Sub

Private Sub CheckNumber_Click()
  If CheckNumber.Value = 1 Then CheckName.Value = 0
End Sub

