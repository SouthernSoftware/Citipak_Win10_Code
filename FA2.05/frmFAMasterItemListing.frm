VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFAMasterItemListing 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Item Listing"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAMasterItemListing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6204
      Left            =   1932
      TabIndex        =   4
      Top             =   1344
      Width           =   7740
      _Version        =   196609
      _ExtentX        =   13652
      _ExtentY        =   10943
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmFAMasterItemListing.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3510
         TabIndex        =   3
         ToolTipText     =   "Select Graphical for a more robust but slower processing report. Select Text for a quick report."
         Top             =   3990
         Width           =   2355
         _Version        =   196608
         _ExtentX        =   4154
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
         AutoSearch      =   2
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
         AutoSearchFill  =   -1  'True
         AutoSearchFillDelay=   200
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmFAMasterItemListing.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbYN 
         Height          =   405
         Left            =   5475
         TabIndex        =   1
         ToolTipText     =   "Enter Y to include disposed of fixed assets or N to exclude disposed of fixed assets."
         Top             =   2490
         Width           =   780
         _Version        =   196608
         _ExtentX        =   1376
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
         Object.TabStop         =   0   'False
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
         MaxEditLen      =   5
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
         AutoSearchFill  =   -1  'True
         AutoSearchFillDelay=   200
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmFAMasterItemListing.frx":0BDD
      End
      Begin LpLib.fpCombo fpcmbOrder 
         Height          =   405
         Left            =   3210
         TabIndex        =   0
         Top             =   1875
         Width           =   3240
         _Version        =   196608
         _ExtentX        =   5715
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
         Object.TabStop         =   0   'False
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
         AutoSearch      =   2
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
         MaxEditLen      =   5
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
         AutoSearchFill  =   -1  'True
         AutoSearchFillDelay=   200
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmFAMasterItemListing.frx":0ED4
      End
      Begin EditLib.fpText fptxtDeptNum 
         Height          =   396
         Left            =   3072
         TabIndex        =   2
         ToolTipText     =   "If DEPARTMENT NUMBER is selected in the Report Order field then enter the desired department on which to report."
         Top             =   3168
         Width           =   1068
         _Version        =   196608
         _ExtentX        =   1884
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
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
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 - A L a l"
         MaxLength       =   14
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
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
      Begin fpBtnAtlLibCtl.fpBtn cmdDept 
         Height          =   390
         Left            =   4416
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to bring up a list of all current department numbers."
         Top             =   3168
         Width           =   1365
         _Version        =   131072
         _ExtentX        =   2408
         _ExtentY        =   688
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         GrayAreaColor   =   13684944
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
         ButtonDesigner  =   "frmFAMasterItemListing.frx":11CB
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   690
         Left            =   1584
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to create the desired report."
         Top             =   4848
         Width           =   1875
         _Version        =   131072
         _ExtentX        =   3307
         _ExtentY        =   1217
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         GrayAreaColor   =   13684944
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
         ButtonDesigner  =   "frmFAMasterItemListing.frx":13AB
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4416
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to create the desired report."
         Top             =   4848
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1217
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         GrayAreaColor   =   13684944
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
         ButtonDesigner  =   "frmFAMasterItemListing.frx":1587
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
         Height          =   348
         Left            =   1824
         TabIndex        =   9
         Top             =   4068
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Master Asset Listing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   2070
         TabIndex        =   8
         Top             =   720
         Width           =   3735
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1488
         Top             =   576
         Width           =   4908
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dept #"
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
         Left            =   1968
         TabIndex        =   7
         Top             =   3264
         Width           =   924
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Include Disposed Of Items (Y/N):"
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
         Left            =   1632
         TabIndex        =   6
         Top             =   2592
         Width           =   3660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Report Order:"
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
         Left            =   1440
         TabIndex        =   5
         Top             =   1968
         Width           =   1548
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6396
      Left            =   1836
      Top             =   1236
      Width           =   7932
   End
End
Attribute VB_Name = "frmFAMasterItemListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim DsplYNFlag As Boolean
  
Private Sub cmdDept_Click()
  frmFADeptList.Show vbModal
End Sub

Private Sub cmdExit_Click()
  KillFile "masteritemlistopen.dat"
  If Exist("itemmaintmenu.dat") Then
    frmFAItemMaintMenu.Show
  Else
    frmFAReportMenu.Show
  End If
  Close
  DoEvents
  Unload frmFAMasterItemListing
End Sub

Private Sub PrintText()
  Dim DHandle As Integer
  Dim DeptRec As FADeptCodeType
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxRecNums As Integer
  Dim x As Integer
  Dim NumOfDepts As Integer
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim ItemCnt As Integer
  Dim Dash80$, FF$
  Dim RptHandle As Integer
  Dim ReportFile$
  Dim cnt As Integer
  Dim Dispose$, Dept$
  Dim DItemCnt#
  Dim DOrigCost#
  Dim DDepTot#
  Dim DBookTot#
  Dim OrigCost#, PValue#
  Dim DepTot#
  Dim BookTot#
  Dim Page As Integer
  Dim TagFlag As Boolean
  Dim TagIdxHandle As Integer
  Dim TagRec As TagNumbSortIdxType
  Dim Nextx As Integer
  Dim ThisDept$, ThisDesc$
  Dim MatchThis As String
  Dim LifeLeft As String * 3
  Dim WholeLife As String * 3
  Dim LifeData As String * 7
  Dim HeaderFlag As Boolean
  Dim Lines80$
  Dim ItemTotal As Long
  
  On Error GoTo ERRORSTUFF
  
  If QPTrim$(fpcmbOrder.Text) = "TAG NUMBER" Then
    HeaderFlag = False 'HeaderFlag is used to modify the
    'header script depending on which print out method is chosen
  Else
    HeaderFlag = True
  End If
  
  If Check4ValidDept = False Then Exit Sub 'user alerts
  'are issued in the Check4ValidDept function if no match
  'was found
  ReportFile$ = "FAMaster.PRN"  'Report File Name
  Dash80$ = String$(80, "=")
  Lines80$ = String$(80, "-")
  FF$ = Chr$(12)
  Dispose$ = QPTrim$(fpcmbYN.Text)
  Dept$ = QPTrim$(fptxtDeptNum.Text)
  MaxLines = 52
  LineCnt = 0
  ItemCnt = 0

  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  OpenFAItemFile FAHandle
  NumOfFARecs = LOF(FAHandle) / Len(FAItemRec)
  
  If NumOfFARecs = 0 Then
    MsgBox "No fixed asset items on record."
    Close FAHandle
    Exit Sub
  End If
  
  TagFlag = False
  If QPTrim$(fpcmbOrder.Text) = "TAG NUMBER" Then
    TagFlag = True
  End If
  
  ReDim ItemNum(1 To NumOfFARecs) As Integer
  
  OpenTagIdxFile TagIdxHandle
  For x = 1 To NumOfFARecs
    Get TagIdxHandle, x, TagRec
      ItemNum(x) = TagRec.DataRecNum 'load array with item
      'tag numbers in numerical order
  Next x
  Close TagIdxHandle
  
  If TagFlag = False Then 'chose to display report by department
    OpenDeptIdxFile DIdxHandle
    DIdxRecNums = LOF(DIdxHandle) \ Len(DeptIdx)
    
    ReDim DIdx(1 To DIdxRecNums) As DeptNumbSortIdxType
    
    For x = 1 To DIdxRecNums
      Get DIdxHandle, x, DeptIdx 'load up dept data arrays
      DIdx(x).DeptRecNum = DeptIdx.DeptRecNum
      DIdx(x).DeptIdxDesc = QPTrim$(DeptIdx.DeptIdxDesc)
      DIdx(x).DeptNumb = QPTrim$(DeptIdx.DeptNumb)
    Next x
    
    MatchThis = QPTrim$(fptxtDeptNum.Text)
    If MatchThis <> "ALL" Then 'user wants a specific dept
      ThisDept = MatchThis
      For x = 1 To DIdxRecNums 'determine the specific dept's
      'description
        If ThisDept = QPTrim$(DIdx(x).DeptNumb) Then
          ThisDesc = QPTrim$(DIdx(x).DeptIdxDesc)
          Exit For
        End If
      Next x
    Else
      ThisDept = QPTrim$(DIdx(1).DeptNumb) 'otherwise just
      'start at the beginning
      ThisDesc = QPTrim$(DIdx(1).DeptIdxDesc)
      Dept$ = "ALL"
    End If
  
    Nextx = 1
  End If
  
  GoSub PrintMasterHeader 'now that we know the dept data
  'we can print an accurate header
  
  frmFAShowPctComp.Label1 = "Loading Master Item List"
  frmFAShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False
  DoEvents
  
  Do
    For cnt = 1 To NumOfFARecs
      Get FAHandle, ItemNum(cnt), FAItemRec
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintMasterHeader
      End If
      'Check For Disposed Of
      If Dispose$ = "N" Then 'if an item has a dsplflag = 1 it will
      'have a disposal date but we still want to print it here because
      'it is still considered in inventory until actually disposed of...
      'items converted from dos may have disposal dates with no dsplflag
        If (FAItemRec.DispDate > 0 And FAItemRec.DsplFlag = 0) Or FAItemRec.DsplFlag = 2 Then GoTo SkipEm
      End If
      If TagFlag = True Then 'if TagFlag is true then after one iteration
      'the program jumps out of the loop
        GoTo PrintIt
      ElseIf TagFlag = False Then 'here we only print those items
      'that are in the current department
        If Val(ThisDept) <> FAItemRec.IDEPT Then GoTo SkipEm
      End If
PrintIt:
      ItemTotal = ItemTotal + 1
      LifeLeft = CStr(FAItemRec.LifeLeft) 'here we are adjusting the
      'way this asset's life data is printed
      If Len(LifeLeft) = 2 Then
        LifeLeft = QPTrim$(LifeLeft)
      ElseIf Len(LifeLeft) = 1 Then
        LifeLeft = " " + QPTrim$(LifeLeft)
      End If
      If FAItemRec.ILIFE = 0 Then
        WholeLife = " 0"
      Else
        WholeLife = CStr(FAItemRec.ILIFE)
      End If
      
      LifeData = Left(QPTrim$(WholeLife), 3) + "/" + Left(LifeLeft, 3) 'assign line to LifeData
      If FAItemRec.DsplFlag = 1 Then
        Print #RptHandle, QPTrim$(FAItemRec.ItemTag) + " *D* on " + MakeRegDate(FAItemRec.DispDate); Tab(22); Left(QPTrim$(FAItemRec.IDESC1), 24); Tab(47); Left(QPTrim$(FAItemRec.SERIALNO), 20);
        Print #RptHandle, Tab(67); Using("$##,###,##0.00", FAItemRec.ORGCOST)
        LineCnt = LineCnt + 1 'the added script causes the print out to drop
        '1 line but this also highlights the item's status so it was left as is
      Else
        Print #RptHandle, QPTrim$(FAItemRec.ItemTag); Tab(22); Left(QPTrim$(FAItemRec.IDESC1), 24); Tab(47); Left(QPTrim$(FAItemRec.SERIALNO), 20);
        Print #RptHandle, Tab(67); Using("$##,###,##0.00", FAItemRec.ORGCOST)
      End If
      If FAItemRec.AQURDATE <= -20000 Then
        Print #RptHandle, Left$(QPTrim$(FAItemRec.ITEMLOC), 20); Tab(22); CStr(FAItemRec.FundNum) + "/" + CStr(FAItemRec.IDEPT); Tab(47); "No Record"; Tab(72); LifeData
      Else
        Print #RptHandle, Left$(QPTrim$(FAItemRec.ITEMLOC), 20); Tab(22); CStr(FAItemRec.FundNum) + "/" + CStr(FAItemRec.IDEPT); Tab(47); MakeRegDate(FAItemRec.AQURDATE); Tab(72); LifeData
      End If
      
      If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
        Print #RptHandle, Using("$######,0.00", FAItemRec.DEP2DATE) + "*";
      Else
        Print #RptHandle, Using("$######,0.00", FAItemRec.DEP2DATE);
      End If
      
'      Print #RptHandle, Using("$######,0.00", FAItemRec.DEP2DATE);
      If FAItemRec.ISTATUS = "I" Then
        Print #RptHandle, Tab(22); "Inactive"; Tab(47); QPTrim$(FAItemRec.ASSETCODE);
      Else
        Print #RptHandle, Tab(22); "Active"; Tab(47); QPTrim$(FAItemRec.ASSETCODE);
      End If
      
      If Dispose$ = "Y" And FAItemRec.DispDate > 0 Then
        If FAItemRec.DsplFlag = 2 Then
          Print #RptHandle, Tab(67); Using$("$##,###,##0.00", FAItemRec.CURRVAL); Tab(86); MakeRegDate(FAItemRec.DispDate)
          Print #RptHandle, String$(95, "-")
        ElseIf FAItemRec.DsplFlag = 1 Then
          Print #RptHandle, Tab(67); Using$("$##,###,##0.00", FAItemRec.CURRVAL); Tab(86); "Set for " + MakeRegDate(FAItemRec.DispDate)
          Print #RptHandle, String$(95, "-")
        End If
      Else
        Print #RptHandle, Tab(67); Using$("$##,###,##0.00", FAItemRec.CURRVAL)
        Print #RptHandle, String$(80, "-")
      End If
  
      'SubTotal Here
      LineCnt = LineCnt + 4
      ItemCnt = ItemCnt + 1
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintMasterHeader
      End If
      OrigCost# = OrigCost# + FAItemRec.ORGCOST
      DepTot# = DepTot# + FAItemRec.DEP2DATE
      BookTot# = BookTot# + FAItemRec.CURRVAL
      DItemCnt = DItemCnt + 1
      DOrigCost# = DOrigCost# + FAItemRec.ORGCOST
      DDepTot# = DDepTot# + FAItemRec.DEP2DATE
      DBookTot# = DBookTot# + FAItemRec.CURRVAL
SkipEm:
      If TagFlag = True Then
        frmFAShowPctComp.ShowPctComp cnt, NumOfFARecs
        If frmFAShowPctComp.Out = True Then
          Close
          frmFAShowPctComp.Out = False
          Unload frmFAShowPctComp
          EnableCloseButton Me.hwnd, True
          Me.cmdExit.Enabled = True
          Me.cmdProcess.Enabled = True
          Exit Sub
        End If
      End If
ExitRpt:
    Next cnt
    If TagFlag = True Then
      Exit Do
    ElseIf Dept$ <> "ALL" Then 'only wanted one dept displayed
      Exit Do
    Else
      If DItemCnt > 0 Then 'been thru one iteration so if DItemCnt
      'is zero there was nothing to report for this dept
        Print #RptHandle, "Subtotal for Dept Number: "; ThisDept; "  "; ThisDesc
        Print #RptHandle, "Total Items Printed: "; Using("#####", DItemCnt)
        Print #RptHandle, "Total Original Cost: "; Using("$########,#.##", DOrigCost#)
        Print #RptHandle, "Total Deprc to Date: "; Using("$########,#.##", DDepTot#)
        Print #RptHandle, "Total Book Value:    "; Using("$########,#.##", DBookTot#)
        If DsplYNFlag = False Then
          Print #RptHandle, Dash80$
        Else
          Print #RptHandle, String$(95, "=")
        End If
        
        Print #RptHandle,
        LineCnt = LineCnt + 5
      Else
        Print #RptHandle, "No fixed assets for Dept Number: "; ThisDept; "  "; ThisDesc
        If DsplYNFlag = False Then
          Print #RptHandle, Lines80$
        Else
          Print #RptHandle, String$(95, "-")
        End If
        
        Print #RptHandle,
        LineCnt = LineCnt + 3
      End If
      Nextx = Nextx + 1 'increment to get the next dept
      If Nextx > DIdxRecNums Then Exit Do 'this signifies that
      'there are no more depts left to gather data for so jump
      'out of loop
      ThisDept = QPTrim$(DIdx(Nextx).DeptNumb)
      ThisDesc = QPTrim$(DIdx(Nextx).DeptIdxDesc)
    End If
    frmFAShowPctComp.ShowPctComp Nextx, DIdxRecNums
    If frmFAShowPctComp.Out = True Then
      Close
      frmFAShowPctComp.Out = False
      Unload frmFAShowPctComp
      EnableCloseButton Me.hwnd, True
      Me.cmdExit.Enabled = True
      Me.cmdProcess.Enabled = True
      Exit Sub
    End If
    
    DItemCnt = 0
    DOrigCost# = 0
    DDepTot# = 0
    DBookTot# = 0
    Print #RptHandle, 'new department data can start in the middle
    'of a page
    Print #RptHandle, "Dept # "; ThisDept$; "  "; ThisDesc
    If DsplYNFlag = False Then
      Print #RptHandle, Dash80$
    Else
      Print #RptHandle, String(95, "=")
    End If
    LineCnt = LineCnt + 3
   
 Loop
  
  Unload frmFAShowPctComp
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True
  
  If ItemTotal = 0 Then
    MsgBox "There are no fixed assets saved for this criteria"
    Close
    GoTo NoItemTotal
  End If
  

  GoSub PrintMasterValueEnding
  
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close FAHandle
  Close         'Close all open files now

  If DsplYNFlag = True Then
    ViewPrint ReportFile$, "Fixed Asset Report", True
  Else
    ViewPrint ReportFile$, "Fixed Asset Report", False
  End If
NoItemTotal:
  KillFile (ReportFile$)

  Exit Sub

PrintMasterHeader:
  Page = Page + 1
  Print #RptHandle, Tab(29); "Master Asset Listing"
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
  If HeaderFlag = True Then
    Print #RptHandle, "Dept # "; ThisDept$; "  "; ThisDesc
  Else 'printed when TAG NUMBER is selected
    Print #RptHandle, "Dept # "; "ALL"
  End If
  Print #RptHandle, "*D* = disposal pending/to be disposed date"; Tab(60); "* = DO NOT DEPRECIATE"
'  Print #RptHandle, "*D* = disposal pending/to be disposed date"
  Print #RptHandle, ""
  Print #RptHandle, "Tag Number"; Tab(22); "Description"; Tab(47); "Serial Num"; Tab(68); "Original Cost"
  Print #RptHandle, "Location"; Tab(22); "Fund/Dept"; Tab(47); "Acquired On"; Tab(69); "Life/Left"
  Print #RptHandle, "Deprec To Date"; Tab(22); "Status"; Tab(47); "Asset Group";
  If Dispose$ = "Y" Then
    Print #RptHandle, Tab(69); "Book Value"; Tab(83); "Disposal Date"
  Else
    Print #RptHandle, Tab(69); "Book Value"
  End If
  If DsplYNFlag = False Then
    Print #RptHandle, Dash80$
  Else
    Print #RptHandle, String(95, "=")
  End If
  LineCnt = 14
  
Return

PrintMasterValueEnding:
  If DsplYNFlag = False Then
    Print #RptHandle, Dash80$
  Else
    Print #RptHandle, String(95, "=")
  End If
  Print #RptHandle, "Dept Number: "; Dept$
  Print #RptHandle, "Total Items Printed: "; Using("#####", ItemCnt)
  Print #RptHandle, "Total Original Cost: "; Using("$########,#.##", OrigCost#)
  Print #RptHandle, "Total Deprc to Date: "; Using("$########,#.##", DepTot#)
  Print #RptHandle, "Total Book Value:    "; Using("$########,#.##", BookTot#)
  Print #RptHandle, FF$
Return

   
ERRORSTUFF:
   Unload frmFAShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAMasterItemListing", gstrcProgName, Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
End Sub

Private Sub cmdProcess_Click()
  If fpcomboPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    If DsplYNFlag = True Then
      MsgBox "Pitch 12 is recommended for this report."
    Else
      MsgBox "Pitch 10 is recommended for this report."
    End If
    Call PrintText
  Else
    Exit Sub
  End If

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  DsplYNFlag = False
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%D"
      Call cmdDept_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      KillFile "masteritemlistopen.dat"
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAMasterItemListing.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim One As Integer
  Dim FileHandle As Integer
  
  One = 1
  FileHandle = FreeFile
  'masteritemlistopen.dat identifies this form to the dept list
  'form so that if the user uses that feature it will know which form
  'to populate the user's choice from that list
  Open "masteritemlistopen.dat" For Output As FileHandle Len = 2
  Print #FileHandle, One
  Close FileHandle
  fpcmbOrder.Text = "TAG NUMBER"
  fpcmbOrder.AddItem "TAG NUMBER"
  fpcmbOrder.AddItem "DEPARTMENT NUMBER"
  fpcmbYN.Text = "N"
  fpcmbYN.AddItem "Y"
  fpcmbYN.AddItem "N"
  fptxtDeptNum.Text = "ALL"
  fptxtDeptNum.Enabled = False
  fpcomboPrintOpt.AddItem "Graphical"
  fpcomboPrintOpt.AddItem "Text"
  fpcomboPrintOpt.Text = "Graphical"
  
End Sub

Private Sub fpcmbOrder_Change()
  'default this field to ALL and disable if TAG NUMBER
  'is chosen
  If QPTrim$(fpcmbOrder.Text) = "TAG NUMBER" Then
    fptxtDeptNum.Enabled = False
    fptxtDeptNum.Text = "ALL"
    cmdDept.Enabled = False
  ElseIf QPTrim$(fpcmbOrder.Text) = "" Then
    fpcmbOrder.Text = "TAG NUMBER"
    fptxtDeptNum.Enabled = False
    cmdDept.Enabled = False
    fptxtDeptNum.Text = "ALL"
  Else
    fptxtDeptNum.Enabled = True
    cmdDept.Enabled = True
  End If

End Sub

Private Sub fpcmbOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbOrder.ListIndex = -1
  End If
  If fpcmbOrder.ListDown <> True Then
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

Private Sub fpcmbYN_Change()
  'default this field to N
  If QPTrim$(fpcmbYN.Text) <> "Y" And QPTrim$(fpcmbYN.Text) <> "N" Then
    fpcmbYN.Text = "N"
  End If
  If QPTrim$(fpcmbYN.Text) = "Y" Then
    DsplYNFlag = True
  ElseIf QPTrim$(fpcmbYN.Text) = "N" Then
    DsplYNFlag = False
  End If
End Sub

Private Sub fpcmbYN_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYN.ListIndex = -1
  End If
  If fpcmbYN.ListDown <> True Then
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

Private Sub fptxtDeptNum_Change()
  'default this field to ALL
  If QPTrim$(fptxtDeptNum.Text) = "" Then
    fptxtDeptNum.Text = "ALL"
  End If
End Sub


Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcomboPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboPrintOpt.ListIndex = -1
  End If
  If fpcomboPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdExit.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub PrintGraphics()
  Const strcProcName As String = "PrintGraphics"
  
  Dim DHandle As Integer
  Dim DeptRec As FADeptCodeType
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxRecNums As Integer
  Dim x As Integer
  Dim NumOfDepts As Integer
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim ItemCnt As Long
  Dim RptHandle As Integer
  Dim ReportFile$
  Dim TagReportFile$
  Dim TagRptHandle As Integer
  Dim ItemRecNo As Integer
  Dim cnt As Integer
  Dim Dispose$, Dept$
  Dim DItemCnt#
  Dim DOrigCost#
  Dim DDepTot#
  Dim DBookTot#
  Dim OrigCost#, PValue#
  Dim DepTot#, Disp$, BookTot#
  Dim Page As Integer
  Dim TagFlag As Boolean
  Dim PCnt As Integer
  Dim PNumOFFARecs As Integer
  Dim TagIdxHandle As Integer
  Dim TagRec As TagNumbSortIdxType
  Dim Nextx As Integer
  Dim ThisDept$, ThisDesc$
  Dim MatchThis As String
  Dim dlm$
  Dim FASetUpRec As FASetupRecType
  Dim FASHandle As Integer
  Dim Employer$
  Dim GTRptHandle As Integer
  Dim GTReportFile$
  Dim DptRptHandle As Integer
  Dim DptReportFile$
  Dim SkipTagPrintFlag As Boolean
  Dim EmptyDeptFlag As Boolean
  Dim OneDept As Integer
  Dim DDate$, PrintDesc As String * 20
  Dim ItemTotal As Long
  
  'this sub runs much like the PrintText sub which is commented
  On Error GoTo ERRORSTUFF
  If Check4ValidDept = False Then Exit Sub
  OneDept = 1
  EmptyDeptFlag = False
  SkipTagPrintFlag = False
  
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  Employer = FASetUpRec.TownName
  
  dlm$ = "~"
  ReportFile$ = "FARPTS\FAMASTER.RPT"  'Report File Name
  TagReportFile$ = "FARPTS\FATAGMASTER.RPT"
  DptReportFile$ = "FARPTS\FADEPTMASTER.RPT"
  GTReportFile$ = "FARPTS\FAGTMASTER.RPT"
  Dispose$ = QPTrim$(fpcmbYN.Text)
  Dept$ = QPTrim$(fptxtDeptNum.Text)

  OpenFAItemFile FAHandle
  NumOfFARecs = LOF(FAHandle) / Len(FAItemRec)
  
  If NumOfFARecs = 0 Then
    MsgBox "No fixed asset items on record."
    Close FAHandle
    Exit Sub
  End If

  frmFAShowPctComp.Label1 = "Loading Master Item List"
  frmFAShowPctComp.Show
  DoEvents
  
  ReDim ItemNum(1 To NumOfFARecs) As Integer
  ReDim ItemDesc(1 To NumOfFARecs) As String
  
  TagFlag = False
  
  If QPTrim$(fpcmbOrder.Text) = "TAG NUMBER" Then
    TagFlag = True
    TagRptHandle = FreeFile
    Open TagReportFile$ For Output As #TagRptHandle
  Else
    RptHandle = FreeFile
    Open ReportFile$ For Output As #RptHandle
  End If
  
  OpenTagIdxFile TagIdxHandle
  For x = 1 To NumOfFARecs
    Get TagIdxHandle, x, TagRec
      ItemNum(x) = TagRec.DataRecNum
  Next x
  Close TagIdxHandle
  
SkipTagPrint:
  If SkipTagPrintFlag = True Then
    ItemCnt = 0
    OrigCost# = 0
    DepTot# = 0
    BookTot# = 0
    DItemCnt = 0
    DOrigCost# = 0
    DDepTot# = 0
    DBookTot# = 0
  End If
  
  If TagFlag = False Then
    OpenDeptIdxFile DIdxHandle
    DIdxRecNums = LOF(DIdxHandle) \ Len(DeptIdx)
    
    ReDim DIdx(1 To DIdxRecNums) As DeptNumbSortIdxType
    
    For x = 1 To DIdxRecNums
      Get DIdxHandle, x, DeptIdx
      DIdx(x).DeptRecNum = DeptIdx.DeptRecNum
      DIdx(x).DeptIdxDesc = DeptIdx.DeptIdxDesc
      DIdx(x).DeptNumb = DeptIdx.DeptNumb
    Next x
    
    MatchThis = QPTrim$(fptxtDeptNum.Text)
    If MatchThis <> "ALL" Then
      ThisDept = MatchThis
      For x = 1 To DIdxRecNums 'need the desc for this dept
        If ThisDept = QPTrim$(DIdx(x).DeptNumb) Then
          ThisDesc = QPTrim$(DIdx(x).DeptIdxDesc)
          Exit For
        End If
      Next x
    Else
      ThisDept = QPTrim$(DIdx(1).DeptNumb) 'start with the first dept
      ThisDesc = QPTrim$(DIdx(1).DeptIdxDesc)
      Dept$ = "ALL"
    End If
  
  End If
  
  If DIdxRecNums = 0 Then DIdxRecNums = 1

  ReDim MLDept(1 To DIdxRecNums) As String
  ReDim MLDesc(1 To DIdxRecNums) As String
  ReDim MLItemCnt(1 To DIdxRecNums) As Long
  ReDim MLPurchPrice(1 To DIdxRecNums) As Double
  ReDim MLDeprToDate(1 To DIdxRecNums) As Double
  ReDim MLBookVal(1 To DIdxRecNums) As Double
  Nextx = 1
  
  Do
    For cnt = 1 To NumOfFARecs
      Get FAHandle, ItemNum(cnt), FAItemRec
      
      If Dispose$ = "N" Then 'user chose to exclude disposed of items
        If FAItemRec.DispDate > 0 And FAItemRec.DsplFlag <> 1 Then GoTo SkipEm 'left off display flag = 2 because of conversions
      Else
        If FAItemRec.DispDate = 0 Then
          DDate = ""
        Else
          DDate = MakeRegDate(FAItemRec.DispDate)
        End If
      End If
      
      If TagFlag = False And MatchThis = "ALL" Then 'all depts = true
        OneDept = 2
        ThisDept = QPTrim$(DIdx(Nextx).DeptNumb)
        ThisDesc = QPTrim$(DIdx(Nextx).DeptIdxDesc)
      ElseIf TagFlag = True Then
        GoTo PrintIt
      End If
      
     If Val(ThisDept) <> FAItemRec.IDEPT Then GoTo SkipEm 'filter out non-matching dept nums
      
PrintIt:
    PrintDesc = QPTrim$(FAItemRec.IDESC1)
    DItemCnt = DItemCnt + 1
    DOrigCost# = DOrigCost# + FAItemRec.ORGCOST
    DDepTot# = DDepTot# + FAItemRec.DEP2DATE
    DBookTot# = DBookTot# + FAItemRec.CURRVAL
    If SkipTagPrintFlag = True Then GoTo NoTagPrint
    If TagFlag = False Then
'      If Dispose$ = "Y" And FAItemRec.DispDate > 0 Then
        If FAItemRec.DsplFlag = 1 Then
          '                          0                                                                  1
          Print #RptHandle, FAItemRec.ItemTag; dlm; FAItemRec.SERIALNO; dlm;
        Else
          '                          0                         1
          Print #RptHandle, FAItemRec.ItemTag; dlm; FAItemRec.SERIALNO; dlm;
        End If
'      Else
'      '                       0        1
'        Print #RptHandle, FAItemRec.ItemTag; dlm; FAItemRec.SERIALNO; dlm;
'      End If
      '                          2                         3
      Print #RptHandle, PrintDesc; dlm; FAItemRec.ITEMMFG; dlm;
      '                          4                         5
      Print #RptHandle, Dispose$; dlm; FAItemRec.CONTACT; dlm;
      '                          6                         7
      Print #RptHandle, QPTrim$(FAItemRec.ITEMLOC); dlm; CStr(FAItemRec.FundNum) + "/" + CStr(FAItemRec.IDEPT); dlm;
      If FAItemRec.AQURDATE <= -20000 Then
        '                          8                         9
        Print #RptHandle, FAItemRec.ORGCOST; dlm; "No Record"; dlm;
      Else
        '                          8                         9
        Print #RptHandle, FAItemRec.ORGCOST; dlm; MakeRegDate(FAItemRec.AQURDATE); dlm;
      End If
      '                         10                         11
      Print #RptHandle, FAItemRec.ILIFE; dlm; FAItemRec.DEP2DATE; dlm;
      '                         12                         13
      Print #RptHandle, FAItemRec.ISTATUS; dlm; FAItemRec.CURRVAL; dlm;
      '                         14                    15              16
      Print #RptHandle, FAItemRec.ASSETCODE; dlm; ThisDept$; dlm; ThisDesc; dlm;
      '                    17             18              19              20             21
      Print #RptHandle, Employer; dlm; DItemCnt; dlm; DOrigCost#; dlm; DDepTot#; dlm; DBookTot#; dlm;
      
      If Dispose$ = "N" Then
        '                    22                 23                     24              25
        Print #RptHandle, OneDept; dlm; FAItemRec.LifeLeft; dlm; FAItemRec.DEPYN; dlm; ""; dlm;
      Else
        If FAItemRec.DsplFlag = 1 Then
          '                    22                       23              24                    25
          Print #RptHandle, OneDept; dlm; FAItemRec.LifeLeft; dlm; FAItemRec.DEPYN; dlm; "P/" + DDate; dlm;
        ElseIf FAItemRec.DsplFlag = 2 Then
          '                    22                       23              24                 25
          Print #RptHandle, OneDept; dlm; FAItemRec.LifeLeft; dlm; FAItemRec.DEPYN; dlm; DDate; dlm;
        Else
          '                    22                       23              24               25
          Print #RptHandle, OneDept; dlm; FAItemRec.LifeLeft; dlm; FAItemRec.DEPYN; dlm; ""; dlm;
        End If
      End If
      If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
        '                    26
        Print #RptHandle, "*"
      Else
        '                    26
        Print #RptHandle, ""
      End If
   Else
'      If Dispose$ = "Y" And FAItemRec.DispDate > 0 Then
        If FAItemRec.DsplFlag = 1 Then
          '                            0                                                                     1
          Print #TagRptHandle, FAItemRec.ItemTag; dlm; FAItemRec.SERIALNO; dlm;
        Else
          '                            0                        1
          Print #TagRptHandle, FAItemRec.ItemTag; dlm; FAItemRec.SERIALNO; dlm;
        End If
'      Else
'      '                        0        1
'        Print #TagRptHandle, FAItemRec.ItemTag; dlm; FAItemRec.SERIALNO; dlm;
'      End If
      '                            2                        3
      Print #TagRptHandle, PrintDesc; dlm; FAItemRec.ITEMMFG; dlm;
      '                            4                        5
      Print #TagRptHandle, Dispose$; dlm; FAItemRec.CONTACT; dlm;
      '                            6                        7
      Print #TagRptHandle, QPTrim$(FAItemRec.ITEMLOC); dlm; CStr(FAItemRec.FundNum) + "/" + CStr(FAItemRec.IDEPT); dlm;
      '                            8                        9
      Print #TagRptHandle, FAItemRec.ORGCOST; dlm; MakeRegDate(FAItemRec.AQURDATE); dlm;
      '                           10                        11
      Print #TagRptHandle, FAItemRec.ILIFE; dlm; FAItemRec.DEP2DATE; dlm;
      '                           12                        13
      Print #TagRptHandle, FAItemRec.ISTATUS; dlm; FAItemRec.CURRVAL; dlm;
      '                           14
      Print #TagRptHandle, FAItemRec.ASSETCODE; dlm; ThisDept$; dlm; ThisDesc; dlm;
      '                    17                18              19              20             21                               22
      Print #TagRptHandle, Employer; dlm; DItemCnt; dlm; DOrigCost#; dlm; DDepTot#; dlm; DBookTot#; dlm;
      If Dispose$ = "N" Then
        '                            22                       23             24
        Print #TagRptHandle, FAItemRec.LifeLeft; dlm; FAItemRec.DEPYN; dlm; ""; dlm;
      Else
        If FAItemRec.DsplFlag = 1 Then
          '                            22                       23              24
          Print #TagRptHandle, FAItemRec.LifeLeft; dlm; FAItemRec.DEPYN; dlm; "P/" + DDate; dlm;
        ElseIf FAItemRec.DsplFlag = 2 Then
          '                            22                       23              24
          Print #TagRptHandle, FAItemRec.LifeLeft; dlm; FAItemRec.DEPYN; dlm; DDate; dlm;
        Else
          '                            22                       23              24
          Print #TagRptHandle, FAItemRec.LifeLeft; dlm; FAItemRec.DEPYN; dlm; ""; dlm;
        End If
      End If
      If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
        '                    25
        Print #TagRptHandle, "*"
      Else
        '                    25
        Print #TagRptHandle, ""
      End If
   End If
   ItemTotal = ItemTotal + 1
NoTagPrint:
      'SubTotal Here
   OrigCost# = OrigCost# + FAItemRec.ORGCOST
   DepTot# = DepTot# + FAItemRec.DEP2DATE
   BookTot# = BookTot# + FAItemRec.CURRVAL
   ItemCnt = ItemCnt + 1
  
SkipEm:
    Next cnt
    
    If TagFlag = True Then
      TagFlag = False
      SkipTagPrintFlag = True
      GoTo SkipTagPrint
    ElseIf Dept$ <> "ALL" Then
      Exit Do
    Else
      MLDept(Nextx) = ThisDept$ 'tally dept data
      MLDesc(Nextx) = ThisDesc$
      MLItemCnt(Nextx) = DItemCnt
      MLPurchPrice(Nextx) = DOrigCost#
      MLDeprToDate(Nextx) = DDepTot#
      MLBookVal(Nextx) = DBookTot#
      Nextx = Nextx + 1 'move to next dept
      If Nextx > DIdxRecNums Then Exit Do 'if there are
      'no more dept numbers then jump out of loop
    End If
    
    DItemCnt = 0 'reset for next dept
    DOrigCost# = 0 'reset for next dept
    DDepTot# = 0 'reset for next dept
    DBookTot# = 0
    
    frmFAShowPctComp.ShowPctComp Nextx, DIdxRecNums
    If frmFAShowPctComp.Out = True Then
      Close
      frmFAShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdExit.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload frmFAShowPctComp
      Exit Sub
    End If
  Loop
  
  Unload frmFAShowPctComp
  
  Close FAHandle
  Close         'Close all open files now
  
  If ItemTotal = 0 Then
    MsgBox "There are no fixed assets saved for this criteria"
    Close
    GoTo NoItemTotal
  End If
  
  If Dept$ = "ALL" Then
    DptRptHandle = FreeFile
    Open DptReportFile For Output As DptRptHandle
    For x = 1 To DIdxRecNums
      '                        0               1                 2                   3                     4
      Print #DptRptHandle, MLDept(x); dlm; MLDesc(x); dlm; MLItemCnt(x); dlm; MLPurchPrice(x); dlm; MLDeprToDate(x); dlm; MLBookVal(x)
    Next x
    Close DptRptHandle
  End If
  
  GTRptHandle = FreeFile
  Open GTReportFile$ For Output As #GTRptHandle
  GoSub PrintMasterValueEnding

  Close
1000:
  If SkipTagPrintFlag = False Then
    arFAMasterRpt.Show
  Else
1100:
    arFATagMaster.Show
  End If
1200:
  frmFALoadReport.Show
1300:
NoItemTotal:
  Exit Sub

PrintMasterValueEnding:
  '                     0            1              2             3
  Print #GTRptHandle, Dept$; dlm; ItemCnt; dlm; OrigCost#; dlm; DepTot#; dlm; BookTot#
Return


ERRORSTUFF:
   Unload frmFAShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAMasterItemListing", strcProcName, Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Function Check4ValidDept() As Boolean
  Dim x As Integer
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxRecNums As Integer
  Dim ThisDept$
  
  On Error GoTo ERRORSTUFF
  
  'this function examines each valid department number
  'saved and compares the department number entered by
  'the user with each until a match is found or if no match
  'is found this function alerts the user
  Check4ValidDept = True
  
  OpenDeptIdxFile DIdxHandle
  DIdxRecNums = LOF(DIdxHandle) \ Len(DeptIdx)
  If DIdxRecNums = 0 Then
    MsgBox "No departments saved in index."
    Close
    Check4ValidDept = False
    Exit Function
  End If
  
  If QPTrim$(fptxtDeptNum.Text) = "ALL" Then
    Close
    Exit Function
  End If
  
  ThisDept$ = QPTrim$(fptxtDeptNum.Text)
  
  For x = 1 To DIdxRecNums
    Get DIdxHandle, x, DeptIdx
    If ThisDept$ = QPTrim$(DeptIdx.DeptNumb) Then
      Close
      Exit Function
    End If
  Next x
  
  MsgBox "No department number matches this entry. Please try again."
  Check4ValidDept = False
  fptxtDeptNum.SetFocus
  Close
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAMasterItemListing", "Check4ValidDept", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    ClearInUse (PWcnt)
    Terminate
    Unload Me
End Function
