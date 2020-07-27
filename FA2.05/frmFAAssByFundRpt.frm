VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFAAssByFundRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets by Fund Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAAssByFundRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6108
      Left            =   1872
      TabIndex        =   6
      Top             =   1368
      Width           =   7884
      _Version        =   196609
      _ExtentX        =   13906
      _ExtentY        =   10774
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmFAAssByFundRpt.frx":08CA
      Begin LpLib.fpCombo fpcmbYN 
         Height          =   405
         Left            =   5565
         TabIndex        =   4
         ToolTipText     =   "Enter Y to include disposed of fixed assets or N to exclude disposed of fixed assets."
         Top             =   3600
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
         ColDesigner     =   "frmFAAssByFundRpt.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbOrder 
         Height          =   405
         Left            =   3210
         TabIndex        =   0
         ToolTipText     =   "Select the order this report will display data."
         Top             =   1350
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
         ColDesigner     =   "frmFAAssByFundRpt.frx":0BDD
      End
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3555
         TabIndex        =   5
         ToolTipText     =   "Select  Graphic for a robust report that takes more time to process. Select Text for a faster report."
         Top             =   4230
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
         ColDesigner     =   "frmFAAssByFundRpt.frx":0ED4
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00D0D0D0&
         Caption         =   " Include Department Breakdown?"
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
         Left            =   2256
         TabIndex        =   2
         Top             =   2496
         Width           =   3900
      End
      Begin EditLib.fpText fptxtFundNum 
         Height          =   396
         Left            =   3072
         TabIndex        =   1
         ToolTipText     =   "If Report Order is DEPARTMENT NUMBER then enter the desired department number which will appear in this report."
         Top             =   1968
         Width           =   1500
         _Version        =   196608
         _ExtentX        =   2646
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
      Begin EditLib.fpText fptxtDeptNum 
         Height          =   396
         Left            =   3072
         TabIndex        =   3
         ToolTipText     =   "If Report Order is DEPARTMENT NUMBER then enter the desired department number which will appear in this report."
         Top             =   2976
         Width           =   1500
         _Version        =   196608
         _ExtentX        =   2646
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
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 - A L L"
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
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   690
         Left            =   1536
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   4944
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
         ButtonDesigner  =   "frmFAAssByFundRpt.frx":11CB
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4560
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to create the report based on the parameters entered above."
         Top             =   4944
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
         ButtonDesigner  =   "frmFAAssByFundRpt.frx":13A7
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdDeptList 
         Height          =   390
         Left            =   4704
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to bring up a list of all current departments."
         Top             =   2976
         Width           =   1350
         _Version        =   131072
         _ExtentX        =   2381
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
         ButtonDesigner  =   "frmFAAssByFundRpt.frx":1586
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdFund 
         Height          =   390
         Left            =   4704
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to bring up a list of all current fund codes."
         Top             =   1968
         Width           =   1350
         _Version        =   131072
         _ExtentX        =   2381
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
         ButtonDesigner  =   "frmFAAssByFundRpt.frx":1766
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fund #"
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
         Left            =   1920
         TabIndex        =   12
         Top             =   2064
         Width           =   924
      End
      Begin VB.Label Label10 
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
         Left            =   1200
         TabIndex        =   11
         Top             =   1392
         Width           =   1836
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1536
         Top             =   336
         Width           =   4908
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Assets By Asset Fund Report"
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
         Left            =   1920
         TabIndex        =   10
         Top             =   480
         Width           =   4215
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
         Left            =   1872
         TabIndex        =   9
         Top             =   4320
         Width           =   1500
      End
      Begin VB.Label Label2 
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
         TabIndex        =   8
         Top             =   3696
         Width           =   3660
      End
      Begin VB.Label Label1 
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
         Left            =   1920
         TabIndex        =   7
         Top             =   3072
         Width           =   924
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6300
      Left            =   1788
      Top             =   1284
      Width           =   8076
   End
End
Attribute VB_Name = "frmFAAssByFundRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim DsplYNFlag As Boolean

Private Sub cmdCode_Click()
  frmFAAssetCodeList.Show vbModal
  DoEvents
End Sub

Private Sub Check1_Click()
  If Check1.Value = 1 Then
    cmdDeptList.Enabled = True
    fptxtDeptNum.Enabled = True
  Else
    cmdDeptList.Enabled = False
    fptxtDeptNum.Enabled = False
  End If
End Sub

Private Sub cmdDeptList_Click()
  frmFADeptList.Show vbModal
End Sub

Private Sub cmdExit_Click()
  frmFAReportMenu.Show
  Close
  DoEvents
  KillFile "assetbyfundrpt.dat"
  Unload frmFAAssByFundRpt
End Sub

Private Sub cmdFund_Click()
  frmFAFundList.Show vbModal
End Sub

Private Sub cmdProcess_Click()
  If fpcomboPrintOpt.Text = "Graphical" Then
    If Check1.Value = 0 Then
      Call PrintGraphics
    Else
      Call PrintDeptGraphics
    End If
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    MsgBox "Pitch 15 or higher is recommended for this report."
    If Check1.Value = 0 Then
      Call PrintText
    ElseIf Check1.Value = 1 Then
      Call PrintDeptText
    Else
      Exit Sub
    End If
  End If
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
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
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%L"
      Call cmdFund_Click
      KeyCode = 0
    Case vbKeyF9:
      SendKeys "%D"
      Call cmdDeptList_Click
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
      KillFile "assetbyfundrpt.dat"
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAAssByFundRpt.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub fpcmbOrder_Change()
  'default is ALL for this combo box
  If QPTrim$(fpcmbOrder.Text) = "TAG NUMBER" Then
    Check1.Enabled = False
    Check1.Value = 0
    fptxtFundNum.Enabled = False
    cmdFund.Enabled = False
    fptxtFundNum.Text = "ALL"
    fptxtDeptNum.Enabled = False
    cmdDeptList.Enabled = False
    fptxtDeptNum.Text = "ALL"
  ElseIf QPTrim$(fpcmbOrder.Text) = "" Then
    Check1.Enabled = False
    Check1.Value = 0
    fpcmbOrder.Text = "TAG NUMBER"
    fptxtFundNum.Enabled = False
    cmdFund.Enabled = False
    fptxtFundNum.Text = "ALL"
    fptxtDeptNum.Enabled = False
    cmdDeptList.Enabled = False
    fptxtDeptNum.Text = "ALL"
  Else
    Check1.Enabled = True
    Check1.Value = 0
    fptxtFundNum.Enabled = True
    cmdFund.Enabled = True
  End If
  
End Sub

Private Sub fpcmbOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  'this prevents the user from inadvertently changing data in the combo box when
  'tabbing through the fields
  If KeyCode = vbKeySpace Then
    fpcmbOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbOrder.ListIndex = -1
  End If
  If fpcmbOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fpcmbOrder.Text = "TAG NUMBER" Then
        fpcmbYN.SetFocus
      Else
        fptxtFundNum.SetFocus
      End If
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub LoadMe()
  Dim One As Integer
  Dim FileHandle As Integer
  One = 1
  FileHandle = FreeFile
  Open "assetbyfundrpt.dat" For Output As FileHandle Len = 2
  Print #FileHandle, One
  Close FileHandle
  fpcmbOrder.Text = "TAG NUMBER"
  fpcmbOrder.AddItem "TAG NUMBER"
  fpcmbOrder.AddItem "FUND NUMBER"
  fptxtFundNum.Text = "ALL"
  fptxtDeptNum.Text = "ALL"
  fptxtDeptNum.Enabled = False
  cmdDeptList.Enabled = False
  fpcomboPrintOpt.AddItem "Graphical"
  fpcomboPrintOpt.AddItem "Text"
  fpcomboPrintOpt.Text = "Graphical"
  fpcmbYN.Text = "N"
  fpcmbYN.AddItem "Y"
  fpcmbYN.AddItem "N"
  Check1.Enabled = False
  Check1.Value = 0
  
End Sub

Private Function Check4ValidFund() As Boolean
  Dim FundRec As FAFundCodeType
  Dim FundHandle As Integer
  Dim NumOfFundRecs As Integer
  Dim ThisFund As Integer
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
        
  OpenFAFundCodeFile FundHandle
  NumOfFundRecs = LOF(FundHandle) / Len(FundRec)
  If NumOfFundRecs = 0 Then
    MsgBox "No fund code records saved."
    Close
    Exit Function
  End If
  
  Check4ValidFund = True
  
  If QPTrim$(fptxtFundNum.Text) = "ALL" Then
    Close
    Exit Function
  End If
  
  ThisFund = Val(fptxtFundNum.Text)
  
  For x = 1 To NumOfFundRecs
    Get FundHandle, x, FundRec
    If ThisFund = FundRec.FundNum Then
      Close
      Exit Function
    End If
  Next x
  
  MsgBox "No fund code number matches this entry. Please try again."
  Check4ValidFund = False
  fptxtFundNum.SetFocus
  Close
  
  Exit Function
        
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAAssByFundRpt", "Check4ValidFund", Erl)
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
    Unload Me
  
End Function

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
  If KeyCode = vbKeySpace Then
    fpcmbYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYN.ListIndex = -1
  End If
  If fpcmbYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcomboPrintOpt.SetFocus
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

Private Sub fpcomboPrintOpt_Change()
  'Graphical is the default
  If QPTrim$(fpcomboPrintOpt.Text) = "" Then
    fpcomboPrintOpt.Text = "Graphical"
  End If
End Sub

Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  'this prevents the user from inadvertently changing data in the combo box when
  'tabbing through the fields
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

Private Sub PrintText()
  Dim ReportFile$
  Dim Dash80$
  Dim FF$
  Dim MaxLines As Integer
  Dim ItemCnt&
  Dim LineCnt&
  Dim code$
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim x As Integer
  Dim Nextx As Integer
  Dim Index$
  Dim Page As Integer
  Dim FundIdxHandle As Integer
  Dim FundIdxCnt As Integer
  Dim FundIdxRec As FundNumbSortIdxType
  Dim FundNumber As Integer
  Dim FundNumDesc$
  Dim FundHeader$
  Dim FundCnt As Integer
  Dim FirstFlag As Boolean
  Dim TagFlag As Boolean
  Dim DataFlag As Boolean
  Dim OrigCost As Double
  Dim BookTotal As Double
  Dim YDep As Double
  Dim YTDDep As Double
  Dim COrigCost As Double
  Dim CBookTotal As Double
  Dim CYDep As Double
  Dim LifeLeft$
  Dim WholeLife$
  Dim LifeData$
  Dim TotalItems As Integer
  Dim TagPrint As Boolean
  Dim FundRec As FAFundCodeType
  Dim FundHandle As Integer
  Dim NumOfFundRecs As Integer
  Dim ThisFundDesc$
  Dim ThisFundNum$
  Dim ItemTotal As Long
  Dim DeptNum As Integer
  
  On Error GoTo ERRORSTUFF
        
  If QPTrim$(fptxtDeptNum.Text) <> "ALL" Then
    If Check4ValidDept = True Then
      DeptNum = CInt(fptxtDeptNum.Text)
    Else
      Exit Sub
    End If
  End If
  
  FirstFlag = True

  If Check4ValidFund = False Then Exit Sub

  ReportFile$ = "FAFUNDRPT.PRN"  'Report File Name
  Dash80$ = String$(80, "=")
  FF$ = Chr$(12)

  MaxLines = 56
  LineCnt& = 0
  ItemCnt& = 0
  code$ = QPTrim$(fptxtFundNum.Text)

  RptHandle = FreeFile
  Index$ = QPTrim$(fpcmbOrder.Text)
  Open ReportFile$ For Output As #RptHandle

  OpenTagIdxFile TagIdxHandle
  NumOfFARecs = LOF(TagIdxHandle) \ Len(TagIdx)
  If NumOfFARecs = 0 Then
    MsgBox "No item records on file."
    Close TagIdxHandle
    Exit Sub
  End If
  
  ReDim TagIdxRecs(1 To NumOfFARecs)
  For x = 1 To NumOfFARecs
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum
  Next x
  Close TagIdxHandle
  
  OpenFundIdxFile FundIdxHandle
  FundIdxCnt = LOF(FundIdxHandle) \ Len(FundIdxRec)
  ReDim FundRecNum(1 To FundIdxCnt) As Integer
  For x = 1 To FundIdxCnt
    Get FundIdxHandle, x, FundIdxRec
    FundRecNum(x) = FundIdxRec.FundRecNum
  Next x
  Close FundIdxHandle
  
  OpenFAFundCodeFile FundHandle
  NumOfFundRecs = LOF(FundHandle) / Len(FundRec)
  If NumOfFundRecs <> FundIdxCnt Then
    Call CreateFundIdx
  End If
  
  If NumOfFundRecs <> FundIdxCnt Then
    MsgBox "Error: The number of fund codes saved is not the same number as the number of fund codes indexed. Please go to Fund Code Maintenance and resave any fund code to reindex."
    Close
    Exit Sub
  End If
  
  ReDim FundNum(1 To FundIdxCnt) As Integer
  ReDim FundDesc(1 To FundIdxCnt) As String
  For x = 1 To FundIdxCnt
    Get FundHandle, FundRecNum(x), FundRec
      FundNum(x) = FundRec.FundNum
      FundDesc(x) = QPTrim$(FundRec.FundDesc)
  Next x
  Close FundHandle
  
  If code$ <> "ALL" Then
    ThisFundNum = QPTrim$(fptxtFundNum.Text)
    For x = 1 To FundIdxCnt
      If ThisFundNum = FundNum(x) Then
        ThisFundDesc = QPTrim$(FundDesc(x))
        Exit For
      End If
    Next x
  Else
    ThisFundNum = FundNum(1)
    ThisFundDesc = QPTrim(FundDesc(1))
  End If

  GoSub PrintMasterHeader1

  ReDim ATagCOrigCost(1 To FundIdxCnt) As Double
  ReDim ATagCBookTotal(1 To FundIdxCnt) As Double
  ReDim ATagCYDep(1 To FundIdxCnt) As Double
  ReDim ATagCCnt(1 To FundIdxCnt) As Long
  
  OpenFAItemFile FAHandle

  TagFlag = False

  frmFAShowPctComp.Label1 = "Gathering Asset Fund Data"
  frmFAShowPctComp.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False

GetTagTotals: 'to print the report for tag numbers only the report
  'runs through all items one time in tag order and prints pertinent
  'data...then it returns to this spot and runs back through the
  'following loop gathering department totals (just like the
  'DEPARTMENT NUMBERS option does) but does not print items by
  'department...it just prints department totals at the end of the
  'report
  Nextx = 1
  If TagFlag = True Then
    Index = "ASSET FUND NUMBER"
    LineCnt = 0
  End If

  Do 'this loop iterates once if there is 1 dept requested, DIdxCnt + 1 if
  'department numbers "ALL" is requested and DIdxCnt + 2 for Tag Numbers
  '...Tag Numbers requires one iteration ignoring departments to get a list of
  'all valid tags in numeric order then DIdxCnt + 1 skipping the itemized tag
  'data print out just to allow the dept totals to assimilate...
    DataFlag = False
    For cnt& = 1 To NumOfFARecs
      Get FAHandle, TagIdxRecs(cnt), FAItemRec
      If fpcmbYN.Text = "N" Then
        If FAItemRec.DsplFlag = 2 Then GoTo SkipEm1
      End If
      
      If LineCnt& >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintMasterHeader1
      End If
      
      YTDDep# = FAItemRec.DEP2DATE

      'TAG NUMBER is actually true (if selected by the user) for the first
      'complete iteration only so that all tag numbers can be printed in numeric
      'order. After that we do not want the tags itemized we only want department
      'totals figured so that at the end of the report a "totals by dept" section
      'can be printed.
      If QPTrim$(Index) = "TAG NUMBER" Then
        GoTo TagOnly1 'prints itemized tag data in numeric order
      ElseIf ThisFundNum <> FAItemRec.FundNum Then 'any time department data
      'is needed then the item falls into this part of the if statement
        GoTo SkipEm1 'if the prevailing department (in the numeric index)
        'doesn't match this item's dept number then we don'y want it now
      End If
      'at this point the item's dept matches the prevailing dept number
      'but since TagFlag (don't want itemized tag data anymore) is true then
      'go ahead and collect dept data
      If DeptNum <> 0 Then
        If FAItemRec.IDEPT <> DeptNum Then
          GoTo SkipEm1
        End If
      End If
      If TagFlag = True Then GoTo TagOnly2
TagOnly1: 'printing valid tag data...skipped if TAG NUMBER is chosen and this is
          'not the first iteration
      If QPTrim$(fpcmbOrder.Text) <> "TAG NUMBER" And FundCnt = 0 Then
        Print #RptHandle, String$(111, "=")
        Print #RptHandle, "Assets for Fund Number: "; CStr(ThisFundNum); " "; ThisFundDesc
        Print #RptHandle, String$(111, "_")
        LineCnt& = LineCnt& + 3
      End If
      DataFlag = True
      LifeLeft = CStr(FAItemRec.LifeLeft)
      'format the asset's life data
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
      LifeData = QPTrim$(WholeLife) + "/" + LifeLeft
      Print #RptHandle, Tab(2); QPTrim$(FAItemRec.ItemTag); Tab(22); Left$(FAItemRec.IDESC1, 28);
      Print #RptHandle, Tab(51); Using("###0", FAItemRec.FundNum);
      Print #RptHandle, Tab(60); LifeData;
      Print #RptHandle, Tab(68); Using("$##,###,##0.00", CStr(FAItemRec.ORGCOST));
      If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
        Print #RptHandle, Tab(82); Using("$##,###,##0.00", CStr(YTDDep#)); "*";
      Else
        Print #RptHandle, Tab(82); Using("$##,###,##0.00", CStr(YTDDep#));
      End If
      Print #RptHandle, Tab(98); Using("$##,###,##0.00", CStr(FAItemRec.CURRVAL))
      If fpcmbYN.Text = "Y" Then
        If FAItemRec.DsplFlag = 2 Then
          Print #RptHandle, Tab(22); "^Disposal Date: "; Tab(40); MakeRegDate(FAItemRec.DispDate)
          LineCnt& = LineCnt& + 1
        End If
      End If
      If FAItemRec.DsplFlag = 1 Then
        Print #RptHandle, Tab(10); "^Scheduled For Disposal On: "; Tab(40); MakeRegDate(FAItemRec.DispDate)
        LineCnt& = LineCnt& + 1
      End If
      LineCnt& = LineCnt& + 1
      ItemCnt& = ItemCnt& + 1
      ItemTotal = ItemTotal + 1
TagOnly2: 'collects data for each department for reporting totals

      'This if statement filters out the first iteration of TAG NUMBER
      'selection because we do not want to start accumulating dept data
      'until the second iteration
      If TagFlag = False And QPTrim$(Index) = "TAG NUMBER" Then GoTo SkipEm1
  
      'collects grand totals
      OrigCost# = OrigCost# + FAItemRec.ORGCOST
      BookTotal# = BookTotal# + (FAItemRec.CURRVAL)
      YDep# = YDep# + YTDDep#
      'collects dept totals
      FundCnt = FundCnt + 1
      ATagCCnt(Nextx) = FundCnt
      TotalItems = TotalItems + 1
      COrigCost# = COrigCost# + FAItemRec.ORGCOST
      ATagCOrigCost(Nextx) = COrigCost#
      CBookTotal# = CBookTotal# + (FAItemRec.CURRVAL)
      ATagCBookTotal(Nextx) = CBookTotal#
      CYDep# = CYDep# + YTDDep#
      ATagCYDep(Nextx) = CYDep#
SkipEm1:

    Next cnt&
    'here we begin the iteration over again but this time TagFlag
    'becomes true so we know that this was originally TAG NUMBERS
    'and the first iteration is done
    If QPTrim$(Index) = "TAG NUMBER" And TagFlag = False Then
      TagFlag = True
      GoTo GetTagTotals
      Exit Do
    End If

    If TagFlag = True Then GoTo NoData 'don't want the next dept
    'data to print

    If DataFlag = False Then
      GoTo NoData
    End If

  'First Print Subtotals
    Print #RptHandle, String$(111, "_")
    Print #RptHandle, "Assets for Fund Number: "; CStr(ThisFundNum); " "; ThisFundDesc;
    Print #RptHandle, Tab(68); Using("$##,###,##0.00", CStr(COrigCost#));
    Print #RptHandle, Tab(82); Using("$##,###,##0.00", CStr(CYDep#));
    Print #RptHandle, Tab(98); Using("$##,###,##0.00", CStr(CBookTotal#))
    Print #RptHandle, "Total Items: "; CStr(FundCnt)
    Print #RptHandle, String$(111, "=")
    Print #RptHandle,
    LineCnt& = LineCnt& + 5
NoData:
    frmFAShowPctComp.ShowPctComp Nextx, FundIdxCnt
    If frmFAShowPctComp.Out = True Then
      Close
      frmFAShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdExit.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload frmFAShowPctComp
      Exit Sub
    End If
    'if "ALL" is not selected then the user has selected a single
    'department so we have all the data we need at this point...exit
    If QPTrim$(code$) <> "ALL" Then Exit Do
    'if all the depts have been examined then time to go
    If Nextx = FundIdxCnt Then Exit Do
    'move to the next dept
    Nextx = Nextx + 1
    'assign new dept to DeptNumber
    ThisFundNum = FundNum(Nextx)
    ThisFundDesc = QPTrim$(FundDesc(Nextx))
    'clear all dept totals
    COrigCost# = 0
    CBookTotal# = 0
    CYDep# = 0
    FundCnt = 0
  Loop

  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True

  If ItemTotal = 0 Then
    MsgBox "There are no fixed assets saved for this criteria."
    Close
    Exit Sub
  End If
  
  GoSub PrintDeptTotals

  If TagPrint = False Then GoSub PrintMasterValueEnding1
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi

  Close         'Close all open files now

  ViewPrint ReportFile$, "Value By Purchase Price", True

  KillFile (ReportFile$)

  Exit Sub

PrintMasterHeader1:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Master Asset Listing : Asset Listing by Fund Code"
  If FirstFlag = False Then
    Print #RptHandle, "Fund # "; CStr(ThisFundNum); " "; ThisFundDesc
  End If
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, "* = DO NOT DEPRECIATE THIS ASSET"
  Print #RptHandle, Tab(1); "Asset Number"; Tab(22); "Description"; Tab(51); "Fund"; Tab(58); "Life/Left"; Tab(68); "Purchase Price"; Tab(84); "Total Deprec"; Tab(102); "Book Value"
  LineCnt& = 5
  If FirstFlag = True Then
    FirstFlag = False
    LineCnt = 4
  End If
  If FundCnt > 0 Then
    Print #RptHandle, String$(111, "=")
    LineCnt = LineCnt + 1
  End If
  If fpcmbOrder.Text = "TAG NUMBER" Then
    Print #RptHandle, String$(111, "=")
    LineCnt = LineCnt + 1
  End If
  Return

PrintMasterValueEnding1:

  Page = Page + 1
  Print #RptHandle, FF$
  Print #RptHandle, Tab(30); "Master Asset Listing : Grand Totals"
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  If fptxtFundNum.Text = "ALL" Then
    Print #RptHandle, "Asset Fund: ALL"
  Else
    Print #RptHandle, "Asset Fund: " + ThisFundNum + "  " + ThisFundDesc
  End If
  Print #RptHandle, Tab(18); "Total Items"; Tab(47); "Purchase Price"; Tab(63); "Total Deprec"; Tab(79); "Book Value"
  Print #RptHandle, String$(88, "=")
  Print #RptHandle, "Total Assets ";
  Print #RptHandle, Tab(21); TotalItems;
  Print #RptHandle, Tab(47); Using("$##,###,##0.00", CStr(OrigCost#));
  Print #RptHandle, Tab(61); Using("$##,###,##0.00", CStr(YDep#));
  Print #RptHandle, Tab(75); Using("$##,###,##0.00", CStr(BookTotal#))
  Print #RptHandle, FF$

  Return

PrintDeptTotals:

  Page = Page + 1
  Print #RptHandle, FF$
  Print #RptHandle, Tab(30); "Master Asset Listing : Asset Fund Totals"
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, Tab(4); "Fund"; Tab(15); "Description"; Tab(40); "Item Count"; Tab(68); "Purchase Price"; Tab(85); "Total Deprec"; Tab(102); "Book Value"
  Print #RptHandle, String$(111, "=")
  LineCnt = 5

  If fptxtFundNum.Text = "ALL" Then
    For x = 1 To FundIdxCnt
      Print #RptHandle, Tab(3); Using("####0", FundNum(x)); Tab(15); FundDesc(x); Tab(40); Using("#####0", ATagCCnt(x)); Tab(68); Using("$##,###,##0.00", CStr(ATagCOrigCost(x))); Tab(83); Using("$##,###,##0.00", CStr(ATagCYDep(x))); Tab(98); Using("$##,###,##0.00", CStr(ATagCBookTotal(x)))
      LineCnt = LineCnt + 1
  
      If LineCnt& >= MaxLines And x <> FundIdxCnt Then
        LineCnt& = 0
        Page = Page + 1
        Print #RptHandle, FF$
        Print #RptHandle, Tab(20); "Master Asset Listing : Asset Fund Totals"
        Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
        Print #RptHandle, Tab(1); "Fund"; Tab(15); "Description"; Tab(40); "Item Count"; Tab(69); "Purchase Price"; Tab(85); "Total Deprec"; Tab(101); "Book Value"
        Print #RptHandle, String$(111, "=")
        LineCnt = LineCnt + 5
      End If
    Next x
  Else
    Print #RptHandle, Tab(3); Using("####0", ThisFundNum); Tab(15); ThisFundDesc; Tab(40); Using("#####0", ATagCCnt(1)); Tab(68); Using("$##,###,##0.00", CStr(ATagCOrigCost(1))); Tab(83); Using("$##,###,##0.00", CStr(ATagCYDep(1))); Tab(98); Using("$##,###,##0.00", CStr(ATagCBookTotal(1)))
    LineCnt = LineCnt + 1
  End If
  
  
  If LineCnt <= 53 Then
    Print #RptHandle, String$(111, "=")
    Print #RptHandle, "Total Assets ";
    Print #RptHandle, Tab(40); Using("#####0", TotalItems);
    Print #RptHandle, Tab(68); Using("$##,###,##0.00", CStr(OrigCost#));
    Print #RptHandle, Tab(83); Using("$##,###,##0.00", CStr(YDep#));
    Print #RptHandle, Tab(98); Using("$##,###,##0.00", CStr(BookTotal#))
  Else
    Print #RptHandle, FF$
    Print #RptHandle, Tab(30); "Master Asset Listing : Asset Fund Totals"
    Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
    Print #RptHandle, Tab(1); "Department"; Tab(15); "Description"; Tab(40); "Item Count"; Tab(63); "Purchase Price"; Tab(80); "Total Deprec"; Tab(97); "Book Value"
    Print #RptHandle, String$(111, "=")
    Print #RptHandle, String$(111, "=")
    Print #RptHandle, "Total Assets ";
    Print #RptHandle, Tab(40); Using("#####0", TotalItems);
    Print #RptHandle, Tab(68); Using("$##,###,##0.00", CStr(OrigCost#));
    Print #RptHandle, Tab(83); Using("$##,###,##0.00", CStr(YDep#));
    Print #RptHandle, Tab(98); Using("$##,###,##0.00", CStr(BookTotal#))
  End If
  Print #RptHandle, FF$
  TagPrint = True

  Return
        
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmAssByFundRpt", "PrintText", Erl)
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
    Unload Me

End Sub

Private Sub fptxtFundNum_Change()
  'ALL is the default value
  If fptxtFundNum.Text = "" Then
    fptxtFundNum = "ALL"
  End If
End Sub

Private Sub PrintGraphics()
  Dim ReportFile$
  Dim Dash80$
  Dim FF$
  Dim MaxLines As Integer
  Dim ItemCnt&
  Dim LineCnt&
  Dim Fund$
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim x As Integer
  Dim Nextx As Integer
  Dim Index$
  Dim Page As Integer
  Dim FundIdxHandle As Integer
  Dim FundIdxCnt As Integer
  Dim FundIdxRec As FundNumbSortIdxType
  Dim FundNumber As Integer
  Dim FundNumDesc$
  Dim FundHeader$
  Dim FundCnt As Integer
  Dim FirstFlag As Boolean
  Dim TagFlag As Boolean
  Dim DataFlag As Boolean
  Dim OrigCost As Double
  Dim BookTotal As Double
  Dim YDep As Double
  Dim YTDDep As Double
  Dim COrigCost As Double
  Dim CBookTotal As Double
  Dim CYDep As Double
  Dim LifeLeft$
  Dim WholeLife$
  Dim LifeData$
  Dim TotalItems As Integer
  Dim TagPrint As Boolean
  Dim FundRec As FAFundCodeType
  Dim FundHandle As Integer
  Dim NumOfFundRecs As Integer
  Dim ThisFundDesc$
  Dim ThisFundNum As Integer
  Dim Employer$
  Dim FASHandle As Integer
  Dim FASetUpRec As FASetupRecType
  Dim dlm$
  Dim SubReportFile$
  Dim SubRptHandle As Integer
  Dim EndRpt As Integer
  Dim TagRptHandle As Integer
  Dim TagReportFile$
  Dim ItemTotal As Long
  Dim DeptNum As Integer
  Dim DeptDesc$
  Dim DeptRec As FADeptCodeType
  Dim DeptHandle As Integer
  Dim NumOfDepts As Integer
  
  On Error GoTo ERRORSTUFF
        
  DeptNum = 0
  If QPTrim$(fptxtDeptNum.Text) <> "ALL" Then
    If Check4ValidDept = True Then
      DeptNum = CInt(fptxtDeptNum.Text)
      OpenFADeptCodeFile DeptHandle
      NumOfDepts = LOF(DeptHandle) / Len(DeptRec)
      For x = 1 To NumOfDepts
        Get DeptHandle, x, DeptRec
        If DeptRec.DeptNum = DeptNum Then
          DeptDesc$ = QPTrim$(DeptRec.DeptDesc)
          Exit For
        End If
      Next x
      Close DeptHandle
    Else
      Exit Sub
    End If
  End If

  dlm = "~"
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle

  Employer = FASetUpRec.TownName

  FirstFlag = True

  If Check4ValidFund = False Then Exit Sub

  ReportFile$ = "FARPTS\FABYFUND.RPT"
  TagReportFile$ = "FARPTS\FATAGFUND.RPT"
  SubReportFile$ = "FARPTS\FASUBFUND.RPT"

  ItemCnt& = 0
  Fund$ = QPTrim$(fptxtFundNum.Text)
  Index$ = QPTrim$(fpcmbOrder.Text)

  If Index$ = "TAG NUMBER" Then
    TagRptHandle = FreeFile
    Open TagReportFile$ For Output As #TagRptHandle
  Else
    RptHandle = FreeFile
    Open ReportFile$ For Output As #RptHandle
  End If

  OpenTagIdxFile TagIdxHandle
  NumOfFARecs = LOF(TagIdxHandle) \ Len(TagIdx)
  If NumOfFARecs = 0 Then
    MsgBox "No item records on file."
    Close TagIdxHandle
    Exit Sub
  End If

  ReDim TagIdxRecs(1 To NumOfFARecs)
  For x = 1 To NumOfFARecs
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum
  Next x
  Close TagIdxHandle

  OpenFundIdxFile FundIdxHandle
  FundIdxCnt = LOF(FundIdxHandle) \ Len(FundIdxRec)
  ReDim FundRecNum(1 To FundIdxCnt) As Integer
  For x = 1 To FundIdxCnt
    Get FundIdxHandle, x, FundIdxRec
    FundRecNum(x) = FundIdxRec.FundRecNum
  Next x
  Close FundIdxHandle

  OpenFAFundCodeFile FundHandle
  NumOfFundRecs = LOF(FundHandle) / Len(FundRec)
  If NumOfFundRecs <> FundIdxCnt Then
    Call CreateFundIdx
  End If

  If NumOfFundRecs <> FundIdxCnt Then
    MsgBox "Error: The number of funds saved is not the same number as the number of funds indexed. Please go to Asset Fund Maintenance and resave any fund code to reindex."
    Close
    Exit Sub
  End If

  ReDim FundNum(1 To FundIdxCnt) As Integer
  ReDim FundDesc(1 To FundIdxCnt) As String
  For x = 1 To FundIdxCnt
    Get FundHandle, FundRecNum(x), FundRec
      FundNum(x) = FundRec.FundNum
      FundDesc(x) = QPTrim$(FundRec.FundDesc)
  Next x
  Close FundHandle

  If Fund$ <> "ALL" Then
    ThisFundNum = Val(fptxtFundNum.Text)
    For x = 1 To FundIdxCnt
      If ThisFundNum = FundNum(x) Then
        ThisFundDesc = QPTrim$(FundDesc(x))
        Exit For
      End If
    Next x
  Else
    ThisFundNum = FundNum(1)
    ThisFundDesc = QPTrim(FundDesc(1))
  End If

  ReDim ATagCOrigCost(1 To FundIdxCnt) As Double
  ReDim ATagCBookTotal(1 To FundIdxCnt) As Double
  ReDim ATagCYDep(1 To FundIdxCnt) As Double
  ReDim ATagCCnt(1 To FundIdxCnt) As Long

  OpenFAItemFile FAHandle

  TagFlag = False

  frmFAShowPctComp.Label1 = "Gathering Asset Fund Data"
  frmFAShowPctComp.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False

GetTagTotals: 'to print the report for tag numbers only the report
  'runs through all items one time in tag order and prints pertinent
  'data...then it returns to this spot and runs back through the
  'following loop gathering department totals (just like the
  'DEPARTMENT NUMBERS option does) but does not print items by
  'department...it just prints department totals at the end of the
  'report
  Nextx = 1
  If TagFlag = True Then
    Index = "ASSET FUND NUMBER"
  End If

  Do
    DataFlag = False
    For cnt& = 1 To NumOfFARecs
      Get FAHandle, TagIdxRecs(cnt), FAItemRec
      
      If fpcmbYN.Text = "N" Then
        If FAItemRec.DsplFlag = 2 Then GoTo SkipEm1
      End If

      YTDDep# = FAItemRec.DEP2DATE

      'TAG NUMBER is actually true (if selected by the user) for the first
      'complete iteration only so that all tag numbers can be printed in numeric
      'order. After that we do not want the tags itemized we only want department
      'totals figured so that at the end of the report a "totals by dept" section
      'can be printed.
      If QPTrim$(Index) = "TAG NUMBER" Then
        GoTo TagOnly1 'prints itemized tag data in numeric order
      ElseIf ThisFundNum <> FAItemRec.FundNum Then 'any time department data
      'is needed then the item falls into this part of the if statement
        GoTo SkipEm1 'if the prevailing department (in the numeric index)
        'doesn't match this item's dept number then we don'y want it now
      End If
      'at this point the item's dept matches the prevailing dept number
      'but since TagFlag (don't want itemized tag data anymore) is true then
      'go ahead and collect dept data
      If DeptNum <> 0 Then
        If FAItemRec.IDEPT <> DeptNum Then
          GoTo SkipEm1
        End If
      End If

      If TagFlag = True Then GoTo TagOnly2

TagOnly1: 'printing valid tag data...skipped if TAG NUMBER is chosen and this is
          'not the first iteration
      DataFlag = True
      If TagRptHandle > 0 Then
        '                     0
        Print #TagRptHandle, Employer; dlm;
        '                          1                        2
        Print #TagRptHandle, QPTrim$(FAItemRec.ItemTag); dlm; QPTrim$(FAItemRec.IDESC1); dlm;
        '                        3
        Print #TagRptHandle, CStr(FAItemRec.FundNum); dlm;
        '                        4
        Print #TagRptHandle, FAItemRec.ILIFE; dlm;
        '                        5
        Print #TagRptHandle, FAItemRec.ORGCOST; dlm;
        If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
          '                       6           7
          Print #TagRptHandle, YTDDep#; dlm; "*"; dlm;
        Else
          '                       6           7
          Print #TagRptHandle, YTDDep#; dlm; " "; dlm;
        End If
        '                          8                          9
        Print #TagRptHandle, FAItemRec.CURRVAL; dlm; FAItemRec.LifeLeft; dlm;
        If DsplYNFlag = False Then
           If FAItemRec.DsplFlag = 1 Then
             '                            10
             Print #TagRptHandle, "P" + MakeRegDate(FAItemRec.DispDate); dlm; fpcmbYN.Text
           Else
             '                 10
             Print #TagRptHandle, ""; dlm; fpcmbYN.Text
           End If
        Else
          If FAItemRec.DsplFlag = 2 Then
            '                            10
            Print #TagRptHandle, MakeRegDate(FAItemRec.DispDate); dlm; fpcmbYN.Text
          ElseIf FAItemRec.DsplFlag = 1 Then
            '                            10
            Print #TagRptHandle, "P" + MakeRegDate(FAItemRec.DispDate); dlm; fpcmbYN.Text
          Else
            '                    10
            Print #TagRptHandle, ""; dlm; fpcmbYN.Text
          End If
        End If
      Else
        '                     0
        Print #RptHandle, Employer; dlm;
        '                          1                        2
        Print #RptHandle, QPTrim$(FAItemRec.ItemTag); dlm; QPTrim$(FAItemRec.IDESC1); dlm;
        '                        3
        Print #RptHandle, CStr(FAItemRec.FundNum); dlm;
        '                        4
        Print #RptHandle, FAItemRec.ILIFE; dlm;
        '                        5
        Print #RptHandle, FAItemRec.ORGCOST; dlm;
        If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
          '                    6           7
          Print #RptHandle, YTDDep#; dlm; "*"; dlm;
        Else
          '                    6           7
          Print #RptHandle, YTDDep#; dlm; " "; dlm;
        End If
        '                         8
        Print #RptHandle, FAItemRec.CURRVAL; dlm;
        '                  9               10                  11
        Print #RptHandle, Fund$; dlm; ThisFundDesc; dlm; ThisFundNum; dlm;
        '                      12                13          14              15
        Print #RptHandle, COrigCost#; dlm; CYDep#; dlm; CBookTotal#; dlm; CStr(FundCnt); dlm;
        '                      16          17             18               19
        Print #RptHandle, OrigCost#; dlm; YDep#; dlm; BookTotal#; dlm; TotalItems; dlm;
        '                         20                   21
        Print #RptHandle, FAItemRec.DEPYN; dlm; FAItemRec.LifeLeft; dlm;
        If DsplYNFlag = False Then
          If FAItemRec.DsplFlag = 1 Then
            '                            22                                    23
            Print #RptHandle, "P" + MakeRegDate(FAItemRec.DispDate); dlm; fpcmbYN.Text
          Else
            '                 22          23
            Print #RptHandle, ""; dlm; fpcmbYN.Text
          End If
        Else
          If FAItemRec.DsplFlag = 2 Then
            '                            22                              23
            Print #RptHandle, MakeRegDate(FAItemRec.DispDate); dlm; fpcmbYN.Text
          ElseIf FAItemRec.DsplFlag = 1 Then
            '                            22                                    23
            Print #RptHandle, "P" + MakeRegDate(FAItemRec.DispDate); dlm; fpcmbYN.Text
          Else
            '                 22          23
            Print #RptHandle, ""; dlm; fpcmbYN.Text
          End If
        End If
      End If

      ItemCnt& = ItemCnt& + 1
      ItemTotal = ItemTotal + 1
TagOnly2: 'collects data for each department for reporting totals

      'This if statement filters out the first iteration of TAG NUMBER
      'selection because we do not want to start accumulating dept data
      'until the second iteration
      If TagFlag = False And QPTrim$(Index) = "TAG NUMBER" Then GoTo SkipEm1

      'collects grand totals
      OrigCost# = OrigCost# + FAItemRec.ORGCOST
      BookTotal# = BookTotal# + (FAItemRec.CURRVAL)
      YDep# = YDep# + YTDDep#

      'collects dept totals
      FundCnt = FundCnt + 1
      ATagCCnt(Nextx) = FundCnt
      TotalItems = TotalItems + 1
      COrigCost# = COrigCost# + FAItemRec.ORGCOST
      ATagCOrigCost(Nextx) = COrigCost#
      CBookTotal# = CBookTotal# + (FAItemRec.CURRVAL)
      ATagCBookTotal(Nextx) = CBookTotal#
      CYDep# = CYDep# + YTDDep#
      ATagCYDep(Nextx) = CYDep#
SkipEm1:

    Next cnt&
    'here we begin the iteration over again but this time TagFlag
    'becomes true so we know that this was originally TAG NUMBERS
    'and the first iteration is done
    If QPTrim$(Index) = "TAG NUMBER" And TagFlag = False Then
      TagFlag = True
      GoTo GetTagTotals
      Exit Do
    End If

    If TagFlag = True Then GoTo NoData 'don't want the next dept
    'data to print

NoData:
    frmFAShowPctComp.ShowPctComp Nextx, FundIdxCnt ' + 1
    If frmFAShowPctComp.Out = True Then
      Close
      frmFAShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdExit.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload frmFAShowPctComp
      Exit Sub
    End If
    'if "ALL" is not selected then the user has selected a single
    'department so we have all the data we need at this point...exit
    If QPTrim$(Fund$) <> "ALL" Then Exit Do
    'if all the depts have been examined then time to go
    If Nextx = FundIdxCnt Then Exit Do
    'move to the next dept
    Nextx = Nextx + 1
    'assign new dept to DeptNumber
    ThisFundNum = FundNum(Nextx)
    ThisFundDesc = QPTrim$(FundDesc(Nextx))
    'clear all dept totals
    COrigCost# = 0
    CBookTotal# = 0
    CYDep# = 0
    FundCnt = 0
   Loop

  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True

  If ItemTotal = 0 Then
    MsgBox "There are no fixed assets saved for this criteria."
    Close
    Exit Sub
  End If

  'only prints if TAG NUMBERS was selected
  Close         'Close all open files now
  GoSub PrintTagFundTotals

  If QPTrim(fpcmbOrder.Text) = "TAG NUMBER" Then
    arFATagByFund.Show
  Else
    arFAAssetsByFundRpt.Show
  End If

  frmFALoadReport.Show

  Exit Sub

PrintTagFundTotals: 'print only if TAG NUMBERS was selected

  SubRptHandle = FreeFile
  Open SubReportFile$ For Output As SubRptHandle
  EndRpt = 1
  If fptxtFundNum = "ALL" Then
    For x = 1 To FundIdxCnt
      '                        0                1                2                   3                      4                    5
      Print #SubRptHandle, CStr(FundNum(x)); dlm; QPTrim$(FundDesc(x)); dlm; ATagCCnt(x); dlm; ATagCOrigCost(x); dlm; ATagCYDep(x); dlm; ATagCBookTotal(x); dlm;
      '                        6                7                8                9
      Print #SubRptHandle, Using$("######", CStr(TotalItems)); dlm; OrigCost#; dlm; YDep#; dlm; BookTotal#; dlm; EndRpt
    Next x
  Else
      '                        0                1                2                   3                      4                    5
      Print #SubRptHandle, ThisFundNum; dlm; ThisFundDesc; dlm; ATagCCnt(1); dlm; ATagCOrigCost(1); dlm; ATagCYDep(1); dlm; ATagCBookTotal(1); dlm;
      '                        6                7                8                9
      Print #SubRptHandle, TotalItems; dlm; OrigCost#; dlm; YDep#; dlm; BookTotal#; dlm; EndRpt
  End If
  Close SubRptHandle

  Return
        
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAAssByFundRpt", "PrintGraphics", Erl)
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
    Unload Me

End Sub

Private Sub fptxtDeptNum_Change()
  If fptxtDeptNum.Text = "" Then
    fptxtDeptNum.Text = "ALL"
  End If

End Sub

Private Function Check4ValidDept() As Boolean
  Dim x As Integer
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxRecNums As Integer
  Dim ThisDept$
  
  On Error GoTo ERRORSTUFF
        
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAAssByFundRpt", "Check4ValidDept", Erl)
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
    Unload Me
End Function

Private Sub PrintDeptGraphics()
  Dim ReportFile$
  Dim Dash80$
  Dim FF$
  Dim MaxLines As Integer
  Dim ItemCnt&
  Dim LineCnt&
  Dim Fund$
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim x As Integer, Y As Integer
  Dim Nextx As Integer
  Dim Index$
  Dim Page As Integer
  Dim FundIdxHandle As Integer
  Dim FundIdxCnt As Integer
  Dim FundIdxRec As FundNumbSortIdxType
  Dim FundNumber As Integer
  Dim FundNumDesc$
  Dim FundHeader$
  Dim FundCnt As Integer
  Dim FirstFlag As Boolean
  Dim TagFlag As Boolean
  Dim DataFlag As Boolean
  Dim OrigCost As Double
  Dim BookTotal As Double
  Dim YDep As Double
  Dim YTDDep As Double
  Dim COrigCost As Double
  Dim CBookTotal As Double
  Dim CYDep As Double
  Dim LifeLeft$
  Dim WholeLife$
  Dim LifeData$
  Dim TotalItems As Integer
  Dim TagPrint As Boolean
  Dim FundRec As FAFundCodeType
  Dim FundHandle As Integer
  Dim NumOfFundRecs As Integer
  Dim ThisFundDesc$
  Dim ThisFundNum As Integer
  Dim Employer$
  Dim FASHandle As Integer
  Dim FASetUpRec As FASetupRecType
  Dim dlm$
  Dim SubReportFile$
  Dim Sub2ReportFile$
  Dim SubRptHandle As Integer
  Dim Sub2RptHandle As Integer
  Dim EndRpt As Integer
  Dim TagRptHandle As Integer
  Dim TagReportFile$
  Dim ItemTotal As Long
  Dim DeptRec As FADeptCodeType
  Dim DeptHandle As Integer
  Dim NumOfDepts As Integer
  Dim DeptIdxRec As DeptNumbSortIdxType
  Dim DeptIdxHandle As Integer
  Dim FundIdx As Integer 'added 8/14/07
  
  On Error GoTo ERRORSTUFF
  frmFALoadReport.Show
  DoEvents
  
  If fptxtDeptNum.Enabled = True Then
    If QPTrim$(fptxtDeptNum.Text) <> "ALL" Then
      If Check4ValidDept = True Then
        OpenDeptIdxFile DeptIdxHandle
        NumOfDepts = LOF(DeptIdxHandle) / Len(DeptIdxRec)
        If NumOfDepts = 0 Then
          Close DeptIdxHandle
          GoTo NoDepts
        End If
        ReDim DeptDesc(1 To 1) As String
        ReDim DeptNum(1 To 1) As Integer
        ReDim DeptRecNum(1 To 1) As Integer
        DeptNum(1) = CInt(fptxtDeptNum.Text)
        For x = 1 To NumOfDepts
          Get DeptIdxHandle, x, DeptIdxRec
          If DeptIdxRec.DeptNumb = DeptNum(1) Then
            DeptDesc(1) = QPTrim$(DeptIdxRec.DeptIdxDesc)
            DeptRecNum(1) = DeptIdxRec.DeptRecNum
            Exit For
          End If
        Next x
        Close DeptIdxHandle
      Else
        MsgBox "The department number entered does not match any on file. Please try again."
        Close
        Exit Sub
      End If
    Else 'does = "ALL"
      OpenDeptIdxFile DeptIdxHandle
      NumOfDepts = LOF(DeptIdxHandle) / Len(DeptIdxRec)
      If NumOfDepts = 0 Then
        Close DeptIdxHandle
        GoTo NoDepts
      End If
      ReDim DeptDesc(1 To NumOfDepts) As String
      ReDim DeptNum(1 To NumOfDepts) As Integer
      ReDim DeptRecNum(1 To NumOfDepts) As Integer
      For x = 1 To NumOfDepts
        Get DeptIdxHandle, x, DeptIdxRec
        DeptDesc(x) = QPTrim$(DeptIdxRec.DeptIdxDesc)
        DeptRecNum(x) = DeptIdxRec.DeptRecNum
        DeptNum(x) = DeptIdxRec.DeptNumb
      Next x
      Close DeptIdxHandle
    End If
  End If
  
NoDepts:
  dlm = "~"
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle

  Employer = FASetUpRec.TownName

  FirstFlag = True

  If Check4ValidFund = False Then Exit Sub

  ReportFile$ = "FARPTS\FADPTFND.RPT"
  SubReportFile$ = "FARPTS\FASUBFUND.RPT"
  Sub2ReportFile$ = "FARPTS\FASUB2FND.RPT"

  ItemCnt& = 0
  Fund$ = QPTrim$(fptxtFundNum.Text)
  Index$ = QPTrim$(fpcmbOrder.Text)

  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  OpenTagIdxFile TagIdxHandle
  NumOfFARecs = LOF(TagIdxHandle) \ Len(TagIdx)
  If NumOfFARecs = 0 Then
    MsgBox "No item records on file."
    Close TagIdxHandle
    Exit Sub
  End If

  ReDim TagIdxRecs(1 To NumOfFARecs)
  For x = 1 To NumOfFARecs
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum
  Next x
  Close TagIdxHandle

  OpenFundIdxFile FundIdxHandle
  FundIdxCnt = LOF(FundIdxHandle) \ Len(FundIdxRec)
  ReDim FundRecNum(1 To FundIdxCnt) As Integer
  For x = 1 To FundIdxCnt
    Get FundIdxHandle, x, FundIdxRec
    FundRecNum(x) = FundIdxRec.FundRecNum
  Next x
  Close FundIdxHandle

  OpenFAFundCodeFile FundHandle
  NumOfFundRecs = LOF(FundHandle) / Len(FundRec)
  If NumOfFundRecs <> FundIdxCnt Then
    Call CreateFundIdx
  End If

  If NumOfFundRecs <> FundIdxCnt Then
    MsgBox "Error: The number of funds saved is not the same number as the number of funds indexed. Please go to Asset Fund Maintenance and resave any fund code to reindex."
    Close
    Exit Sub
  End If

  ReDim FundNum(1 To FundIdxCnt) As Integer
  ReDim FundDesc(1 To FundIdxCnt) As String
  For x = 1 To FundIdxCnt
    Get FundHandle, FundRecNum(x), FundRec
      FundNum(x) = FundRec.FundNum
      FundDesc(x) = QPTrim$(FundRec.FundDesc)
  Next x
  Close FundHandle

  If Fund$ <> "ALL" Then
    ThisFundNum = Val(fptxtFundNum.Text)
    For x = 1 To FundIdxCnt
      If ThisFundNum = FundNum(x) Then
        ThisFundDesc = QPTrim$(FundDesc(x))
        FundIdx = x 'added 8/14/07
        Exit For
      End If
    Next x
  Else
    ThisFundNum = FundNum(1)
    ThisFundDesc = QPTrim(FundDesc(1))
  End If

  ReDim ATagCOrigCost(1 To FundIdxCnt) As Double
  ReDim ATagCBookTotal(1 To FundIdxCnt) As Double
  ReDim ATagCYDep(1 To FundIdxCnt) As Double
  ReDim ATagCCnt(1 To FundIdxCnt) As Long

  OpenFAItemFile FAHandle

  TagFlag = False

  frmFAShowPctComp.Label1 = "Gathering Asset Fund Data"
  frmFAShowPctComp.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False

  Nextx = 1
  If QPTrim$(fptxtDeptNum.Text) <> "ALL" Then NumOfDepts = 1
  Unload frmFALoadReport
  
  ReDim FundDeptCnt(1 To FundIdxCnt, 1 To NumOfDepts) As Integer
  ReDim FundDeptPurPrice(1 To FundIdxCnt, 1 To NumOfDepts) As Double
  ReDim FundDeptDpr2Date(1 To FundIdxCnt, 1 To NumOfDepts) As Double
  ReDim FundDeptCurValue(1 To FundIdxCnt, 1 To NumOfDepts) As Double
  
  Do
    DataFlag = False
    For x = 1 To NumOfDepts
      For cnt& = 1 To NumOfFARecs
        Get FAHandle, TagIdxRecs(cnt), FAItemRec
        If fpcmbYN.Text = "N" Then
          If FAItemRec.DsplFlag = 2 Then GoTo SkipEm1Dept
        End If

        YTDDep# = FAItemRec.DEP2DATE

        If ThisFundNum <> FAItemRec.FundNum Then 'any time department data
        'is needed then the item falls into this part of the if statement
          GoTo SkipEm1Dept 'if the prevailing department (in the numeric index)
          'doesn't match this item's dept number then we don'y want it now
        End If

        If QPTrim$(fptxtDeptNum.Text) = "ALL" Then
          If FAItemRec.IDEPT <> DeptNum(x) Then
            GoTo SkipEm1Dept
          End If
        Else
          If FAItemRec.IDEPT <> DeptNum(1) Then
            GoTo SkipEm1Dept
          End If
        End If
        '                     0
        Print #RptHandle, Employer; dlm;
        '                          1                        2
        Print #RptHandle, QPTrim$(FAItemRec.ItemTag); dlm; QPTrim$(FAItemRec.IDESC1); dlm;
        '                        3
        Print #RptHandle, CStr(FAItemRec.FundNum); dlm;
        '                        4
        Print #RptHandle, FAItemRec.ILIFE; dlm;
        '                        5
        Print #RptHandle, FAItemRec.ORGCOST; dlm;
        If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
          '                    6           7
          Print #RptHandle, YTDDep#; dlm; "*"; dlm;
        Else
          '                    6           7
          Print #RptHandle, YTDDep#; dlm; " "; dlm;
        End If
        '                         8
        Print #RptHandle, FAItemRec.CURRVAL; dlm;
        '                  9               10                  11
        Print #RptHandle, Fund$; dlm; ThisFundDesc; dlm; ThisFundNum; dlm;
        '                      12                13          14              15
        Print #RptHandle, COrigCost#; dlm; CYDep#; dlm; CBookTotal#; dlm; CStr(FundCnt); dlm;
        '                      16          17             18               19
        Print #RptHandle, OrigCost#; dlm; YDep#; dlm; BookTotal#; dlm; TotalItems; dlm;
        '                         20                   21
        Print #RptHandle, FAItemRec.DEPYN; dlm; FAItemRec.LifeLeft; dlm;
        If DsplYNFlag = False Then
          If QPTrim$(fptxtDeptNum.Text) = "ALL" Then
            '                 22          23                   24                       25
            Print #RptHandle, ""; dlm; fpcmbYN.Text; dlm; CStr(DeptNum(x)); dlm; QPTrim$(DeptDesc$(x))
          Else
            '                 22          23                   24                       25
            Print #RptHandle, ""; dlm; fpcmbYN.Text; dlm; CStr(DeptNum(1)); dlm; QPTrim$(DeptDesc$(1))
          End If
        Else
          If FAItemRec.DsplFlag = 2 Then
            If QPTrim$(fptxtDeptNum.Text) = "ALL" Then
              '                            22                              23                  24                    25
              Print #RptHandle, MakeRegDate(FAItemRec.DispDate); dlm; fpcmbYN.Text; dlm; CStr(DeptNum(x)); dlm; QPTrim$(DeptDesc$(x))
            Else
              '                            22                              23                  24                    25
              Print #RptHandle, MakeRegDate(FAItemRec.DispDate); dlm; fpcmbYN.Text; dlm; CStr(DeptNum(1)); dlm; QPTrim$(DeptDesc$(1))
            End If
          ElseIf FAItemRec.DsplFlag = 1 Then
            If QPTrim$(fptxtDeptNum.Text) = "ALL" Then
              '                            22                                    23                 24                   25
              Print #RptHandle, "P" + MakeRegDate(FAItemRec.DispDate); dlm; fpcmbYN.Text; dlm; CStr(DeptNum(x)); dlm; QPTrim$(DeptDesc$(x))
            Else
              '                            22                                    23                 24                   25
              Print #RptHandle, "P" + MakeRegDate(FAItemRec.DispDate); dlm; fpcmbYN.Text; dlm; CStr(DeptNum(1)); dlm; QPTrim$(DeptDesc$(1))
            End If
          Else
            If QPTrim$(fptxtDeptNum.Text) = "ALL" Then
              '                 22          23                   24                     25
              Print #RptHandle, ""; dlm; fpcmbYN.Text; dlm; CStr(DeptNum(x)); dlm; QPTrim$(DeptDesc$(x))
            Else
              '                 22          23                   24                     25
              Print #RptHandle, ""; dlm; fpcmbYN.Text; dlm; CStr(DeptNum(1)); dlm; QPTrim$(DeptDesc$(1))
            End If
          End If
        End If
        
        FundDeptCnt(Nextx, x) = FundDeptCnt(Nextx, x) + 1
        ItemCnt& = ItemCnt& + 1
        ItemTotal = ItemTotal + 1
        
        FundDeptPurPrice(Nextx, x) = FundDeptPurPrice(Nextx, x) + FAItemRec.ORGCOST
        FundDeptDpr2Date(Nextx, x) = FundDeptDpr2Date(Nextx, x) + YTDDep
        FundDeptCurValue(Nextx, x) = FundDeptCurValue(Nextx, x) + FAItemRec.CURRVAL
        
        OrigCost# = OrigCost# + FAItemRec.ORGCOST
        BookTotal# = BookTotal# + (FAItemRec.CURRVAL)
        YDep# = YDep# + YTDDep#

        'collects dept totals
        FundCnt = FundCnt + 1
        TotalItems = TotalItems + 1
        COrigCost# = COrigCost# + FAItemRec.ORGCOST
        CBookTotal# = CBookTotal# + (FAItemRec.CURRVAL)
        CYDep# = CYDep# + YTDDep#
        If Fund$ <> "ALL" Then 'added 8/14/07
          ATagCCnt(FundIdx) = FundCnt
          ATagCOrigCost(FundIdx) = COrigCost#
          ATagCBookTotal(FundIdx) = CBookTotal#
          ATagCYDep(FundIdx) = CYDep#
        Else
          ATagCCnt(Nextx) = FundCnt
          ATagCOrigCost(Nextx) = COrigCost#
          ATagCBookTotal(Nextx) = CBookTotal#
          ATagCYDep(Nextx) = CYDep#
        End If
        
SkipEm1Dept:
      Next cnt&
    Next x

NoDataDept:
    frmFAShowPctComp.ShowPctComp Nextx, FundIdxCnt ' + 1
    If frmFAShowPctComp.Out = True Then
      Close
      frmFAShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdExit.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload frmFAShowPctComp
      Exit Sub
    End If
    'if "ALL" is not selected then the user has selected a single
    'department so we have all the data we need at this point...exit
    If QPTrim$(Fund$) <> "ALL" Then Exit Do
    'if all the depts have been examined then time to go
    If Nextx = FundIdxCnt Then Exit Do
    'move to the next dept
    Nextx = Nextx + 1
    'assign new dept to DeptNumber
    ThisFundNum = FundNum(Nextx)
    ThisFundDesc = QPTrim$(FundDesc(Nextx))
    'clear all dept totals
    COrigCost# = 0
    CBookTotal# = 0
    CYDep# = 0
    FundCnt = 0
  Loop 'get the next fund number
  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True

  If ItemTotal = 0 Then
    MsgBox "There are no fixed assets saved for this criteria."
    Close
    Exit Sub
  End If

  Close
  GoSub PrintSummary
  If QPTrim(fpcmbOrder.Text) = "TAG NUMBER" Then
    arFATagByFund.Show
  Else
    arFAFundAndDeptRpt.Show
  End If

  frmFALoadReport.Show

  Exit Sub

PrintSummary:
  SubRptHandle = FreeFile
  Open SubReportFile$ For Output As #SubRptHandle
  For x = 1 To FundIdxCnt
    '                              0                      1                    2                    3                   4                    5
    Print #SubRptHandle, QPTrim$(FundDesc(x)); dlm; CStr(FundNum(x)); dlm; ATagCCnt(x); dlm; ATagCOrigCost(x); dlm; ATagCYDep(x); dlm; ATagCBookTotal(x)
  Next x
  Close
  
  Sub2RptHandle = FreeFile
  Open Sub2ReportFile$ For Output As #Sub2RptHandle
  For x = 1 To FundIdxCnt
    For Y = 1 To NumOfDepts
      '                                0                     1                     2                3                 4                          5                          6                              7
      Print #Sub2RptHandle, QPTrim$(FundDesc(x)); dlm; CStr(FundNum(x)); dlm; DeptDesc(Y); dlm; DeptNum(Y); dlm; FundDeptCnt(x, Y); dlm; FundDeptPurPrice(x, Y); dlm; FundDeptDpr2Date(x, Y); dlm; FundDeptCurValue(x, Y)
    Next Y
  Next x
  Close
  
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAAssByFundRpt", "PrintDeptGraphics", Erl)
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
    Unload Me
End Sub

Private Sub PrintDeptText()
  Dim ReportFile$
  Dim Dash80$
  Dim FF$
  Dim MaxLines As Integer
  Dim ItemCnt&
  Dim LineCnt&
  Dim code$
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim x As Integer, Y As Integer, z As Integer
  Dim Fundx As Integer
  Dim Index$
  Dim Page As Integer
  Dim FundIdxHandle As Integer
  Dim FundIdxCnt As Integer
  Dim FundIdxRec As FundNumbSortIdxType
  Dim FundNumber As Integer
  Dim FundNumDesc$
  Dim FundHeader$
  Dim FundCnt As Integer
  Dim FirstFlag As Boolean
  Dim TagFlag As Boolean
  Dim DataFlag As Boolean
  Dim OrigCost As Double
  Dim BookTotal As Double
  Dim YDep As Double
  Dim YTDDep As Double
  Dim COrigCost As Double
  Dim CBookTotal As Double
  Dim CYDep As Double
  Dim LifeLeft$
  Dim WholeLife$
  Dim LifeData$
  Dim TotalItems As Integer
  Dim TagPrint As Boolean
  Dim FundRec As FAFundCodeType
  Dim FundHandle As Integer
  Dim NumOfFundRecs As Integer
  Dim ThisFundDesc$
  Dim ThisFundNum$
  Dim ItemTotal As Long
  Dim DeptRec As FADeptCodeType
  Dim DeptHandle As Integer
  Dim NumOfDepts As Integer
  Dim DeptIdxRec As DeptNumbSortIdxType
  Dim DeptIdxHandle As Integer
  Dim NextDep As Integer
  Dim ThisCnt As Integer
  Dim Thisy As Integer
  Dim PurPrice As Double
  Dim Dpr2Date As Double
  Dim BookVal As Double
  Dim FundItemCnt As Integer
  
  'If fptxtDeptNum.Text is enabled then this sub is activated
  
'  On Error GoTo ERRORSTUFF
  If QPTrim$(fptxtDeptNum.Text) <> "ALL" Then
    If Check4ValidDept = True Then
      OpenDeptIdxFile DeptIdxHandle
      NumOfDepts = LOF(DeptIdxHandle) / Len(DeptIdxRec)
      If NumOfDepts = 0 Then
        Close DeptIdxHandle
        GoTo NoDepts
      End If
      ReDim DeptDesc(1 To 1) As String
      ReDim DeptNum(1 To 1) As Integer
      ReDim DeptRecNum(1 To 1) As Integer
      DeptNum(1) = CInt(fptxtDeptNum.Text)
      For x = 1 To NumOfDepts
        Get DeptIdxHandle, x, DeptIdxRec
        If DeptIdxRec.DeptNumb = DeptNum(1) Then
          DeptDesc(1) = QPTrim$(DeptIdxRec.DeptIdxDesc)
          DeptRecNum(1) = DeptIdxRec.DeptRecNum
          Thisy = x
          Exit For
        End If
      Next x
      Close DeptIdxHandle
    Else
      MsgBox "The department number entered does not match any on file. Please try again."
      Close
      Exit Sub
    End If
  Else 'does = "ALL"
    OpenDeptIdxFile DeptIdxHandle
    NumOfDepts = LOF(DeptIdxHandle) / Len(DeptIdxRec)
    If NumOfDepts = 0 Then
      Close DeptIdxHandle
      GoTo NoDepts
    End If
    ReDim DeptDesc(1 To NumOfDepts) As String
    ReDim DeptNum(1 To NumOfDepts) As Integer
    ReDim DeptRecNum(1 To NumOfDepts) As Integer
    For x = 1 To NumOfDepts
      Get DeptIdxHandle, x, DeptIdxRec
      DeptDesc(x) = QPTrim$(DeptIdxRec.DeptIdxDesc)
      DeptRecNum(x) = DeptIdxRec.DeptRecNum
      DeptNum(x) = DeptIdxRec.DeptNumb
    Next x
    Close DeptIdxHandle
  End If
  
NoDepts:
  If Check4ValidFund = False Then Exit Sub

  ReportFile$ = "FAFNDBYD.PRN"  'Report File Name
  Dash80$ = String$(80, "=")
  FF$ = Chr$(12)

  MaxLines = 56
  LineCnt& = 0
  ItemCnt& = 0
  code$ = QPTrim$(fptxtFundNum.Text)

  RptHandle = FreeFile
  Index$ = QPTrim$(fpcmbOrder.Text)
  Open ReportFile$ For Output As #RptHandle

  OpenTagIdxFile TagIdxHandle
  NumOfFARecs = LOF(TagIdxHandle) \ Len(TagIdx)
  If NumOfFARecs = 0 Then
    MsgBox "No item records on file."
    Close TagIdxHandle
    Exit Sub
  End If
  
  ReDim TagIdxRecs(1 To NumOfFARecs)
  For x = 1 To NumOfFARecs
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum
  Next x
  Close TagIdxHandle
  
  OpenFundIdxFile FundIdxHandle
  FundIdxCnt = LOF(FundIdxHandle) \ Len(FundIdxRec)
  ReDim FundRecNum(1 To FundIdxCnt) As Integer
  For x = 1 To FundIdxCnt
    Get FundIdxHandle, x, FundIdxRec
    FundRecNum(x) = FundIdxRec.FundRecNum
  Next x
  Close FundIdxHandle
  
  OpenFAFundCodeFile FundHandle
  NumOfFundRecs = LOF(FundHandle) / Len(FundRec)
  If NumOfFundRecs <> FundIdxCnt Then
    Call CreateFundIdx
  End If
  
  If NumOfFundRecs <> FundIdxCnt Then
    MsgBox "Error: The number of fund codes saved is not the same number as the number of fund codes indexed. Please go to Fund Code Maintenance and resave any fund code to reindex."
    Close
    Exit Sub
  End If
  
  ReDim FundNum(1 To FundIdxCnt) As Integer
  ReDim FundDesc(1 To FundIdxCnt) As String
  For x = 1 To FundIdxCnt
    Get FundHandle, FundRecNum(x), FundRec
      FundNum(x) = FundRec.FundNum
      FundDesc(x) = QPTrim$(FundRec.FundDesc)
  Next x
  Close FundHandle
  
  If code$ <> "ALL" Then
    ThisFundNum = QPTrim$(fptxtFundNum.Text)
    For x = 1 To FundIdxCnt
      If ThisFundNum = FundNum(x) Then
        ThisFundDesc = QPTrim$(FundDesc(x))
        Exit For
      End If
    Next x
  Else
    ThisFundNum = FundNum(1)
    ThisFundDesc = QPTrim(FundDesc(1))
  End If

  GoSub PrintMasterHeader1

  ReDim ATagCOrigCost(1 To FundIdxCnt) As Double
  ReDim ATagCBookTotal(1 To FundIdxCnt) As Double
  ReDim ATagCYDep(1 To FundIdxCnt) As Double
  ReDim ATagCCnt(1 To FundIdxCnt) As Long
  ReDim ThisDeptOrigCost(1 To NumOfDepts) As Double
  ReDim ThisDeptBookTotal(1 To NumOfDepts) As Double
  ReDim ThisDeptDep(1 To NumOfDepts) As Double
  ReDim ThisDeptCnt(1 To NumOfDepts) As Integer
  
  OpenFAItemFile FAHandle

  ReDim ThisFundDept(1 To FundIdxCnt, 1 To NumOfDepts) As String
  
  frmFALoadReport.Show
  DoEvents
  Fundx = 1
  If fptxtDeptNum.Text = "ALL" Then
    ThisCnt = 0
    If fptxtFundNum.Text = "ALL" Then
      For x = 1 To FundIdxCnt
        For Y = 1 To NumOfDepts
          For z = 1 To NumOfFARecs
            Get FAHandle, TagIdxRecs(z), FAItemRec
            If fpcmbYN.Text = "N" And FAItemRec.DsplFlag = 2 Then GoTo DisposedOf1
            If FAItemRec.FundNum = FundNum(x) Then
              If FAItemRec.IDEPT = DeptNum(Y) Then
                ThisCnt = ThisCnt + 1
              End If
            End If
DisposedOf1:
          Next z
          If ThisCnt > 0 Then
            ThisFundDept(x, Y) = "Full"
          Else
            ThisFundDept(x, Y) = "Empty"
          End If
          ThisCnt = 0
        Next Y
      Next x
    Else 'fptxtFundNum.Text <> "ALL"
      For x = 1 To FundIdxCnt
        If FundNum(x) = Val(fptxtFundNum.Text) Then
          Fundx = x
          Exit For
        End If
      Next x
      For Y = 1 To NumOfDepts
        For z = 1 To NumOfFARecs
          Get FAHandle, TagIdxRecs(z), FAItemRec
          If fpcmbYN.Text = "N" And FAItemRec.DsplFlag = 2 Then GoTo DisposedOf2
          If FAItemRec.FundNum = FundNum(x) Then
            If FAItemRec.IDEPT = DeptNum(Y) Then
              ThisCnt = ThisCnt + 1
            End If
          End If
DisposedOf2:
        Next z
        If ThisCnt > 0 Then
          ThisFundDept(x, Y) = "Full"
        Else
          ThisFundDept(x, Y) = "Empty"
        End If
        ThisCnt = 0
      Next Y
      For x = 1 To FundIdxCnt
        For Y = 1 To NumOfDepts
          If x <> Fundx Then
            ThisFundDept(x, Y) = "EMPTY"
          End If
        Next Y
      Next x
    End If
  Else 'fptxtDeptNum.Text <> "ALL"
    If fptxtFundNum.Text = "ALL" Then
      For x = 1 To FundIdxCnt
        For Y = 1 To NumOfDepts
          ThisFundDept(x, Y) = "Empty"
        Next Y
      Next x
      For x = 1 To FundIdxCnt
        Y = 1
        For z = 1 To NumOfFARecs
          Get FAHandle, TagIdxRecs(z), FAItemRec
          If fpcmbYN.Text = "N" And FAItemRec.DsplFlag = 2 Then GoTo DisposedOf3
          If FAItemRec.FundNum = FundNum(x) Then
            If FAItemRec.IDEPT = DeptNum(1) Then
              ThisCnt = ThisCnt + 1
            End If
          End If
DisposedOf3:
        Next z
        If ThisCnt > 0 Then
          ThisFundDept(x, Y) = "Full"
        Else
          ThisFundDept(x, Y) = "Empty"
        End If
        ThisCnt = 0
      Next x
    Else
      For x = 1 To FundIdxCnt
        If FundNum(x) = Val(fptxtFundNum.Text) Then
          Fundx = x
          Exit For
        End If
      Next x
      For Y = 1 To NumOfDepts
        If DeptNum(Y) = Val(fptxtDeptNum.Text) Then
          Exit For
        End If
      Next Y
      For z = 1 To NumOfFARecs
        Get FAHandle, TagIdxRecs(z), FAItemRec
        If fpcmbYN.Text = "N" And FAItemRec.DsplFlag = 2 Then GoTo DisposedOf4
        If FAItemRec.FundNum = FundNum(x) Then
          If FAItemRec.IDEPT = DeptNum(Y) Then
            ThisCnt = ThisCnt + 1
          End If
        End If
DisposedOf4:
      Next z
      If ThisCnt > 0 Then
        ThisFundDept(x, Y) = "Full"
      Else
        ThisFundDept(x, Y) = "Empty"
      End If
      ThisCnt = 0
    End If
  End If
  Unload frmFALoadReport
  
  frmFAShowPctComp.Label1 = "Gathering Asset Fund Data"
  frmFAShowPctComp.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False
  
  ReDim FundDeptCnt(1 To FundIdxCnt, 1 To NumOfDepts) As Integer
  ReDim FundDeptPurPrice(1 To FundIdxCnt, 1 To NumOfDepts) As Double
  ReDim FundDeptDpr2Date(1 To FundIdxCnt, 1 To NumOfDepts) As Double
  ReDim FundDeptCurValue(1 To FundIdxCnt, 1 To NumOfDepts) As Double
  
  Do
    GoSub PrintFundHeader
    For x = 1 To NumOfDepts
      If fptxtDeptNum.Enabled = True And fptxtDeptNum.Text = "ALL" Then
        If ThisFundDept(Fundx, x) = "Empty" Then
          GoTo NutinHoney
        End If
      End If
      GoSub PrintDeptHeader
      For cnt& = 1 To NumOfFARecs
        Get FAHandle, TagIdxRecs(cnt), FAItemRec
        If fpcmbYN.Text = "N" And FAItemRec.DsplFlag = 2 Then GoTo SkipEm1Dept
        If LineCnt& >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintMasterHeader1
        End If
          
        YTDDep# = FAItemRec.DEP2DATE
  
        If ThisFundNum <> FAItemRec.FundNum Then
          GoTo SkipEm1Dept
        End If
        
        If QPTrim$(fptxtDeptNum.Text) = "ALL" Then
          If FAItemRec.IDEPT <> DeptNum(x) Then
            GoTo SkipEm1Dept
          End If
        Else
          If FAItemRec.IDEPT <> DeptNum(1) Then
            GoTo SkipEm1Dept
          End If
        End If
        
        LifeLeft = CStr(FAItemRec.LifeLeft)
        'format the asset's life data
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
        LifeData = QPTrim$(WholeLife) + "/" + LifeLeft
        Print #RptHandle, Tab(2); QPTrim$(FAItemRec.ItemTag); Tab(22); Left$(FAItemRec.IDESC1, 28);
        Print #RptHandle, Tab(51); Using("###0", FAItemRec.FundNum);
        Print #RptHandle, Tab(60); LifeData;
        Print #RptHandle, Tab(68); Using("$##,###,##0.00", CStr(FAItemRec.ORGCOST));
        If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
          Print #RptHandle, Tab(82); Using("$##,###,##0.00", CStr(YTDDep#)); "*";
        Else
          Print #RptHandle, Tab(82); Using("$##,###,##0.00", CStr(YTDDep#));
        End If
        Print #RptHandle, Tab(98); Using("$##,###,##0.00", CStr(FAItemRec.CURRVAL))
        If fpcmbYN.Text = "Y" Then
          If FAItemRec.DsplFlag = 2 Then
            Print #RptHandle, Tab(22); "^Disposal Date: "; Tab(40); MakeRegDate(FAItemRec.DispDate)
            LineCnt& = LineCnt& + 1
          End If
        End If
        If FAItemRec.DsplFlag = 1 Then
          Print #RptHandle, Tab(10); "^Scheduled For Disposal On: "; Tab(40); MakeRegDate(FAItemRec.DispDate)
          LineCnt& = LineCnt& + 1
        End If
        LineCnt& = LineCnt& + 1
        
        FundDeptCnt(Fundx, x) = FundDeptCnt(Fundx, x) + 1
        ItemCnt& = ItemCnt& + 1
        ItemTotal = ItemTotal + 1
        'collects grand totals
        
        FundDeptPurPrice(Fundx, x) = FundDeptPurPrice(Fundx, x) + FAItemRec.ORGCOST
        FundDeptDpr2Date(Fundx, x) = FundDeptDpr2Date(Fundx, x) + YTDDep
        FundDeptCurValue(Fundx, x) = FundDeptCurValue(Fundx, x) + FAItemRec.CURRVAL
        
        OrigCost# = OrigCost# + FAItemRec.ORGCOST
        BookTotal# = BookTotal# + (FAItemRec.CURRVAL)
        YDep# = YDep# + YTDDep#
        'collects dept totals
        FundCnt = FundCnt + 1
        ATagCCnt(Fundx) = FundCnt
        TotalItems = TotalItems + 1
        COrigCost# = COrigCost# + FAItemRec.ORGCOST
        ATagCOrigCost(Fundx) = COrigCost#
        CBookTotal# = CBookTotal# + (FAItemRec.CURRVAL)
        ATagCBookTotal(Fundx) = CBookTotal#
        CYDep# = CYDep# + YTDDep#
        ATagCYDep(Fundx) = CYDep#
        ThisDeptOrigCost(x) = ThisDeptOrigCost(x) + FAItemRec.ORGCOST
        ThisDeptBookTotal(x) = ThisDeptBookTotal(x) + FAItemRec.CURRVAL
        ThisDeptDep(x) = ThisDeptDep(x) + YTDDep#
        ThisDeptCnt(x) = ThisDeptCnt(x) + 1
SkipEm1Dept:
      Next cnt&
      GoSub PrintDeptFooter
      If fptxtDeptNum.Enabled = True And fptxtDeptNum.Text <> "ALL" Then Exit For
NutinHoney:
    Next x
    For x = 1 To NumOfDepts
      ThisDeptOrigCost(x) = 0
      ThisDeptBookTotal(x) = 0
      ThisDeptDep(x) = 0
      ThisDeptCnt(x) = 0
    Next x

  'First Print Subtotals
    Print #RptHandle, String$(111, "-")
    Print #RptHandle, "TOTALS FOR FUND NUMBER: "; CStr(ThisFundNum); " "; ThisFundDesc;
    Print #RptHandle, Tab(68); Using("$##,###,##0.00", CStr(COrigCost#));
    Print #RptHandle, Tab(82); Using("$##,###,##0.00", CStr(CYDep#));
    Print #RptHandle, Tab(98); Using("$##,###,##0.00", CStr(CBookTotal#))
    Print #RptHandle, "Total Items: "; CStr(FundCnt)
    Print #RptHandle, String$(111, "=")
    Print #RptHandle,
    LineCnt& = LineCnt& + 5
NoData:
    frmFAShowPctComp.ShowPctComp Fundx, FundIdxCnt
    If frmFAShowPctComp.Out = True Then
      Close
      frmFAShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdExit.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload frmFAShowPctComp
      Exit Sub
    End If
    If QPTrim$(code$) <> "ALL" Then Exit Do
    If Fundx = FundIdxCnt Then Exit Do
    Fundx = Fundx + 1
    ThisFundNum = FundNum(Fundx)
    ThisFundDesc = QPTrim$(FundDesc(Fundx))
    'clear all dept totals
    COrigCost# = 0
    CBookTotal# = 0
    CYDep# = 0
    FundCnt = 0
  Loop

  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True
  If ItemTotal = 0 Then
    MsgBox "There are no fixed assets saved for this criteria."
    Close
    Exit Sub
  End If
  
  Print #RptHandle, FF$
  GoSub PrintFundTotals2

  If TagPrint = False Then GoSub PrintMasterValueEnding1
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi

  Close         'Close all open files now
  
  ViewPrint ReportFile$, "Value By Purchase Price", True

  KillFile (ReportFile$)

  Exit Sub

PrintMasterHeader1:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Master Asset Listing : Asset Listing by Fund Code"
  If FirstFlag = False Then
    Print #RptHandle, "Fund # "; CStr(ThisFundNum); " "; ThisFundDesc
  End If
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, "* = DO NOT DEPRECIATE THIS ASSET"
  LineCnt& = 5
  Return

PrintMasterValueEnding1:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Master Asset Listing : Grand Totals"
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, Tab(23); "Total Items"; Tab(40); "Purchase Price"; Tab(59); "Total Deprec"; Tab(78); "Book Value"
  Print #RptHandle, String$(88, "=")
  Print #RptHandle, "Total Assets ";
  Print #RptHandle, Tab(21); Using$("##,###,##0", TotalItems);
  Print #RptHandle, Tab(40); Using("$##,###,##0.00", CStr(OrigCost#));
  Print #RptHandle, Tab(57); Using("$##,###,##0.00", CStr(YDep#));
  Print #RptHandle, Tab(74); Using("$##,###,##0.00", CStr(BookTotal#))
  Print #RptHandle, FF$

  Return

PrintDeptHeader:
  Print #RptHandle, String$(111, "=")
  If fptxtDeptNum.Enabled = True And fptxtDeptNum.Text = "ALL" Then
    Print #RptHandle, Tab(1); "Department: "; Tab(20); QPTrim$(DeptDesc(x)); Tab(55); "Dept # " + CStr(DeptNum(x))
  Else
    Print #RptHandle, Tab(1); "Department: "; Tab(20); QPTrim$(DeptDesc(1)); Tab(55); "Dept # " + CStr(DeptNum(1))
  End If
  Print #RptHandle, ""
  Print #RptHandle, Tab(1); "Asset Number"; Tab(22); "Description"; Tab(51); "Fund"; Tab(58); "Life/Left"; Tab(68); "Purchase Price"; Tab(84); "Total Deprec"; Tab(102); "Book Value"
  Print #RptHandle, String$(111, "-")
  LineCnt = LineCnt + 5
  
  Return

PrintDeptFooter:
  Print #RptHandle, String$(111, "-")
  If fptxtDeptNum.Enabled = True And fptxtDeptNum.Text = "ALL" Then
    Print #RptHandle, "Dept Total For: "; Tab(17); QPTrim$(DeptDesc(x)); Tab(40); "Dept # " + CStr(DeptNum(x))
    Print #RptHandle, "Number Of Department Items: "; Tab(37); QPTrim$(Using("######", ThisDeptCnt(x))); Tab(68); Using("$##,###,##0.00", ThisDeptOrigCost(x)); Tab(82); Using("$##,###,##0.00", ThisDeptDep(x)); Tab(98); Using("$##,###,##0.00", ThisDeptBookTotal(x))
  Else
    Print #RptHandle, "Dept Total For: "; Tab(17); QPTrim$(DeptDesc(1)); Tab(40); "Dept # " + CStr(DeptNum(1))
    Print #RptHandle, "Number Of Department Items: "; Tab(37); QPTrim$(Using("######", ThisDeptCnt(1))); Tab(68); Using("$##,###,##0.00", ThisDeptOrigCost(1)); Tab(82); Using("$##,###,##0.00", ThisDeptDep(1)); Tab(98); Using("$##,###,##0.00", ThisDeptBookTotal(1))
  End If
  Print #RptHandle, String$(111, "=")
  Print #RptHandle, ""
  Print #RptHandle, ""
  
  LineCnt = LineCnt + 6
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintMasterHeader1
  End If
  Return

PrintFundHeader:
  If Fundx <> 1 Then
    Print #RptHandle, FF$
    GoSub PrintMasterHeader1
  End If
  Print #RptHandle, String$(111, "=")
  If fptxtFundNum.Text = "ALL" Then
    Print #RptHandle, Tab(1); "ASSET FUND :"; Tab(20); QPTrim$(FundDesc(Fundx)); Tab(55); "Fund Number: "; Tab(70); Using("####0", FundNum(Fundx))
  Else
    Print #RptHandle, Tab(1); "ASSET FUND :"; Tab(20); ThisFundDesc; Tab(55); "Fund Number: "; Tab(70); QPTrim$(Using("####0", CStr(ThisFundNum)))
  End If
  Print #RptHandle, String$(111, "=")
  Print #RptHandle, ""
  LineCnt = LineCnt + 4
  Return
  
PrintFundTotals:
  Print #RptHandle, Tab(30); "Master Asset Listing : Asset Fund Totals"
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, Tab(4); "Fund"; Tab(15); "Description"; Tab(40); "Item Count"; Tab(68); "Purchase Price"; Tab(85); "Total Deprec"; Tab(102); "Book Value"
  Print #RptHandle, String$(111, "=")
  LineCnt = LineCnt + 4

  If fptxtFundNum.Text = "ALL" Then
    For x = 1 To FundIdxCnt
      Print #RptHandle, Tab(3); Using("####0", FundNum(x)); Tab(15); FundDesc(x); Tab(40); Using("#####0", ATagCCnt(x)); Tab(68); Using("$##,###,##0.00", CStr(ATagCOrigCost(x))); Tab(83); Using("$##,###,##0.00", CStr(ATagCYDep(x))); Tab(98); Using("$##,###,##0.00", CStr(ATagCBookTotal(x)))
      LineCnt = LineCnt + 1
  
      If LineCnt& >= MaxLines And x <> FundIdxCnt Then
        LineCnt& = 0
        Page = Page + 1
        Print #RptHandle, FF$
        Print #RptHandle, Tab(20); "Master Asset Listing : Asset Fund Totals"
        Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
        Print #RptHandle, Tab(1); "Fund"; Tab(15); "Description"; Tab(40); "Item Count"; Tab(69); "Purchase Price"; Tab(85); "Total Deprec"; Tab(101); "Book Value"
        Print #RptHandle, String$(111, "=")
        LineCnt = LineCnt + 5
      End If
    Next x
  
    If LineCnt <= 53 Then
      Print #RptHandle, String$(111, "=")
      Print #RptHandle, "Total Assets ";
      Print #RptHandle, Tab(40); Using("#####0", TotalItems);
      Print #RptHandle, Tab(68); Using("$##,###,##0.00", CStr(OrigCost#));
      Print #RptHandle, Tab(83); Using("$##,###,##0.00", CStr(YDep#));
      Print #RptHandle, Tab(98); Using("$##,###,##0.00", CStr(BookTotal#))
    Else
      Print #RptHandle, FF$
      Print #RptHandle, Tab(30); "Master Asset Listing : Asset Fund Totals"
      Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
      Print #RptHandle, Tab(1); "Department"; Tab(15); "Description"; Tab(40); "Item Count"; Tab(63); "Purchase Price"; Tab(80); "Total Deprec"; Tab(97); "Book Value"
      Print #RptHandle, String$(111, "=")
      Print #RptHandle, String$(111, "=")
      Print #RptHandle, "Total Assets ";
      Print #RptHandle, Tab(40); Using("#####0", TotalItems);
      Print #RptHandle, Tab(68); Using("$##,###,##0.00", CStr(OrigCost#));
      Print #RptHandle, Tab(83); Using("$##,###,##0.00", CStr(YDep#));
      Print #RptHandle, Tab(98); Using("$##,###,##0.00", CStr(BookTotal#))
    End If
  Else
    Print #RptHandle, Tab(3); Using("####0", CStr(ThisFundNum)); Tab(15); ThisFundDesc; Tab(40); Using("#####0", ATagCCnt(Fundx)); Tab(68); Using("$##,###,##0.00", CStr(ATagCOrigCost(Fundx))); Tab(83); Using("$##,###,##0.00", CStr(ATagCYDep(Fundx))); Tab(98); Using("$##,###,##0.00", CStr(ATagCBookTotal(Fundx)))
    LineCnt = LineCnt + 1

  End If
  Print #RptHandle, FF$
  TagPrint = True

  Return
  
PrintFundTotals2:
  Print #RptHandle, Tab(30); "Master Asset Listing : Asset Fund Totals"
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, String$(88, "=")
  LineCnt = 3
  
  Fundx = 1
  Do
    PurPrice = 0
    Dpr2Date = 0
    BookVal = 0
    FundItemCnt = 0
    GoSub PrintFundTotalsHeader
    For Y = 1 To NumOfDepts
      Print #RptHandle, QPTrim$(DeptDesc(Y)); Tab(15); CStr(DeptNum(Y)); Tab(22); Using$("#,###,##0", FundDeptCnt(Fundx, Y)); Tab(40); Using("$##,###,##0.00", FundDeptPurPrice(Fundx, Y)); Tab(57); Using$("$##,###,##0.00", FundDeptDpr2Date(Fundx, Y)); Tab(74); Using$("$##,###,##0.00", FundDeptCurValue(Fundx, Y))
      LineCnt = LineCnt + 1
      PurPrice = PurPrice + FundDeptPurPrice(Fundx, Y)
      Dpr2Date = Dpr2Date + FundDeptDpr2Date(Fundx, Y)
      BookVal = BookVal + FundDeptCurValue(Fundx, Y)
      FundItemCnt = FundItemCnt + FundDeptCnt(Fundx, Y)
      If LineCnt& >= MaxLines Then '  And x <> FundIdxCnt Then
        Print #RptHandle, FF$
        Page = Page + 1
        Print #RptHandle, Tab(30); "Master Asset Listing : Asset Fund Totals"
        Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
        Print #RptHandle, String$(88, "=")
        LineCnt = 3
        GoSub PrintFundTotalsHeader
      End If
    Next Y
    
    If LineCnt& >= MaxLines - 5 Then
      Print #RptHandle, FF$
      Page = Page + 1
      Print #RptHandle, Tab(30); "Master Asset Listing : Asset Fund Totals"
      Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
      Print #RptHandle, String$(88, "=")
      LineCnt = 3
      GoSub PrintFundTotalsHeader
    End If
      
    Print #RptHandle, String$(88, "-")
    Print #RptHandle, "Totals For: "; Tab(15); QPTrim$(FundDesc(Fundx)); Tab(35); "Fund # "; CStr(FundNum(Fundx))
    Print #RptHandle, Tab(21); Using$("##,###,##0", FundItemCnt); Tab(40); Using("$##,###,##0.00", PurPrice); Tab(57); Using$("$##,###,##0.00", Dpr2Date); Tab(74); Using$("$##,###,##0.00", BookVal)
    Print #RptHandle, String$(88, "=")
    Print #RptHandle, ""
    Print #RptHandle, ""
    LineCnt = LineCnt + 6
    If Fundx = FundIdxCnt Then Exit Do
    Fundx = Fundx + 1
  Loop
  Return
  
PrintFundTotalsHeader:
  If LineCnt& >= MaxLines Then
    Print #RptHandle, FF$
    Page = Page + 1
    Print #RptHandle, Tab(30); "Master Asset Listing : Asset Fund Totals"
    Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
    Print #RptHandle, String$(88, "=")
    LineCnt = 3
  End If
  Print #RptHandle, "Fund Description: "; QPTrim$(FundDesc(Fundx)); Tab(35); "Fund # "; CStr(FundNum(Fundx))
  Print #RptHandle, Tab(1); "Dept Desc"; Tab(14); "Dept #"; Tab(27); "Count"; Tab(40); "Purchase Price"; Tab(59); "Total Deprec"; Tab(78); "Book Value"
  Print #RptHandle, String$(88, "-")
  LineCnt = LineCnt + 3
  Return


ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAAssByFundRpt", "PrintDeptText", Erl)
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
    Unload Me

End Sub

