VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmWarrantyRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets Warranty Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmWarrantyRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6156
      Left            =   1932
      TabIndex        =   5
      Top             =   1344
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   10858
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmWarrantyRpt.frx":08CA
      Begin LpLib.fpCombo fpcmbOrder 
         Height          =   405
         Left            =   3165
         TabIndex        =   0
         ToolTipText     =   "Select the order this report will display."
         Top             =   1530
         Width           =   3225
         _Version        =   196608
         _ExtentX        =   5689
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
         ColDesigner     =   "frmWarrantyRpt.frx":08E6
      End
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3510
         TabIndex        =   4
         ToolTipText     =   "Select Graphical for a more robust report that may process more slowly. Select Text for a quicker report."
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
         ColDesigner     =   "frmWarrantyRpt.frx":0BDD
      End
      Begin EditLib.fpDateTime fpDateEnd 
         Height          =   444
         Left            =   3840
         TabIndex        =   3
         ToolTipText     =   "Enter the ending date for all warranty expiration dates for this report."
         Top             =   3552
         Width           =   1836
         _Version        =   196608
         _ExtentX        =   3238
         _ExtentY        =   783
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
         Text            =   "2/27/2003"
         DateCalcMethod  =   0
         DateTimeFormat  =   0
         UserDefinedFormat=   ""
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
         ButtonColor     =   13684944
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fptxtDeptNum 
         Height          =   396
         Left            =   3024
         TabIndex        =   1
         ToolTipText     =   "If DEPARTMENT NUMBER is selected in the Report Order field then enter the desired department this report will gather data for."
         Top             =   2208
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
      Begin EditLib.fpDateTime fpDateBegin 
         Height          =   444
         Left            =   3840
         TabIndex        =   2
         ToolTipText     =   "Enter the starting date for all warranty expiration dates for this report."
         Top             =   2880
         Width           =   1836
         _Version        =   196608
         _ExtentX        =   3238
         _ExtentY        =   783
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
         Text            =   "2/27/2003"
         DateCalcMethod  =   0
         DateTimeFormat  =   0
         UserDefinedFormat=   ""
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
         ButtonColor     =   13684944
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdDept 
         Height          =   405
         Left            =   4656
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to bring up a list of all current departments."
         Top             =   2208
         Width           =   1350
         _Version        =   131072
         _ExtentX        =   2381
         _ExtentY        =   714
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
         ButtonDesigner  =   "frmWarrantyRpt.frx":0ED4
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   690
         Left            =   1440
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   5040
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
         ButtonDesigner  =   "frmWarrantyRpt.frx":10B4
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4560
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to create the report based on the parameters entered above."
         Top             =   5040
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
         ButtonDesigner  =   "frmWarrantyRpt.frx":1290
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
         TabIndex        =   11
         Top             =   4320
         Width           =   1500
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Warranty Expiration Report"
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
         Left            =   1770
         TabIndex        =   10
         Top             =   570
         Width           =   4335
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1488
         Top             =   432
         Width           =   4908
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Left            =   1872
         TabIndex        =   9
         Top             =   3024
         Width           =   1692
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Left            =   2016
         TabIndex        =   8
         Top             =   3696
         Width           =   1548
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
         Left            =   1152
         TabIndex        =   7
         Top             =   1584
         Width           =   1836
      End
      Begin VB.Label Label11 
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
         Left            =   1872
         TabIndex        =   6
         Top             =   2304
         Width           =   924
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6348
      Left            =   1836
      Top             =   1260
      Width           =   7980
   End
End
Attribute VB_Name = "frmWarrantyRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdDept_Click()
  frmFADeptList.Show vbModal

End Sub

Private Sub cmdExit_Click()
  frmFAReportMenu.Show
  Close
  KillFile "Wrntyrpt.dat"
  DoEvents
  Unload frmWarrantyRpt
End Sub

Private Sub cmdProcess_Click()
  If fpcomboPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    MsgBox "Pitch 17 is recommended for this report."
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
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
'    'Me.Visible = False
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
      KillFile "Wrntyrpt.dat"
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmWarrantyRpt.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub fpcmbOrder_Change()
  If QPTrim$(fpcmbOrder.Text) = "TAG NUMBER" Then
    fptxtDeptNum.Enabled = False
    cmdDept.Enabled = False
    fptxtDeptNum.Text = "ALL"
  ElseIf QPTrim$(fpcmbOrder.Text) = "" Then
    fpcmbOrder.Text = "TAG NUMBER"
    fptxtDeptNum.Enabled = False
    fptxtDeptNum.Text = "ALL"
    cmdDept.Enabled = False
  Else
    fptxtDeptNum.Enabled = True
    cmdDept.Enabled = True
  End If
  
End Sub

Private Sub fpcmbOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  'prevents the user from inadvertently changing data in this
  'combo box when the user is tabbing through the fields
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

Private Sub LoadMe()
  Dim One As Integer
  Dim FileHandle As Integer
  One = 1
  FileHandle = FreeFile
  'this .dat file is used by the dept list form to
  'determine where to populate data selected on that
  'form to appear on this form
  
  Open "Wrntyrpt.dat" For Output As FileHandle Len = 2
  Print #FileHandle, One
  Close FileHandle
  fpcmbOrder.Text = "TAG NUMBER"
  fpcmbOrder.AddItem "TAG NUMBER"
  fpcmbOrder.AddItem "DEPARTMENT NUMBER"
  fptxtDeptNum.Text = "ALL"
  fpcomboPrintOpt.AddItem "Graphical"
  fpcomboPrintOpt.AddItem "Text"
  fpcomboPrintOpt.Text = "Graphical"
  fpDateBegin = Date
  fpDateEnd = Date
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAWarrantyRpt", "Check4alidDept", Erl)
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

Private Sub fpcomboPrintOpt_Change()
  'Graphical is the default value
  If QPTrim$(fpcomboPrintOpt.Text) = "" Then
    fpcomboPrintOpt.Text = "Graphical"
  End If
End Sub

Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  'prevents the user from inadvertently changing data in
  'this combo box when he is tabbing through the fields
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
  Dim DOrigCost#(2), DBookTotal#(2), DCDep#(2), DYDep#(2), OrigCost#(2), BookTotal#(2), CDep#(2), YDep#(2)
  Dim ReportFile$
  Dim Dash80$
  Dim FF$
  Dim MaxLines As Integer
  Dim ItemCnt&
  Dim LineCnt&
  Dim Dept$
  Dim BDate As Integer
  Dim EDate As Integer
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim ThisDate As Integer
  Dim DeptNumber As Integer
  Dim YTDDep#
  Dim x As Integer
  Dim Nextx As Integer
  Dim DIdxRec As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxCnt As Integer
  Dim Index$
  Dim Page As Integer
  Dim DataFlag As Boolean
  Dim TagFlag As Boolean
  Dim DeptCnt As Integer
  Dim DeptDescription$
  Dim DeptDescHeader$
  Dim TotalItems As Long
  Dim TagPrint As Boolean
  Dim LifeLeft As String * 3
  Dim WholeLife As String * 3
  Dim LifeData As String * 7
  Dim FirstFlag As Boolean
  Dim ItemTotal As Long
  'this procedure is much like other graphics reports
  'that are commented
  
  On Error GoTo ERRORSTUFF
  
  FirstFlag = True
  TagPrint = False
  If fpDateBegin.Text = "" Then
    MsgBox "Please enter a Start Date."
    fpDateBegin.SetFocus
    Exit Sub
  End If
  
  If fpDateEnd.Text = "" Then
    MsgBox "Please enter an End Date."
    fpDateEnd.SetFocus
    Exit Sub
  End If
  
  BDate = Date2Num(fpDateBegin)
  EDate = Date2Num(fpDateEnd)
  If EDate < BDate Then
    MsgBox "The End Date is before the Start Date. Please reenter these values"
    fpDateBegin.SetFocus
    Exit Sub
  End If
  
  If Check4ValidDept = False Then Exit Sub
  
  ReportFile$ = "FAWARRANTY.PRN"  'Report File Name
  Dash80$ = String$(80, "=")
  FF$ = Chr$(12)
  
  MaxLines = 56
  LineCnt& = 0
  ItemCnt& = 0
  Dept$ = QPTrim$(fptxtDeptNum.Text)
  
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
  
  OpenDeptIdxFile DIdxHandle
  DIdxCnt = LOF(DIdxHandle) \ Len(DIdxRec)
  ReDim DeptNum(1 To DIdxCnt) As String
  ReDim DeptDesc(1 To DIdxCnt) As String
  
  For x = 1 To DIdxCnt
    Get DIdxHandle, x, DIdxRec
    DeptNum(x) = QPTrim$(DIdxRec.DeptNumb)
    DeptDesc(x) = QPTrim$(DIdxRec.DeptIdxDesc)
  Next x
  Close DIdxHandle
  
  If Dept$ <> "ALL" Then
    DeptNumber = Val(QPTrim$(fptxtDeptNum.Text))
    For x = 1 To DIdxCnt + 1
      If DeptNumber = DeptNum(x) Then
        DeptDescription = QPTrim$(DeptDesc(x))
        DeptDescHeader$ = DeptDescription
        Exit For
      End If
    Next x
  Else
    DeptNumber = Val(QPTrim(DeptNum(1)))
    DeptDescription = QPTrim(DeptDesc(1))
    DeptDescHeader$ = ""
  End If
  
  GoSub PrintMasterHeader1
  
  ReDim ATagDOrigCost(1 To DIdxCnt) As Double
  ReDim ATagDBookTotal(1 To DIdxCnt) As Double
  ReDim ATagDYDep(1 To DIdxCnt) As Double
  ReDim ATagDCnt(1 To DIdxCnt) As Long
  
  OpenFAItemFile FAHandle
  
  TagFlag = False
  
  frmFAShowPctComp.Label1 = "Gathering Item Data"
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
    Index = "DEPARTMENT NUMBERS"
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
      If FAItemRec.DsplFlag = 2 Then GoTo SkipEm1 'no reason to see
      'the warranty data if the item is gone
      If LineCnt& >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintMasterHeader1
      End If
      ThisDate = FAItemRec.WARRXDAT
      
      If ThisDate < BDate Or ThisDate > EDate Then
      'filter out items that don't fall inside the date parameters
        GoTo SkipEm1
      End If
      'Can't depreciate something that has no asset life...this should
      'not be necessary unless the asset life was changed to zero after
      'depreciation had already started
      If FAItemRec.ILIFE > 0 Then
        YTDDep# = FAItemRec.DEP2DATE
      Else
        YTDDep# = 0
      End If
      
      'TAG NUMBER is actually true (if selected by the user) for the first
      'complete iteration only so that all tag numbers can be printed in numeric
      'order. After that we do not want the tags itemized we only want department
      'totals figured so that at the end of the report a "totals by dept" section
      'can be printed.
      If QPTrim$(Index) = "TAG NUMBER" Then
        GoTo TagOnly1 'prints itemized tag data in numeric order
      ElseIf DeptNumber <> FAItemRec.IDEPT Then 'any time department data
      'is needed then the item falls into this part of the if statement
        GoTo SkipEm1 'if the prevailing department (in the numeric index)
        'doesn't match this item's dept number then we don'y want it now
      End If
      'at this point the item's dept matches the prevailing dept number
      'but since TagFlag (don't want itemized tag data anymore) is true then
      'go ahead and collect dept data
      If TagFlag = True Then GoTo TagOnly2
      
TagOnly1: 'printing valid tag data...skipped if TAG NUMBER is chosen and this is
          'not the first iteration
      DataFlag = True
      If DeptCnt = 0 And QPTrim$(fpcmbOrder.Text) = "DEPARTMENT NUMBER" Then
         Print #RptHandle,
         Print #RptHandle, "Dept # "; DeptNumber; " "; DeptDescription
         Print #RptHandle, String$(111, "=")
         LineCnt = LineCnt + 3
      End If
      LifeLeft = CStr(FAItemRec.LifeLeft)
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
      
      RSet LifeData = QPTrim$(WholeLife) + "/" + LifeLeft
      ItemTotal = ItemTotal + 1
      Print #RptHandle, Tab(1); FAItemRec.ItemTag; Tab(22); Left$(FAItemRec.IDESC1, 28);
      Print #RptHandle, Tab(51); Using("###0", FAItemRec.IDEPT);
      Print #RptHandle, Tab(59); LifeData;
      Print #RptHandle, Tab(68); Using("###,###,##0.00", CStr(FAItemRec.ORGCOST));
      If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
        Print #RptHandle, Tab(82); Using("###,###,##0.00", CStr(FAItemRec.CURRVAL)); "*";
      Else
        Print #RptHandle, Tab(82); Using("###,###,##0.00", CStr(FAItemRec.CURRVAL));
      End If
      Print #RptHandle, Tab(102); MakeRegDate(ThisDate)
      LineCnt& = LineCnt& + 1
      ItemCnt& = ItemCnt& + 1
      
TagOnly2: 'collects data for each department for reporting totals

      'This if statement filters out the first iteration of TAG NUMBER
      'selection because we do not want to start accumulating dept data
      'until the second iteration
      If TagFlag = False And QPTrim$(Index) = "TAG NUMBER" Then GoTo SkipEm1
      
      'collects grand totals
      OrigCost#(2) = OrigCost#(2) + FAItemRec.ORGCOST
      BookTotal#(2) = BookTotal#(2) + (FAItemRec.CURRVAL)
      YDep#(2) = YDep#(2) + YTDDep#
      
      'collects dept totals
      DeptCnt = DeptCnt + 1
      ATagDCnt(Nextx) = DeptCnt
      TotalItems = TotalItems + 1
      DOrigCost#(2) = DOrigCost#(2) + FAItemRec.ORGCOST
      ATagDOrigCost(Nextx) = DOrigCost#(2)
      DBookTotal#(2) = DBookTotal#(2) + (FAItemRec.CURRVAL)
      ATagDBookTotal(Nextx) = DBookTotal#(2)
      DYDep#(2) = DYDep#(2) + YTDDep#
      ATagDYDep(Nextx) = DYDep#(2)
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
    Print #RptHandle, String$(111, "-")
    Print #RptHandle, "Assets for Dept Number: "; DeptNumber; " "; DeptDescription;
    Print #RptHandle, Tab(68); Using("###,###,##0.00", CStr(DOrigCost#(2)));
    Print #RptHandle, Tab(82); Using("###,###,##0.00", CStr(DBookTotal#(2)))
    Print #RptHandle, "Total Items: "; DeptCnt
    Print #RptHandle, String$(111, "=")
    LineCnt& = LineCnt& + 4
NoData:
    frmFAShowPctComp.ShowPctComp Nextx, DIdxCnt
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
    If QPTrim$(Dept$) <> "ALL" Then Exit Do
    'if all the depts have been examined then time to go
    If Nextx = DIdxCnt Then Exit Do
    'move to the next dept
    Nextx = Nextx + 1
    'assign new dept to DeptNumber
    DeptNumber = Val(QPTrim$(DeptNum(Nextx)))
    DeptDescription = QPTrim$(DeptDesc(Nextx))
    'clear all dept totals
    DOrigCost#(1) = 0
    DBookTotal#(1) = 0
    DYDep#(1) = 0
    DOrigCost#(2) = 0
    DBookTotal#(2) = 0
    DYDep#(2) = 0
    DeptCnt = 0
  
  Loop
  
  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True
  
  If ItemTotal = 0 Then
    MsgBox "There are no Warranties Expiring for this Time Period."
    Close
    Exit Sub
  End If
  
  'only prints if TAG NUMBERS was selected
  If TagFlag = True Then
    GoSub PrintTagDeptTotals
  End If
  
  
  If TagPrint = False Then GoSub PrintMasterValueEnding1
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  
  Close         'Close all open files now
  
  ViewPrint ReportFile$, "Warrantys By Expiration Date", True
  
  KillFile (ReportFile$)
  
  Exit Sub
  
PrintMasterHeader1:
  Page = Page + 1
  
  Print #RptHandle, Tab(30); "Master Asset Listing : Warranties by Expiration Date"
  If QPTrim(fpcmbOrder.Text) = "DEPARTMENT NUMBER" And FirstFlag = False Then
    Print #RptHandle, "Dept # "; CStr(DeptNumber); " "; DeptDescription
  Else
    Print #RptHandle,
  End If
  
  Print #RptHandle, "Warranty Expiration Range "; fpDateBegin; " to "; fpDateEnd
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, "* = DO NOT DEPRECIATE THIS ASSET"
  Print #RptHandle, Tab(1); "Asset Number"; Tab(26); "Description"; Tab(51); "Dept"; Tab(58); "Life/Left"; Tab(68); "Purchase Price"; Tab(86); "Book Value"; Tab(102); "Expires On"
  Print #RptHandle, String$(111, "=")
  LineCnt& = 7
  If FirstFlag = True Then
    FirstFlag = False
  End If
  Return
  
PrintMasterValueEnding1:
  
  Page = Page + 1
  Print #RptHandle, FF$
  Print #RptHandle, Tab(30); "Master Asset Listing : Grand Totals"
  Print #RptHandle, "Warranty Expiration Range "; fpDateBegin; " to "; fpDateEnd
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, Tab(18); "Total Items"; Tab(54); "Purchase Price"; Tab(77); "Book Value"
  Print #RptHandle, String$(86, "=")
  Print #RptHandle, "Total Assets ";
  Print #RptHandle, Tab(21); TotalItems;
  Print #RptHandle, Tab(54); Using("###,###,##0.00", CStr(OrigCost#(2)));
  Print #RptHandle, Tab(73); Using("###,###,##0.00", CStr(BookTotal#(2)))
  
  Print #RptHandle, FF$
  
  Return
  
PrintTagDeptTotals: 'print only if TAG NUMBERS was selected
  
  Page = Page + 1
  Print #RptHandle, FF$
  Print #RptHandle, Tab(30); "Master Asset Listing : Department Totals"
  Print #RptHandle, "Warranty Expiration Range "; fpDateBegin; " to "; fpDateEnd
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, Tab(1); "Department"; Tab(15); "Description"; Tab(40); "Item Count"; Tab(59); "Purchase Price"; Tab(78); "Book Value"
  Print #RptHandle, String$(87, "=")
  LineCnt = 6
  
  For x = 1 To DIdxCnt ' + 1
    If QPTrim$(DeptNum(x)) = "" Then DeptNum(x) = "0"
    Print #RptHandle, Tab(3); Using("####0", DeptNum(x)); Tab(15); DeptDesc(x); Tab(40); Using("#####0", ATagDCnt(x)); Tab(59); Using("###,###,##0.00", CStr(ATagDOrigCost(x))); Tab(74); Using("###,###,##0.00", CStr(ATagDBookTotal(x)))
    LineCnt = LineCnt + 1
    
    If LineCnt& >= MaxLines And x <> DIdxCnt + 1 Then
      LineCnt& = 0
      Page = Page + 1
      Print #RptHandle, FF$
      Print #RptHandle, Tab(20); "Master Asset Listing : Department Totals"
      Print #RptHandle, "Warranty Expiration Range "; fpDateBegin; " to "; fpDateEnd
      Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
      Print #RptHandle, Tab(1); "Department"; Tab(15); "Description"; Tab(40); "Item Count"; Tab(59); "Purchase Price"; Tab(76); "Book Value"
      Print #RptHandle, String$(87, "=")
      LineCnt = LineCnt + 5
    End If
  Next x
  
  If LineCnt <= 53 Then
    Print #RptHandle, String$(87, "=")
    Print #RptHandle, "Total Assets ";
    Print #RptHandle, Tab(40); Using("#####0", TotalItems);
    Print #RptHandle, Tab(59); Using("###,###,##0.00", CStr(OrigCost#(2)));
    Print #RptHandle, Tab(74); Using("###,###,##0.00", CStr(BookTotal#(2)))
  Else
    Print #RptHandle, FF$
    Print #RptHandle, Tab(30); "Master Asset Listing : Department Totals"
    Print #RptHandle, "Warranty Expiration Range "; fpDateBegin; " to "; fpDateEnd
    Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
    Print #RptHandle, Tab(1); "Department"; Tab(15); "Description"; Tab(40); "Item Count"; Tab(54); "Purchase Price"; Tab(71); "Book Value"
    Print #RptHandle, String$(87, "=")
    Print #RptHandle, String$(87, "=")
    Print #RptHandle, "Total Assets ";
    Print #RptHandle, Tab(40); Using("#####0", TotalItems);
    Print #RptHandle, Tab(59); Using("###,###,##0.00", CStr(OrigCost#(2)));
    Print #RptHandle, Tab(74); Using("###,###,##0.00", CStr(BookTotal#(2)))
  End If
  Print #RptHandle, FF$
  TagPrint = True
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmWarrantyRpt", "PrintText", Erl)
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

End Sub

Private Sub fptxtDeptNum_Change()
  If fptxtDeptNum.Text = "" Then
    fptxtDeptNum = "ALL"
  End If
End Sub

Private Sub PrintGraphics()
  Dim DOrigCost#(2), DBookTotal#(2), DCDep#(2), DYDep#(2), OrigCost#(2), BookTotal#(2), CDep#(2), YDep#(2)
  Dim ReportFile$
  Dim TagReportFile$
  Dim TagRptHandle As Integer
  Dim TDReportFile$
  Dim TDRptHandle As Integer
  Dim GTReportFile$
  Dim GTRptHandle As Integer
  Dim ItemCnt&
  Dim Dept$
  Dim BDate As Integer
  Dim EDate As Integer
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim ThisDate As Integer
  Dim DeptNumber As Integer
  Dim YTDDep#
  Dim x As Integer
  Dim Nextx As Integer
  Dim DIdxRec As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxCnt As Integer
  Dim Index$
  Dim DataFlag As Boolean
  Dim TagFlag As Boolean
  Dim DeptCnt As Integer
  Dim DeptDescription$
  Dim DeptDescHeader$
  Dim TotalItems As Long
  Dim TagPrint As Boolean
  Dim FASetUpRec As FASetupRecType
  Dim FASHandle As Integer
  Dim Employer$
  Dim dlm$
  Dim EndRpt As Integer
  Dim DateRange$
  Dim ItemTotal As Long
  
  On Error GoTo ERRORSTUFF
  
  'this procedure is much like other graphics reports
  'that are commented
  If QPTrim$(fpDateBegin.Text) = "" Then
    MsgBox "Please enter a date for Start Date."
    fpDateBegin.SetFocus
    Exit Sub
  End If

  If QPTrim$(fpDateEnd.Text) = "" Then
    MsgBox "Please enter a date for End Date."
    fpDateEnd.SetFocus
    Exit Sub
  End If
  
  DateRange = fpDateBegin.Text + " to " + fpDateEnd.Text
  dlm$ = "~"
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  Employer = FASetUpRec.TownName

  TagPrint = False

  BDate = Date2Num(fpDateBegin.Text)
  EDate = Date2Num(fpDateEnd.Text)
  If EDate < BDate Then
    MsgBox "The End Date is before the Start Date. Please reenter these values"
    fpDateBegin.SetFocus
    Exit Sub
  End If

  If Check4ValidDept = False Then Exit Sub

  ReportFile$ = "FARPTS\FAWRNTYRPT.RPT"  'Report File Name
  TagReportFile$ = "FARPTS\FATAGWRNTYRPT.RPT"
  TDReportFile$ = "FARPTS\FATAGDEPTWRNTY.RPT"
  GTReportFile$ = "FARPTS\FAGTWRNTYRPT.RPT"
  ItemCnt& = 0
  Dept$ = QPTrim$(fptxtDeptNum.Text)

  Index$ = QPTrim$(fpcmbOrder.Text)

  If QPTrim$(Index$) = "TAG NUMBER" Then
    TagRptHandle = FreeFile
    Open TagReportFile For Output As #TagRptHandle
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

  OpenDeptIdxFile DIdxHandle
  DIdxCnt = LOF(DIdxHandle) \ Len(DIdxRec)
  ReDim DeptNum(1 To DIdxCnt) As String
  ReDim DeptDesc(1 To DIdxCnt) As String

  For x = 1 To DIdxCnt
    Get DIdxHandle, x, DIdxRec
    DeptNum(x) = QPTrim$(DIdxRec.DeptNumb)
    DeptDesc(x) = QPTrim$(DIdxRec.DeptIdxDesc)
  Next x
  Close DIdxHandle

  If Dept$ <> "ALL" Then
    DeptNumber = Val(QPTrim$(fptxtDeptNum.Text))
    For x = 1 To DIdxCnt + 1
      If DeptNumber = DeptNum(x) Then
        DeptDescription = QPTrim$(DeptDesc(x))
        DeptDescHeader$ = DeptDescription
        Exit For
      End If
    Next x
  Else
    DeptNumber = Val(QPTrim(DeptNum(1)))
    DeptDescription = QPTrim(DeptDesc(1))
    DeptDescHeader$ = ""
  End If

  ReDim ATagDOrigCost(1 To DIdxCnt) As Double
  ReDim ATagDBookTotal(1 To DIdxCnt) As Double
  ReDim ATagDCnt(1 To DIdxCnt) As Long
  
  OpenFAItemFile FAHandle

  TagFlag = False

  frmFAShowPctComp.Label1 = "Gathering Warranty Data"
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
    Index = "DEPARTMENT NUMBERS"
  End If

  Do 'this loop iterates once if there is 1 dept requested, DIdxCnt + 1 if
  'department numbers "ALL" is requested and DIdxCnt + 2 for Tag Numbers
  '...Tag Numbers requires one iteration ignoring departments to get a list of
  'all valid tags in numeric order then DIdxCnt + 1 skipping the itemized tag
  'data print out just to allow the dept totals to assimilate...
    DataFlag = False
    For cnt& = 1 To NumOfFARecs
      Get FAHandle, TagIdxRecs(cnt), FAItemRec
      If FAItemRec.DsplFlag = 2 Then GoTo SkipEm1 'no reason to see
      'the warranty data if the item is gone
      ThisDate = FAItemRec.WARRXDAT

      If ThisDate < BDate Or ThisDate > EDate Then
      'filter out items that don't fall inside the date parameters
        GoTo SkipEm1
      End If
      'Can't depreciate something that has no asset life...this should
      'not be necessary unless the asset life was changed to zero after
      'depreciation had already started
      If FAItemRec.ILIFE > 0 Then
        YTDDep# = FAItemRec.DEP2DATE
      Else
        YTDDep# = 0
      End If

      'TAG NUMBER is actually true (if selected by the user) for the first
      'complete iteration only so that all tag numbers can be printed in numeric
      'order. After that we do not want the tags itemized we only want department
      'totals figured so that at the end of the report a "totals by dept" section
      'can be printed.
      If QPTrim$(Index) = "TAG NUMBER" Then
        GoTo TagOnly1 'prints itemized tag data in numeric order
      ElseIf DeptNumber <> FAItemRec.IDEPT Then 'any time department data
      'is needed then the item falls into this part of the if statement
        GoTo SkipEm1 'if the prevailing department (in the numeric index)
        'doesn't match this item's dept number then we don'y want it now
      End If
      'at this point the item's dept matches the prevailing dept number
      'but since TagFlag (don't want itemized tag data anymore) is true then
      'go ahead and collect dept data
      If TagFlag = True Then GoTo TagOnly2

TagOnly1: 'printing valid tag data...skipped if TAG NUMBER is chosen and this is
          'not the first iteration
      DataFlag = True
      If TagRptHandle > 0 Then
        '                        0              1
        Print #TagRptHandle, Employer; dlm; DateRange; dlm;
        '                             2                       3
        Print #TagRptHandle, FAItemRec.ItemTag; dlm; FAItemRec.IDESC1; dlm;
        '                            4
        Print #TagRptHandle, FAItemRec.IDEPT; dlm;
        '                            5
        Print #TagRptHandle, FAItemRec.ILIFE; dlm;
        '                            6
        Print #TagRptHandle, FAItemRec.ORGCOST; dlm;
        If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
          '                                         7                        8
          Print #TagRptHandle, FAItemRec.CURRVAL; dlm; "*"; dlm;
        Else
          '                                         7                        8
          Print #TagRptHandle, FAItemRec.CURRVAL; dlm; " "; dlm;
        End If
        '                     9              10                  11                  12
        Print #TagRptHandle, Dept$; dlm; DeptDescHeader$; dlm; DeptNumber; dlm; DeptDescription; dlm;
        '                         13               14               15
        Print #TagRptHandle, OrigCost#(2); dlm; BookTotal#(2); dlm; TotalItems; dlm;
        '                         16                      17                        18
        Print #TagRptHandle, FAItemRec.DEPYN; dlm; FAItemRec.LifeLeft; dlm; MakeRegDate(ThisDate)

      Else
        '                     0              1
        Print #RptHandle, Employer; dlm; DateRange; dlm;
        '                          2                        3
        Print #RptHandle, FAItemRec.ItemTag; dlm; FAItemRec.IDESC1; dlm;
        '                        4
        Print #RptHandle, FAItemRec.IDEPT; dlm;
        '                        5
        Print #RptHandle, FAItemRec.ILIFE; dlm;
        '                        6
        Print #RptHandle, FAItemRec.ORGCOST; dlm;
        If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
          '                                     7                         8
          Print #RptHandle, FAItemRec.CURRVAL; dlm; "*"; dlm;
        Else
          '                                     7                         8
          Print #RptHandle, FAItemRec.CURRVAL; dlm; " "; dlm;
        End If
        '                  9                 10                  11                 12
        Print #RptHandle, Dept$; dlm; DeptDescHeader$; dlm; DeptNumber; dlm; DeptDescription; dlm;
        '                      13                14                 15
        Print #RptHandle, DOrigCost#(2); dlm; DBookTotal#(2); dlm; DeptCnt; dlm;
        '                      16               17               18                 20
        Print #RptHandle, OrigCost#(2); dlm; BookTotal#(2); dlm; TotalItems; dlm;
        '                         19                      20                21
        Print #RptHandle, FAItemRec.DEPYN; dlm; FAItemRec.LifeLeft; dlm; MakeRegDate(ThisDate)
      End If

      ItemCnt& = ItemCnt& + 1
      ItemTotal = ItemTotal + 1
TagOnly2: 'collects data for each department for reporting totals

      'This if statement filters out the first iteration of TAG NUMBER
      'selection because we do not want to start accumulating dept data
      'until the second iteration
      If TagFlag = False And QPTrim$(Index) = "TAG NUMBER" Then GoTo SkipEm1

      'collects grand totals
      OrigCost#(2) = OrigCost#(2) + FAItemRec.ORGCOST
      BookTotal#(2) = BookTotal#(2) + (FAItemRec.CURRVAL)
      YDep#(2) = YDep#(2) + YTDDep#

      'collects dept totals
      DeptCnt = DeptCnt + 1
      ATagDCnt(Nextx) = DeptCnt
      TotalItems = TotalItems + 1
      DOrigCost#(2) = DOrigCost#(2) + FAItemRec.ORGCOST
      ATagDOrigCost(Nextx) = DOrigCost#(2)
      DBookTotal#(2) = DBookTotal#(2) + (FAItemRec.CURRVAL)
      ATagDBookTotal(Nextx) = DBookTotal#(2)
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
    frmFAShowPctComp.ShowPctComp Nextx, DIdxCnt ' + 1
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
    If QPTrim$(Dept$) <> "ALL" Then Exit Do
    'if all the depts have been examined then time to go
'    If Nextx = DIdxCnt + 1 Then Exit Do
    If Nextx = DIdxCnt Then Exit Do
    'move to the next dept
    Nextx = Nextx + 1
    'assign new dept to DeptNumber
    DeptNumber = Val(QPTrim$(DeptNum(Nextx)))
    DeptDescription = QPTrim$(DeptDesc(Nextx))
    'clear all dept totals
    DOrigCost#(1) = 0
    DBookTotal#(1) = 0
    DYDep#(1) = 0
    DOrigCost#(2) = 0
    DBookTotal#(2) = 0
    DYDep#(2) = 0
    DeptCnt = 0
   Loop

  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True

  'only prints if TAG NUMBERS was selected
  Close         'Close all open files now
  
  If ItemTotal = 0 Then
    MsgBox "There are no Warranties Expiring for this Time Period."
    Exit Sub
  End If
  
  If TagFlag = True Then
    GoSub PrintTagDeptTotals
  End If

  If TagFlag = True Then
    arFATagWrntyRpt.Show
  Else
    arFAWrntyRpt.Show
  End If

  frmFALoadReport.Show

  Exit Sub


PrintTagDeptTotals: 'print only if TAG NUMBERS was selected

  GTRptHandle = FreeFile
  Open GTReportFile$ For Output As GTRptHandle
  EndRpt = 1
  For x = 1 To DIdxCnt
    '                        0                1                2                   3                      4
    Print #GTRptHandle, DeptNum(x); dlm; DeptDesc(x); dlm; ATagDCnt(x); dlm; ATagDOrigCost(x); dlm; ATagDBookTotal(x); dlm;
    '                        5                6                7                8
    Print #GTRptHandle, TotalItems; dlm; OrigCost#(2); dlm; BookTotal#(2); dlm; EndRpt
  Next x

  Close GTRptHandle

  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmWarrantyRpt", "PrintGraphics", Erl)
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

End Sub


