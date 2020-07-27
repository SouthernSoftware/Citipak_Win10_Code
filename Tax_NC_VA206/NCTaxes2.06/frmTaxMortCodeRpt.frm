VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmTaxMortCodeRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Mortgage Code Report"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxMortCodeRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6150
      Left            =   1920
      TabIndex        =   4
      Top             =   1290
      Width           =   7785
      _Version        =   196609
      _ExtentX        =   13732
      _ExtentY        =   10848
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmTaxMortCodeRpt.frx":08CA
      Begin LpLib.fpCombo fpcmbMortCodes 
         Height          =   384
         Left            =   1128
         TabIndex        =   0
         Top             =   2280
         Width           =   5724
         _Version        =   196608
         _ExtentX        =   10096
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
         Columns         =   3
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
         ColDesigner     =   "frmTaxMortCodeRpt.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   384
         Left            =   2928
         TabIndex        =   3
         Top             =   4236
         Width           =   3564
         _Version        =   196608
         _ExtentX        =   6286
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
         Object.TabStop         =   0   'False
         BackColor       =   16777215
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
         ColDesigner     =   "frmTaxMortCodeRpt.frx":0D45
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   2928
         TabIndex        =   2
         Top             =   3660
         Width           =   3564
         _Version        =   196608
         _ExtentX        =   6286
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
         ColDesigner     =   "frmTaxMortCodeRpt.frx":1120
      End
      Begin LpLib.fpCombo fpcmbTownship 
         Height          =   384
         Left            =   2928
         TabIndex        =   1
         Top             =   3096
         Width           =   3564
         _Version        =   196608
         _ExtentX        =   6286
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
         ColDesigner     =   "frmTaxMortCodeRpt.frx":14FB
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   645
         Left            =   2040
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   5250
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmTaxMortCodeRpt.frx":18D6
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   645
         Left            =   4275
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   $"frmTaxMortCodeRpt.frx":1AB4
         Top             =   5250
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmTaxMortCodeRpt.frx":1B5F
      End
      Begin VB.Shape Shape2 
         Height          =   1455
         Left            =   960
         Top             =   1440
         Width           =   6015
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00D0D0D0&
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         TabIndex        =   13
         Top             =   2040
         Width           =   1260
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00D0D0D0&
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Top             =   2040
         Width           =   660
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Mortgage Codes:"
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
         Left            =   2880
         TabIndex        =   11
         Top             =   1560
         Width           =   1980
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Township:"
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
         Left            =   1515
         TabIndex        =   10
         Top             =   3180
         Width           =   1260
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3855
         Left            =   645
         Top             =   1170
         Width           =   6570
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Print Order:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1470
         TabIndex        =   9
         Top             =   4320
         Width           =   1305
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Mortgage Code Report"
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
         Height          =   390
         Left            =   1800
         TabIndex        =   8
         Top             =   450
         Width           =   4335
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   1530
         Top             =   315
         Width           =   4905
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Report Type:"
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
         Left            =   1275
         TabIndex        =   7
         Top             =   3750
         Width           =   1500
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6420
      Left            =   1800
      Top             =   1155
      Width           =   8055
   End
End
Attribute VB_Name = "frmTaxMortCodeRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim Town$
  Dim UseOpt As String * 1
  Dim ThisOpt$
  Dim MortCodeArray() As String
  Dim MortCodeArrayCnt As Integer
  Dim MortCodeBank() As String
  
Private Sub cmdExit_Click()
  frmTaxReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  If fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  Else
    frmTaxMsg.Label1.Caption = "Pitch 10 is recommended for this printout."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Call PrintText
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
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpMortgageCode
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxMortCodeRpt.")
      Call Terminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TSRec As TownshipType
  Dim TSCnt As Integer
  Dim TSHandle As Integer
  Dim x As Integer
  Dim MortRec As MortCodeRecType
  Dim MHandle As Integer
  Dim NumOfMCodes As Integer
  
  'on error goto ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Close TMHandle
  Town$ = QPTrim$(TaxMasterRec.Name)
  If Exist(TaxTownships) Then
    fpcmbTownship.Text = "All"
    fpcmbTownship.AddItem "All"
    OpenTownshipFile TSHandle, TSCnt
    For x = 1 To TSCnt
      Get TSHandle, x, TSRec
      fpcmbTownship.AddItem QPTrim$(TSRec.TownShip)
    Next x
    Close TSHandle
  Else
    fpcmbTownship.Text = "No Townships Saved"
  End If
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town = QPTrim$(TaxMasterRec.Name)
  
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  
  fpcmbPrintOrder.Text = "Name Order"
  fpcmbPrintOrder.AddItem "Name Order"
  fpcmbPrintOrder.AddItem "Acct Number Order"
  fpcmbPrintOrder.AddItem "Search Name"
  ThisOpt = QPTrim$(TaxMasterRec.OptSrchCust)
  If ThisOpt <> "" Then
    fpcmbPrintOrder.AddItem ThisOpt + " Order"
  End If
  
  fpcmbMortCodes.InsertRow = "All" + Chr(9) + "No Description" + Chr(9) + "0"
  fpcmbMortCodes.SearchText = "All" + Chr(9) + "No Description" + Chr(9) + "0"
  fpcmbMortCodes.Action = 0
  If fpcmbMortCodes.SearchIndex <> -1 Then
    fpcmbMortCodes.ListIndex = fpcmbMortCodes.SearchIndex
  Else
    fpcmbMortCodes.ListIndex = 0
  End If
  
  OpenMortCodeFile MHandle, NumOfMCodes
  
  MortCodeArrayCnt = NumOfMCodes
  ReDim MortCodeArray(1 To MortCodeArrayCnt) As String
  ReDim MortCodeBank(1 To MortCodeArrayCnt) As String
  For x = 1 To NumOfMCodes
    Get MHandle, x, MortRec
    If MortRec.Deleted = True Then GoTo ItsDeleted
    fpcmbMortCodes.InsertRow = QPTrim$(MortRec.MORTCODE) + Chr(9) + QPTrim$(MortRec.BName) + Chr(9) + CStr(x)
    MortCodeArray(x) = QPTrim$(MortRec.MORTCODE)
    MortCodeBank(x) = QPTrim$(MortRec.BName)
ItsDeleted:
  Next x
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxMortCodeRpt", "LoadMe", Erl)
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

Private Sub fpcmbMortCodes_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbMortCodes.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbMortCodes.ListIndex = -1
  End If
  If fpcmbMortCodes.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbTownship.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOpt.ListIndex = -1
  End If
  If fpcmbPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPrintOrder.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPrintOrder_Change()
  If ThisOpt <> "" Then
    If InStr(fpcmbPrintOrder.Text, ThisOpt) Then
      UseOpt = "Y"
    Else
      UseOpt = "N"
    End If
  End If
End Sub

Private Sub fpcmbPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOrder.ListIndex = -1
  End If
  If fpcmbPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbMortCodes.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub fpcmbTownship_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTownship.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTownship.ListIndex = -1
  End If
  If fpcmbTownship.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPrintOpt.SetFocus
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
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Integer, z As Integer
  Dim dlm$
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim ThisTownship$
  Dim MORTCODE$
  Dim ThisMortRec As Integer
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim NextRec As Long
  Dim MortCnt As Integer
  Dim ThisMortCode As Integer
  Dim CustRec As Long
  Dim PrintCnt As Long
  Dim TotCnt As Integer
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  
  'on error goto ERRORSTUFF
  
  ReDim CodeCnt(1 To MortCodeArrayCnt) As Integer
  OptFlag = False
    
  fpcmbMortCodes.Col = 0
  MORTCODE$ = fpcmbMortCodes.ColText
  fpcmbMortCodes.Col = 2
  ThisMortCode = CInt(fpcmbMortCodes.ColText)
  dlm$ = "~"
  ThisTownship = fpcmbTownship.Text
  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no customers saved."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close IdxHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Search Name" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no search names indexed."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = ThisOpt + " Order" Then
    OpenCustOptSearchFile OHandle, NumOfIdx
    If NumOfIdx = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    
    For x = 1 To NumOfIdx
      Get OHandle, x, OptRec
      IdxArray(x) = OptRec.CustRec
    Next x
    Close OHandle
    IdxFlag = True
    OptFlag = True
  End If

  RptFile$ = "TAXRPTS\MORTCODE.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  frmTaxShowPctComp.Label1 = "Gathering Mortgage Code Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    
    If ThisTownship <> "No Townships Saved" Then
      If ThisTownship <> "All" Then
        If ThisTownship <> QPTrim$(TaxCust.TownShip) Then
          GoTo SkipIt
        End If
      End If
    End If
    MortCnt = 0
    ReDim MortCodes(1 To 1) As String
    ReDim PropPin(1 To 1) As String
    ReDim PropVal(1 To 1) As Double
    ReDim BankName(1 To 1) As String
    If TaxCust.FirstPropRec > 0 Then
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted = -1 Then GoTo NotThisOne
'        'check to be sure that the mortgage code saved is listed
        For z = 1 To MortCodeArrayCnt
          If QPTrim$(RealRec.MORTCODE) = MortCodeArray(z) Then
            Exit For
          End If
        Next z
        If z > MortCodeArrayCnt Then GoTo NotThisOne
        MortCnt = MortCnt + 1
        ReDim Preserve MortCodes(1 To MortCnt) As String
        MortCodes(MortCnt) = QPTrim(RealRec.MORTCODE)
        ReDim Preserve PropPin(1 To MortCnt) As String
        PropPin(MortCnt) = QPTrim$(RealRec.RealPin)
        ReDim Preserve PropVal(1 To MortCnt) As Double
        PropVal(MortCnt) = RealRec.PROPVALU
        ReDim Preserve BankName(1 To MortCnt) As String
        BankName(MortCnt) = MortCodeBank(z)
NotThisOne:
        NextRec = RealRec.NextRec
      Loop
    Else
      GoTo SkipIt
    End If
    
    For y = 1 To MortCnt
      If MORTCODE <> "All" Then
        If MortCodes(y) <> MORTCODE$ Then
          GoTo MoveOn
        Else
'          TotCnt = TotCnt + 1
''          CodeCnt(z) = CodeCnt(z) + 1 'changed to below on 8/29/06
'          CodeCnt(y) = CodeCnt(y) + 1
          For z = 1 To MortCodeArrayCnt 'added on 8/29/06
            If MortCodes(y) = MortCodeArray(z) Then
              TotCnt = TotCnt + 1
              CodeCnt(z) = CodeCnt(z) + 1
            End If
          Next z
        End If
      Else
        For z = 1 To MortCodeArrayCnt
          If MortCodes(y) = MortCodeArray(z) Then
            TotCnt = TotCnt + 1
            CodeCnt(z) = CodeCnt(z) + 1
            Exit For
          End If
        Next z
        If z > MortCodeArrayCnt Then GoTo MoveOn
      End If
      '                   0                     1
      Print #RptHandle, Town$; dlm; QPTrim$(TaxCust.CustName); dlm;
      '                      2                        3
      Print #RptHandle, TaxCust.Acct; dlm; QPTrim$(TaxCust.Addr1); dlm;
      '                           4                     5               6
      Print #RptHandle, QPTrim$(TaxCust.Addr2); dlm; MORTCODE; dlm; BankName(y); dlm;
      '                          7
      Print #RptHandle, QPTrim$(TaxCust.City) + ", " + QPTrim$(TaxCust.State) + "  " + QPTrim$(TaxCust.Zip); dlm;
      '                     8                 9                10               11                12
      Print #RptHandle, MortCodes(y); dlm; PropPin(y); dlm; PropVal(y); dlm; ThisTownship; dlm; TotCnt; dlm;
      If UseOpt = "Y" Then
        '                    13                     14
        Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc)
      Else
        '                 13       14
        Print #RptHandle, ""; dlm; ""
      End If
      PrintCnt = PrintCnt + 1
MoveOn:
    Next y
    
SkipIt:
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  
  SubRptFile$ = "TAXRPTS\SUBMORTCODE.RPT"
  SubRptHandle = FreeFile
  Open SubRptFile For Output As #SubRptHandle
  For x = 1 To MortCodeArrayCnt
    If CodeCnt(x) > 0 Then
      Print #SubRptHandle, CodeCnt(x); dlm; MortCodeArray(x); dlm; MortCodeBank(x)
    End If
  Next x
  
  Close
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  If PrintCnt = 0 Then
    Call TaxMsg(900, "There are no customers using the mortgage code parameters entered.")
    Exit Sub
  End If
  
  arMortCodeRpt.Show
  
  Exit Sub

ERRORSTUFF:
   Unload frmTaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxMortCodeRpt", "PrintGraphics", Erl)
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

Private Sub PrintText()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Integer, z As Integer
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim ThisTownship$
  Dim MORTCODE$
  Dim ThisMortRec As Integer
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim NextRec As Long
  Dim MortCnt As Integer
  Dim ThisMortCode As Integer
  Dim CustRec As Long
  Dim PrintCnt As Long
  Dim TotCnt As Integer
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim Page As Integer
  Dim FF$
  Dim Printed As Boolean
  Dim PntMortCode As String * 8
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  
  'on error goto ERRORSTUFF
  
  Printed = False
  FF$ = Chr(12)
  MaxLines = 58
  OptFlag = False
  ReDim CodeCnt(1 To MortCodeArrayCnt) As Integer
    
  fpcmbMortCodes.Col = 0
  MORTCODE$ = fpcmbMortCodes.ColText
  fpcmbMortCodes.Col = 2
  ThisMortCode = CInt(fpcmbMortCodes.ColText)
  ThisTownship = fpcmbTownship.Text
  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no customers saved."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close IdxHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Search Name" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no search names indexed."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = ThisOpt + " Order" Then
    OpenCustOptSearchFile OHandle, NumOfIdx
    If NumOfIdx = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    
    For x = 1 To NumOfIdx
      Get OHandle, x, OptRec
      IdxArray(x) = OptRec.CustRec
    Next x
    Close OHandle
    IdxFlag = True
    OptFlag = True
  End If

  RptFile$ = "TAXRPTS\MORTCODE.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  frmTaxShowPctComp.Label1 = "Gathering Mortgage Code Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  GoSub PrintHeader
  
  For x = 1 To NumOfTCRecs
    Printed = False
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    
    If ThisTownship <> "No Townships Saved" Then
      If ThisTownship <> "All" Then
        If ThisTownship <> QPTrim$(TaxCust.TownShip) Then
          GoTo SkipIt
        End If
      End If
    End If
    MortCnt = 0
    ReDim MortCodes(1 To 1) As String
    ReDim PropPin(1 To 1) As String
    ReDim PropVal(1 To 1) As Double
    ReDim BankName(1 To 1) As String
    If TaxCust.FirstPropRec > 0 Then
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted = -1 Then GoTo NotThisOne
'        'check to be sure that the mortgage code saved is listed
        For z = 1 To MortCodeArrayCnt
          If QPTrim$(RealRec.MORTCODE) = MortCodeArray(z) Then
            Exit For
          End If
        Next z
        If z > MortCodeArrayCnt Then GoTo NotThisOne
        MortCnt = MortCnt + 1
        ReDim Preserve MortCodes(1 To MortCnt) As String
        MortCodes(MortCnt) = QPTrim(RealRec.MORTCODE)
        ReDim Preserve PropPin(1 To MortCnt) As String
        PropPin(MortCnt) = QPTrim$(RealRec.RealPin)
        ReDim Preserve PropVal(1 To MortCnt) As Double
        PropVal(MortCnt) = RealRec.PROPVALU
        ReDim Preserve BankName(1 To MortCnt) As String
        BankName(MortCnt) = MortCodeBank(z)
NotThisOne:
        NextRec = RealRec.NextRec
      Loop
    Else
      GoTo SkipIt
    End If
    
    For y = 1 To MortCnt
      If MORTCODE <> "All" Then
        If MortCodes(y) <> MORTCODE$ Then
          GoTo MoveOn
        Else
          GoSub PrintCustHeader
          Printed = True
'          TotCnt = TotCnt + 1
''          CodeCnt(z) = CodeCnt(z) + 1 'changed to below on 8/29/06
'          CodeCnt(y) = CodeCnt(y) + 1
          For z = 1 To MortCodeArrayCnt 'added on 8/29/06
            If MortCodes(y) = MortCodeArray(z) Then
              TotCnt = TotCnt + 1
              CodeCnt(z) = CodeCnt(z) + 1
            End If
          Next z
        End If
      Else
        GoSub PrintCustHeader
        Printed = True
        For z = 1 To MortCodeArrayCnt
          If MortCodes(y) = MortCodeArray(z) Then
            TotCnt = TotCnt + 1
            CodeCnt(z) = CodeCnt(z) + 1
            Exit For
          End If
        Next z
        If z > MortCodeArrayCnt Then GoTo MoveOn
      End If
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
      Print #RptHandle, Tab(5); "Property Pin #: "; Tab(23); QPTrim$(PropPin(y)); Tab(47); "Property Value: " + Using$("$###,###,##0.00", PropVal(y))
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
        Printed = False
        GoSub PrintCustHeader
        Printed = True
      End If
      Print #RptHandle, Tab(5); "Mortgage Code: "; Tab(22); QPTrim$(MortCodes(y)); Tab(32); "Bank Name: "; BankName(y)
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
        If y <> MortCnt Then
          Printed = False
          GoSub PrintCustHeader
          Printed = True
        End If
      End If
      PrintCnt = PrintCnt + 1
MoveOn:
    Next y
    If MortCnt > 0 And Printed = True Then
      If LineCnt >= MaxLines - 2 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      ElseIf LineCnt <> 6 Then
        Print #RptHandle, String(77, "-")
        Print #RptHandle,
        LineCnt = LineCnt + 2
      End If
    End If
SkipIt:
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      Exit Sub
    End If

  Next x
  
  If LineCnt >= MaxLines - (MortCodeArrayCnt + 5) Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  If LineCnt <> 6 Then
    Print #RptHandle, String(77, "-")
  End If
  Print #RptHandle, "Summary:"
  Print #RptHandle,
  Print #RptHandle, Tab(6); "Code #"; Tab(15); "Bank Name"; Tab(50); "Count"
  Print #RptHandle, String(77, "-")
  LineCnt = LineCnt + 4
  For x = 1 To MortCodeArrayCnt
    If CodeCnt(x) > 0 Then
      RSet PntMortCode = MortCodeArray(x)
      Print #RptHandle, Tab(4); PntMortCode; Tab(15); MortCodeBank(x); Tab(50); Using$("####0", CodeCnt(x))
      LineCnt = LineCnt + 1
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
        Print #RptHandle, "Summary:"
        Print #RptHandle,
        Print #RptHandle, Tab(5); "Code #"; Tab(15); "Bank Name"; Tab(50); "Count"
        Print #RptHandle, String(77, "-")
        LineCnt = LineCnt + 4
      End If
    End If
  Next x
  If LineCnt >= MaxLines + 2 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  Print #RptHandle, Tab(5); String(50, "-")
  Print #RptHandle, Tab(38); "Total Count:"; Tab(50); Using$("####0", TotCnt)
  
  Print #RptHandle, FF$
  Close
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  If PrintCnt = 0 Then
    Call TaxMsg(900, "There are no customers using the mortgage code parameters entered.")
    Exit Sub
  End If
  
  ViewPrint RptFile, "Tax Mortgage Code Report"
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Mortgage Code Report"
  Print #RptHandle, Town; Tab(65); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Township:"; Tab(12); ThisTownship
  Print #RptHandle, "Mortgage Code:"; Tab(16); MORTCODE
  Print #RptHandle, String(77, "-")
  LineCnt = 6
  
  Return

PrintCustHeader:
  If Printed = True Then Return
  If LineCnt >= MaxLines - 6 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  Print #RptHandle, QPTrim$(TaxCust.CustName); Tab(65); "Acct #: " + Using$("####0", TaxCust.Acct)
  Print #RptHandle, QPTrim$(TaxCust.Addr1)
  Print #RptHandle, QPTrim$(TaxCust.Addr2)
  Print #RptHandle, QPTrim$(TaxCust.City) + ", " + QPTrim$(TaxCust.State) + "  " + QPTrim$(TaxCust.Zip)
  If UseOpt = "Y" Then
    Print #RptHandle, ThisOpt + ": " + QPTrim$(TaxCust.OptSrchDesc)
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, String(77, "-")
  LineCnt = LineCnt + 5
  Return

ERRORSTUFF:
   Unload frmTaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxMortCodeRpt", "PrintText", Erl)
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

