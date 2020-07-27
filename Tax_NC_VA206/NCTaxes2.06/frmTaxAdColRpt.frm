VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmTaxAdColRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Advertising Report"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxAdColRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6750
      Left            =   1920
      TabIndex        =   0
      Top             =   990
      Width           =   7785
      _Version        =   196609
      _ExtentX        =   13732
      _ExtentY        =   11906
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmTaxAdColRpt.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   2928
         TabIndex        =   9
         Top             =   4248
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
         ColDesigner     =   "frmTaxAdColRpt.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   384
         Left            =   2928
         TabIndex        =   10
         Top             =   4836
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
         ColDesigner     =   "frmTaxAdColRpt.frx":0CC1
      End
      Begin LpLib.fpCombo fpcmbTaxYear 
         Height          =   384
         Left            =   2928
         TabIndex        =   7
         Top             =   3684
         Width           =   1380
         _Version        =   196608
         _ExtentX        =   2434
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
         ColDesigner     =   "frmTaxAdColRpt.frx":109C
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00D0D0D0&
         Caption         =   "*Select Revenues to Include In Tax Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   1080
         TabIndex        =   16
         Top             =   1320
         Width           =   5775
         Begin VB.CheckBox chkOpt3 
            BackColor       =   &H00D0D0D0&
            Caption         =   "Opt3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   8
            Top             =   1320
            Width           =   3375
         End
         Begin VB.CheckBox chkOpt2 
            BackColor       =   &H00D0D0D0&
            Caption         =   "Opt2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   6
            Top             =   960
            Width           =   3375
         End
         Begin VB.CheckBox chkOpt1 
            BackColor       =   &H00D0D0D0&
            Caption         =   "Opt1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   5
            Top             =   600
            Width           =   3375
         End
         Begin VB.CheckBox chkLateList 
            BackColor       =   &H00D0D0D0&
            Caption         =   "Late Listing"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   2
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox chkAdv 
            BackColor       =   &H00D0D0D0&
            Caption         =   "Advertising"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   4
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CheckBox chkInt 
            BackColor       =   &H00D0D0D0&
            Caption         =   "Interest"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   3
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox chkPrinc 
            BackColor       =   &H00D0D0D0&
            Caption         =   "Principle"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   1
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00D0D0D0&
            Caption         =   "Optional Revenues"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   17
            Top             =   360
            Width           =   1860
         End
         Begin VB.Shape Shape5 
            Height          =   1215
            Left            =   2040
            Top             =   480
            Width           =   3615
         End
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   636
         Left            =   1800
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   5760
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
         _ExtentY        =   1122
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
         ButtonDesigner  =   "frmTaxAdColRpt.frx":1477
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   636
         Left            =   4560
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   $"frmTaxAdColRpt.frx":1655
         Top             =   5760
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
         _ExtentY        =   1122
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
         ButtonDesigner  =   "frmTaxAdColRpt.frx":1700
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Year:"
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
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   3755
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   2100
         Left            =   1005
         Top             =   1200
         Width           =   5970
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   2100
         Left            =   1005
         Top             =   3405
         Width           =   5970
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
         TabIndex        =   15
         Top             =   4920
         Width           =   1305
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Advertising Report"
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   4350
         Width           =   1500
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   0  'Transparent
      Caption         =   "*Report Includes Real Estate Only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   1800
      TabIndex        =   18
      Top             =   8040
      Width           =   3060
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   7020
      Left            =   1800
      Top             =   855
      Width           =   8055
   End
End
Attribute VB_Name = "frmTaxAdColRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim UseOpt As String * 1
  Dim ThisOpt$
  Dim Town$
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$

Private Sub cmdExit_Click()
  frmTaxReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  If fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  Else
    frmTaxMsg.Label1.Caption = "Pitch 17 is recommended for this printout."
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
  Me.HelpContextID = hlpPrintAdvertising
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxAdColRpt.")
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
  Dim TSCnt As Integer
  Dim TSHandle As Integer
  Dim BigYr As Integer
  Dim ThisBigYr As Integer
  Dim HoldYr As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
  Dim YrCnt As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Integer
  
  'on error goto ERRORSTUFF
  
  frmTaxLoadReport.Label1.Caption = "Loading Years"
  frmTaxLoadReport.Show
  DoEvents
  ReDim Years(1 To 1) As Integer
  YrCnt = 0
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If YrCnt = 0 Then
      If TaxTrans.TaxYear > 0 Then
        YrCnt = YrCnt + 1
        ReDim Preserve Years(1 To YrCnt) As Integer
        Years(YrCnt) = TaxTrans.TaxYear
      End If
    Else
      For y = 1 To YrCnt
        If TaxTrans.TaxYear = Years(y) Then
          Exit For
        End If
      Next y
      If y > YrCnt Then
        YrCnt = YrCnt + 1
        ReDim Preserve Years(1 To YrCnt) As Integer
        Years(YrCnt) = TaxTrans.TaxYear
      End If
    End If
  Next x
  Close TTHandle
  
  
  BigYr = 0
  For x = 1 To YrCnt
    If Years(x) > BigYr Then
      BigYr = Years(x)
    End If
  Next x
  
  Nextx = 1
  ThisBigYr = BigYr + 1
  Do While Nextx <= YrCnt
    For x = Nextx To YrCnt
      If Years(x) < ThisBigYr Then
        ThisBigYr = Years(x)
        Thisx = x
      End If
    Next x
    HoldYr = Years(Nextx)
    Years(Nextx) = Years(Thisx)
    Years(Thisx) = HoldYr
    Nextx = Nextx + 1
    ThisBigYr = BigYr + 1
  Loop
    
  fpcmbTaxYear.Text = "All"
  fpcmbTaxYear.AddItem "All"
  
  For x = YrCnt To 1 Step -1
    fpcmbTaxYear.AddItem CStr(Years(x))
  Next x
  
  Unload frmTaxLoadReport
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town = QPTrim$(TaxMasterRec.Name)
  Opt1Desc = QPTrim$(TaxMasterRec.OptRev1)
  Opt2Desc = QPTrim$(TaxMasterRec.OptRev2)
  Opt3Desc = QPTrim$(TaxMasterRec.OptRev3)
  If Opt1Desc <> "" Then
    chkOpt1.Caption = Opt1Desc
    chkOpt1.Enabled = True
  Else
    chkOpt1.Caption = "Not Being Used"
    chkOpt1.Enabled = False
  End If
  If Opt2Desc <> "" Then
    chkOpt2.Caption = Opt2Desc
    chkOpt2.Enabled = True
  Else
    chkOpt2.Caption = "Not Being Used"
    chkOpt2.Enabled = False
  End If
  If Opt3Desc <> "" Then
    chkOpt3.Caption = Opt3Desc
    chkOpt3.Enabled = True
  Else
    chkOpt3.Caption = "Not Being Used"
    chkOpt3.Enabled = False
  End If
  chkPrinc.Value = 1
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
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxAdColRpt", "LoadMe", Erl)
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
      chkPrinc.SetFocus
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
'  If KeyCode = vbKeySpace Then
'    fpcmbTownship.ListDown = True
'  End If
'  If KeyCode = vbKeyDelete Then
'    fpcmbTownship.ListIndex = -1
'  End If
'  If fpcmbTownship.ListDown <> True Then
'    If KeyCode = vbKeyDown Then
'      fpcmbPrintOpt.SetFocus
'      KeyCode = 0
'    Else
'      If KeyCode = vbKeyUp Then
'        SendKeys "+{Tab}"
'        KeyCode = 0
'      End If
'    End If
'  End If

End Sub

Private Sub PrintGraphics()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim dlm$
  Dim InactiveFlag As Boolean
  Dim x As Long, y As Long
  Dim NextRec As Long
  Dim PropAdd$, PropTownShip$
  Dim CustRec As Long
  Dim CustName$
  Dim CustAcct$
  Dim ThisTownship$
  Dim RealTotVal As Double
  Dim PersTotVal As Double
  Dim TotVal As Double
  Dim TotLLCnt As Long
  Dim TotRealLLCnt As Long
  Dim TotPersLLCnt As Long
  Dim ThisPersVal As Double
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Balance As Double
  Dim RecCnt As Integer
  Dim NextPropRec As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim Charged#
  Dim Paid#
  Dim UnknownCnt As Long
  Dim TaxYear$
  Dim ThisTaxYear As Integer
  Dim PrintCnt As Long
  Dim Revenues$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  
  'on error goto ERRORSTUFF
  
  TaxYear = fpcmbTaxYear.Text
  OptFlag = False
  If chkPrinc.Value = 1 Then
    Revenues$ = "Princ"
  End If
  If chkInt.Value = 1 Then
    If Len(Revenues) > 0 Then
      Revenues = Revenues + "/Int"
    Else
      Revenues = "Int"
    End If
  End If
  If chkAdv.Value = 1 Then
    If Len(Revenues) > 0 Then
      Revenues = Revenues + "/Adv"
    Else
      Revenues = "Adv"
    End If
  End If
  If chkLateList.Value = 1 Then
    If Len(Revenues) > 0 Then
      Revenues = Revenues + "/Late List"
    Else
      Revenues = "Late List"
    End If
  End If
  If chkOpt1.Value = 1 Then
    If Len(Revenues) > 0 Then
      Revenues = Revenues + "/" + Mid(Opt1Desc, 1, 5)
    Else
      Revenues = Mid(Opt1Desc, 1, 5)
    End If
  End If
  If chkOpt2.Value = 1 Then
    If Len(Revenues) > 0 Then
      Revenues = Revenues + "/" + Mid(Opt2Desc, 1, 5)
    Else
      Revenues = Mid(Opt2Desc, 1, 5)
    End If
  End If
  If chkOpt3.Value = 1 Then
    If Len(Revenues) > 0 Then
      Revenues = Revenues + "/" + Mid(Opt3Desc, 1, 5)
    Else
      Revenues = Mid(Opt3Desc, 1, 5)
    End If
  End If
  
  IdxFlag = False
  dlm$ = "~"
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

  RptFile$ = "TAXRPTS\ADVLIST.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  frmTaxShowPctComp.Label1 = "Gathering Advertising Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  ReDim UnknownName(1 To 1) As String
  ReDim UnknownDesc(1 To 1) As String
  ReDim UnknownAmt(1 To 1) As Double
  ReDim UnknownAcct(1 To 1) As String
  ReDim UnknownPin(1 To 1) As String
  
  UnknownCnt = 0
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    CustName = QPTrim$(TaxCust.CustName)
    CustAcct = CStr(TaxCust.Acct)
    NextRec = TaxCust.LastTrans
    If NextRec = 0 Then GoTo SkipIt
    RecCnt = 0
    ReDim ThisPropRec(1 To 1) As Long
    ReDim ThisPropDesc(1 To 1) As String
    ReDim ThisPropPin(1 To 1) As String
    
    NextPropRec = TaxCust.FirstPropRec
    If NextPropRec > 0 Then
      Do While NextPropRec > 0
        Get RHandle, NextPropRec, RealRec
        RecCnt = RecCnt + 1
        ReDim Preserve ThisPropRec(1 To RecCnt) As Long
        ThisPropRec(RecCnt) = NextPropRec
        ReDim Preserve ThisPropDesc(1 To RecCnt) As String
        If QPTrim$(RealRec.PropAddr) <> "" Then
          ThisPropDesc(RecCnt) = QPTrim$(RealRec.PropAddr)
        ElseIf QPTrim$(RealRec.PROPNOT1) <> "" Then
          ThisPropDesc(RecCnt) = QPTrim$(RealRec.PROPNOT1)
        Else
          ThisPropDesc(RecCnt) = "No Description Available"
        End If
        ReDim Preserve ThisPropPin(1 To RecCnt) As String
        ThisPropPin(RecCnt) = QPTrim$(RealRec.RealPin)
        NextPropRec = RealRec.NextRec
      Loop
    Else
      GoTo SkipIt
    End If
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
      If TaxYear = "All" Then
        ThisTaxYear = TaxTrans.TaxYear
      Else
        ThisTaxYear = CInt(TaxYear)
      End If
      TaxTrans.TranType = TaxTrans.TranType
      If TaxTrans.TranType = 1 And TaxTrans.TaxYear = ThisTaxYear Then
        If chkPrinc.Value = 0 Then
          TaxTrans.Revenue.Principle1 = 0
          TaxTrans.Revenue.Principle1Pd = 0
          TaxTrans.Revenue.Principle2 = 0
          TaxTrans.Revenue.Principle2Pd = 0
          TaxTrans.Revenue.Principle3 = 0
          TaxTrans.Revenue.Principle3Pd = 0
          TaxTrans.Revenue.Principle4 = 0
          TaxTrans.Revenue.Principle4Pd = 0
          TaxTrans.Revenue.Principle5 = 0
          TaxTrans.Revenue.Principle5Pd = 0
        End If
        If chkInt.Value = 0 Then
          TaxTrans.Revenue.Interest = 0
          TaxTrans.Revenue.InterestPd = 0
        End If
        If chkAdv.Value = 0 Then
          TaxTrans.Revenue.Collection = 0
          TaxTrans.Revenue.CollectionPd = 0
        End If
        If chkLateList.Value = 0 Then
          TaxTrans.Revenue.LateList = 0
          TaxTrans.Revenue.LateListPd = 0
        End If
        If chkOpt1.Enabled = True Then
          If chkOpt1.Value = 0 Then
            TaxTrans.Revenue.RevOpt1 = 0
            TaxTrans.Revenue.RevOpt1Pd = 0
          End If
        Else
          TaxTrans.Revenue.RevOpt1 = 0
          TaxTrans.Revenue.RevOpt1Pd = 0
        End If
        If chkOpt2.Enabled = True Then
          If chkOpt2.Value = 0 Then
            TaxTrans.Revenue.RevOpt2 = 0
            TaxTrans.Revenue.RevOpt2Pd = 0
          End If
        Else
          TaxTrans.Revenue.RevOpt2 = 0
          TaxTrans.Revenue.RevOpt2Pd = 0
        End If
        If chkOpt3.Enabled = True Then
          If chkOpt3.Value = 0 Then
            TaxTrans.Revenue.RevOpt3 = 0
            TaxTrans.Revenue.RevOpt3Pd = 0
          End If
        Else
          TaxTrans.Revenue.RevOpt3 = 0
          TaxTrans.Revenue.RevOpt3Pd = 0
        End If
          
        Charged# = OldRound#(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
        Charged# = OldRound#(Charged# + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.Interest)
        Charged# = OldRound(Charged# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        Charged# = OldRound(Charged# + TaxTrans.Revenue.RevOpt3)
        Paid# = OldRound#(TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd)
        Paid# = OldRound#(Paid# + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.InterestPd)
        Paid# = OldRound(Paid# + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd)
        Paid# = OldRound(Paid# + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt) 'added .DiscAmt on 2/2/07)
        Balance# = OldRound#(Charged# - Paid#)
        If Balance# <= 0 Then GoTo ZeroBalance
        
        If TaxCust.FirstPersRec = 0 Then 'we know that this customer owns no personal
        'property and this transaction is definitely for real estate...although, without a real pin
        'number we can't be sure of which property this is if there are multiple listings
          If QPTrim$(TaxTrans.RealPin) = "0" Then
            PrintCnt = PrintCnt + 1
            'this is from old trans recs and uses the description of the last property
            'for any property past the first one...best that can be done since no link exists
            '                   0            1                2
            Print #RptHandle, Town; dlm; CustName; dlm; ThisPropDesc(RecCnt); dlm;
            '                    3             4             5              6
            Print #RptHandle, Balance; dlm; "Known"; dlm; TaxYear; dlm; Revenues; dlm;
            If UseOpt = "Y" Then
              '                    7                       8                       9                     10
              Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); ThisPropPin(RecCnt); dlm; CustAcct
            Else
              '                  7        8               9                   10
              Print #RptHandle, ""; dlm; ""; dlm; ThisPropPin(RecCnt); dlm; CustAcct
            End If
          Else
            For y = 1 To RecCnt
              If QPTrim$(TaxTrans.RealPin) = "0" Then
                GoTo NotSure
              End If
              If QPTrim$(TaxTrans.RealPin) = ThisPropPin(y) Then
                PrintCnt = PrintCnt + 1
                '                   0            1           2
                Print #RptHandle, Town; dlm; CustName; dlm; ThisPropDesc(y); dlm;
                '                    3             4             5             6
                Print #RptHandle, Balance; dlm; "Known"; dlm; TaxYear; dlm; Revenues; dlm;
                If UseOpt = "Y" Then
                  '                    7                       8                          9                 10
                  Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); dlm; ThisPropPin(y); dlm; CustAcct
                Else
                  '                  7        8              9                10
                  Print #RptHandle, ""; dlm; ""; dlm; ThisPropPin(y); dlm; CustAcct
                End If
                Exit For
              End If
NotSure:
            Next y
            If y > RecCnt Then
              UnknownCnt = UnknownCnt + 1
              ReDim Preserve UnknownName(1 To UnknownCnt) As String
              UnknownName(UnknownCnt) = QPTrim(TaxCust.CustName) + "/" + CStr(TaxCust.Acct)
              ReDim Preserve UnknownDesc(1 To UnknownCnt) As String
              UnknownDesc(UnknownCnt) = "" 'ThisPropDesc(RecCnt)
              ReDim Preserve UnknownAmt(1 To UnknownCnt) As Double
              UnknownAmt(UnknownCnt) = Balance
              ReDim Preserve UnknownAcct(1 To UnknownCnt) As String
              UnknownAcct(UnknownCnt) = CustAcct
              ReDim Preserve UnknownPin(1 To UnknownCnt) As String
              UnknownPin(UnknownCnt) = QPTrim$(TaxTrans.RealPin)
            End If
          End If
        Else
'          If TaxTrans.CustomerRec = 7039 Then
'            Stop
'          End If
          For y = 1 To RecCnt
            If QPTrim$(TaxTrans.RealPin) = "0" Then
              GoTo NotSure2
            End If
            If QPTrim$(TaxTrans.RealPin) = ThisPropPin(y) Then
              PrintCnt = PrintCnt + 1
              '                   0            1           2
              Print #RptHandle, Town; dlm; CustName; dlm; ThisPropDesc(y); dlm;
              '                    3             4             5             6
              Print #RptHandle, Balance; dlm; "Known"; dlm; TaxYear; dlm; Revenues; dlm;
              If UseOpt = "Y" Then
                '                    7                       8                             9                    10
                Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); dlm; ThisPropPin(RecCnt); dlm; CustAcct
              Else
                '                  7        8            9                  10
                Print #RptHandle, ""; dlm; ""; dlm; ThisPropPin(y); dlm; CustAcct
              End If
              Exit For
            End If
NotSure2:
          Next y
          If y > RecCnt Then
            UnknownCnt = UnknownCnt + 1
            ReDim Preserve UnknownName(1 To UnknownCnt) As String
            UnknownName(UnknownCnt) = QPTrim(TaxCust.CustName) + "/" + CStr(TaxCust.Acct)
            ReDim Preserve UnknownDesc(1 To UnknownCnt) As String
            UnknownDesc(UnknownCnt) = "" 'ThisPropDesc(RecCnt)
            ReDim Preserve UnknownAmt(1 To UnknownCnt) As Double
            UnknownAmt(UnknownCnt) = Balance
            ReDim Preserve UnknownAcct(1 To UnknownCnt) As String
            UnknownAcct(UnknownCnt) = CustAcct
            ReDim Preserve UnknownPin(1 To UnknownCnt) As String
            UnknownPin(UnknownCnt) = QPTrim$(TaxTrans.RealPin)
          End If
        End If
      End If
ZeroBalance:
      NextRec = TaxTrans.LastTrans
    Loop

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
  
  For x = 1 To UnknownCnt
    PrintCnt = PrintCnt + 1
    '                   0              1                      2                 3                  4              5              6
    Print #RptHandle, Town; dlm; UnknownName(x); dlm; UnknownDesc(x); dlm; UnknownAmt(x); dlm; "Unknown"; dlm; TaxYear; dlm; Revenues; dlm;
    If UseOpt = "Y" Then
      '                    7                       8                          9                  10
      Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); dlm; UnknownPin(x); dlm; UnknownAcct(x)
    Else
      '                  7        8           9                   10
      Print #RptHandle, ""; dlm; ""; dlm; UnknownPin(x); dlm; UnknownAcct(x)
    End If
  Next x
  
  Close
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  If PrintCnt = 0 Then
    Call TaxMsg(900, "There are no advertising listings for the parameters entered.")
    Exit Sub
  End If
  
  arTaxAdvRpt.Show
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxAdColRpt", "PrintGraphics", Erl)
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
  
End Sub

Private Sub PrintText()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim InactiveFlag As Boolean
  Dim x As Long, y As Long
  Dim NextRec As Long
  Dim PropAdd$, PropTownShip$
  Dim CustRec As Long
  Dim CustName$
  Dim CustAcct$
  Dim ThisTownship$
  Dim RealTotVal As Double
  Dim PersTotVal As Double
  Dim TotVal As Double
  Dim TotLLCnt As Long
  Dim TotRealLLCnt As Long
  Dim TotPersLLCnt As Long
  Dim ThisPersVal As Double
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Balance As Double
  Dim RecCnt As Integer
  Dim NextPropRec As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim Charged#
  Dim Paid#
  Dim UnknownCnt As Long
  Dim TaxYear$
  Dim ThisTaxYear As Integer
  Dim PrintCnt As Long
  Dim FF$
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim Page As Integer
  Dim Revenues$
  Dim ThisDesc As String * 30
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  
  'on error goto ERRORSTUFF
  
  OptFlag = False
  TaxYear = fpcmbTaxYear.Text
  MaxLines = 58
  FF$ = Chr(12)
  If chkPrinc.Value = 1 Then
    Revenues$ = "Princ"
  End If
  If chkInt.Value = 1 Then
    If Len(Revenues) > 0 Then
      Revenues = Revenues + "/Int"
    Else
      Revenues = "Int"
    End If
  End If
  If chkAdv.Value = 1 Then
    If Len(Revenues) > 0 Then
      Revenues = Revenues + "/Adv"
    Else
      Revenues = "Adv"
    End If
  End If
  If chkLateList.Value = 1 Then
    If Len(Revenues) > 0 Then
      Revenues = Revenues + "/Late List"
    Else
      Revenues = "Late List"
    End If
  End If
  If chkOpt1.Value = 1 Then
    If Len(Revenues) > 0 Then
      Revenues = Revenues + "/" + Mid(Opt1Desc, 1, 5)
    Else
      Revenues = Mid(Opt1Desc, 1, 5)
    End If
  End If
  If chkOpt2.Value = 1 Then
    If Len(Revenues) > 0 Then
      Revenues = Revenues + "/" + Mid(Opt2Desc, 1, 5)
    Else
      Revenues = Mid(Opt2Desc, 1, 5)
    End If
  End If
  If chkOpt3.Value = 1 Then
    If Len(Revenues) > 0 Then
      Revenues = Revenues + "/" + Mid(Opt3Desc, 1, 5)
    Else
      Revenues = Mid(Opt3Desc, 1, 5)
    End If
  End If
  
  IdxFlag = False
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

  RptFile$ = "TAXRPTS\ADVLIST.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle

  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenRealPropFile RHandle, NumOfRealRecs

  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If

  frmTaxShowPctComp.Label1 = "Gathering Advertising Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False

  ReDim UnknownName(1 To 1) As String
  ReDim UnknownDesc(1 To 1) As String
  ReDim UnknownAmt(1 To 1) As Double
  ReDim UnknownAcct(1 To 1) As String
  ReDim UnknownPin(1 To 1) As String
  
  UnknownCnt = 0
  GoSub PrintHeader
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    CustName = QPTrim$(TaxCust.CustName)
    CustAcct = CStr(TaxCust.Acct)
    NextRec = TaxCust.LastTrans
    If NextRec = 0 Then GoTo SkipIt
    RecCnt = 0
    ReDim ThisPropRec(1 To 1) As Long
    ReDim ThisPropDesc(1 To 1) As String
    ReDim ThisPropPin(1 To 1) As String
    NextPropRec = TaxCust.FirstPropRec
    If NextPropRec > 0 Then
      Do While NextPropRec > 0
        Get RHandle, NextPropRec, RealRec
        RecCnt = RecCnt + 1
        ReDim Preserve ThisPropRec(1 To RecCnt) As Long
        ThisPropRec(RecCnt) = NextPropRec
        ReDim Preserve ThisPropDesc(1 To RecCnt) As String
        If QPTrim$(RealRec.PropAddr) <> "" Then
          ThisDesc = QPTrim$(RealRec.PropAddr)
          ThisPropDesc(RecCnt) = ThisDesc
        ElseIf QPTrim$(RealRec.PROPNOT1) <> "" Then
          ThisDesc = QPTrim$(RealRec.PROPNOT1)
          ThisPropDesc(RecCnt) = ThisDesc
        Else
          ThisPropDesc(RecCnt) = "No Description Available"
        End If
        ReDim Preserve ThisPropPin(1 To RecCnt) As String
        ThisPropPin(RecCnt) = QPTrim$(RealRec.RealPin)
        NextPropRec = RealRec.NextRec
      Loop
    Else
      GoTo SkipIt
    End If
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
      If TaxYear = "All" Then
        ThisTaxYear = TaxTrans.TaxYear
      Else
        ThisTaxYear = CInt(TaxYear)
      End If
      If TaxTrans.TranType = 1 And TaxTrans.TaxYear = ThisTaxYear Then
        If chkPrinc.Value = 0 Then
          TaxTrans.Revenue.Principle1 = 0
          TaxTrans.Revenue.Principle1Pd = 0
          TaxTrans.Revenue.Principle2 = 0
          TaxTrans.Revenue.Principle2Pd = 0
          TaxTrans.Revenue.Principle3 = 0
          TaxTrans.Revenue.Principle3Pd = 0
          TaxTrans.Revenue.Principle4 = 0
          TaxTrans.Revenue.Principle4Pd = 0
          TaxTrans.Revenue.Principle5 = 0
          TaxTrans.Revenue.Principle5Pd = 0
        End If
        If chkInt.Value = 0 Then
          TaxTrans.Revenue.Interest = 0
          TaxTrans.Revenue.InterestPd = 0
        End If
        If chkAdv.Value = 0 Then
          TaxTrans.Revenue.Collection = 0
          TaxTrans.Revenue.CollectionPd = 0
        End If
        If chkLateList.Value = 0 Then
          TaxTrans.Revenue.LateList = 0
          TaxTrans.Revenue.LateListPd = 0
        End If
        If chkOpt1.Enabled = True Then
          If chkOpt1.Value = 0 Then
            TaxTrans.Revenue.RevOpt1 = 0
            TaxTrans.Revenue.RevOpt1Pd = 0
          End If
        Else
          TaxTrans.Revenue.RevOpt1 = 0
          TaxTrans.Revenue.RevOpt1Pd = 0
        End If
        If chkOpt2.Enabled = True Then
          If chkOpt2.Value = 0 Then
            TaxTrans.Revenue.RevOpt2 = 0
            TaxTrans.Revenue.RevOpt2Pd = 0
          End If
        Else
          TaxTrans.Revenue.RevOpt2 = 0
          TaxTrans.Revenue.RevOpt2Pd = 0
        End If
        If chkOpt3.Enabled = True Then
          If chkOpt3.Value = 0 Then
            TaxTrans.Revenue.RevOpt3 = 0
            TaxTrans.Revenue.RevOpt3Pd = 0
          End If
        Else
          TaxTrans.Revenue.RevOpt3 = 0
          TaxTrans.Revenue.RevOpt3Pd = 0
        End If

        Charged# = OldRound#(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
        Charged# = OldRound#(Charged# + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.Interest)
        Charged# = OldRound(Charged# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        Charged# = OldRound(Charged# + TaxTrans.Revenue.RevOpt3)
        Paid# = OldRound#(TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd)
        Paid# = OldRound#(Paid# + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.InterestPd)
        Paid# = OldRound(Paid# + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd)
        Paid# = OldRound(Paid# + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt) 'added .DiscAmt on 2/2/07)
        Balance# = OldRound#(Charged# - Paid#)
        If Balance# <= 0 Then GoTo ZeroBalance

        If TaxCust.FirstPersRec = 0 Then 'we know that this customer owns no personal
        'property and this transaction is definitely for real estate...although, without a real pin
        'number we can't be sure of which property this is if there are multiple listings
          If QPTrim$(TaxTrans.RealPin) = "0" Then 'this is a DOS transaction
            PrintCnt = PrintCnt + 1
            'this is from old trans recs and uses the description of the last property
            'for any property past the first one...best that can be done since no link exists
            If LineCnt >= MaxLines - 2 Then
              Print #RptHandle, FF$
              GoSub PrintHeader
            End If
            Print #RptHandle, CustAcct; Tab(11); CustName; Tab(61); ThisPropDesc(RecCnt); Tab(96); ThisPropPin(RecCnt); Tab(116); Using$("$###,###,##0.00", Balance#)
            LineCnt = LineCnt + 1
            If UseOpt = "Y" Then
              Print #RptHandle, ThisOpt + ": " + QPTrim$(TaxCust.OptSrchDesc)
              LineCnt = LineCnt + 1
            End If
          Else
            For y = 1 To RecCnt
              If QPTrim$(TaxTrans.RealPin) = "0" Then GoTo NotSure
              If QPTrim$(TaxTrans.RealPin) = ThisPropPin(y) Then
                PrintCnt = PrintCnt + 1
                If LineCnt >= MaxLines - 2 Then
                  Print #RptHandle, FF$
                  GoSub PrintHeader
                End If
                Print #RptHandle, CustAcct; Tab(11); CustName; Tab(61); ThisPropDesc(y); Tab(96); ThisPropPin(RecCnt); Tab(116); Using$("$###,###,##0.00", Balance#)
                LineCnt = LineCnt + 1
                If UseOpt = "Y" Then
                  Print #RptHandle, ThisOpt + ": " + QPTrim$(TaxCust.OptSrchDesc)
                  LineCnt = LineCnt + 1
                End If
                Exit For
              End If
NotSure:
            Next y
            If y > RecCnt Then
              UnknownCnt = UnknownCnt + 1
              ReDim Preserve UnknownName(1 To UnknownCnt) As String
              UnknownName(UnknownCnt) = QPTrim(TaxCust.CustName)
              ReDim Preserve UnknownDesc(1 To UnknownCnt) As String
              UnknownDesc(UnknownCnt) = "" 'ThisPropDesc(RecCnt)
              ReDim Preserve UnknownAmt(1 To UnknownCnt) As Double
              UnknownAmt(UnknownCnt) = Balance
              ReDim Preserve UnknownAcct(1 To UnknownCnt) As String
              UnknownAcct(UnknownCnt) = CustAcct
              ReDim Preserve UnknownPin(1 To UnknownCnt) As String
              UnknownPin(UnknownCnt) = QPTrim$(TaxTrans.RealPin)
            End If
          End If
        Else
          For y = 1 To RecCnt
            If QPTrim$(TaxTrans.RealPin) = "0" Then GoTo NotSure2
            If QPTrim$(TaxTrans.RealPin) = ThisPropPin(y) Then
              PrintCnt = PrintCnt + 1
              If LineCnt >= MaxLines - 2 Then
                Print #RptHandle, FF$
                GoSub PrintHeader
              End If
              Print #RptHandle, CustAcct; Tab(11); CustName; Tab(61); ThisPropDesc(y); Tab(96); ThisPropPin(y); Tab(116); Using$("$###,###,##0.00", Balance#)
              LineCnt = LineCnt + 1
              If UseOpt = "Y" Then
                Print #RptHandle, ThisOpt + ": " + QPTrim$(TaxCust.OptSrchDesc)
                LineCnt = LineCnt + 1
              End If
              Exit For
            End If
NotSure2:
          Next y
          If y > RecCnt Then
            UnknownCnt = UnknownCnt + 1
            ReDim Preserve UnknownName(1 To UnknownCnt) As String
            UnknownName(UnknownCnt) = QPTrim(TaxCust.CustName)
            ReDim Preserve UnknownDesc(1 To UnknownCnt) As String
            UnknownDesc(UnknownCnt) = ""
            ReDim Preserve UnknownAmt(1 To UnknownCnt) As Double
            UnknownAmt(UnknownCnt) = Balance
            ReDim Preserve UnknownAcct(1 To UnknownCnt) As String
            UnknownAcct(UnknownCnt) = CustAcct
            ReDim Preserve UnknownPin(1 To UnknownCnt) As String
            UnknownPin(UnknownCnt) = QPTrim$(TaxTrans.RealPin)
          End If
        End If
      End If
ZeroBalance:
      NextRec = TaxTrans.LastTrans
    Loop

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

  If UnknownCnt > 0 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    Print #RptHandle, "***The following listings could not be positively identified as either "
    Print #RptHandle, "Real Property or Personal Property.***"
    Print #RptHandle, String(130, "-")
    LineCnt = LineCnt + 3
  End If
  For x = 1 To UnknownCnt
    PrintCnt = PrintCnt + 1
    If LineCnt >= MaxLines - 2 Then
      Print #RptHandle, FF$
      GoSub PrintHeader
      Print #RptHandle, "***The following listings could not be positively identified as either "
      Print #RptHandle, "Real Property or Personal Property.***"
      Print #RptHandle, String(130, "-")
      LineCnt = LineCnt + 3
    End If
    Print #RptHandle, UnknownAcct(x); Tab(11); UnknownName(x); Tab(61); UnknownDesc(x); Tab(96); UnknownPin(x); Tab(116); Using$("$###,###,##0.00", UnknownAmt(x))
    LineCnt = LineCnt + 1
    If UseOpt = "Y" Then
      Print #RptHandle, ThisOpt + ": " + QPTrim$(TaxCust.OptSrchDesc)
      LineCnt = LineCnt + 1
    End If
  Next x

  Print #RptHandle, FF$
  Close
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  If PrintCnt = 0 Then
    Call TaxMsg(900, "There are no advertising listings for the parameters entered.")
    Exit Sub
  End If

  ViewPrint RptFile, "Tax Advertising List Report", True

  Exit Sub

PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Advertising List Report"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Tax Year:"; Tab(12); CStr(TaxYear)
  Print #RptHandle, "Revenues Included:"; Tab(20); Revenues
  Print #RptHandle, "Cust Acct"; Tab(11); "Customer Name"; Tab(61); "Property Description"; Tab(96); "Pin #"; Tab(125); "Amount"
  Print #RptHandle, String(130, "-")
  LineCnt = 7

  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxAdColRpt", "PrintText", Erl)
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
  
End Sub

Private Sub fpcmbTaxYear_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTaxYear.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTaxYear.ListIndex = -1
  End If
  If fpcmbTaxYear.ListDown <> True Then
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

