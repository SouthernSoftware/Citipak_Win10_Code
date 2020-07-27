VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLGLList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Of General Ledger Numbers"
   ClientHeight    =   7125
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8400
   Icon            =   "frmBLGLList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin LpLib.fpList fpList1 
      Height          =   3375
      Left            =   1740
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2010
      Width           =   4905
      _Version        =   196608
      _ExtentX        =   8652
      _ExtentY        =   5953
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      Columns         =   2
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
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
      BorderColor     =   0
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
      DataField       =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "frmBLGLList.frx":08CA
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008F8265&
      Caption         =   "G/L Number Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1050
      Left            =   528
      TabIndex        =   10
      Top             =   840
      Width           =   7404
      Begin EditLib.fpText fptxtRev 
         Height          =   420
         Left            =   285
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   525
         Width           =   1890
         _Version        =   196608
         _ExtentX        =   3334
         _ExtentY        =   741
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.5
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
         ControlType     =   1
         Text            =   ""
         CharValidationText=   ""
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
      Begin VB.OptionButton optCash 
         BackColor       =   &H008F8265&
         Caption         =   "Cash Receipt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   252
         Left            =   5376
         TabIndex        =   2
         ToolTipText     =   "Press F3 to bring up assistance for this field."
         Top             =   240
         Width           =   1548
      End
      Begin VB.OptionButton optAccts 
         BackColor       =   &H008F8265&
         Caption         =   "Accts Receivable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   252
         Left            =   2760
         TabIndex        =   1
         ToolTipText     =   "Press F3 to bring up assistance for this field."
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton optRev 
         BackColor       =   &H008F8265&
         Caption         =   "Revenue "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   252
         Left            =   672
         TabIndex        =   0
         ToolTipText     =   "Press F3 to bring up assistance for this field."
         Top             =   240
         Width           =   1164
      End
      Begin EditLib.fpText fptxtAccts 
         Height          =   420
         Left            =   2730
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   525
         Width           =   1890
         _Version        =   196608
         _ExtentX        =   3334
         _ExtentY        =   741
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.5
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
         ControlType     =   1
         Text            =   ""
         CharValidationText=   ""
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
      Begin EditLib.fpText fptxtCash 
         Height          =   420
         Left            =   5235
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   525
         Width           =   1890
         _Version        =   196608
         _ExtentX        =   3334
         _ExtentY        =   741
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.5
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
         ControlType     =   1
         Text            =   ""
         CharValidationText=   ""
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
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClose 
      Height          =   480
      Left            =   3127
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5850
      Width           =   2175
      _Version        =   131072
      _ExtentX        =   3836
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmBLGLList.frx":0CA2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   480
      Left            =   615
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5850
      Width           =   2175
      _Version        =   131072
      _ExtentX        =   3836
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmBLGLList.frx":0E80
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdApply 
      Height          =   480
      Left            =   5640
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5850
      Width           =   2160
      _Version        =   131072
      _ExtentX        =   3810
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmBLGLList.frx":1063
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   624
      TabIndex        =   12
      Top             =   5376
      Width           =   780
      _Version        =   131072
      _ExtentX        =   1376
      _ExtentY        =   783
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   9405029
      ForeColor       =   -2147483634
      Text            =   ""
      Shape           =   0
      ShapeRoundWidth =   195
      ShapeRoundHeight=   195
      BorderWidth     =   -1
      BorderColor     =   -2147483630
      ThreeDWidth     =   -1
      ThreeDShadowColor=   -2147483632
      ThreeDHighlightColor=   16777215
      ThreeDText      =   0
      ThreeDTextHighlightColor=   16777215
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignPictureH   =   0
      AlignPictureV   =   0
      PictureStyle    =   0
      WordWrap        =   -1  'True
      ScaleMode       =   1
      ThreeDStyle     =   2
      Position        =   0
      PosBaseX        =   0
      PosBaseY        =   0
      PosOffsetX      =   -100
      PosOffsetY      =   300
      MaxWidth        =   3000
      CloudInset      =   100
      CloudMinWidth   =   600
      TailShape       =   2
      TailType        =   2
      TailBaseOffsetOutside=   300
      TailBaseOffsetInside=   100
      TailBaseAxisOutside=   0
      TailBaseAxisInside=   0
      TailBubbleCount =   3
      AlignTextH      =   1
      AlignTextV      =   1
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      AutoScan        =   0
      ScanAllDescendants=   -1  'True
      Interval        =   500
      IntervalNext    =   200
      AutoSize        =   -1  'True
      UseTagProp      =   -1  'True
      HideOnInactiveApp=   0   'False
      HideOnMouseDown =   2
      HideOnKeyDown   =   2
      HideOnFocus     =   0   'False
      ScanDisabledControls=   -1  'True
      ThreeDAppearance=   0
      FollowFocus     =   0   'False
      TemplateName    =   ""
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   6924
      Left            =   144
      Top             =   96
      Width           =   8124
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "#5) Press the 'F10 Apply/Close' button and GL numbers selected will appear in the appropriate fields on the main screen."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   1056
      TabIndex        =   17
      Top             =   6336
      Width           =   6252
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "#4) Repeat the process for all enabled G/L number fields"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1308
      Left            =   6864
      TabIndex        =   16
      Top             =   4368
      Width           =   1212
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "#3) Single click your GL number selection and the number appears in the selected 'G/L Number Options' field"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   6870
      TabIndex        =   15
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "#2) Next select a General Ledger number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1308
      Left            =   288
      TabIndex        =   14
      Top             =   3744
      Width           =   1212
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "#1) Click on one of the enabled 'G/L Number Options' buttons"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1548
      Left            =   288
      TabIndex        =   13
      Top             =   1968
      Width           =   1212
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   576
      X2              =   2352
      Y1              =   3792
      Y2              =   3552
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1344
      X2              =   3936
      Y1              =   2016
      Y2              =   1872
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   555
      Left            =   2175
      Top             =   240
      Width           =   4050
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "General Ledger Numbers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   396
      Left            =   2292
      TabIndex        =   9
      Top             =   348
      Width           =   3900
   End
End
Attribute VB_Name = "frmBLGLList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim AMeth As Integer

Private Sub cmdApply_Click()
  Call fpList1_DblClick
End Sub

Private Sub cmdClose_Click()
  Unload frmBLGLList
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Line1.Visible = True
    Line2.Visible = True
    frmBLMessageBoxJr.Label1.Caption = "The General Ledger number fields are enabled only if the accounting method saved requires that General Ledger number."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Line1.Visible = False
    Line2.Visible = False
  End If
  
  
'  If Exist("categoryedit.dat") Or Exist("changeaccmeth.dat") Then
'    frmBLMessageBox.Label1.Height = 2000
'    frmBLMessageBox.Label1.Top = 1000
'    frmBLMessageBox.Label1.Caption = "Highlight either 'Revenue', 'Accounts Receivable' (if the Cash accrual method is selected in the Town Setup screen the Accounts Receivable will be disabled) or 'Cash Receipt' shown at the top of the screen. Then when you click on one of the General Ledger numbers in the list that selection will appear underneath the choice made at the top of the screen. When you have completed the General Ledger number assignments then press F10 to send this data to the Category Edit screen and close this screen."
'    frmBLMessageBox.Show vbModal
'  ElseIf Exist("townsetup.dat") Then
'    frmBLMessageBox.Label1.Height = 2000
'    frmBLMessageBox.Label1.Top = 1200
'    frmBLMessageBox.Label1.Caption = "Highlight either 'Revenue', 'Accounts Receivable' (disabled if the Cash accrual method is selected) or 'Cash Receipt' shown at the top of the screen. Then when you click on one of the General Ledger numbers in the list that selection will appear underneath the choice made at the top of the screen. When you have completed the General Ledger number assignments then press F10 to send this data to the Town Setup screen and close this screen."
'    frmBLMessageBox.Show vbModal
'  End If
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyTab Then
    If optRev.Value = True Then
      If optAccts.Enabled = True Then
        optAccts.SetFocus
        KeyCode = 0
        Exit Sub
      ElseIf optCash.Enabled = True Then
        optCash.SetFocus
        KeyCode = 0
        Exit Sub
      End If
    End If
    
    If optAccts.Value = True Then
      optCash.SetFocus
        KeyCode = 0
        Exit Sub
    End If
    
    If optCash.Value = True Then
      If optRev.Enabled = True Then
        optRev.SetFocus
        KeyCode = 0
        Exit Sub
      End If
    End If
  End If
  
  Select Case KeyCode
    Case vbKeyReturn
      Call fpList1_DblClick
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdClose_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%A"
      Call fpList1_DblClick
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%T"
      Call cmdHelp_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub LoadMe()
  Dim GLIdxRec As JGLAcctIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdxRecs As Integer
  Dim GLAcctRec As GLAcctRecType
  Dim AcctHandle As Integer
  Dim x As Integer
  Dim Number$
  Dim DESC$
  Dim Found As Boolean
  Dim ThisMeth$
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  Label2.Visible = False
  Label3.Visible = False
  Label4.Visible = False
  Label5.Visible = False
  Label6.Visible = False
  Line1.Visible = False
  Line2.Visible = False
  
  If Exist("artownsu.dat") Then
    OpenTownFile THandle
    Get THandle, 1, TownRec
    Close THandle
    ThisMeth$ = Mid(TownRec.AcctMeth, 1, 1)
  ElseIf Exist("townsetup.dat") Then
    ThisMeth$ = Mid(frmBLTownSetup.fpcmbAcctMethod.Text, 1, 1)
  Else
    frmBLMessageBoxJr.Label1.Caption = "No accounting method saved. Please go to the Town Setup screen and save an accounting method before continuing."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  Select Case ThisMeth
    Case "C"
      AMeth = 2
    Case "A"
      AMeth = 3
    Case Else
      AMeth = 0
  End Select
  
  If Exist("categoryedit.dat") Then
    If AMeth = 2 Then
      fptxtAccts.Enabled = False
      optAccts.Enabled = False
      fptxtAccts.Text = ""
    Else
      fptxtAccts.Text = QPTrim$(frmBLCatEdit.fptxtAcctsRec)
    End If
    fptxtRev.Text = QPTrim$(frmBLCatEdit.fptxtRevGLAcctNum)
    fptxtCash.Text = QPTrim$(frmBLCatEdit.fptxtCashReceipt)
  ElseIf Exist("changeaccmeth.dat") Then
    If AMeth = 2 Then
      fptxtAccts.Enabled = False
      optAccts.Enabled = False
      fptxtAccts.Text = ""
    Else
      fptxtAccts.Text = QPTrim$(frmBLChangeAcctMeth.fptxtPenARGL)
    End If
    fptxtRev.Text = QPTrim$(frmBLChangeAcctMeth.fptxtPenRevGL)
    fptxtCash.Text = QPTrim$(frmBLChangeAcctMeth.fptxtPenCashGL)
  ElseIf Exist("townsetup.dat") Then
    If AMeth = 2 Then
      fptxtAccts.Enabled = False
      optAccts.Enabled = False
      fptxtAccts.Text = ""
    Else
      fptxtAccts.Text = QPTrim$(frmBLTownSetup.fptxtAcctsRec)
    End If
    fptxtRev.Text = QPTrim$(frmBLTownSetup.fptxtRevGLAcctNum)
    fptxtCash.Text = QPTrim$(frmBLTownSetup.fptxtCashReceipt)
  End If
  
  optRev.BackColor = &H80FFFF
  optRev.ForeColor = &H0&
  
  OpenGLIdxFile IdxHandle
  NumOfIdxRecs = LOF(IdxHandle) \ Len(GLIdxRec)

  If NumOfIdxRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no General Ledger numbers indexed."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  ReDim IdxRec(1 To NumOfIdxRecs) As Integer

  For x = 1 To NumOfIdxRecs
    Get IdxHandle, x, GLIdxRec
    IdxRec(x) = GLIdxRec.RecNo
  Next x
  Close IdxHandle

  OpenGLAcctFile AcctHandle
  For x = 1 To NumOfIdxRecs
    Get AcctHandle, IdxRec(x), GLAcctRec
    fpList1.InsertRow = QPTrim$(GLAcctRec.Num) + Chr(9) + QPTrim$(GLAcctRec.Title)
   Next x
  Close AcctHandle
   
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLGLList", "LoadMe", Erl)
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
'  --- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
    Unload Me
End Sub

Private Sub fpList1_Click()
  fpList1.Col = 0
  
  If optRev.Value = True Then
    fptxtRev.Text = QPTrim$(fpList1.ColText)
  ElseIf optAccts.Value = True Then
    fptxtAccts.Text = QPTrim$(fpList1.ColText)
  ElseIf optCash.Value = True Then
    fptxtCash.Text = QPTrim$(fpList1.ColText)
  End If
End Sub

Private Sub fpList1_DblClick()
  Dim Number$
  Dim DESC$
  Dim x As Integer
  Dim Nextx As Integer
  
  On Error Resume Next
  
  fpList1.Col = 0
  Number$ = QPTrim(fpList1.ColText)
  fpList1.Col = 1
  DESC$ = QPTrim$(fpList1.ColText)
  

  If Exist("categoryedit.dat") Then
    frmBLCatEdit.fptxtRevGLAcctNum.Text = QPTrim$(fptxtRev.Text)
    frmBLCatEdit.fptxtAcctsRec.Text = QPTrim$(fptxtAccts.Text)
    frmBLCatEdit.fptxtCashReceipt.Text = QPTrim$(fptxtCash.Text)
  ElseIf Exist("changeaccmeth.dat") Then
    If frmBLChangeAcctMeth.fptxtPenX.Text = "X" Then
      frmBLChangeAcctMeth.fptxtPenRevGL.Text = QPTrim$(fptxtRev.Text)
      frmBLChangeAcctMeth.fptxtPenARGL.Text = QPTrim$(fptxtAccts.Text)
      frmBLChangeAcctMeth.fptxtPenCashGL.Text = QPTrim$(fptxtCash.Text)
      frmBLChangeAcctMeth.fptxtPenX = ""
    End If
    ReDim ActiveX(1 To 1) As Integer
    
    frmBLChangeAcctMeth.vaSpread1.Col = 6
    Nextx = 0
    
    For x = 1 To 500
      frmBLChangeAcctMeth.vaSpread1.Row = x
      If frmBLChangeAcctMeth.vaSpread1.Text = "X" Then
        Nextx = Nextx + 1
        ReDim Preserve ActiveX(1 To Nextx) As Integer
        ActiveX(Nextx) = x
      End If
    Next x
    For x = 1 To Nextx
      frmBLChangeAcctMeth.vaSpread1.Row = ActiveX(x)
      frmBLChangeAcctMeth.vaSpread1.Col = 3
      frmBLChangeAcctMeth.vaSpread1.Text = QPTrim$(fptxtRev.Text)
      frmBLChangeAcctMeth.vaSpread1.Col = 4
      frmBLChangeAcctMeth.vaSpread1.Text = QPTrim$(fptxtAccts.Text)
      frmBLChangeAcctMeth.vaSpread1.Col = 5
      frmBLChangeAcctMeth.vaSpread1.Text = QPTrim$(fptxtCash.Text)
      frmBLChangeAcctMeth.vaSpread1.Col = 6
      frmBLChangeAcctMeth.vaSpread1.Text = ""
    Next x
  ElseIf Exist("townsetup.dat") Then
    frmBLTownSetup.fptxtRevGLAcctNum.Text = QPTrim$(fptxtRev.Text)
    frmBLTownSetup.fptxtAcctsRec.Text = QPTrim$(fptxtAccts.Text)
    frmBLTownSetup.fptxtCashReceipt.Text = QPTrim$(fptxtCash.Text)
  End If
  
  Unload frmBLGLList

End Sub

Private Sub fptxtAccts_GotFocus()
  optAccts.Value = True
End Sub

Private Sub fptxtCash_GotFocus()
  optCash.Value = True
End Sub

Private Sub fptxtRev_GotFocus()
  optRev.Value = True
End Sub

Private Sub optAccts_Click()
  optAccts.BackColor = &H80FFFF
  optAccts.ForeColor = &H0&
  optRev.BackColor = &H8F8265
  optRev.ForeColor = &HFFFFFF
  optCash.BackColor = &H8F8265
  optCash.ForeColor = &HFFFFFF
End Sub

Private Sub optCash_Click()
  optCash.BackColor = &H80FFFF
  optCash.ForeColor = &H0&
  optRev.BackColor = &H8F8265
  optRev.ForeColor = &HFFFFFF
  optAccts.BackColor = &H8F8265
  optAccts.ForeColor = &HFFFFFF
End Sub

Private Sub optRev_Click()
  optRev.BackColor = &H80FFFF
  optRev.ForeColor = &H0&
  optAccts.BackColor = &H8F8265
  optAccts.ForeColor = &HFFFFFF
  optCash.BackColor = &H8F8265
  optCash.ForeColor = &HFFFFFF

End Sub
