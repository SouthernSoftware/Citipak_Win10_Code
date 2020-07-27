VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmMiscAddEdit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Edit Miscellaneous Codes"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmMiscAddEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboInactive 
      Height          =   384
      Left            =   4680
      TabIndex        =   4
      Top             =   4920
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   1
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
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
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
      ColumnEdit      =   0
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmMiscAddEdit.frx":08CA
   End
   Begin LpLib.fpCombo fpcboAcctNumNa 
      Height          =   384
      Left            =   4656
      TabIndex        =   2
      Top             =   4296
      Width           =   5856
      _Version        =   196608
      _ExtentX        =   10329
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
      Columns         =   4
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   3
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   100
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmMiscAddEdit.frx":0CA4
   End
   Begin EditLib.fpText fpMiscCode 
      Height          =   372
      Left            =   4656
      TabIndex        =   0
      Top             =   3168
      Width           =   1752
      _Version        =   196608
      _ExtentX        =   3090
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
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
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   7
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
   Begin EditLib.fpText fpDescription 
      Height          =   372
      Left            =   4656
      TabIndex        =   1
      Top             =   3732
      Width           =   5664
      _Version        =   196608
      _ExtentX        =   9991
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
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
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   25
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
   Begin EditLib.fpText fpGLAcct 
      Height          =   372
      Left            =   4656
      TabIndex        =   3
      Top             =   4296
      Width           =   3312
      _Version        =   196608
      _ExtentX        =   5842
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
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
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   14
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7154
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "2:07 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "6/18/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdSave 
      Height          =   480
      Left            =   7080
      TabIndex        =   5
      Top             =   7224
      Width           =   1380
      _Version        =   131072
      _ExtentX        =   2434
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmMiscAddEdit.frx":111B
   End
   Begin fpBtnAtlLibCtl.fpBtn CmdExit 
      Height          =   480
      Left            =   8844
      TabIndex        =   6
      Top             =   7224
      Width           =   1380
      _Version        =   131072
      _ExtentX        =   2434
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmMiscAddEdit.frx":12F7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdList 
      Height          =   480
      Left            =   4008
      TabIndex        =   7
      Top             =   7200
      Width           =   1740
      _Version        =   131072
      _ExtentX        =   3069
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmMiscAddEdit.frx":14D3
   End
   Begin EditLib.fpLongInteger fpMisc 
      Height          =   252
      Left            =   1080
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1656
      Visible         =   0   'False
      Width           =   684
      _Version        =   196608
      _ExtentX        =   1206
      _ExtentY        =   444
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
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
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   0
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "0"
      MaxValue        =   "2147483647"
      MinValue        =   "-2147483648"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Inactive Flag:"
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
      Index           =   3
      Left            =   2496
      TabIndex        =   13
      Top             =   4944
      Width           =   1944
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      FillColor       =   &H8000000E&
      Height          =   3108
      Left            =   1644
      Top             =   2568
      Width           =   8940
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GL Account Number:"
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
      Index           =   2
      Left            =   1752
      TabIndex        =   12
      Top             =   4360
      Width           =   2688
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Code Description:"
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
      Index           =   0
      Left            =   1752
      TabIndex        =   11
      Top             =   3776
      Width           =   2688
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Miscellaneous Code:"
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
      Index           =   1
      Left            =   1752
      TabIndex        =   10
      Top             =   3192
      Width           =   2688
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add/Edit Miscellaneous Code "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3576
      TabIndex        =   8
      Top             =   1188
      Width           =   5100
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   840
      Left            =   2592
      Top             =   984
      Width           =   7020
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   960
      Left            =   2604
      Top             =   864
      Width           =   7020
   End
End
Attribute VB_Name = "frmMiscAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim LinkGL As Boolean, Oper As String, MiscRec As Long
Dim EditFlag As Boolean, BeenDone As Boolean
Dim AcctRecNo As Boolean, Ms0 As String
Dim Ms1 As String, Ms2 As String, Ms3 As String, Ms4 As String, Ms5 As String
Dim ChkOKFlag As Boolean, stuff As Boolean
Dim TempAcct As String, DupFound As Boolean
'Dim noreset As Boolean, CmNum As Long, MiscCode As String
'Dim Oper As String, PayListRec As Long, RecpPort As String
'Dim fromform As Form, toform As Form, codeopt As Integer
'Dim DefPayDate As String
Private Sub cmdExit_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
  FntSize = frmMsgDialog.Label(1).FontSize
 'check to see if have anything started on screen then if not go ahead on
  StuffBeenEntered
  If stuff = True Then
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "WARNING:Entry in Progress"
    MsgText(1) = ""
    MsgText(2) = "Do You Want to Abandon this Entry?"
    MsgText(3) = "Ok to Abandon,"
    MsgText(4) = "Cancel to Remain on Current Entry."
    MsgText(5) = ""
    If GetOKorNot(MsgText()) Then
      UBLog "USER WANTS TO Abandon"
      Unload Me
      frmCMSetupMenu.Show
      DoEvents
    Else
     UBLog "USER Canceled"
    End If
  Else
    Unload Me
    frmCMSetupMenu.Show
    DoEvents
  End If

End Sub
Private Sub Chk4Change()
'  Answer = 0
'  If fpTotReceived <> 0 Or fpTotPaid <> 0 Then
'    frmChangedWarning.Show vbModal, Me
'    Select Case SaveFlag
'    Case False
'      Answer = 3
'    Case True
'      Answer = 2
'    Case 1
'      Answer = 1
'    End Select
'  Else
'    Answer = 0
'  End If
End Sub
Private Sub ChkOK2Save()
'  Dim FntSize As Integer
'  Dim cntout As Integer, cnt As Integer
    DupFound = False
    ChkOKFlag = True
    If EditFlag = False Then
      Srch4DupCode
      If DupFound = True Then
        Ms0 = "DUPLICATE CODE"
        Ms1 = ""
        Ms2 = "PLEASE CHECK YOUR CODE"
        Ms3 = ""
        Ms4 = "Enter a Different Misc Code"
        Ms5 = ""
        Dothemsg
        ChkOKFlag = False
        Exit Sub
      End If
    End If
    If Len(fpMiscCode.Text) = 0 Then
      Ms0 = "INVALID CODE"
      Ms1 = ""
      Ms2 = "PLEASE CHECK YOUR CODE"
      Ms3 = ""
      Ms4 = "Enter a Valid Misc Code"
      Ms5 = ""
      Dothemsg
      ChkOKFlag = False
      Exit Sub
    End If
    If Len(fpDescription.Text) = 0 Then
      Ms0 = "INVALID DESCRIPTION"
      Ms1 = ""
      Ms2 = "PLEASE CHECK DESCRIPTION"
      Ms3 = ""
      Ms4 = "Enter a Valid Description"
      Ms5 = ""
      Dothemsg
      ChkOKFlag = False
      Exit Sub
    End If
    If LinkGL Then
      If fpcboAcctNumNa.ListIndex = -1 Then
        Ms0 = "INVALID ACCOUNT NUMBER"
        Ms1 = ""
        Ms2 = "PLEASE CHECK YOUR ACCT"
        Ms3 = ""
        Ms4 = "Enter a Valid GL Account"
        Ms5 = ""
        Dothemsg
        ChkOKFlag = False
        Exit Sub
      End If
    End If
    If fpcboInactive.ListIndex = -1 Then
      Ms0 = "INVALID ENTRY"
      Ms1 = ""
      Ms2 = "PLEASE CHECK INACTIVE FLAG"
      Ms3 = ""
      Ms4 = "Enter a Valid Inactive Option"
      Ms5 = ""
      Dothemsg
      ChkOKFlag = False
      Exit Sub
    End If
'  cntout = 0
'  Answer = 0
'    If fpTotReceived <> 0 Or fpTotPaid <> 0 Then cntout = cntout + 1
'
'  If cntout > 0 Then
'    ReDim MsgText(0 To 5) As String
'    FntSize = frmMsgDialog.Label(1).FontSize
'    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
'    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
'    MsgText(0) = "WARNING:Payment In Progress"
'    MsgText(1) = ""
'    MsgText(2) = "Do You Want to Abandon this Payment?"
'    MsgText(3) = "Ok to Abandon,"
'    MsgText(4) = "Cancel to Remain on Current Payment."
'    MsgText(5) = ""
'    If GetOKorNot(MsgText()) Then
'     UBLog "USER WANTS TO Abandon"
'     Answer = 2
'    Else
'     UBLog "USER Canceled"
'     Answer = 1
'    End If
'  Else
'    Answer = 0
'  End If
'  End If
End Sub

Private Sub StuffBeenEntered()
  Dim cnt As Integer
  cnt = 0
  stuff = False
  If Len(fpMiscCode.Text) <> 0 Then cnt = cnt + 1
  If Len(fpDescription.Text) <> 0 Then cnt = cnt + 1
  If LinkGL Then
    If fpcboAcctNumNa.ListIndex > -1 Then cnt = cnt + 1
  End If
  If fpcboInactive.ListIndex > -1 Then cnt = cnt + 1
  
  If cnt > 0 Then
    stuff = True
  End If
End Sub
Private Sub cmdList_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
  FntSize = frmMsgDialog.Label(1).FontSize
 'check to see if have anything started on screen then if not go ahead on
  StuffBeenEntered
  If stuff = True Then
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "WARNING:Entry in Progress"
    MsgText(1) = ""
    MsgText(2) = "Do You Want to Abandon this Entry?"
    MsgText(3) = "Ok to Abandon,"
    MsgText(4) = "Cancel to Remain on Current Entry."
    MsgText(5) = ""
    If GetOKorNot(MsgText()) Then
     UBLog "USER WANTS TO Abandon"
     ClearScn
     frmMiscCodeList.Show 1
    Else
     UBLog "USER Canceled"
    End If
  Else
    frmMiscCodeList.Show 1
  End If
  fpDescription.SetFocus
End Sub

Private Sub fpCmdSave_Click()
  ChkOK2Save
  If ChkOKFlag Then
'   'DeActivateControls Me
    MiscCodeSave
    MsgBox "Transaction Complete.", vbOKOnly, "Complete"
    ClearScn
    fpMiscCode.SetFocus
  End If
End Sub
Private Sub MiscCodeSave()
  Dim MiscCodeRecLen  As Integer, MCFile As Integer, NumOfMiscRecs As Long
  ReDim MiscCodeRec(1) As MiscCodeRecType
  MiscCodeRecLen = Len(MiscCodeRec(1))
  MCFile = FreeFile
  Open UBPath$ + "CMMISCCD.DAT" For Random Shared As MCFile Len = MiscCodeRecLen
  NumOfMiscRecs = LOF(MCFile) \ MiscCodeRecLen
  
    MiscCodeRec(1).MiscCode = QPTrim$(fpMiscCode)
    MiscCodeRec(1).Description = QPTrim$(fpDescription)
    If LinkGL Then
      If fpcboAcctNumNa.ListIndex > 0 Then
        fpcboAcctNumNa.col = 1
        MiscCodeRec(1).GlAcctNumb = QPTrim$(fpcboAcctNumNa.ColText)
      Else
        MiscCodeRec(1).GlAcctNumb = TempAcct$
      End If
    Else
      MiscCodeRec(1).GlAcctNumb = QPTrim$(fpGLAcct)
    End If
    If fpcboInactive.ListIndex = 1 Then
      MiscCodeRec(1).InActiveFlag = "Y"
    Else
      MiscCodeRec(1).InActiveFlag = "N"
    End If
    If MiscRec > 0 Then
      Put MCFile, MiscRec, MiscCodeRec(1)
    Else
      Put MCFile, NumOfMiscRecs + 1, MiscCodeRec(1)
    End If
    Close MCFile
'  End If
  EditFlag = False
End Sub
Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        CMLog "Closed via CMMiscAddEdit by " + PWUser$ + " operator-" + Oper$
        CitiTerminate
      End If
    End If
  End If
End Sub
Private Sub fpcboInactive_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboInactive.ListDown = True
  End If
  If fpcboInactive.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpCmdSave.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        If fpcboAcctNumNa.Enabled = True Then
          fpcboAcctNumNa.SetFocus
        ElseIf fpGLAcct.Enabled = True Then
          fpGLAcct.SetFocus
        End If
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboAcctNumNa_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboAcctNumNa.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboAcctNumNa.ListIndex = -1
    fpcboAcctNumNa.Action = ActionClearSearchBuffer
  End If
  If fpcboAcctNumNa.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboInactive.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpDescription.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      DoEvents
      If cmdExit.Enabled Then
        Call cmdExit_Click
      End If
    Case vbKeyF5:
      KeyCode = 0
      DoEvents
      If cmdList.Enabled = True Then
        cmdList_Click
      End If
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      If fpCmdSave.Enabled Then
        Call fpCmdSave_Click
      End If
    Case Else:
  End Select
End Sub
Private Sub Srch4DupCode()
  Dim MiscCodeRecLen  As Integer, MCFile As Integer, NumOfMiscRecs As Long
  Dim cnt As Long
  ReDim MiscCodeRec(1) As MiscCodeRecType
  MiscCodeRecLen = Len(MiscCodeRec(1))
  MCFile = FreeFile
  Open UBPath$ + "CMMISCCD.DAT" For Random Shared As MCFile Len = MiscCodeRecLen
  NumOfMiscRecs = LOF(MCFile) \ MiscCodeRecLen

  If NumOfMiscRecs > 0 Then
  For cnt = 1 To NumOfMiscRecs
    Get MCFile, cnt, MiscCodeRec(1)
    If QPTrim$(fpMiscCode.Text) = QPTrim$(MiscCodeRec(1).MiscCode) Then
      DupFound = True
      Exit For
    End If
  Next cnt
 End If
End Sub

Private Sub Form_Load()
  Dim CMSetUpRec(1) As CMSetupType
  Dim RecLen As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TownName$
  LoadCMSetUpFile CMSetUpRec(), RecLen
  If QPTrim$(CMSetUpRec(1).GLInterface) = "Y" Then
    LinkGL = True
  Else
    LinkGL = False
  End If
  If LinkGL = True Then
    fpcboAcctNumNa.Visible = True
    fpGLAcct.Visible = False
    fpGLAcct.Enabled = False
    FillAcctNumName fpcboAcctNumNa
  Else
    fpcboAcctNumNa.Visible = False
    fpcboAcctNumNa.Enabled = False
    fpGLAcct.Visible = True
  End If
  fpcboInactive.AddItem "No"
  fpcboInactive.AddItem "Yes"
  fpcboInactive.ListIndex = -1
  Oper$ = Str$(OperNum)
  CMLog " IN Oper " + Oper$ + ": CMMiscAddEdit"
  'GetRcpInfo
  EditFlag = False
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Temp_Class.ResizeControls Me
  End If
  DoEvents
End Sub
Private Sub ClearScn()
  fpMiscCode = ""
  fpDescription = ""
  If LinkGL = True Then
    fpcboAcctNumNa.ListIndex = -1
  Else
    fpGLAcct = ""
  End If
  fpcboInactive.ListIndex = -1
  TempAcct$ = ""
  MiscRec = 0
  AcctRecNo = 0
  EditFlag = False
  fpMiscCode.Enabled = True
End Sub
Public Sub LoadScreen(MiscRecT)
  Dim MiscCodeRecLen  As Integer, MCFile As Integer
  ReDim MiscCodeRec(1) As MiscCodeRecType
  MiscCodeRecLen = Len(MiscCodeRec(1))
  MCFile = FreeFile
  Open UBPath$ + "CMMISCCD.DAT" For Random Shared As MCFile Len = MiscCodeRecLen
  MiscRec = MiscRecT
  If MiscRec > 0 Then
    Get MCFile, MiscRec, MiscCodeRec(1)
    fpMiscCode = QPTrim$(MiscCodeRec(1).MiscCode)
    fpDescription = QPTrim$(MiscCodeRec(1).Description)
    If LinkGL Then
      fpcboAcctNumNa.SearchText = QPStrip(MiscCodeRec(1).GlAcctNumb)
      fpcboAcctNumNa.Action = 0
      If fpcboAcctNumNa.SearchIndex <> -1 Then
        fpcboAcctNumNa.ListIndex = fpcboAcctNumNa.SearchIndex
      Else
        TempAcct$ = QPTrim(MiscCodeRec(1).GlAcctNumb)
        fpcboAcctNumNa.ListIndex = 0
      End If
    Else
      fpGLAcct = QPTrim$(MiscCodeRec(1).GlAcctNumb)
    End If
    If MiscCodeRec(1).InActiveFlag = "Y" Then
      fpcboInactive.ListIndex = 1
    ElseIf MiscCodeRec(1).InActiveFlag = "N" Then
      fpcboInactive.ListIndex = 0
    Else
      fpcboInactive.ListIndex = -1
    End If
    fpMiscCode.Enabled = False
    EditFlag = True
    BeenDone = True
  Else
    fpMiscCode.Enabled = True
  End If
  Close
End Sub
Private Sub SrchMiscCode()
  Dim MiscCodeRecLen  As Integer, MCFile As Integer, NumOfMiscRecs As Long
  Dim cnt As Long
  ReDim MiscCodeRec(1) As MiscCodeRecType
  MiscCodeRecLen = Len(MiscCodeRec(1))
  MCFile = FreeFile
  Open UBPath$ + "CMMISCCD.DAT" For Random Shared As MCFile Len = MiscCodeRecLen
  NumOfMiscRecs = LOF(MCFile) \ MiscCodeRecLen

  If NumOfMiscRecs > 0 Then
  For cnt = 1 To NumOfMiscRecs
    Get MCFile, cnt, MiscCodeRec(1)
    If QPTrim$(fpMiscCode.Text) = MiscCodeRec(1).MiscCode Then
      fpMiscCode = QPTrim$(MiscCodeRec(1).MiscCode)
      fpDescription = QPTrim$(MiscCodeRec(1).Description)
      If LinkGL Then
        fpcboAcctNumNa.SearchText = QPStrip(MiscCodeRec(1).GlAcctNumb)
        fpcboAcctNumNa.Action = 0
        If fpcboAcctNumNa.SearchIndex <> -1 Then
          fpcboAcctNumNa.ListIndex = fpcboAcctNumNa.SearchIndex
        Else
          fpcboAcctNumNa.ListIndex = 0
          TempAcct$ = QPTrim(MiscCodeRec(1).GlAcctNumb)
        End If
      Else
        fpGLAcct = QPTrim$(MiscCodeRec(1).GlAcctNumb)
      End If
      If MiscCodeRec(1).InActiveFlag = "Y" Then
        fpcboInactive.ListIndex = 1
      ElseIf MiscCodeRec(1).InActiveFlag = "N" Then
        fpcboInactive.ListIndex = 0
      Else
        fpcboInactive.ListIndex = -1
      End If
      MiscRec = cnt
      EditFlag = True
      Exit For
    End If
  Next cnt
  
  End If
  If EditFlag = False Then
    MiscRec = 0
  End If
End Sub
'Private Sub ChkGLAcct()
'  Dim SetupFileNum As Integer, Fund As Integer, Accnt As Integer, Det As Integer
'  Dim GLNumber As String
'    OpenSetupFile SetupFileNum
'
'    If LOF(SetupFileNum) > 0 Then
'      Get SetupFileNum, 1, GLSetup
'      Fund = GLSetup.FundLen
'      Accnt = GLSetup.AcctLen
'      Det = GLSetup.DetLen
'      GLNumber$ = Left$(fpGLAcct, Fund) + "-" + Mid$(fpGLAcct, Fund + 1, Accnt) + "-" + Mid$(fpGLAcct, Fund + Accnt + 1, Det)
'      AcctRecNo = AcctFind(GLNumber$)
'    Else
'      AcctRecNo = True
'    End If
'    Close SetupFileNum
'End Sub

Private Sub fpDescription_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpDescription_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If LinkGL Then
      If fpcboAcctNumNa.Enabled = True Then
        fpcboAcctNumNa.SetFocus
      End If
    Else
      If fpGLAcct.Enabled = True Then
        fpGLAcct.SetFocus
      End If
    End If
  End If
End Sub
Private Sub fpGLAcct_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpGLAcct_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboInactive.SetFocus
  End If
End Sub

'Private Sub fpGLAcct_LostFocus()
'  If LinkGL = True Then
'    If EditFlag = False Then
'      If Len(fpGLAcct.Text) > 0 Then
'        ChkGLAcct
'        If AcctRecNo = False Then
'          Ms0 = "INVALID ACCOUNT NUMBER"
'          Ms1 = ""
'          Ms2 = "PLEASE CHECK YOUR ACCT"
'          Ms3 = ""
'          Ms4 = "Enter a Valid GL Account"
'          Ms5 = ""
'          Dothemsg
'          Exit Sub
'        End If
'      Else
'        'message here about no blank gl's
'        Ms0 = "Blank Account Not Valid"
'        Ms1 = ""
'        Ms2 = "PLEASE CHECK YOUR ACCT"
'        Ms3 = ""
'        Ms4 = "Enter a Valid GL Account"
'        Ms5 = ""
'        Dothemsg
'      End If
'    End If
'  End If
'End Sub


Private Sub fpMiscCode_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpMiscCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpDescription.SetFocus
  End If
End Sub

Private Sub fpMiscCode_LostFocus()
  If BeenDone = False Then
    If Len(fpMiscCode.Text) > 0 Then
      SrchMiscCode
    End If
  End If
End Sub
Private Sub Dothemsg()
    Dim FntSize As Integer
    ReDim MsgText(0 To 5) As String
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(1).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = Ms0
    MsgText(1) = Ms1
    MsgText(2) = Ms2
    MsgText(3) = Ms3
    MsgText(4) = Ms4
    MsgText(5) = Ms5
    GetOKorNot MsgText(), True
End Sub

