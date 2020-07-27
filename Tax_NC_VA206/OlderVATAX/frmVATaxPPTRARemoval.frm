VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxPPTRARemoval 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PPTRA Removal"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxPPTRARemoval.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList 
      Height          =   825
      Left            =   2220
      TabIndex        =   6
      Top             =   4920
      Width           =   7215
      _Version        =   196608
      _ExtentX        =   12726
      _ExtentY        =   1455
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
      Columns         =   5
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   1
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
      ColDesigner     =   "frmVATaxPPTRARemoval.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbPrintOpt 
      Height          =   405
      Left            =   4920
      TabIndex        =   4
      Top             =   6960
      Width           =   3330
      _Version        =   196608
      _ExtentX        =   5874
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
      ColDesigner     =   "frmVATaxPPTRARemoval.frx":0CA2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSelect 
      Height          =   732
      Left            =   2280
      TabIndex        =   16
      Top             =   6120
      Width           =   972
      _Version        =   131072
      _ExtentX        =   1714
      _ExtentY        =   1291
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   9405029
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
      ButtonDesigner  =   "frmVATaxPPTRARemoval.frx":1009
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0D0D0&
      Height          =   1572
      Left            =   3720
      TabIndex        =   14
      Top             =   1800
      Width           =   4212
      Begin VB.OptionButton OptByBill 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Remove By Individual Billing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   600
         TabIndex        =   2
         Top             =   1200
         Width           =   3492
      End
      Begin VB.OptionButton OptYear 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Remove By Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   600
         TabIndex        =   1
         Top             =   720
         Width           =   3492
      End
      Begin VB.OptionButton OptAll 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Remove All PPTRA Discounts"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   600
         TabIndex        =   0
         Top             =   240
         Width           =   3372
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintList 
      Height          =   492
      Left            =   4380
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6120
      Width           =   2892
      _Version        =   131072
      _ExtentX        =   5101
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmVATaxPPTRARemoval.frx":11EA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   495
      Left            =   6135
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7920
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmVATaxPPTRARemoval.frx":13D4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   3468
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7920
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmVATaxPPTRARemoval.frx":15B3
   End
   Begin EditLib.fpDateTime fptxtYear 
      Height          =   372
      Left            =   6216
      TabIndex        =   3
      Top             =   3720
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
      _ExtentY        =   656
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   -2147483643
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "2018"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "20010101"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
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
   Begin fpBtnAtlLibCtl.fpBtn cmdDeselect 
      Height          =   732
      Left            =   8400
      TabIndex        =   17
      Top             =   6120
      Width           =   972
      _Version        =   131072
      _ExtentX        =   1714
      _ExtentY        =   1291
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   9405029
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
      ButtonDesigner  =   "frmVATaxPPTRARemoval.frx":178F
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Removal Tax Year:"
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
      Height          =   276
      Left            =   4176
      TabIndex        =   15
      Top             =   3828
      Width           =   1860
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3132
      Left            =   1920
      Top             =   4440
      Width           =   7812
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PPTRA Removal Options"
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
      Height          =   360
      Left            =   4320
      TabIndex        =   13
      Top             =   1440
      Width           =   2892
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2892
      Left            =   3000
      Top             =   1320
      Width           =   5652
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   3120
      TabIndex        =   10
      Top             =   7080
      Width           =   1812
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Post Date"
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
      Height          =   276
      Left            =   2280
      TabIndex        =   9
      Top             =   4560
      Width           =   1620
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Backup File Name"
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
      Height          =   276
      Left            =   4320
      TabIndex        =   8
      Top             =   4560
      Width           =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   3000
      X2              =   8640
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PPTRA Removal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3420
      TabIndex        =   5
      Top             =   600
      Width           =   4848
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   600
      Left            =   2316
      Top             =   480
      Width           =   7008
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   756
      Left            =   2316
      Top             =   360
      Width           =   7020
   End
End
Attribute VB_Name = "frmVATaxPPTRARemoval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim TaxYear As Integer
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim GCustList() As Long
  Dim GCustCnt As Long
  Dim GFirstTrans As Long
  Dim GLastTrans As Long
  Dim TownName$

Private Sub cmdDeselect_Click()
  If fpList.Enabled = False Then Exit Sub
  fpList.Action = ActionDeselectAll
End Sub

Private Sub cmdExit_Click()
  frmVATaxPPTRAMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPrintList_Click()
  Dim PPRec As VAPPTaxBillType
  Dim PPHandle As Integer
  Dim NumOfPPRecs As Long
  Dim MyPath$, ThisFile$
  Dim x As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim dlm$
  Dim TotAmt#, TotPPTRA#
  Dim PCnt As Long
  Dim TransDate$
  
  If fpList.Enabled = False Then
    Call TaxMsg(900, "The billing list is disabled. Billing detail cannot be displayed.")
    Exit Sub
  End If
  If fpList.SelCount = 0 Then
    Call TaxMsg(900, "Please highlight a selection from the list.")
    Exit Sub
  ElseIf fpList.SelCount > 1 Then
    Call TaxMsg(900, "Please select only one billing from the list.")
    Exit Sub
  End If
  
  dlm$ = "~"
  MyPath = StartPath + "\TAXBILLBU\"
  fpList.Col = 1
  For x = 0 To fpList.ListCount - 1
    If fpList.Selected(x) = True Then
      fpList.Row = x
      Exit For
    End If
  Next x
  If x > fpList.ListCount - 1 Then
    Call TaxMsg(900, "ERROR: List selection could not be determined. Please select again.")
    Exit Sub
  End If
  ThisFile = QPTrim$(fpList.ColList)
  fpList.Col = 2
  GFirstTrans = CLng(fpList.ColList)
  fpList.Col = 3
  GLastTrans = CLng(fpList.ColList)
  
  RptFile = "TAXRPTS\PPTRARMV.RPT"
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  PCnt = 0
  TotAmt = 0
  TotPPTRA# = 0
  TransDate = ""
  OpenPersPostedReprintFile PPHandle, NumOfPPRecs, ThisFile
  For x = 1 To NumOfPPRecs
    Get PPHandle, x, PPRec
    If CLng(PPRec.BillNumber) < 0 Then GoTo SkipP
    If TransDate = "" Then
      TransDate = MakeRegDate(PPRec.PostDate)
    End If
    PCnt = PCnt + 1
    TotPPTRA = OldRound(TotPPTRA# + PPRec.PPTRADiscnt)
    TotAmt = OldRound(TotAmt + PPRec.TotalBillDue)
    '                       0                      1                   2
    Print #RptHandle, PPRec.BillNumber; dlm; PPRec.CustRec; dlm; QPTrim$(PPRec.CustName); dlm;
    '                       3                     4                    5              6           7
    Print #RptHandle, PPRec.TaxYear; dlm; PPRec.TotalBillDue; dlm; "PERSONAL"; dlm; TotAmt; dlm; PCnt; dlm;
    '                     8                   9                  10
    Print #RptHandle, TransDate; dlm; PPRec.PPTRADiscnt; dlm; TotPPTRA#
SkipP:
  Next x
  
  Close
  
  arVATaxPPTRARmvlList.Show
  
End Sub

Private Sub cmdProcess_Click()
  If OptAll.Value = True Then
    Call ProcessAll
  ElseIf OptYear.Value = True Then
    Call ProcessByYear
  ElseIf OptByBill.Value = True Then
    Call ProcessByBill
  End If
End Sub

Private Sub DWProcessAll()


'  Dim TransRec As TaxTransactionType
'  Dim TRHandle As Integer
'  Dim NumOfTRRecs As Long
'  Dim x As Long
'  Dim RmvlRec As TaxPPTRARemovalType
'  Dim RHandle As Integer
'  Dim NumOfRmvlRecs As Long
'  Dim TaxCust As TaxCustType
'  Dim TCHandle As Integer
'  Dim NumOfTCRecs As Long
'  Dim NextRec As Long
'  Dim PPRec As VAPPTaxBillType
'  Dim PPHandle As Integer
'  Dim NumOfPPRecs As Long
'  Dim ThisFile$
'  Dim y As Integer
'  Dim TotAmt#, TotPPTRA#
'  Dim PCnt As Long
'  Dim TransDate$
'  Dim SaveHere As Long
'  Dim t$
'
'  Dim SaveYear$
'  Dim PPTRATot As Double
'
'  Open "KEN_PPTRA.TXT" For Output As #1
'  t$ = "|"
'  Print #1, " CustomerRec "; t$; " PropPin "; t$; " Bill Num "; t$; " PPTRADisc "
'  OpenTaxCustFile TCHandle, NumOfTCRecs
'  OpenTaxTransFile TRHandle, NumOfTRRecs
'  For PCnt = 1 To NumOfTRRecs
'    Get TRHandle, PCnt, TransRec
'    If TransRec.TaxYear = 2012 Then
'      If TransRec.TranType = 1 Then
'        If TransRec.PPTRADisc > 0 Then
'          Get TCHandle, TransRec.CustomerRec, TaxCust
'          Print #1, QPTrim(CStr(TransRec.CustomerRec)); t$; QPTrim(TransRec.PersPin); t$;
'          Print #1, QPTrim(ParseBillNum$(TransRec.Description)); t$; QPTrim(CStr(TransRec.PPTRADisc))
'
''          RmvlRec.CustName = TaxCust.CustName
''          RmvlRec.PPTRADisc = TransRec.PPTRADisc
''          PPTRATot = OldRound(PPTRATot + TransRec.PPTRADisc)
''          RmvlRec.PPTRAValue = TransRec.PPTRAVal
''          RmvlRec.TaxAmount = TransRec.Amount
''          RmvlRec.TransRec = PPRec.TransRec
''          RmvlRec.BillDate = PPRec.PostDate
''          RmvlRec.BelongTo = PPRec.TransRec
''          RmvlRec.RmvlFile = ThisFile
'          't$
'          'Stop
'        End If
'      End If
'    End If
'  Next
'  Close
'  End
'
''  KillFile PPTRARemovalFile
''  OpenPPTRARmvlFile RHandle, NumOfRmvlRecs
'  'fpList.Col = 1
''  For y = 0 To fpList.ListCount - 1
''    fpList.Row = y
''    ThisFile = QPTrim$(fpList.ColList)
''    OpenPersPostedReprintFile PPHandle, NumOfPPRecs, ThisFile
''    For x = 1 To NumOfPPRecs
''      Get PPHandle, x, PPRec
''      If CDbl(PPRec.BillNumber) < 0 Then GoTo SkipIt
''        If PPRec.TransRec = 0 Then
''          GoTo SkipIt
''        End If
''        Get TRHandle, PPRec.TransRec, TransRec
''        If TransRec.TranType = 1 And TransRec.PPTRADisc > 0 Then
''          If BillBal(PPRec.TransRec, TRHandle) = False Then GoTo SkipIt
''          RmvlRec.BillNum = ParseBillNum$(TransRec.Description)
''          RmvlRec.CustAcct = TransRec.CustomerRec
''          Get TCHandle, TransRec.CustomerRec, TaxCust
''          RmvlRec.CustName = TaxCust.CustName
''          RmvlRec.PPTRADisc = TransRec.PPTRADisc
''          PPTRATot = OldRound(PPTRATot + TransRec.PPTRADisc)
''          RmvlRec.PPTRAValue = TransRec.PPTRAVal
''          RmvlRec.TaxAmount = TransRec.Amount
''          RmvlRec.TransRec = PPRec.TransRec
''          RmvlRec.BillDate = PPRec.PostDate
''          RmvlRec.BelongTo = PPRec.TransRec
''          RmvlRec.RmvlFile = ThisFile
''          RmvlRec.TxBillPostRec = x
''          SaveHere = SaveHere + 1
''          Put RHandle, SaveHere, RmvlRec
''        End If
''SkipIt:
''    Next x
''    Close PPHandle
''NextRow:
''  Next y
''
''  Close
''
''  If PPTRATot = 0 Then
''    Call TaxMsg(900, "There were no PPTRA discount values more than zero in the processed file.")
''    Exit Sub
''  End If
''
''  If fpcmbPrintOpt.Text = "Text" Then
''    Call PrintText
''  ElseIf fpcmbPrintOpt.Text = "Graphical" Then
''    Call PrintGraphics
''  End If

End Sub

Private Sub ProcessAll()
  Dim TransRec As TaxTransactionType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Long
  Dim x As Long
  Dim RmvlRec As TaxPPTRARemovalType
  Dim RHandle As Integer
  Dim NumOfRmvlRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim NextRec As Long
  Dim PPRec As VAPPTaxBillType
  Dim PPHandle As Integer
  Dim NumOfPPRecs As Long
  Dim ThisFile$
  Dim y As Integer
  Dim TotAmt#, TotPPTRA#
  Dim PCnt As Long
  Dim TransDate$
  Dim SaveHere As Long
  Dim SaveYear$
  Dim PPTRATot As Double
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TRHandle, NumOfTRRecs
  
  KillFile PPTRARemovalFile
  OpenPPTRARmvlFile RHandle, NumOfRmvlRecs
  fpList.Col = 1
  For y = 0 To fpList.ListCount - 1
    fpList.Row = y
    ThisFile = QPTrim$(fpList.ColList)
    OpenPersPostedReprintFile PPHandle, NumOfPPRecs, ThisFile
    For x = 1 To NumOfPPRecs
      Get PPHandle, x, PPRec
      If CDbl(PPRec.BillNumber) < 0 Then GoTo SkipIt
        If PPRec.TransRec = 0 Then
          GoTo SkipIt
        End If
        Get TRHandle, PPRec.TransRec, TransRec
        If TransRec.TranType = 1 And TransRec.PPTRADisc > 0 Then
          If BillBal(PPRec.TransRec, TRHandle) = False Then GoTo SkipIt
          RmvlRec.BillNum = ParseBillNum$(TransRec.Description)
          RmvlRec.CustAcct = TransRec.CustomerRec
          Get TCHandle, TransRec.CustomerRec, TaxCust
          RmvlRec.CustName = TaxCust.CustName
          RmvlRec.PPTRADisc = TransRec.PPTRADisc
          PPTRATot = OldRound(PPTRATot + TransRec.PPTRADisc)
          RmvlRec.PPTRAValue = TransRec.PPTRAVal
          RmvlRec.TaxAmount = TransRec.Amount
          RmvlRec.TransRec = PPRec.TransRec
          RmvlRec.BillDate = PPRec.PostDate
          RmvlRec.BelongTo = PPRec.TransRec
          RmvlRec.RmvlFile = ThisFile
          RmvlRec.TxBillPostRec = x
          SaveHere = SaveHere + 1
          Put RHandle, SaveHere, RmvlRec
        End If
SkipIt:
    Next x
    Close PPHandle
NextRow:
  Next y
  
  Close
  
  If PPTRATot = 0 Then
    Call TaxMsg(900, "There were no PPTRA discount values more than zero in the processed file.")
    Exit Sub
  End If
  
  If fpcmbPrintOpt.Text = "Text" Then
    Call PrintText
  ElseIf fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  End If
End Sub
Private Sub ProcessByBill()
  Dim TransRec As TaxTransactionType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Long
  Dim x As Long
  Dim RmvlRec As TaxPPTRARemovalType
  Dim RHandle As Integer
  Dim NumOfRmvlRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim NextRec As Long
  Dim PPRec As VAPPTaxBillType
  Dim PPHandle As Integer
  Dim NumOfPPRecs As Long
  Dim ThisFile$
  Dim y As Integer
  Dim TotAmt#, TotPPTRA#
  Dim PCnt As Long
  Dim TransDate$
  Dim SaveHere As Long
  Dim ThisRec As Integer
  Dim PPTRATot As Double
  
  If fpList.Enabled = False Then
    Call TaxMsg(800, "The billing list is disabled. Billing detail cannot be displayed.")
    Exit Sub
  End If
  If fpList.SelCount = 0 Then
    Call TaxMsg(900, "Please highlight a selection from the list.")
    Exit Sub
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TRHandle, NumOfTRRecs
  
  KillFile PPTRARemovalFile
  OpenPPTRARmvlFile RHandle, NumOfRmvlRecs
  For y = 0 To fpList.ListCount - 1
'    fpList.Col = 1
    fpList.Row = y
    If fpList.Selected(y) = True Then
      fpList.Col = 4
      ThisRec = CInt(fpList.ColList)
      fpList.Col = 1
      ThisFile = QPTrim$(fpList.ColList)
      OpenPersPostedReprintFile PPHandle, NumOfPPRecs, ThisFile
      For x = 1 To NumOfPPRecs
        Get PPHandle, x, PPRec
        If PPRec.TransRec = 0 Then
          GoTo SkipIt
        End If
        If CDbl(PPRec.BillNumber) < 0 Then GoTo SkipIt 'changed from Cint to CDbl 5/6/09
          Get TRHandle, PPRec.TransRec, TransRec
          If TransRec.TranType = 1 And TransRec.PPTRADisc > 0 Then
            If BillBal(PPRec.TransRec, TRHandle) = False Then GoTo SkipIt
            RmvlRec.BillNum = ParseBillNum$(TransRec.Description)
            RmvlRec.CustAcct = TransRec.CustomerRec
            Get TCHandle, TransRec.CustomerRec, TaxCust
            RmvlRec.CustName = TaxCust.CustName
            RmvlRec.PPTRADisc = TransRec.PPTRADisc
            PPTRATot = OldRound(PPTRATot + TransRec.PPTRADisc)
            RmvlRec.PPTRAValue = TransRec.PPTRAVal
            RmvlRec.TaxAmount = TransRec.Amount
            RmvlRec.TransRec = PPRec.TransRec
            RmvlRec.BillDate = PPRec.PostDate
            RmvlRec.RmvlFile = ThisFile
            RmvlRec.BelongTo = PPRec.TransRec
            RmvlRec.TxBillPostRec = ThisRec
            SaveHere = SaveHere + 1
            Put RHandle, SaveHere, RmvlRec
          End If
SkipIt:
      Next x
      Close PPHandle
    End If
  Next y
  
  Close
  If PPTRATot = 0 Then
    Call TaxMsg(900, "There were no PPTRA discount values more than zero in the processed file.")
    Exit Sub
  End If
  
  If fpcmbPrintOpt.Text = "Text" Then
    Call PrintText
  ElseIf fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  End If
  
End Sub
Private Sub ProcessByYear()
  Dim TransRec As TaxTransactionType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Long
  Dim x As Long
  Dim RmvlRec As TaxPPTRARemovalType
  Dim RHandle As Integer
  Dim NumOfRmvlRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim NextRec As Long
  Dim PPRec As VAPPTaxBillType
  Dim PPHandle As Integer
  Dim NumOfPPRecs As Long
  Dim ThisFile$
  Dim y As Integer
  Dim TotAmt#, TotPPTRA#
  Dim PCnt As Long
  Dim TransDate$
  Dim SaveHere As Long
  Dim SaveYear$
  Dim PPTRATot As Double
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TRHandle, NumOfTRRecs
  
  KillFile PPTRARemovalFile
  OpenPPTRARmvlFile RHandle, NumOfRmvlRecs
  fpList.Col = 1
  For y = 0 To fpList.ListCount - 1
    fpList.Row = y
    ThisFile = QPTrim$(fpList.ColList)
    SaveYear = Mid(ThisFile, 25, 4)
    If SaveYear <> fptxtYear.Text Then GoTo NextRow
    OpenPersPostedReprintFile PPHandle, NumOfPPRecs, ThisFile
    For x = 1 To NumOfPPRecs
      Get PPHandle, x, PPRec
      If CInt(PPRec.BillNumber) < 0 Then GoTo SkipIt
        If PPRec.TransRec = 0 Then
          GoTo SkipIt
        End If
        Get TRHandle, PPRec.TransRec, TransRec
        If TransRec.TranType = 1 And TransRec.PPTRADisc > 0 Then
          If BillBal(PPRec.TransRec, TRHandle) = False Then GoTo SkipIt
          RmvlRec.BillNum = ParseBillNum$(TransRec.Description)
          RmvlRec.CustAcct = TransRec.CustomerRec
          Get TCHandle, TransRec.CustomerRec, TaxCust
          RmvlRec.CustName = TaxCust.CustName
          RmvlRec.PPTRADisc = TransRec.PPTRADisc
          PPTRATot = OldRound(PPTRATot + TransRec.PPTRADisc)
          RmvlRec.PPTRAValue = TransRec.PPTRAVal
          RmvlRec.TaxAmount = TransRec.Amount
          RmvlRec.TransRec = PPRec.TransRec
          RmvlRec.BillDate = PPRec.PostDate
          RmvlRec.BelongTo = PPRec.TransRec
          RmvlRec.RmvlFile = ThisFile
          RmvlRec.TxBillPostRec = x
          SaveHere = SaveHere + 1
          Put RHandle, SaveHere, RmvlRec
        End If
SkipIt:
    Next x
    Close PPHandle
NextRow:
  Next y
  
  Close
  
  If PPTRATot = 0 Then
    Call TaxMsg(900, "There were no PPTRA discount values more than zero in the processed file.")
    Exit Sub
  End If
  
  If SaveHere = 0 Then
    Call TaxMsg(900, "There are no billing files available for year " + fptxtYear.Text + ".")
    fptxtYear.SetFocus
    Exit Sub
  End If
  
  If fpcmbPrintOpt.Text = "Text" Then
    Call PrintText
  ElseIf fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  End If
  
End Sub
Private Sub cmdSelect_Click()
  If fpList.Enabled = False Then Exit Sub
  fpList.Action = ActionSelectAll
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
    Case vbKeyF9:
      SendKeys "%S"
      Call cmdPrintList_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpPPTRARemoval
  Call LoadMe
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    'Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
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
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub LoadList(ThisYear As Integer)
  Dim PostRec As TaxBillPostDateType
  Dim PostHandle As Integer
  Dim NumOfPostRecs As Long
  Dim x As Long
  Dim CurrYear As Integer
  Dim ThisType As String * 1
  Dim CompareYr As Integer
  
  fpList.Action = ActionClear
  If Exist(TaxBillPostDateFile) Then
    OpenBillPostDateFile PostHandle, NumOfPostRecs
    For x = 1 To NumOfPostRecs
      Get PostHandle, x, PostRec
      If PostRec.PPTRAPosted = "Y" Then GoTo SkipIt
      If ThisYear = 0 Then GoTo GoAhead
      If QPTrim$(PostRec.BackUpName) = "" Then GoTo SkipIt
      CompareYr = CInt(Mid(PostRec.BackUpName, 25, 4))
      If CompareYr <> CInt(fptxtYear.Text) Then GoTo SkipIt
GoAhead:
      If PostRec.BillType = "P" Then
        fpList.InsertRow = MakeRegDate(PostRec.PostDate) + Chr(9) + PostRec.BackUpName + Chr(9) + CStr(PostRec.FirstTrans) + Chr(9) + CStr(PostRec.LastTrans) + Chr(9) + CStr(x)
      End If
SkipIt:
    Next x
    Close PostHandle
  End If
  
'  If fpList.ListCount > 0 Then
'    fpList.Selected(0) = True
'    fpList.ListIndex = 0
'  End If
  
End Sub

Private Sub GetGCustList()
  Dim PPRec As VAPPTaxBillType
  Dim PPHandle As Integer
  Dim NumOfPPRecs As Long
  Dim MyPath$, ThisFile$
  Dim x As Integer
  
  MyPath = StartPath + "\TAXBILLBU\"
  fpList.Col = 1
  fpList.Selected(fpList.ListIndex) = True
  fpList.Row = fpList.ListIndex
  ThisFile = QPTrim$(fpList.ColText)
  fpList.Col = 2
  GFirstTrans = CLng(fpList.ColText)
  fpList.Col = 3
  GLastTrans = CLng(fpList.ColText)
  
  OpenPersPostedReprintFile PPHandle, NumOfPPRecs, ThisFile
  ReDim GCustList(1 To NumOfPPRecs) As Long
  For x = 1 To NumOfPPRecs
    Get PPHandle, x, PPRec
    GCustList(x) = PPRec.CustRec
  Next x
  GCustCnt = NumOfPPRecs
  Close PPHandle

End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim x As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TownName = QPTrim$(TaxMasterRec.Name)
  
  OptAll.Value = True
  fptxtYear.Enabled = True
  fpList.Enabled = False
  
  fptxtYear.Text = CStr(TaxMasterRec.RTaxYear)
  TaxYear = TaxMasterRec.RTaxYear
  
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  Call LoadList(0)
End Sub

Private Sub fptxtYear_Change()
  LoadList (CInt(fptxtYear.Text))
End Sub

Private Sub OptAll_Click()
  fptxtYear.Enabled = False
  fpList.Enabled = False
  fpList.Action = ActionDeselectAll
End Sub

Private Sub OptByBill_Click()
  fptxtYear.Enabled = False
  fpList.Enabled = True
  Call LoadList(0)
  If fpList.ListCount > 0 Then
    fpList.Selected(0) = True
    fpList.ListIndex = 0
  End If
End Sub

Private Sub OptYear_Click()
  fptxtYear.Enabled = True
  fpList.Enabled = False
  fpList.Action = ActionDeselectAll
End Sub

Private Sub PrintText()
  Dim x As Long
  Dim RmvlRec As TaxPPTRARemovalType
  Dim RHandle As Integer
  Dim NumOfRmvlRecs As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim MaxLines As Integer
  Dim FF$
  Dim LineCnt As Integer
  Dim Page As Integer
  Dim ThisName$
  Dim ThatName$
  Dim NewTotal As Double
  Dim SubNewTotal As Double
  Dim GNewTotal As Double
  Dim SubOldTotal As Double
  Dim GOldTotal As Double
  Dim SubPPTRAAmt As Double
  Dim GPPTRAAmt As Double
  Dim SubCnt As Long
  Dim GCnt As Long
  Dim PostDate$
  
  MaxLines = 58
  FF$ = Chr(12)
  
  RptFile$ = "TAXRPTS\RMVBYBILL.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  ThisName = ""
  GoSub PrintHeader
  SubNewTotal = 0
  SubOldTotal = 0
  SubPPTRAAmt = 0
  SubCnt = 0
  GNewTotal = 0
  GOldTotal = 0
  GPPTRAAmt = 0
  GCnt = 0
  OpenPPTRARmvlFile RHandle, NumOfRmvlRecs
  For x = 1 To NumOfRmvlRecs
    Get RHandle, x, RmvlRec
    ThisName = QPTrim$(RmvlRec.RmvlFile)
    If ThisName <> ThatName Then
      LineCnt = LineCnt + 1
      If LineCnt > MaxLines - 8 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
      If SubCnt > 0 Then GoSub PrintSummary
      ThatName = ThisName
      SubNewTotal = 0
      SubOldTotal = 0
      SubPPTRAAmt = 0
      SubCnt = 0
      If LineCnt > MaxLines - 4 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
      GoSub PrintSubHeader
    End If
    PostDate = MakeRegDate(RmvlRec.BillDate)
    NewTotal = OldRound(RmvlRec.TaxAmount + RmvlRec.PPTRADisc)
    SubNewTotal = OldRound(SubNewTotal + NewTotal)
    SubOldTotal = OldRound(SubOldTotal + RmvlRec.TaxAmount)
    SubPPTRAAmt = OldRound(SubPPTRAAmt + RmvlRec.PPTRADisc)
    SubCnt = SubCnt + 1
    GNewTotal = OldRound(GNewTotal + NewTotal)
    GOldTotal = OldRound(GOldTotal + RmvlRec.TaxAmount)
    GPPTRAAmt = OldRound(GPPTRAAmt + RmvlRec.PPTRADisc)
    GCnt = GCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
      GoSub PrintSubHeader
    End If
    Print #RptHandle, Using$("#####0", RmvlRec.CustAcct); Tab(12); QPTrim$(RmvlRec.CustName); Tab(53); Using("######", RmvlRec.BillNum); Tab(65); Using$("$##,##0.00", RmvlRec.PPTRADisc); Tab(77); Using$("$##,##0.00", RmvlRec.TaxAmount); Tab(93); Using$("$##,##0.00", NewTotal)
    LineCnt = LineCnt + 1
  Next x
  
  If LineCnt > MaxLines - 9 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintSubHeader
  End If
  GoSub PrintSummary
  
  If LineCnt > MaxLines - 7 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  
  GoSub PrintTotals
  Close
  
  ViewPrint RptFile$, "Printing PPTRA Removal Report", True
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(40); "PPTRA Removal Report"
  Print #RptHandle, "Date: "; CStr(Date); Tab(83); "Page #"; CStr(Page)
  Print #RptHandle,
  Print #RptHandle, "Cust #"; Tab(12); "Customer Name"; Tab(53); "Bill #"; Tab(63); "PPTRA Amount"; Tab(77); "Bill Total"; Tab(89); "New Bill Total"
  Print #RptHandle, String$(102, "-")
  LineCnt = 5
  
  Return
  
PrintSubHeader:
  Print #RptHandle, "FileName: " + ThisName; Tab(60); "Billing Date: " + PostDate
  Print #RptHandle, String(102, "-")
  LineCnt = LineCnt + 3
  
  Return
  
PrintSummary:
  Print #RptHandle, "Summary for FileName: " + ThatName
  Print #RptHandle, String(102, "-")
  Print #RptHandle, "Transaction Count: " + Using$("##,##0", SubCnt)
  Print #RptHandle, "Total PPTRA Discounts: " + Using$("$###,###,##0.00", SubPPTRAAmt)
  Print #RptHandle, "Bill Totals:           " + Using$("$###,###,##0.00", SubOldTotal)
  Print #RptHandle, "New Bill Totals:       " + Using$("$###,###,##0.00", SubNewTotal)
  Print #RptHandle, String(102, "-")
  Print #RptHandle,
  Print #RptHandle,
  LineCnt = LineCnt + 8

  Return
  
PrintTotals:
  Print #RptHandle,
  Print #RptHandle, String$(102, "-")
  Print #RptHandle, "Grand Totals"
  Print #RptHandle, "Grand Transaction Count: " + Using$("##,##0", GCnt)
  Print #RptHandle, "Grand Total PPTRA Discounts: " + Using$("$###,###,##0.00", GPPTRAAmt)
  Print #RptHandle, "Bill Grand Totals:           " + Using$("$###,###,##0.00", GOldTotal)
  Print #RptHandle, "New Bill Grand Totals:       " + Using$("$###,###,##0.00", GNewTotal)
  Print #RptHandle, String(102, "-")
  LineCnt = LineCnt + 8

  Return

End Sub

Private Sub PrintGraphics()
  Dim x As Long
  Dim RmvlRec As TaxPPTRARemovalType
  Dim RHandle As Integer
  Dim NumOfRmvlRecs As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim ThisName$
  Dim ThatName$
  Dim NewTotal As Double
  Dim SubNewTotal As Double
  Dim GNewTotal As Double
  Dim SubOldTotal As Double
  Dim GOldTotal As Double
  Dim SubPPTRAAmt As Double
  Dim GPPTRAAmt As Double
  Dim SubCnt As Long
  Dim GCnt As Long
  Dim dlm$, PostDate$
  
  dlm$ = "~"
  RptFile$ = "TAXRPTS\RMVBYBILL.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  ThisName = ""
  
  SubNewTotal = 0
  SubOldTotal = 0
  SubPPTRAAmt = 0
  SubCnt = 0
  GNewTotal = 0
  GOldTotal = 0
  GPPTRAAmt = 0
  GCnt = 0
  OpenPPTRARmvlFile RHandle, NumOfRmvlRecs
  For x = 1 To NumOfRmvlRecs
    Get RHandle, x, RmvlRec
    ThisName = QPTrim$(RmvlRec.RmvlFile)
    If ThisName <> ThatName Then
      ThatName = ThisName
      SubNewTotal = 0
      SubOldTotal = 0
      SubPPTRAAmt = 0
      SubCnt = 0
    End If
    PostDate = MakeRegDate(RmvlRec.BillDate)
    NewTotal = OldRound(RmvlRec.TaxAmount + RmvlRec.PPTRADisc)
    SubNewTotal = OldRound(SubNewTotal + NewTotal)
    SubOldTotal = OldRound(SubOldTotal + RmvlRec.TaxAmount)
    SubPPTRAAmt = OldRound(SubPPTRAAmt + RmvlRec.PPTRADisc)
    SubCnt = SubCnt + 1
    GNewTotal = OldRound(GNewTotal + NewTotal)
    GOldTotal = OldRound(GOldTotal + RmvlRec.TaxAmount)
    GPPTRAAmt = OldRound(GPPTRAAmt + RmvlRec.PPTRADisc)
    GCnt = GCnt + 1
    GoSub PrintIt
  Next x
  
  Close
  
  arVATaxPPTRARmvlByBill.Show
  
  Exit Sub

PrintIt:
    '                    0                    1                             2
    Print #RptHandle, TownName; dlm; QPTrim$(RmvlRec.RmvlFile); dlm; RmvlRec.CustAcct; dlm;
    '                            3                           4
    Print #RptHandle, QPTrim$(RmvlRec.CustName); dlm; RmvlRec.BillNum; dlm;
    '                         5                      6                   7
    Print #RptHandle, RmvlRec.PPTRADisc; dlm; RmvlRec.TaxAmount; dlm; NewTotal; dlm;
    '                     8                  9                10
    Print #RptHandle, SubNewTotal; dlm; SubOldTotal; dlm; SubPPTRAAmt; dlm;
    '                   11             12             13              14            15
    Print #RptHandle, SubCnt; dlm; GNewTotal; dlm; GOldTotal; dlm; GPPTRAAmt; dlm; GCnt; dlm;
    '                    16
    Print #RptHandle, PostDate

  Return

End Sub

Private Function BillBal(TRec As Long, TRHandle As Integer) As Boolean
  Dim TransRec As TaxTransactionType
  Dim BillOwed#, BillPaid#
  BillBal = True
  BillOwed# = 0
  BillPaid# = 0
  Get TRHandle, TRec, TransRec
  BillOwed# = OldRound(TransRec.Revenue.Collection + TransRec.Revenue.Interest + TransRec.Revenue.LateList)
  BillOwed# = OldRound(BillOwed# + TransRec.Revenue.Penalty + TransRec.Revenue.Principle1 + TransRec.Revenue.Principle2)
  BillOwed# = OldRound(BillOwed# + TransRec.Revenue.Principle3 + TransRec.Revenue.Principle4 + TransRec.Revenue.Principle5)
  BillOwed# = OldRound(BillOwed# + TransRec.Revenue.RevOpt1 + TransRec.Revenue.RevOpt2 + TransRec.Revenue.RevOpt3 - TransRec.PPTRADisc)
  BillPaid# = OldRound(TransRec.Revenue.CollectionPd + TransRec.Revenue.InterestPd + TransRec.Revenue.LateListPd)
  BillPaid# = OldRound(BillPaid# + TransRec.Revenue.PenaltyPd + TransRec.Revenue.Principle1Pd + TransRec.Revenue.Principle2Pd)
  BillPaid# = OldRound(BillPaid# + TransRec.Revenue.Principle3Pd + TransRec.Revenue.Principle4Pd + TransRec.Revenue.Principle5Pd)
  BillPaid# = OldRound(BillPaid# + TransRec.Revenue.RevOpt1Pd + TransRec.Revenue.RevOpt2Pd + TransRec.Revenue.RevOpt3Pd)
  If BillOwed# = BillPaid# Then
    BillBal = False
  End If
  
End Function
