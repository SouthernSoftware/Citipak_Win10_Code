VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmTaxSeniorDscRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Discount Analysis Report"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxSeniorDscRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5790
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   7785
      _Version        =   196609
      _ExtentX        =   13732
      _ExtentY        =   10213
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmTaxSeniorDscRpt.frx":08CA
      Begin LpLib.fpCombo fpcmbTownship 
         Height          =   384
         Left            =   2928
         TabIndex        =   4
         Top             =   2232
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
         ColDesigner     =   "frmTaxSeniorDscRpt.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   2928
         TabIndex        =   3
         Top             =   2928
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
         ColDesigner     =   "frmTaxSeniorDscRpt.frx":0CC1
      End
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   384
         Left            =   2928
         TabIndex        =   2
         Top             =   3636
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
         ColDesigner     =   "frmTaxSeniorDscRpt.frx":109C
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   645
         Left            =   2040
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   4770
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
         ButtonDesigner  =   "frmTaxSeniorDscRpt.frx":1477
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   645
         Left            =   4275
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   $"frmTaxSeniorDscRpt.frx":1655
         Top             =   4770
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
         ButtonDesigner  =   "frmTaxSeniorDscRpt.frx":1700
      End
      Begin EditLib.fpDoubleSingle fpDblTaxRate 
         Height          =   375
         Left            =   3480
         TabIndex        =   0
         Top             =   1605
         Width           =   1095
         _Version        =   196608
         _ExtentX        =   1931
         _ExtentY        =   661
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
         AlignTextH      =   2
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
         Text            =   "0.00"
         DecimalPlaces   =   2
         DecimalPoint    =   "."
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "0"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "(per $100)"
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
         Left            =   4800
         TabIndex        =   12
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Tax Rate:"
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
         Left            =   1800
         TabIndex        =   11
         Top             =   1680
         Width           =   1500
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
         Left            =   1275
         TabIndex        =   10
         Top             =   2340
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   1005
         Top             =   1290
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
         TabIndex        =   9
         Top             =   3720
         Width           =   1305
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Exemptions Analysis Report"
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
         Left            =   1320
         TabIndex        =   8
         Top             =   480
         Width           =   5415
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   1080
         Top             =   315
         Width           =   5865
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
         Top             =   3030
         Width           =   1500
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6120
      Left            =   1800
      Top             =   1305
      Width           =   8055
   End
End
Attribute VB_Name = "frmTaxSeniorDscRpt"
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

Private Sub cmdExit_Click()
  frmTaxReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  If fpDblTaxRate.Value = 0 Then
    Call TaxMsg(900, "The tax rate entered is zero.")
    fpDblTaxRate.SetFocus
    Exit Sub
  End If
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
  Me.HelpContextID = hlpExemption
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxSeniorDscRpt.")
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
  
  'on error goto ERRORSTUFF
  
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
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxSeniorDscRpt", "LoadMe", Erl)
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
      fpDblTaxRate.SetFocus
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
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PropAdd$, PropTownShip$
  Dim CustCnt As Long
  Dim CustRec As Long
  Dim CustName$
  Dim ThisTownship$
  Dim TotBothDsc As Double
  Dim ThisBothDsc As Double
  Dim ThisLoss As Double
  Dim TotLoss As Double
  Dim SourceFlag As Integer
  Dim ThisRate As Double
  Dim ThisSnrDsc As Double
  Dim TotSnrDsc As Double
  Dim ThisSnrLoss As Double
  Dim TotSnrLoss As Double
  Dim ThisOthDsc As Double
  Dim TotOthDsc As Double
  Dim ThisOthLoss As Double
  Dim TotOthLoss As Double
  Dim ThisVal As Double
  Dim SnrX As Double
  Dim OthX As Double
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  
  'on error goto ERRORSTUFF
  
  ThisRate = fpDblTaxRate.Value
  ThisTownship = fpcmbTownship.Text
  OptFlag = False
  
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

  RptFile$ = "TAXRPTS\DISCOUNT.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenPersPropFile PHandle, NumOfPersRecs
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  frmTaxShowPctComp.Label1 = "Gathering Discount Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  TotLoss = 0
  ThisLoss = 0
  ThisSnrDsc = 0
  TotSnrDsc = 0
  ThisOthDsc = 0
  TotOthDsc = 0
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
    If TaxCust.FirstPersRec > 0 Then
      NextRec = TaxCust.FirstPersRec
      Do While NextRec > 0
        Get PHandle, NextRec, PersRec
        If PersRec.Deleted = 0 Then
          If PersRec.EXMPOTHR > 0 Or PersRec.EXMPSENI > 0 Then
            CustCnt = CustCnt + 1
            NextRec = 0
            GoTo GotCount
            Exit Do
          Else
            NextRec = PersRec.NextRec
          End If
        Else
          NextRec = PersRec.NextRec
        End If
      Loop
    End If

    If TaxCust.FirstPropRec > 0 Then
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted = 0 Then
          If RealRec.EXMPOTHR > 0 Or RealRec.EXMPSENI > 0 Then
            CustCnt = CustCnt + 1
            NextRec = 0
            Exit Do
          Else
            NextRec = RealRec.NextRec
          End If
        Else
          NextRec = RealRec.NextRec
        End If
      
      Loop
    End If
    
GotCount:
    CustName = QPTrim$(TaxCust.CustName)
    ThisSnrDsc = 0
    ThisOthDsc = 0
    ThisBothDsc = 0
    If TaxCust.FirstPropRec > 0 Then
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted = -1 Then GoTo Deleted1
        SnrX = RealRec.EXMPSENI
        OthX = RealRec.EXMPOTHR
        ThisVal = OldRound(SnrX + OthX)
        ThisBothDsc = OldRound(ThisBothDsc + ThisVal)
        TotBothDsc = OldRound(TotBothDsc + ThisVal)
        ThisVal = SnrX
        ThisSnrDsc = OldRound(ThisSnrDsc + ThisVal)
        TotSnrDsc = OldRound(TotSnrDsc + ThisVal)
        ThisVal = OthX
        ThisOthDsc = OldRound(ThisOthDsc + ThisVal)
        TotOthDsc = OldRound(TotOthDsc + ThisVal)
Deleted1:
        NextRec = RealRec.NextRec
      Loop
    End If
    
    If TaxCust.FirstPersRec > 0 Then
      NextRec = TaxCust.FirstPersRec
      Do While NextRec > 0
        Get PHandle, NextRec, PersRec
        If PersRec.Deleted = -1 Then GoTo Deleted2
        SnrX = PersRec.EXMPSENI
        OthX = PersRec.EXMPOTHR
        ThisVal = OldRound(SnrX + OthX)
        ThisBothDsc = OldRound(ThisBothDsc + ThisVal)
        TotBothDsc = OldRound(TotBothDsc + ThisVal)
        ThisVal = SnrX
        ThisSnrDsc = OldRound(ThisSnrDsc + ThisVal)
        TotSnrDsc = OldRound(TotSnrDsc + ThisVal)
        ThisVal = OthX
        ThisOthDsc = OldRound(ThisOthDsc + ThisVal)
        TotOthDsc = OldRound(TotOthDsc + ThisVal)
Deleted2:
        NextRec = PersRec.NextRec
      Loop
    End If
    ThisLoss = OldRound((ThisBothDsc * ThisRate) / 100)
    TotLoss = OldRound(ThisLoss + TotLoss)
    ThisSnrLoss = OldRound((ThisSnrDsc * ThisRate) / 100)
    TotSnrLoss = OldRound(TotSnrLoss + ThisSnrLoss)
    ThisOthLoss = OldRound((ThisOthDsc * ThisRate) / 100)
    TotOthLoss = OldRound(TotOthLoss + ThisOthLoss)
    If ThisLoss > 0 Then
      '                   0              1                2
      Print #RptHandle, Town; dlm; TaxCust.Acct; dlm; CustName; dlm;
      '                 3                4                   5
      Print #RptHandle, ""; dlm; ThisBothDsc; dlm; TotBothDsc; dlm;
       '                     6              7             8                9
      Print #RptHandle, ThisLoss; dlm; TotLoss; dlm; ThisSnrDsc; dlm; TotSnrDsc; dlm;
      '                      10              11                 12             13
      Print #RptHandle, ThisSnrLoss; dlm; TotSnrLoss; dlm; ThisOthDsc; dlm; TotOthDsc; dlm;
      '                     14                15             16               17
      Print #RptHandle, ThisOthLoss; dlm; TotOthLoss; dlm; CustCnt; dlm; ThisTownship; dlm;
      '                           18                                   19
      Print #RptHandle, QPTrim$(TaxCust.TownShip); dlm; QPTrim$(TaxCust.Addr1) + "  " + QPTrim$(TaxCust.City) + ", " + QPTrim(TaxCust.State) + "  " + QPTrim$(TaxCust.Zip); dlm;
      '                    20
      Print #RptHandle, ThisRate; dlm;
      If UseOpt = "Y" Then
        '                    21                     22
        Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc)
      Else
        '                 21       22
        Print #RptHandle, ""; dlm; ""
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
  
  Close
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  If CustCnt = 0 Then
    Call TaxMsg(900, "There are no customers with discounts with the parameters entered.")
    Exit Sub
  End If
  
  arTaxDiscountRpt.Show
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxSeniorDscRpt", "PrintGraphics", Erl)
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
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PropAdd$, PropTownShip$
  Dim CustCnt As Long
  Dim CustRec As Long
  Dim CustName$
  Dim ThisTownship$
  Dim TotBothDsc As Double
  Dim ThisBothDsc As Double
  Dim ThisLoss As Double
  Dim TotLoss As Double
  Dim ThisRate As Double
  Dim ThisSnrDsc As Double
  Dim TotSnrDsc As Double
  Dim ThisSnrLoss As Double
  Dim TotSnrLoss As Double
  Dim ThisOthDsc As Double
  Dim TotOthDsc As Double
  Dim ThisOthLoss As Double
  Dim TotOthLoss As Double
  Dim ThisVal As Double
  Dim SnrX As Double
  Dim OthX As Double
  Dim FF$
  Dim Page As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  
  'on error goto ERRORSTUFF
  
  FF$ = Chr(12)
  MaxLines = 58
  ThisRate = fpDblTaxRate.Value
  ThisTownship = fpcmbTownship.Text
  OptFlag = False
  
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

  RptFile$ = "TAXRPTS\DISCOUNT.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenPersPropFile PHandle, NumOfPersRecs
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  frmTaxShowPctComp.Label1 = "Gathering Discount Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  TotLoss = 0
  ThisLoss = 0
  ThisSnrDsc = 0
  TotSnrDsc = 0
  ThisOthDsc = 0
  TotOthDsc = 0
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
    
    If ThisTownship <> "No Townships Saved" Then
      If ThisTownship <> "All" Then
        If ThisTownship <> QPTrim$(TaxCust.TownShip) Then
          GoTo SkipIt
        End If
      End If
    End If
    If TaxCust.FirstPersRec > 0 Then
      NextRec = TaxCust.FirstPersRec
      Do While NextRec > 0
        Get PHandle, NextRec, PersRec
        If PersRec.Deleted = 0 Then
          If PersRec.EXMPOTHR > 0 Or PersRec.EXMPSENI > 0 Then
            CustCnt = CustCnt + 1
            NextRec = 0
            GoTo GotCount
            Exit Do
          Else
            NextRec = PersRec.NextRec
          End If
        Else
          NextRec = PersRec.NextRec
        End If
      Loop
    End If

    If TaxCust.FirstPropRec > 0 Then
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted = 0 Then
          If RealRec.EXMPOTHR > 0 Or RealRec.EXMPSENI > 0 Then
            CustCnt = CustCnt + 1
            NextRec = 0
            Exit Do
          Else
            NextRec = RealRec.NextRec
          End If
        Else
          NextRec = RealRec.NextRec
        End If
      
      Loop
    End If
    
GotCount:
    CustName = QPTrim$(TaxCust.CustName)
    ThisSnrDsc = 0
    ThisOthDsc = 0
    ThisBothDsc = 0
    If TaxCust.FirstPropRec > 0 Then
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted = -1 Then GoTo Deleted1
        SnrX = RealRec.EXMPSENI
        OthX = RealRec.EXMPOTHR
        ThisVal = OldRound(SnrX + OthX)
        ThisBothDsc = OldRound(ThisBothDsc + ThisVal)
        TotBothDsc = OldRound(TotBothDsc + ThisVal)
        ThisVal = SnrX
        ThisSnrDsc = OldRound(ThisSnrDsc + ThisVal)
        TotSnrDsc = OldRound(TotSnrDsc + ThisVal)
        ThisVal = OthX
        ThisOthDsc = OldRound(ThisOthDsc + ThisVal)
        TotOthDsc = OldRound(TotOthDsc + ThisVal)
Deleted1:
        NextRec = RealRec.NextRec
      Loop
    End If
    
    If TaxCust.FirstPersRec > 0 Then
      NextRec = TaxCust.FirstPersRec
      Do While NextRec > 0
        Get PHandle, NextRec, PersRec
        If PersRec.Deleted = -1 Then GoTo Deleted2
        SnrX = PersRec.EXMPSENI
        OthX = PersRec.EXMPOTHR
        ThisVal = OldRound(SnrX + OthX)
        ThisBothDsc = OldRound(ThisBothDsc + ThisVal)
        TotBothDsc = OldRound(TotBothDsc + ThisVal)
        ThisVal = SnrX
        ThisSnrDsc = OldRound(ThisSnrDsc + ThisVal)
        TotSnrDsc = OldRound(TotSnrDsc + ThisVal)
        ThisVal = OthX
        ThisOthDsc = OldRound(ThisOthDsc + ThisVal)
        TotOthDsc = OldRound(TotOthDsc + ThisVal)
Deleted2:
        NextRec = PersRec.NextRec
      Loop
    End If
    ThisLoss = OldRound((ThisBothDsc * ThisRate) / 100)
    TotLoss = OldRound(ThisLoss + TotLoss)
    ThisSnrLoss = OldRound((ThisSnrDsc * ThisRate) / 100)
    TotSnrLoss = OldRound(TotSnrLoss + ThisSnrLoss)
    ThisOthLoss = OldRound((ThisOthDsc * ThisRate) / 100)
    TotOthLoss = OldRound(TotOthLoss + ThisOthLoss)
    If ThisLoss > 0 Then
      If LineCnt >= MaxLines - 4 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
      Print #RptHandle, Using("####0", TaxCust.Acct); Tab(8); QPTrim$(TaxCust.CustName);
      Print #RptHandle, Tab(52); QPTrim$(TaxCust.Addr1) + "  " + QPTrim$(TaxCust.City) + ", " + QPTrim$(TaxCust.State) + "  " + QPTrim$(TaxCust.Zip)
      If UseOpt = "Y" Then
        Print #RptHandle, Tab(8); ThisOpt + ": " + QPTrim$(TaxCust.OptSrchDesc)
        LineCnt = LineCnt + 1
      End If
      Print #RptHandle, Tab(5); Using$("$###,##0.00", ThisOthDsc); Tab(21); Using$("$###,##0.00", ThisOthLoss); Tab(37); Using$("$###,##0.00", ThisSnrDsc);
      Print #RptHandle, Tab(53); Using$("$###,##0.00", ThisSnrLoss); Tab(69); Using$("$###,##0.00", ThisBothDsc); Tab(85); Using$("$###,##0.00", ThisLoss)
      Print #RptHandle, String(99, "-")
      LineCnt = LineCnt + 3

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
  
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  If CustCnt = 0 Then
    Call TaxMsg(900, "There are no customers with discounts with the parameters entered.")
    Exit Sub
  End If
  GoSub PrintSummary
  Print #RptHandle, FF$
  Close
  ViewPrint RptFile, "Tax Discount Analysis Report", True
  Exit Sub

PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Discount Analysis Report"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Township:"; Tab(11); ThisTownship
  Print #RptHandle, "Tax Rate (per 100): " + Using$("##0.00", ThisRate)
  Print #RptHandle, "Acct#"; Tab(8); "Customer Name"; Tab(52); "Address"
  Print #RptHandle, Tab(6); "Other Disc"; Tab(22); "Other Loss"; Tab(37); "Senior Disc"; Tab(53); "Senior Loss"; Tab(70); "Total Disc"; Tab(86); "Total Loss"
  Print #RptHandle, String(99, "-")
  LineCnt = 8
  
  Return

PrintSummary:
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  Print #RptHandle,
  Print #RptHandle, "Total Customers Printed: " + Using$("####0", CustCnt)
  Print #RptHandle,
  Print #RptHandle, Tab(31); "Discount Value"; Tab(59); "Tax  Losses"
  Print #RptHandle, Tab(5); "Other Discounts"; Tab(30); Using$("$###,###,##0.00", TotOthDsc); Tab(55); Using$("$###,###,##0.00", TotOthLoss)
  Print #RptHandle, Tab(5); "Senior Discount"; Tab(30); Using$("$###,###,##0.00", TotSnrDsc); Tab(55); Using$("$###,###,##0.00", TotSnrLoss)
  Print #RptHandle, Tab(31); String(39, "-")
  Print #RptHandle, Tab(5); "Grand Totals"; Tab(30); Using$("$###,###,##0.00", TotBothDsc); Tab(55); Using$("$###,###,##0.00", TotLoss)
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxSeniorDscRpt", "PrintText", Erl)
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
