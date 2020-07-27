VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmPayRateRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Pay Rate Report"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmPayRateRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6345
      Left            =   2160
      TabIndex        =   1
      Top             =   1170
      Width           =   7350
      _Version        =   196609
      _ExtentX        =   12965
      _ExtentY        =   11192
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmPayRateRpt.frx":08CA
      Begin LpLib.fpCombo fpcmbParameters 
         Height          =   405
         Left            =   3120
         TabIndex        =   4
         Top             =   3720
         Width           =   2595
         _Version        =   196608
         _ExtentX        =   4577
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
         ColDesigner     =   "frmPayRateRpt.frx":08E6
      End
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3120
         TabIndex        =   5
         Top             =   4395
         Width           =   2595
         _Version        =   196608
         _ExtentX        =   4577
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
         ColDesigner     =   "frmPayRateRpt.frx":0C15
      End
      Begin LpLib.fpCombo fpcmbLast 
         Height          =   390
         Left            =   2160
         TabIndex        =   2
         Top             =   2400
         Width           =   4455
         _Version        =   196608
         _ExtentX        =   7858
         _ExtentY        =   688
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
         Columns         =   3
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
         EditAlignH      =   0
         EditAlignV      =   0
         ColDesigner     =   "frmPayRateRpt.frx":0F44
      End
      Begin LpLib.fpCombo fpcmbFirst 
         Height          =   390
         Left            =   2160
         TabIndex        =   0
         Top             =   1800
         Width           =   4455
         _Version        =   196608
         _ExtentX        =   7858
         _ExtentY        =   688
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
         Columns         =   3
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
         EditAlignH      =   0
         EditAlignV      =   0
         ColDesigner     =   "frmPayRateRpt.frx":12F7
      End
      Begin VB.CheckBox chkTerm 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Include Terminated Employees"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   3000
         Width           =   3495
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4200
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to process the 'Employee Pay Rate Report' report."
         Top             =   5280
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
         ButtonDesigner  =   "frmPayRateRpt.frx":16AA
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1320
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   5280
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
         ButtonDesigner  =   "frmPayRateRpt.frx":1889
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Last Employee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   12
         Top             =   2445
         Width           =   1500
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "First Employee"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   11
         Top             =   1875
         Width           =   1620
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Type:"
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
         Left            =   960
         TabIndex        =   10
         Top             =   3840
         Width           =   1950
      End
      Begin VB.Label Label4 
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
         Left            =   1395
         TabIndex        =   9
         Top             =   4485
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   1230
         Top             =   525
         Width           =   5055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Pay Rate History Report"
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
         Height          =   495
         Left            =   1260
         TabIndex        =   8
         Top             =   675
         Width           =   4995
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   6660
      Left            =   1980
      Top             =   1035
      Width           =   7695
   End
End
Attribute VB_Name = "frmPayRateRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim FirstThru As Boolean

Private Sub chkTerm_Click()
  Call LoadFirstCmb
  Call LoadLastCmb
End Sub

Private Sub cmdEscape_Click()
  frmReportsProcessing.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  Dim First$
  Dim Last$
  
  fpcmbFirst.Col = 0
  First = QPTrim$(fpcmbFirst.ColText)
  fpcmbLast.Col = 0
  Last = QPTrim$(fpcmbLast.ColText)
  
  If Val(First) > Val(Last) Then
    MsgBox "The first employee number must be before the last employee number. Please re-select the employee range."
    fpcmbFirst.SetFocus
    Exit Sub
  End If
  
  If fpcmbFirst.Text = "None" Then
    MsgBox "There are no employees that fit the parameters entered. Report processing aborted."
    Exit Sub
  End If
  
  If fpcomboPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    Call PrintText
  Else
    Exit Sub
  End If
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      Call cmdProcess_Click
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
  FirstThru = True
  Call LoadMe
  FirstThru = False
  Me.HelpContextID = hlpEmployeePayRate
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub fpcmbFirst_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbFirst.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbFirst.ListIndex = -1
  End If
  If fpcmbFirst.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fpcmbLast.Enabled = True Then
        fpcmbLast.SetFocus
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

Private Sub fpcmbLast_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbLast.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbLast.ListIndex = -1
  End If
  If fpcmbLast.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If chkTerm.Enabled = True Then
        chkTerm.SetFocus
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

Private Sub fpcmbParameters_Change()
  If FirstThru = True Then Exit Sub
  If QPTrim$(fpcmbParameters.Text) <> "Full-Time" And _
    QPTrim$(fpcmbParameters.Text) <> "Part-Time" And _
    QPTrim$(fpcmbParameters.Text) <> "Seasonal" And _
    QPTrim$(fpcmbParameters.Text) <> "ALL" And _
    QPTrim$(fpcmbParameters.Text) <> "Temporary" Then
      fpcmbParameters.Text = "ALL"
  End If
  Call LoadFirstCmb
  Call LoadLastCmb
End Sub

Private Sub fpcmbParameters_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbParameters.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbParameters.ListIndex = -1
  End If
  If fpcmbParameters.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fpcomboPrintOpt.Enabled = True Then
        fpcomboPrintOpt.SetFocus
        fpcomboPrintOpt.ListIndex = 0
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

Private Sub fpcomboPrintOpt_Change()
  If QPTrim$(fpcomboPrintOpt.Text) <> "Text" And QPTrim$(fpcomboPrintOpt.Text) <> "Graphical" Then
    fpcomboPrintOpt.Text = "Graphical"
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
      If fpcmbFirst.Enabled = True Then
        fpcmbFirst.SetFocus
        fpcmbFirst.ListIndex = 0
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
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim IdxRec As NumbSortIdxType
  Dim XHandle As Integer
  Dim x As Integer
  Dim NumOfEmpRecs As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenEmpIdxNNameFile XHandle
  NumOfEmpRecs = LOF(XHandle) \ 2
  
  If NumOfEmpRecs = 0 Then 'file is there but there is nothing in it
    MsgBox "No employee index built. No employee list available."
    Close
    Exit Sub
  End If
   
  ReDim EmpIdx(1 To NumOfEmpRecs) As Integer
  For x = 1 To NumOfEmpRecs
    Get XHandle, x, IdxRec.DataRecNum
    EmpIdx(x) = IdxRec.DataRecNum
  Next x
  Close XHandle
  
  If Exist(PRData + EmpData2Name) Then
    OpenEmpData2File EHandle
  Else
    MsgBox "No employee records have been saved."
    Close
    Exit Sub
  End If
  
  For x = 1 To NumOfEmpRecs
    Get EHandle, EmpIdx(x), EmpRec
    If EmpRec.Deleted = -1 Then GoTo BadEmp
    If EmpRec.EMPTDATE > 0 Then GoTo BadEmp
    If Len(QPTrim$(EmpRec.EmpNo)) = 0 Then GoTo BadEmp
    If QPTrim$(fpcmbFirst.Text) = "" Then
      fpcmbFirst.Text = QPTrim$(EmpRec.EmpNo) + Chr(9) + QPTrim$(EmpRec.EmpLName) & ", " & QPTrim$(EmpRec.EmpFName) + Chr(9) + CStr(EmpIdx(x))
    End If
    fpcmbFirst.AddItem QPTrim$(EmpRec.EmpNo) + Chr(9) + QPTrim$(EmpRec.EmpLName) & ", " & QPTrim$(EmpRec.EmpFName) + Chr(9) + CStr(EmpIdx(x))
    fpcmbLast.AddItem QPTrim$(EmpRec.EmpNo) + Chr(9) + QPTrim$(EmpRec.EmpLName) & ", " & QPTrim$(EmpRec.EmpFName) + Chr(9) + CStr(EmpIdx(x))
BadEmp:
  Next x
  fpcmbFirst.ListIndex = 0

  For x = NumOfEmpRecs To 1 Step -1
    Get EHandle, EmpIdx(x), EmpRec
    If EmpRec.Deleted = -1 Then GoTo BadEmp2
    If Len(QPTrim$(EmpRec.EmpNo)) = 0 Then GoTo BadEmp2
    If EmpRec.EMPTDATE > 0 Then GoTo BadEmp2
    fpcmbLast.Text = QPTrim$(EmpRec.EmpNo) + Chr(9) + QPTrim$(EmpRec.EmpLName) & ", " & QPTrim$(EmpRec.EmpFName) + Chr(9) + CStr(EmpIdx(x))
    Exit For
BadEmp2:
  Next x
  
  Close EHandle
'  fpList.Row = 0
'
'  fpList.Selected = True 'set focus to first line
  
ZeroText:
  
  chkTerm.Value = 0
  fpcomboPrintOpt.AddItem "Graphical"
  fpcomboPrintOpt.AddItem "Text"
  fpcomboPrintOpt.Text = "Graphical"
  fpcmbParameters.Text = "ALL"
  fpcmbParameters.AddItem "ALL"
  fpcmbParameters.AddItem "Full-Time"
  fpcmbParameters.AddItem "Part-Time"
  fpcmbParameters.AddItem "Seasonal"
  fpcmbParameters.AddItem "Temporary"
  
  Exit Sub
   

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmPayRateRpt", "LoadMe", Erl)
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmPayRateRpt.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintGraphics()
  Dim dlm$
  Dim PayRec As PayRateType
  Dim PHandle As Integer
  Dim NumOfPayRecs As Integer
  Dim x As Integer
  Dim y As Integer
  Dim RptName$
  Dim RptHandle As Integer
'  Dim AllFlag As Boolean
  Dim Unit As UnitFileRecType
  Dim UHandle As Integer
  Dim TFlag As Boolean
  Dim EmpType As Integer
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
'  Dim IdxRec As PayRateIndexType
  Dim IdxRec As PayRateIdxNumType
  Dim XHandle As Integer
  Dim OldPay As Double
  Dim NewPay As Double
  Dim ThisPct As Double
  Dim PayTypeOld$
  Dim PayTypeNew$
  Dim PayHrSalOld$
  Dim PayHrSalNew$
  Dim RptCnt As Integer
  Dim StartRec$
  Dim EndRec$
  Dim TotEmpCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  fpcmbFirst.Row = fpcmbFirst.ListIndex
  fpcmbFirst.Col = 0
  StartRec = QPTrim$(fpcmbFirst.ColText)
  fpcmbLast.Row = fpcmbLast.ListIndex
  fpcmbLast.Col = 0
  EndRec = QPTrim$(fpcmbLast.ColText)
  
  RptCnt = 0
  If fpcmbParameters.Enabled = True Then
    If QPTrim$(fpcmbParameters.Text) = "ALL" Then
      EmpType = 1
    ElseIf QPTrim$(fpcmbParameters.Text) = "Full-Time" Then
      EmpType = 2
    ElseIf QPTrim$(fpcmbParameters.Text) = "Part-Time" Then
      EmpType = 3
    ElseIf QPTrim$(fpcmbParameters.Text) = "Seasonal" Then
      EmpType = 4
    ElseIf QPTrim$(fpcmbParameters.Text) = "Temporary" Then
      EmpType = 5
    Else
      EmpType = 0
    End If
    If EmpType = 0 Then
      MsgBox "Please make a valid selection from the Parameters drop down list."
      Close
      fpcmbParameters.SetFocus
      Exit Sub
    End If
  End If
  
  OpenUnitFile UHandle
  Get UHandle, 1, Unit
  Close UHandle
  
  TFlag = False
  If chkTerm.Value = 1 Then TFlag = True
  
  dlm$ = "~"
'  AllFlag = False
'  If GEmpNum = -1 Then
'    AllFlag = True
'  End If
  
'  OpenPayRateIdxFile XHandle
  OpenPayRateNumIdxFile XHandle
  NumOfPayRecs = LOF(XHandle) / Len(IdxRec)
  If NumOfPayRecs = 0 Then
    MsgBox "No pay rate records are on file. Unable to print report."
    Close
    Exit Sub
  End If
  
  ReDim PayIdx(1 To NumOfPayRecs) As Integer
  
  For x = 1 To NumOfPayRecs
    Get XHandle, x, IdxRec
    PayIdx(x) = IdxRec.PayRateRec
  Next x
  
  Close XHandle
  
  frmLoadingRpt.Show
  DoEvents
  OpenPayRateFile PHandle
  
  RptName$ = "PRRPTS\PayRate.RPT"
  RptHandle = FreeFile
  Open RptName$ For Output As RptHandle
  
  OpenEmpData2File EHandle
  TotEmpCnt = 0
  For x = 1 To NumOfPayRecs
    Get PHandle, PayIdx(x), PayRec
    If Val(PayRec.EmpNo) < Val(StartRec) Or Val(PayRec.EmpNo) > Val(EndRec) Then GoTo SkipIt
    Get EHandle, PayRec.EmpRecNum, EmpRec
    If Len(QPTrim$(EmpRec.EmpNo)) = 0 Then GoTo SkipIt
    If EmpRec.Deleted = -1 Then GoTo SkipIt
    If TFlag = False Then
      If EmpRec.EMPTDATE > 0 Then GoTo SkipIt '7/15/2010 made PayRec into EmpRec
    End If
    OldPay = 0
    NewPay = 0
    ThisPct = 0
    PayTypeOld = ""
    PayTypeNew = ""
    PayHrSalOld = ""
    PayHrSalNew = ""
    TotEmpCnt = TotEmpCnt + 1
    Select Case EmpType
      Case 2:
        If QPTrim$(EmpRec.EMPSTATS) <> "Full-Time" Then GoTo SkipIt
      Case 3:
        If QPTrim$(EmpRec.EMPSTATS) <> "Part-Time" Then GoTo SkipIt
      Case 4:
        If QPTrim$(EmpRec.EMPSTATS) <> "Seasonal" Then GoTo SkipIt
      Case 5:
        If QPTrim$(EmpRec.EMPSTATS) <> "Temporary" Then GoTo SkipIt
    End Select
    For y = 1 To 30
      If PayRec.RegPayRate(y) > 0 Then
      If NewPay = 0 Then
        OldPay = OldRound(PayRec.RegPayRate(y))
        NewPay = OldRound(PayRec.RegPayRate(y))
        PayTypeOld = QPTrim$(PayRec.EMPPFREQ(y))
        PayTypeNew = QPTrim$(PayRec.EMPPFREQ(y))
        PayHrSalOld = QPTrim$(PayRec.EMPPTYPE(y))
        PayHrSalNew = QPTrim$(PayRec.EMPPTYPE(y))
      Else
        OldPay = NewPay
        NewPay = OldRound(PayRec.RegPayRate(y))
        PayTypeOld = PayTypeNew
        PayTypeNew = QPTrim$(PayRec.EMPPFREQ(y))
        PayHrSalOld = PayHrSalNew
        PayHrSalNew = QPTrim$(PayRec.EMPPTYPE(y))
        ThisPct = FigurePayIncPct(PayHrSalNew, PayHrSalOld, PayTypeOld, PayTypeNew, OldPay, NewPay)
      End If
        RptCnt = RptCnt + 1
        '                            0
        Print #RptHandle, QPTrim$(PayRec.EmpLName) + ", " + QPTrim$(PayRec.EmpFName); dlm;
        '                            1
        Print #RptHandle, QPTrim$(PayRec.EmpNo); dlm;
        If PayRec.EMPTDATE = 0 Then
          '                  2
          Print #RptHandle, ""; dlm;
        Else
          '                                2
          Print #RptHandle, MakeRegDate(PayRec.EMPTDATE); dlm;
        End If
        
        '                            3                               4
        Print #RptHandle, MakeRegDate(PayRec.EMPHDATE); dlm; UCase(QPTrim$(PayRec.EMPPFREQ(y))); dlm;
        '                         5                      6                          7
        Print #RptHandle, PayRec.OTPayRate(y); dlm; PayRec.RegPayRate(y); dlm; QPTrim$(UCase(PayRec.EMPJOB(y))); dlm;
        '              8
        If PayRec.PayChngDate(y) = 0 Then
          Print #RptHandle, "Initial"; dlm;
        Else
          Print #RptHandle, MakeRegDate(PayRec.PayChngDate(y)); dlm;
        End If
        Print #RptHandle, QPTrim$(Unit.UFEMPR); dlm; ThisPct; dlm; TotEmpCnt
      Else
        Exit For
      End If
    Next y
SkipIt:
  Next x
  
  Close
  Unload frmLoadingRpt
  If RptCnt = 0 Then
    MsgBox "There are no records saved for the current report parameters."
    Exit Sub
  End If
  arPayRate.Show

  Exit Sub
  
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmPayRateRpt", "PrintGraphics", Erl)
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

Private Sub fpList_DblClick()
  Call cmdProcess_Click
End Sub

Private Sub PrintText()
  Dim MaxLines As Integer, LineCnt As Integer
  Dim PayRec As PayRateType
  Dim PHandle As Integer
  Dim NumOfPayRecs As Integer
  Dim x As Integer
  Dim y As Integer, RptTitle$
  Dim RptName$, Page As Integer
  Dim RptHandle As Integer
'  Dim AllFlag As Boolean
  Dim Unit As UnitFileRecType
  Dim UHandle As Integer
  Dim TFlag As Boolean, FF$
  Dim EmpType As Integer
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim Dash As String * 80
'  Dim IdxRec As PayRateIndexType
  Dim IdxRec As PayRateIdxNumType
  Dim XHandle As Integer
  Dim NewPay As Double
  Dim OldPay As Double
  Dim ThisPct As Double
  Dim PayTypeOld$
  Dim PayTypeNew$
  Dim PayHrSalOld$
  Dim PayHrSalNew$
  Dim RptCnt As Integer
  Dim JobTitle As String * 26
  Dim StartRec$
  Dim EndRec$
  Dim TotEmpCnt As Integer
  
  On Error GoTo ERRORSTUFF
  fpcmbFirst.Row = fpcmbFirst.ListIndex
  fpcmbFirst.Col = 0
  StartRec = QPTrim$(fpcmbFirst.ColText)
  fpcmbLast.Row = fpcmbLast.ListIndex
  fpcmbLast.Col = 0
  EndRec = QPTrim$(fpcmbLast.ColText)
  
  RptCnt = 0
  FF$ = Chr(12)
  MaxLines = 57
  LineCnt = 0
  
  If fpcmbParameters.Enabled = True Then
    If QPTrim$(fpcmbParameters.Text) = "ALL" Then
      EmpType = 1
    ElseIf QPTrim$(fpcmbParameters.Text) = "Full-Time" Then
      EmpType = 2
    ElseIf QPTrim$(fpcmbParameters.Text) = "Part-Time" Then
      EmpType = 3
    ElseIf QPTrim$(fpcmbParameters.Text) = "Seasonal" Then
      EmpType = 4
    ElseIf QPTrim$(fpcmbParameters.Text) = "Temporary" Then
      EmpType = 5
    Else
      EmpType = 0
    End If
    If EmpType = 0 Then
      MsgBox "Please make a valid selection from the Parameters drop down list."
      Close
      fpcmbParameters.SetFocus
      Exit Sub
    End If
  End If
  
  OpenUnitFile UHandle
  Get UHandle, 1, Unit
  Close UHandle
  
  TFlag = False
  If chkTerm.Value = 1 Then TFlag = True
  
'  AllFlag = False
'  If GEmpNum = -1 Then
'    AllFlag = True
'  End If
  
'  OpenPayRateIdxFile XHandle
  OpenPayRateNumIdxFile XHandle
  
  NumOfPayRecs = LOF(XHandle) / Len(IdxRec)
  If NumOfPayRecs = 0 Then
    MsgBox "No pay rate records are on file. Unable to print report."
    Close
    Exit Sub
  End If
  
  ReDim PayIdx(1 To NumOfPayRecs) As Integer
  
  For x = 1 To NumOfPayRecs
    Get XHandle, x, IdxRec
    PayIdx(x) = IdxRec.PayRateRec
  Next x
  
  Close XHandle
  
  OpenPayRateFile PHandle
  
  RptTitle$ = "Employee Pay Rate Report"
  RptName$ = "PRRPTS\PayRate.RPT"
  RptHandle = FreeFile
  Open RptName$ For Output As RptHandle
  GoSub PrintHeader
  
  OpenEmpData2File EHandle
  frmLoadingRpt.Show
  DoEvents
  TotEmpCnt = 0
  
  For x = 1 To NumOfPayRecs
    Get PHandle, PayIdx(x), PayRec
    If Val(PayRec.EmpNo) < Val(StartRec) Or Val(PayRec.EmpNo) > Val(EndRec) Then GoTo SkipIt
    Get EHandle, PayRec.EmpRecNum, EmpRec
    If TFlag = False Then
      If EmpRec.EMPTDATE > 0 Then GoTo SkipIt '7/15/2010 made PayRec into EmpRec
    End If
    If EmpRec.Deleted = 1 Then GoTo SkipIt
    If QPTrim$(EmpRec.EmpNo) = "0" Then GoTo SkipIt
    Select Case EmpType
      Case 2:
        If QPTrim$(EmpRec.EMPSTATS) <> "Full-Time" Then GoTo SkipIt
      Case 3:
        If QPTrim$(EmpRec.EMPSTATS) <> "Part-Time" Then GoTo SkipIt
      Case 4:
        If QPTrim$(EmpRec.EMPSTATS) <> "Seasonal" Then GoTo SkipIt
      Case 5:
        If QPTrim$(EmpRec.EMPSTATS) <> "Temporary" Then GoTo SkipIt
    End Select
    
    GoSub PrintEmpHeader
    NewPay = 0
    OldPay = 0
    ThisPct = 0
    PayTypeOld = ""
    PayTypeNew = ""
    PayHrSalOld = ""
    PayHrSalNew = ""
    For y = 1 To 30
      If NewPay = 0 Then
        OldPay = OldRound(PayRec.RegPayRate(y))
        NewPay = OldRound(PayRec.RegPayRate(y))
        PayTypeOld = QPTrim$(PayRec.EMPPFREQ(y))
        PayTypeNew = QPTrim$(PayRec.EMPPFREQ(y))
        PayHrSalOld = QPTrim$(PayRec.EMPPTYPE(y))
        PayHrSalNew = QPTrim$(PayRec.EMPPTYPE(y))
      Else
        OldPay = NewPay
        NewPay = OldRound(PayRec.RegPayRate(y))
        PayTypeOld = PayTypeNew
        PayTypeNew = QPTrim$(PayRec.EMPPFREQ(y))
        PayHrSalOld = PayHrSalNew
        PayHrSalNew = QPTrim$(PayRec.EMPPTYPE(y))
        ThisPct = FigurePayIncPct(PayHrSalNew, PayHrSalOld, PayTypeOld, PayTypeNew, OldPay, NewPay)
      End If
        
      If PayRec.RegPayRate(y) > 0 Then
        If LineCnt > MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          GoSub PrintEmpHeader
        End If
        Select Case UCase(QPTrim$(PayRec.EMPPFREQ(y)))
          Case "WEEKLY"
            Print #RptHandle, "Weekly";
          Case "BI-WEEKLY"
            Print #RptHandle, "Bi-Wkly";
          Case "SEMI-MONTHLY"
            Print #RptHandle, "Semi-Mth";
          Case "MONTHLY"
            Print #RptHandle, "Monthly";
          Case "QUARTERLY"
            Print #RptHandle, "Qrtly";
          Case "SEMI-ANNUALLY"
            Print #RptHandle, "Semi-Ann";
          Case "ANNUALLY"
            Print #RptHandle, "Annually";
          Case Else
            Print #RptHandle, "Unknown";
        End Select
          
        RptCnt = RptCnt + 1
        If PayRec.PayChngDate(y) = 0 Then
          Print #RptHandle, Tab(11); "Initial";
        Else
          Print #RptHandle, Tab(11); MakeRegDate(PayRec.PayChngDate(y));
        End If
        LineCnt = LineCnt + 1
        
        RSet JobTitle = QPTrim$(PayRec.EMPJOB(y))
        Print #RptHandle, Tab(23); Using$("$##,##0.00", PayRec.RegPayRate(y)); Tab(35); Using("##0.00%", ThisPct); Tab(42); Using$("$##,##0.00", PayRec.OTPayRate(y)); Tab(55); JobTitle
        LineCnt = LineCnt + 1
      Else
        Exit For
      End If
    Next y
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle,
    LineCnt = LineCnt + 3
    TotEmpCnt = TotEmpCnt + 1
SkipIt:
  Next x
  Print #RptHandle,
  Print #RptHandle, Tab(5); "Total Employees: " + Using$("###0", TotEmpCnt)
  
  Print #RptHandle, FF$
  
  Close
  Unload frmLoadingRpt
  If RptCnt = 0 Then
    MsgBox "There are no records saved for the current report parameters."
    Exit Sub
  End If
  ViewPrint RptName$, RptTitle$
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(27); "Employee Pay Rate Report"
  Print #RptHandle,
  Print #RptHandle, "Employer: " + QPTrim$(Unit.UFEMPR); Tab(71); "Page# " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  If fpcmbParameters.Enabled = True Then
    Print #RptHandle, "Employee Type: " + QPTrim$(fpcmbParameters.Text)
    LineCnt = 1
  End If
  Print #RptHandle, String$(82, "-")
  Print #RptHandle,
  LineCnt = LineCnt + 6
  Return
  
PrintEmpHeader:
  If LineCnt >= MaxLines - 5 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  Print #RptHandle, QPTrim$(PayRec.EmpLName) + ", " + QPTrim$(PayRec.EmpFName); Tab(60); "Emp # " + QPTrim$(PayRec.EmpNo)
  If PayRec.EMPTDATE <> 0 Then
    Print #RptHandle, "Hire Date: " + MakeRegDate(PayRec.EMPHDATE);
    Print #RptHandle, Tab(55); "Termination Date: " + MakeRegDate(PayRec.EMPTDATE)
  Else
    Print #RptHandle, "Hire Date: " + MakeRegDate(PayRec.EMPHDATE)
  End If
  Print #RptHandle,
  Print #RptHandle, "Frequency"; Tab(13); "Date"; Tab(25); "Pay Rate"; Tab(36); "% +/-"; Tab(45); "OT Rate"; Tab(65); "Job Title"
  Print #RptHandle, String$(82, "-")
  LineCnt = LineCnt + 6
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmPayRateRpt", "PrintText", Erl)
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

Private Sub LoadFirstCmb()
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim IdxRec As NumbSortIdxType
  Dim XHandle As Integer
  Dim x As Integer
  Dim NumOfEmpRecs As Integer
  Dim ValidCnt As Integer
  
  OpenEmpIdxNNameFile XHandle
  NumOfEmpRecs = LOF(XHandle) \ 2
  
  If NumOfEmpRecs = 0 Then 'file is there but there is nothing in it
    MsgBox "No employee index built. No employee list available."
    Close
    Exit Sub
  End If
   
  ReDim EmpIdx(1 To NumOfEmpRecs) As Integer
  For x = 1 To NumOfEmpRecs
    Get XHandle, x, IdxRec.DataRecNum
    EmpIdx(x) = IdxRec.DataRecNum
  Next x
  Close XHandle
  
  OpenEmpData2File EHandle
  fpcmbFirst.Clear
  ValidCnt = 0
  For x = 1 To NumOfEmpRecs
    Get EHandle, EmpIdx(x), EmpRec
    If EmpRec.Deleted = -1 Then GoTo BadEmp
    If Len(QPTrim$(EmpRec.EmpNo)) = 0 Then GoTo BadEmp
    If chkTerm.Value = 0 Then
      If EmpRec.EMPTDATE > 0 Then GoTo BadEmp
    End If
    Select Case Mid(fpcmbParameters.Text, 1, 1)
      Case "A"
      Case "F"
        If Mid(EmpRec.EMPSTATS, 1, 1) <> "F" Then GoTo BadEmp
      Case "P"
        If Mid(EmpRec.EMPSTATS, 1, 1) <> "P" Then GoTo BadEmp
      Case "T"
        If Mid(EmpRec.EMPSTATS, 1, 1) <> "T" Then GoTo BadEmp
      Case "S"
        If Mid(EmpRec.EMPSTATS, 1, 1) <> "S" Then GoTo BadEmp
      Case Else
        GoTo BadEmp
    End Select
    If QPTrim$(fpcmbFirst.Text) = "" Or QPTrim$(fpcmbFirst.Text) = "None" Then
      fpcmbFirst.Text = QPTrim$(EmpRec.EmpNo) + Chr(9) + QPTrim$(EmpRec.EmpLName) & ", " & QPTrim$(EmpRec.EmpFName) + Chr(9) + CStr(EmpIdx(x))
    End If
    fpcmbFirst.AddItem QPTrim$(EmpRec.EmpNo) + Chr(9) + QPTrim$(EmpRec.EmpLName) & ", " & QPTrim$(EmpRec.EmpFName) + Chr(9) + CStr(EmpIdx(x))
    ValidCnt = ValidCnt + 1
BadEmp:
  Next x
  
  If ValidCnt = 0 Then
    fpcmbFirst.Text = "None"
  End If
  
  fpcmbFirst.ListIndex = 0

  Close EHandle

End Sub

Private Sub LoadLastCmb()
  Dim EmpRec As EmpData2Type
  Dim EHandle As Integer
  Dim IdxRec As NumbSortIdxType
  Dim XHandle As Integer
  Dim x As Integer
  Dim NumOfEmpRecs As Integer
  Dim ValidCnt As Integer
  Dim ValidX As Integer
  Dim ValidEmp$
  Dim ValidNum$
  
  OpenEmpIdxNNameFile XHandle
  NumOfEmpRecs = LOF(XHandle) \ 2
  
  If NumOfEmpRecs = 0 Then 'file is there but there is nothing in it
    MsgBox "No employee index built. No employee list available."
    Close
    Exit Sub
  End If
   
  ReDim EmpIdx(1 To NumOfEmpRecs) As Integer
  For x = 1 To NumOfEmpRecs
    Get XHandle, x, IdxRec.DataRecNum
    EmpIdx(x) = IdxRec.DataRecNum
  Next x
  Close XHandle
  ValidCnt = 0
  
  OpenEmpData2File EHandle
  
  fpcmbLast.Clear
  
  For x = 1 To NumOfEmpRecs
    Get EHandle, EmpIdx(x), EmpRec
    If EmpRec.Deleted = -1 Then GoTo BadEmp2
    If Len(QPTrim$(EmpRec.EmpNo)) = 0 Then GoTo BadEmp2
    If chkTerm.Value = 0 Then
      If EmpRec.EMPTDATE > 0 Then GoTo BadEmp2
    End If
    Select Case Mid(fpcmbParameters.Text, 1, 1)
      Case "A"
      Case "F"
        If Mid(EmpRec.EMPSTATS, 1, 1) <> "F" Then GoTo BadEmp2
      Case "P"
        If Mid(EmpRec.EMPSTATS, 1, 1) <> "P" Then GoTo BadEmp2
      Case "T"
        If Mid(EmpRec.EMPSTATS, 1, 1) <> "T" Then GoTo BadEmp2
      Case "S"
        If Mid(EmpRec.EMPSTATS, 1, 1) <> "S" Then GoTo BadEmp2
      Case Else
        GoTo BadEmp2
      End Select
      fpcmbLast.AddItem QPTrim$(EmpRec.EmpNo) + Chr(9) + QPTrim$(EmpRec.EmpLName) & ", " & QPTrim$(EmpRec.EmpFName) + Chr(9) + CStr(EmpIdx(x))
      ValidCnt = ValidCnt + 1
      ValidEmp = QPTrim$(EmpRec.EmpLName) & ", " & QPTrim$(EmpRec.EmpFName)
      ValidNum = QPTrim$(EmpRec.EmpNo)
      ValidX = x
BadEmp2:
  Next x
  
  If ValidCnt = 0 Then
    fpcmbLast.Text = "None"
  Else
    fpcmbLast.Text = ValidNum + Chr(9) + ValidEmp + Chr(9) + CStr(EmpIdx(ValidX))
  End If
  
  Close EHandle

End Sub
