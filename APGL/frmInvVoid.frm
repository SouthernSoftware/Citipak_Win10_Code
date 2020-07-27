VERSION 5.00
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvVoid 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Void An Open Invoice "
   ClientHeight    =   8850
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmInvVoid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   405
      Left            =   4560
      TabIndex        =   3
      Top             =   7245
      Width           =   1905
      _Version        =   196608
      _ExtentX        =   3360
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
      ColumnEdit      =   0
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   3504
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
      ColDesigner     =   "frmInvVoid.frx":08CA
   End
   Begin LpLib.fpCombo fpcboVendName 
      Height          =   375
      Left            =   5130
      TabIndex        =   0
      Top             =   2475
      Width           =   4140
      _Version        =   196608
      _ExtentX        =   7302
      _ExtentY        =   661
      Text            =   ""
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
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
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
      EditMarginLeft  =   5
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmInvVoid.frx":0CA0
   End
   Begin LpLib.fpList lstInvoices 
      Height          =   1860
      Left            =   3315
      TabIndex        =   1
      Top             =   3390
      Width           =   5745
      _Version        =   196608
      _ExtentX        =   10134
      _ExtentY        =   3281
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
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
      Columns         =   0
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
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
      ColDesigner     =   "frmInvVoid.frx":10C3
   End
   Begin LpLib.fpCombo fpcboVoidDate 
      Height          =   405
      Left            =   4560
      TabIndex        =   2
      Top             =   6570
      Width           =   2895
      _Version        =   196608
      _ExtentX        =   5106
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
      Columns         =   1
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
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
      EditMarginLeft  =   5
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmInvVoid.frx":13F7
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Esc E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   8760
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7224
      Width           =   1236
   End
   Begin VB.CommandButton cmdVoid 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Void"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   6972
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7224
      Width           =   1236
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   8484
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "1:24 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "5/14/2018"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Void Date:"
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
      Height          =   372
      Left            =   2880
      TabIndex        =   14
      Top             =   6648
      Width           =   1428
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Type:"
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
      Height          =   372
      Left            =   1920
      TabIndex        =   13
      Top             =   7320
      Width           =   2388
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To Select Invoice to Void, Press SpaceBar or Arrow Key To Highlight."
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
      Height          =   684
      Left            =   2376
      TabIndex        =   12
      Top             =   5568
      Width           =   7452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice "
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
      Index           =   7
      Left            =   3528
      TabIndex        =   11
      Top             =   3072
      Width           =   732
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Date"
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
      Index           =   1
      Left            =   5448
      TabIndex        =   10
      Top             =   3072
      Width           =   1308
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Index           =   3
      Left            =   8088
      TabIndex        =   9
      Top             =   3072
      Width           =   852
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2196
      Left            =   3204
      Top             =   3288
      Width           =   5940
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Code:"
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
      Height          =   300
      Left            =   2952
      TabIndex        =   8
      Top             =   2520
      Width           =   1692
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4236
      Left            =   2328
      Top             =   2040
      Width           =   7548
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   852
      Left            =   2580
      Top             =   792
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Void Open Invoice"
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
      Left            =   3684
      TabIndex        =   7
      Top             =   1032
      Width           =   4836
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   996
      Left            =   2592
      Top             =   672
      Width           =   7020
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmInvVoid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Temp_Class As Resize_Class
Dim TempRec As Long, txtInv As String
Dim APAcct As String, LPDate As Integer, HPDate As Integer

Private Sub cmdExit_Click()
  frmInvProcessMenu.Show
  Unload frmInvVoid
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = True Then
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
          Cancel = True
      Else
        MainLog "Close AP"
        ClearInUse PWcnt
      End If
    Else
      Cancel = True
    End If
  End If
End Sub

Private Sub lstInvoices_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
   If lstInvoices.Selected = True Then
    lstInvoices.Selected = False
   Else
    lstInvoices.Selected = True
   End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%V"
      KeyCode = 0
    Case vbKeyF2:
      SendKeys "%S"
      KeyCode = 0
    Case Else:
  End Select

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
'Private Sub cmdSelect_Click()
'  If Not lstInvoices = "" Then
'    cmdVoid.SetFocus
'  Else
'    MsgBox "You Must Choose An Invoice.", vbOKOnly, "Invalid Selection"
'    lstInvoices.SetFocus
'  End If
'End Sub

Private Sub cmdVoid_Click()
  Dim cnt As Integer, Pcnt As Integer, TmpDate As Integer
  If fpcboRptType.ListIndex = 0 Then
    rptopt = 1
  ElseIf fpcboRptType.ListIndex = 1 Then
    rptopt = 2
  End If
  If rptopt > 0 Then
   If lstInvoices.ListCount > 0 Then
     cnt = 0
     For Pcnt = 0 To lstInvoices.ListCount - 1
       If lstInvoices.Selected(Pcnt) Then
        cnt = 1
        lstInvoices.Row = Pcnt
        txtInv = Mid$(lstInvoices.Text, 1, 10)
        TempRec = Val(Mid$(lstInvoices.Text, 82))
        If fpcboVoidDate.ListIndex <> 0 Then
          TmpDate = DateDiff("d", "12/31/1979", Date$)
        Else
          TmpDate = DateDiff("d", "12/31/1979", (Mid$(lstInvoices.Text, 17, 10)))
        End If
        If (TmpDate < LPDate) Or (TmpDate > HPDate) Then
          If MsgBox("This Date Is Not Within Allowable Posting Range.Ok to Continue or Cancel?", vbOKCancel, "Invalid Date") = vbCancel Then
            Call MainLog("APInvVoid Cancel - DateOutofRange, Canceled")
            Close
            Exit Sub
          End If
        End If
        If MsgBox("Are You Sure You Wish To Void This Invoice?", vbYesNo, "Void Invoice") = vbNo Then
          Exit Sub
        End If
        Me.cmdExit.Enabled = False
        If rptopt = 1 Then
          VoidInvTrans TempRec
        ElseIf rptopt = 2 Then
          VoidInvTrans2 TempRec
        End If
        lstInvoices.Clear
        fpcboVendName.ListIndex = -1
        Me.cmdExit.Enabled = True
        If cnt > 0 Then Exit For
        End If
      Next
      If cnt = 0 Then
        MsgBox "You Must First Select An Invoice.", vbOKOnly, "Invalid Selection"
      End If
    End If
  End If
End Sub
Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdVoid.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboVoidDate.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  GetAPAcct APAcct
  GetPostDates LPDate, HPDate
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpVoidInv
  VendCodeName fpcboVendName
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  fpcboVoidDate.InsertRow = "Invoice Date"
  fpcboVoidDate.InsertRow = "Current Date"
  fpcboVoidDate.ListIndex = 0
End Sub
Private Sub fpcboVendName_Click()
  If fpcboVendName.ListIndex <> -1 Then
    ListEntries
  End If
End Sub
Private Sub fpcboVoidDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVoidDate.ListDown = True
  End If
  If fpcboVoidDate.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        lstInvoices.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub fpcboVendName_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVendName.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboVendName.ListIndex = -1
    fpcboVendName.Action = ActionClearSearchBuffer
  End If
  If fpcboVendName.ListDown <> True Then
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

Private Sub fpcboVendName_LostFocus()
  fpcboVendName.Action = ActionClearSearchBuffer
End Sub

Private Sub ListEntries()
'this fills the listbox with iNVOICE Entries for User to Select
  Dim cnt As Integer, VoidVendorRecNum As Long, disdate As String
  Dim fmt As String, tempstr As String, strInfo As String
  Dim LdRecLen As Integer, VendorFile As Integer, NumVRecs As Integer
  Dim APLedgerFile As Integer, NumTran As Long, NextTrans As Long
  Dim APLedgerRec(1) As APLedger81RecType
  Dim VendorIdx As VendorIdxRecType
  Dim Vendor As VendorRecType
  LdRecLen = Len(APLedgerRec(1))
  fmt = "$########.##"
  fpcboVendName.col = 2
  VoidVendorRecNum = fpcboVendName.ColText
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTran&, LdRecLen
  Get VendorFile, VoidVendorRecNum&, Vendor
  NextTrans& = Vendor.FrstTran
  lstInvoices.Clear
  
  Do Until NextTrans& = 0
    Get APLedgerFile, NextTrans&, APLedgerRec(1)
    If APLedgerRec(1).TRCode = 1 And APLedgerRec(1).PAYCODE = 1 Then
      tempstr = Space$(90)
      Mid$(tempstr, 1, 10) = QPTrim$(APLedgerRec(1).DOCNum)
      Mid$(tempstr, 17) = Format(DateAdd("d", (APLedgerRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
      Mid$(tempstr, 32) = Using("$###,###,###.##", Str$(APLedgerRec(1).Amt))
      Mid$(tempstr, 82) = NextTrans
      lstInvoices.AddItem tempstr
    End If
    NextTrans& = APLedgerRec(1).NextTrans
  Loop
  Close
End Sub

'Private Sub lstInvoices_Click()
'  lstInvoices_DblClick
'End Sub

'Private Sub lstInvoices_DblClick()
'  TempRec = Val(Mid$(lstInvoices, 82))
'  fptxtInv = Mid$(lstInvoices, 1, 10)
'End Sub


Private Sub VoidInvTrans(VoidTransRecNum&)
  Dim Interfacefile As String, LdRecLen As Integer, DistRecLEn As Integer
  Dim GLIFRecLen As Integer, PRNfileName As String, PrnTitle As String
  Dim fmt As String, GLIFFile As Integer, PRNFile As Integer, ToPrint As String
  Dim NextDist As Long, APLedgerFile As Integer, NumTran As Long
  Dim APDistFile As Integer, NumDistRecs As Long, IFRec As Integer
  Dim GTDebit As Double, GTCredit As Double, BadAcct As Integer
  Dim Linecnt As Integer
  '-If this file exists this routine did not finish last time.
  Interfacefile$ = "APVDIF.DAT"
  If Exist(Interfacefile$) Then
    'KILL InterfaceFile$
    Close
    MsgBox "The file APVDIF.DAT already exits. Cannot continue.", vbOKOnly, "Void Canceled"
    Exit Sub
  End If
  Call MainLog("Start AP Void Inv - " + txtInv + fpcboVoidDate.Text)
  ReDim ApLedger(1) As APLedger81RecType
  ReDim DistRec(1) As APDistRecType
  ReDim GLifRec(1) As GLTransRecType
  LdRecLen = Len(ApLedger(1))
  DistRecLEn = Len(DistRec(1))
  GLIFRecLen = Len(GLifRec(1))
  Linecnt = 0
  PRNfileName$ = "VOIDINV.PRN"
  PrnTitle$ = "Voided Invoice Report"
  fmt$ = "###,###,###.##"
  
  OpenAPLedgerFile APLedgerFile, NumTran&, LdRecLen

  Get APLedgerFile, VoidTransRecNum&, ApLedger(1)
  Close APLedgerFile

  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLEn

  GLIFFile = FreeFile
  Open Interfacefile$ For Random As GLIFFile Len = GLIFRecLen

  PRNFile = FreeFile
  Open PRNfileName$ For Output As PRNFile
  'GoSub PrintReportHeader

  NextDist& = ApLedger(1).FrstDist
  Do Until NextDist& = 0
    Get APDistFile, NextDist&, DistRec(1)
    IFRec = IFRec + 1

    '--Set Entry Description Defaults
    GLifRec(1).Src = "VI" + Format$(Now, "mmddyy")
    If fpcboVoidDate.ListIndex <> 0 Then
      GLifRec(1).TRDATE = DateDiff("d", "12/31/1979", Date$)
    Else
      GLifRec(1).TRDATE = ApLedger(1).TRDATE
    End If
    GLifRec(1).Desc = ApLedger(1).DOCNum
    GLifRec(1).LDesc = ApLedger(1).Comment
    GLifRec(1).Ref = "TO VOID"

    '--Make Debit side of entry
    GLifRec(1).AcctNum = Left$(DistRec(1).DistAcctNum, GLFundLen) + APAcct$
    GLifRec(1).CrAmt = 0
    GLifRec(1).DrAmt = DistRec(1).DistAmt
    Put GLIFFile, IFRec, GLifRec(1)
    GTDebit# = GTDebit# + GLifRec(1).DrAmt
    GoSub PrintGLIF

    '--Make Credit side
    IFRec = IFRec + 1
    GLifRec(1).AcctNum = DistRec(1).DistAcctNum
    GLifRec(1).CrAmt = DistRec(1).DistAmt
    GLifRec(1).DrAmt = 0
    Put GLIFFile, IFRec, GLifRec(1)
    GTCredit# = GTCredit# + GLifRec(1).CrAmt
    GoSub PrintGLIF

    NextDist& = DistRec(1).NextDist

  Loop
'  Mid$(ToPrint$, 1) = "Grand Total"
'  Mid$(ToPrint$, 53) = Using(fmt$, GTDebit#)
'  Mid$(ToPrint$, 67) = Using(fmt$, GTCredit#)
'  Print #PRNFile, ToPrint$
  Close

  Post2GL Interfacefile$, BadAcct, frmInvVoid, False
  If BadAcct Then
    Call MainLog("APInvVoid Error - Accts Not Update, Canceled")
    MsgBox "Void Error.  Review GL Utility Log for details.", vbOKOnly, "Void aborted"
    ARptErrorLog.GetName PRNfileName$
    ARptErrorLog.startrpt
    MsgBox "Accounts were not updated.", vbOKOnly, "Void Canceled"
    Exit Sub
  End If
 
  Post2GL Interfacefile$, BadAcct, frmInvVoid, True
  KillFile Interfacefile$
  
  OpenAPLedgerFile APLedgerFile, NumTran&, LdRecLen
  Get APLedgerFile, VoidTransRecNum&, ApLedger(1)
  ApLedger(1).TRCode = -1
  Put APLedgerFile, VoidTransRecNum&, ApLedger(1)
  Close APLedgerFile
  Call MainLog("APInvVoid Completed.")
  ARptVoidEntries.totCred = Using(fmt$, GTCredit#)
  ARptVoidEntries.totDeb = Using(fmt$, GTDebit#)

  ARptVoidEntries.Caption = PrnTitle$
  ARptVoidEntries.Label1.Caption = PrnTitle$
  ARptVoidEntries.txtDate.Caption = Date$
  ARptVoidEntries.txtTown.Caption = GLUserName$
  ARptVoidEntries.GetName PRNfileName$
  ARptVoidEntries.startrpt

  'ViewPrint PRNfileName$, PrnTitle$
  MsgBox "Void Procedure Completed.", vbOKOnly, "Void Complete"
Exit Sub

  

PrintGLIF:
  ToPrint$ = ""
  ToPrint$ = QPTrim(GLifRec(1).AcctNum)
  ToPrint$ = ToPrint$ + "~" + Format(DateAdd("d", (GLifRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
  ToPrint$ = ToPrint$ + "~" + Left$(GLifRec(1).Desc, 15)
  ToPrint$ = ToPrint$ + "~" + QPTrim(GLifRec(1).Ref)
  ToPrint$ = ToPrint$ + "~" + Using(fmt$, GLifRec(1).DrAmt)
  ToPrint$ = ToPrint$ + "~" + Using(fmt$, GLifRec(1).CrAmt)
  Print #PRNFile, ToPrint$
  Return

End Sub
Private Sub VoidInvTrans2(VoidTransRecNum&)
  Dim Interfacefile As String, LdRecLen As Integer, DistRecLEn As Integer
  Dim GLIFRecLen As Integer, PRNfileName As String, PrnTitle As String
  Dim fmt As String, GLIFFile As Integer, PRNFile As Integer, ToPrint As String
  Dim NextDist As Long, APLedgerFile As Integer, NumTran As Long
  Dim APDistFile As Integer, NumDistRecs As Long, IFRec As Integer
  Dim GTDebit As Double, GTCredit As Double, BadAcct As Integer
  Dim Linecnt As Integer
  '-If this file exists this routine did not finish last time.
  Interfacefile$ = "APVDIF.DAT"
  If Exist(Interfacefile$) Then
    'KILL InterfaceFile$
    Close
    MsgBox "The file APVDIF.DAT already exits. Cannot continue.", vbOKOnly, "Void Canceled"
    Exit Sub
  End If
  Call MainLog("Start AP Void Inv - " + txtInv + fpcboVoidDate.Text)
  ReDim ApLedger(1) As APLedger81RecType
  ReDim DistRec(1) As APDistRecType
  ReDim GLifRec(1) As GLTransRecType
  LdRecLen = Len(ApLedger(1))
  DistRecLEn = Len(DistRec(1))
  GLIFRecLen = Len(GLifRec(1))
  Linecnt = 0
  PRNfileName$ = "VOIDINV.PRN"
  PrnTitle$ = "Voided Invoice Report"
  fmt$ = "###,###,###.##"
  
  OpenAPLedgerFile APLedgerFile, NumTran&, LdRecLen

  Get APLedgerFile, VoidTransRecNum&, ApLedger(1)
  Close APLedgerFile

  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLEn

  GLIFFile = FreeFile
  Open Interfacefile$ For Random As GLIFFile Len = GLIFRecLen

  PRNFile = FreeFile
  Open PRNfileName$ For Output As PRNFile
  GoSub PrintReportHeader

  NextDist& = ApLedger(1).FrstDist
  Do Until NextDist& = 0
    Get APDistFile, NextDist&, DistRec(1)
    IFRec = IFRec + 1

    '--Set Entry Description Defaults
    GLifRec(1).Src = "VI" + Format$(Now, "mmddyy")
    If fpcboVoidDate.ListIndex <> 0 Then
      GLifRec(1).TRDATE = DateDiff("d", "12/31/1979", Date$)
    Else
      GLifRec(1).TRDATE = ApLedger(1).TRDATE
    End If
    GLifRec(1).Desc = ApLedger(1).DOCNum
    GLifRec(1).Ref = "TO VOID"
    GLifRec(1).LDesc = ApLedger(1).Comment

    '--Make Debit side of entry
    GLifRec(1).AcctNum = Left$(DistRec(1).DistAcctNum, GLFundLen) + APAcct$
    GLifRec(1).CrAmt = 0
    GLifRec(1).DrAmt = DistRec(1).DistAmt
    Put GLIFFile, IFRec, GLifRec(1)
    GTDebit# = GTDebit# + GLifRec(1).DrAmt
    GoSub PrintGLIF

    '--Make Credit side
    IFRec = IFRec + 1
    GLifRec(1).AcctNum = DistRec(1).DistAcctNum
    GLifRec(1).CrAmt = DistRec(1).DistAmt
    GLifRec(1).DrAmt = 0
    Put GLIFFile, IFRec, GLifRec(1)
    GTCredit# = GTCredit# + GLifRec(1).CrAmt
    GoSub PrintGLIF

    NextDist& = DistRec(1).NextDist

  Loop
  ToPrint$ = Space$(80)
  Print #PRNFile,
  Mid$(ToPrint$, 1) = "Grand Total"
  Mid$(ToPrint$, 53) = Using(fmt$, GTDebit#)
  Mid$(ToPrint$, 67) = Using(fmt$, GTCredit#)
  Print #PRNFile, ToPrint$
  Close

  Post2GL Interfacefile$, BadAcct, frmInvVoid, False
  If BadAcct Then
    Call MainLog("APInvVoid Error - Accts Not Update, Canceled")
    MsgBox "Void Error.  Review GL Utility Log for details.", vbOKOnly, "Void aborted"
    ViewPrint PRNfileName$, PrnTitle$
    MsgBox "Accounts were not updated.", vbOKOnly, "Void Canceled"
    Exit Sub
  End If
 
  Post2GL Interfacefile$, BadAcct, frmInvVoid, True
  KillFile Interfacefile$
  
  OpenAPLedgerFile APLedgerFile, NumTran&, LdRecLen
  Get APLedgerFile, VoidTransRecNum&, ApLedger(1)
  ApLedger(1).TRCode = -1
  Put APLedgerFile, VoidTransRecNum&, ApLedger(1)
  Close APLedgerFile
  Call MainLog("APInvVoid Completed.")
  ViewPrint PRNfileName$, PrnTitle$
  MsgBox "Void Procedure Completed.", vbOKOnly, "Void Complete"
Exit Sub

PrintReportHeader:
  Print #PRNFile, PrnTitle$
  Print #PRNFile, "Run Date: " + Date$
  Print #PRNFile,
  Print #PRNFile, "Acct No          Date       Description        Ref           Debit        Credit"
  Print #PRNFile, "========================================================================================"
  Linecnt = Linecnt + 5
  Return

PrintGLIF:
  ToPrint$ = Space$(80)
  Mid$(ToPrint$, 1) = GLifRec(1).AcctNum
  Mid$(ToPrint, 16) = Format(DateAdd("d", (GLifRec(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
  Mid$(ToPrint$, 28) = Left$(GLifRec(1).Desc, 15)
  Mid$(ToPrint$, 45) = GLifRec(1).Ref
  Mid$(ToPrint$, 53) = Using(fmt$, GLifRec(1).DrAmt)
  Mid$(ToPrint$, 67) = Using(fmt$, GLifRec(1).CrAmt)
  Print #PRNFile, ToPrint$
  Linecnt = Linecnt + 2
  Return

End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
