VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmTaxExpRealInfo 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Real Property Information Export"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmTaxExpRealInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.OptionButton OptDelimiter3 
      BackColor       =   &H008F8265&
      Caption         =   "Tab"
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
      Height          =   372
      Left            =   6948
      TabIndex        =   26
      Top             =   2634
      Width           =   1692
   End
   Begin VB.OptionButton OptDelimiter2 
      BackColor       =   &H008F8265&
      Caption         =   "Pipe Symbol  |"
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
      Height          =   372
      Left            =   6948
      TabIndex        =   25
      Top             =   2286
      Width           =   1740
   End
   Begin VB.OptionButton OptDelimiter1 
      BackColor       =   &H008F8265&
      Caption         =   "Comma  ,"
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
      Height          =   372
      Left            =   6948
      TabIndex        =   24
      Top             =   1914
      Width           =   1692
   End
   Begin VB.CheckBox chkQuotes 
      BackColor       =   &H008F8265&
      Caption         =   "Double Quotes  """
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
      Height          =   372
      Left            =   6900
      TabIndex        =   23
      Top             =   3714
      Width           =   2052
   End
   Begin VB.CheckBox chkFileUnique 
      BackColor       =   &H008F8265&
      Caption         =   "Unique File Name"
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
      Height          =   324
      Left            =   3300
      TabIndex        =   22
      Top             =   2160
      Width           =   2412
   End
   Begin VB.CheckBox chkAcctNum 
      BackColor       =   &H008F8265&
      Caption         =   "Pin Number"
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
      Left            =   2340
      TabIndex        =   21
      Top             =   5034
      Width           =   1908
   End
   Begin VB.CheckBox chkAddress 
      BackColor       =   &H008F8265&
      Caption         =   "Address"
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
      Left            =   2340
      TabIndex        =   20
      Top             =   5754
      Width           =   1908
   End
   Begin VB.CheckBox chkLandRec 
      BackColor       =   &H008F8265&
      Caption         =   "Land Rec/GIS Key"
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
      Left            =   2340
      TabIndex        =   19
      Top             =   6114
      Width           =   2025
   End
   Begin VB.CheckBox chkRealVal 
      BackColor       =   &H008F8265&
      Caption         =   "Real Value"
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
      Left            =   2340
      TabIndex        =   18
      Top             =   6474
      Width           =   1908
   End
   Begin VB.CheckBox chkMap 
      BackColor       =   &H008F8265&
      Caption         =   "Map"
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
      Left            =   2340
      TabIndex        =   17
      Top             =   6834
      Width           =   1908
   End
   Begin VB.CheckBox chkBlock 
      BackColor       =   &H008F8265&
      Caption         =   "Block"
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
      Left            =   2340
      TabIndex        =   16
      Top             =   7194
      Width           =   1908
   End
   Begin VB.CheckBox chkLot 
      BackColor       =   &H008F8265&
      Caption         =   "Lot"
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
      Left            =   4740
      TabIndex        =   15
      Top             =   5034
      Width           =   1908
   End
   Begin VB.CheckBox chkLotAcre 
      BackColor       =   &H008F8265&
      Caption         =   "Lot/Acre"
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
      Left            =   4740
      TabIndex        =   14
      Top             =   5394
      Width           =   1908
   End
   Begin VB.CheckBox chkSize 
      BackColor       =   &H008F8265&
      Caption         =   "Size"
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
      Left            =   4740
      TabIndex        =   13
      Top             =   5754
      Width           =   1908
   End
   Begin VB.CheckBox chkSnrDisc 
      BackColor       =   &H008F8265&
      Caption         =   "Senior Discount"
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
      Left            =   4740
      TabIndex        =   12
      Top             =   6114
      Width           =   1908
   End
   Begin VB.CheckBox chkOthDisc 
      BackColor       =   &H008F8265&
      Caption         =   "Other Discount"
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
      Left            =   4740
      TabIndex        =   11
      Top             =   6474
      Width           =   1908
   End
   Begin VB.CheckBox chkDiscovery 
      BackColor       =   &H008F8265&
      Caption         =   "Discovery Y/N?"
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
      Left            =   4740
      TabIndex        =   10
      Top             =   6834
      Width           =   1788
   End
   Begin VB.CheckBox chkLateList 
      BackColor       =   &H008F8265&
      Caption         =   "Late List Y/N?"
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
      Left            =   4740
      TabIndex        =   9
      Top             =   7194
      Width           =   1908
   End
   Begin VB.CheckBox chkLien 
      BackColor       =   &H008F8265&
      Caption         =   "Lien Y/N?"
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
      Left            =   7260
      TabIndex        =   8
      Top             =   5034
      Width           =   1908
   End
   Begin VB.CheckBox chkLienDesc 
      BackColor       =   &H008F8265&
      Caption         =   "Lien Description"
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
      Left            =   7260
      TabIndex        =   7
      Top             =   5394
      Width           =   1908
   End
   Begin VB.CheckBox chkMortCode 
      BackColor       =   &H008F8265&
      Caption         =   "Mortgage Code"
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
      Left            =   7260
      TabIndex        =   6
      Top             =   5754
      Width           =   1908
   End
   Begin VB.CheckBox chkOpt1 
      BackColor       =   &H008F8265&
      Caption         =   "Opt Rev 1"
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
      Left            =   7260
      TabIndex        =   5
      Top             =   6114
      Width           =   2148
   End
   Begin VB.CheckBox chkOpt2 
      BackColor       =   &H008F8265&
      Caption         =   "Opt Rev 2"
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
      Left            =   7260
      TabIndex        =   4
      Top             =   6474
      Width           =   1908
   End
   Begin VB.CheckBox chkOpt3 
      BackColor       =   &H008F8265&
      Caption         =   "Opt Rev 3"
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
      Left            =   7260
      TabIndex        =   3
      Top             =   6834
      Width           =   1908
   End
   Begin VB.CheckBox chkBalance 
      BackColor       =   &H008F8265&
      Caption         =   "Balance"
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
      Left            =   7260
      TabIndex        =   2
      Top             =   7194
      Width           =   1908
   End
   Begin VB.CheckBox chkCurrOwner 
      BackColor       =   &H008F8265&
      Caption         =   "Current Owner"
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
      Left            =   2340
      TabIndex        =   0
      Top             =   5394
      Width           =   1908
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdUntagAll 
      Height          =   315
      Left            =   7680
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1470
      _Version        =   131072
      _ExtentX        =   2593
      _ExtentY        =   556
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
      ButtonDesigner  =   "frmTaxExpRealInfo.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   480
      Left            =   3840
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
      Top             =   7920
      Width           =   1740
      _Version        =   131072
      _ExtentX        =   3069
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
      ButtonDesigner  =   "frmTaxExpRealInfo.frx":0AA7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   480
      Left            =   6075
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   $"frmTaxExpRealInfo.frx":0C85
      Top             =   7920
      Width           =   1740
      _Version        =   131072
      _ExtentX        =   3069
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
      ButtonDesigner  =   "frmTaxExpRealInfo.frx":0D30
   End
   Begin EditLib.fpText fptxtFileName 
      Height          =   396
      Left            =   2820
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "The program will generate a unique file name and display the name in this text box."
      Top             =   2520
      Width           =   2892
      _Version        =   196608
      _ExtentX        =   5101
      _ExtentY        =   698
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
      AutoCase        =   1
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
      MaxLength       =   50
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
   Begin fpBtnAtlLibCtl.fpBtn cmdTagAll 
      Height          =   315
      Left            =   6120
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1470
      _Version        =   131072
      _ExtentX        =   2593
      _ExtentY        =   556
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
      ButtonDesigner  =   "frmTaxExpRealInfo.frx":0F0F
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Export Real Property Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3348
      TabIndex        =   36
      Top             =   642
      Width           =   5004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   708
      Left            =   2940
      Top             =   474
      Width           =   5772
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Text Qualifier:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   6660
      TabIndex        =   35
      Top             =   3354
      Width           =   2004
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2772
      Left            =   2340
      Top             =   1434
      Width           =   6912
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   6300
      X2              =   6300
      Y1              =   1434
      Y2              =   4194
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Select for Unique File Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   2820
      TabIndex        =   34
      Top             =   1800
      Width           =   3132
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   """TaxRealEx.ASC"" will be used if Option above is not selected."
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
      Height          =   552
      Left            =   2820
      TabIndex        =   33
      Top             =   3480
      Width           =   3012
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   6300
      X2              =   2340
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Delimiter:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   6660
      TabIndex        =   32
      Top             =   1554
      Width           =   2268
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   9242
      X2              =   6312
      Y1              =   3114
      Y2              =   3114
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Fields to Include in Export:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   2400
      TabIndex        =   31
      Top             =   4458
      Width           =   3732
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2652
      Left            =   2100
      Top             =   4920
      Width           =   7392
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   4440
      X2              =   4440
      Y1              =   4920
      Y2              =   7560
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   6900
      X2              =   6900
      Y1              =   4914
      Y2              =   7554
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   828
      Left            =   2940
      Top             =   354
      Width           =   5772
   End
End
Attribute VB_Name = "frmTaxExpRealInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim OptDesc1$
  Dim OptDesc2$
  Dim OptDesc3$

Private Sub chkFileUnique_Click()
  Dim ThisFile$
  Dim Ext$
  Dim Cnt As Integer
  Dim chkCust$
  
  If chkFileUnique.Value = 1 Then
    Ext$ = ".ASC"
    ThisFile$ = "RPX"
    For Cnt = 1 To 5
      GetRPTName ThisFile$
      chkCust$ = ThisFile$ + Ext$
      If Exist(chkCust$) = False Then
        ThisFile$ = chkCust$
        Exit For
      End If
    Next Cnt
    fptxtFileName.Text = ThisFile$
  Else
    fptxtFileName.Text = ""
  End If
  
End Sub

Private Sub cmdProcess_Click()
  Dim q$
  Dim qc$
  Dim qcq$
  Dim Ext$, x As Long
  Dim ThisFile$, chkCust$
  Dim TaxRpt As Integer
  Dim Cnt As Integer
  Dim ThisRec As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim XCnt As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PrintType$
  Dim ThisBal As Double
  Dim CustName$
  Dim RealBal As Double
  
  If QPTrim$(fptxtFileName.Text) <> "" Then
    ThisFile$ = fptxtFileName.Text
  Else
    ThisFile$ = "TaxRealEx.ASC"
    KillFile ThisFile$
  End If
  
  TaxRpt = FreeFile
  Open ThisFile$ For Output As TaxRpt
  
  q$ = ""
  qc$ = ""
  qcq$ = ""
  If OptDelimiter1.Value = True And chkQuotes.Value <> 1 Then
    qcq$ = ","
  ElseIf OptDelimiter2.Value = True And chkQuotes.Value <> 1 Then
    qcq$ = "|"
  ElseIf OptDelimiter3.Value = True And chkQuotes.Value <> 1 Then
    qcq$ = Chr$(9)  'this is tab
  ElseIf chkQuotes.Value = 1 Then
    q$ = Chr$(34) 'this is one quote (")
    If OptDelimiter1.Value = True Then
      qc$ = q$ + ","
      qcq$ = q$ + "," + q$
    ElseIf OptDelimiter2.Value = True Then
      qc$ = q$ + "|"
      qcq$ = q$ + "|" + q$
    ElseIf OptDelimiter3.Value = True Then
      qc$ = q$ + Chr$(9)
      qcq$ = q$ + Chr$(9) + q$ 'this is tab
    End If
  End If
  GoSub DoHeaders
  
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  frmTaxShowPctComp.Label1 = "Printing Real Property Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  For x = 1 To NumOfRealRecs
    Get RHandle, x, RealRec
    If RealRec.Deleted <> 0 Then GoTo NotThisOne
    GoSub ExportThisOne
NotThisOne:
    frmTaxShowPctComp.ShowPctComp x, NumOfRealRecs
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
  
  Close
  
  If XCnt > 0 Then
    Call Savemsg(800, "The real property export file has been saved as " + ThisFile$ + " in the Citipak folder.")
  Else
    Call TaxMsg(900, "There are no real property records that fit the parameters entered.")
  End If
  
  Exit Sub
  
DoHeaders:
  If chkAcctNum.Value = 1 Then Print #TaxRpt, q$; "Pin #";
  If chkCurrOwner.Value = 1 Then Print #TaxRpt, qcq$; "Current Owner";
  If chkAddress.Value = 1 Then Print #TaxRpt, qcq$; "Address";
  If chkLandRec.Value = 1 Then Print #TaxRpt, qcq$; "Land Rec/GIS Key";
  If chkRealVal.Value = 1 Then Print #TaxRpt, qcq$; "Real Value";
  If chkMap.Value = 1 Then Print #TaxRpt, qcq$; "Map";
  If chkBlock.Value = 1 Then Print #TaxRpt, qcq$; "Block";
  If chkLot.Value = 1 Then Print #TaxRpt, qcq$; "Lot";
  If chkLotAcre.Value = 1 Then Print #TaxRpt, qcq$; "Lot/Acre";
  If chkSize.Value = 1 Then Print #TaxRpt, qcq$; "Size";
  If chkSnrDisc.Value = 1 Then Print #TaxRpt, qcq$; "Senior Discount";
  If chkOthDisc.Value = 1 Then Print #TaxRpt, qcq$; "Other Discount";
  If chkDiscovery.Value = 1 Then Print #TaxRpt, qcq$; "Discovery Y/N?";
  If chkLateList.Value = 1 Then Print #TaxRpt, qcq$; "Late List Y/N?";
  If chkLien.Value = 1 Then Print #TaxRpt, qcq$; "Lien Y/N?";
  If chkLienDesc.Value = 1 Then Print #TaxRpt, qcq$; "Lien Description";
  If chkMortCode.Value = 1 Then Print #TaxRpt, qcq$; "Mortgage Code";
  If chkOpt1.Value = 1 Then Print #TaxRpt, qcq$; OptDesc1;
  If chkOpt2.Value = 1 Then Print #TaxRpt, qcq$; OptDesc2;
  If chkOpt3.Value = 1 Then Print #TaxRpt, qcq$; OptDesc3;
  If chkBalance.Value = 1 Then Print #TaxRpt, qcq$; "Balance";
  Print #TaxRpt, q$
  
  Return
  
ExportThisOne:
  XCnt = XCnt + 1
  If chkAcctNum.Value = 1 Then Print #TaxRpt, q$; QPTrim$(RealRec.RealPin);
  If chkCurrOwner = 1 Then
    If RealRec.CustPin > 0 Then
      Get TCHandle, RealRec.CustPin, TaxCust
      CustName = QPTrim$(TaxCust.CustName)
      RealBal = 0
      If QPTrim$(RealRec.RealPin) <> "" And chkBalance.Value = 1 Then
        RealBal = GetRealBalance(RealRec.RealPin)
      End If
    Else
      CustName = "Could Not Be Determined"
    End If
    Print #TaxRpt, qcq$; CustName;
  End If
  If chkAddress.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(RealRec.PropAddr);
  If chkLandRec.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(RealRec.GISPOS);
  If chkRealVal.Value = 1 Then Print #TaxRpt, qcq$; Using$("$########0.00", RealRec.PROPVALU);
  If chkMap.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(RealRec.Map);
  If chkBlock.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(RealRec.BLOCK);
  If chkLot.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(RealRec.LOTNUMB);
  If chkLotAcre.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(RealRec.LOTACRE);
  If chkSize.Value = 1 Then Print #TaxRpt, qcq$; Using$("###0.00", RealRec.PropSize);
  If chkSnrDisc.Value = 1 Then Print #TaxRpt, qcq$; Using$("$########0.00", RealRec.EXMPSENI);
  If chkOthDisc.Value = 1 Then Print #TaxRpt, qcq$; Using$("$########0.00", RealRec.EXMPOTHR);
  If chkDiscovery.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(RealRec.PROPDISC);
  If chkLateList.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(RealRec.LateList);
  If chkLien.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(RealRec.LienYN);
  If chkLienDesc.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(RealRec.LienDesc);
  If chkMortCode.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(RealRec.MORTCODE);
  If chkOpt1.Enabled = True And chkOpt1.Value = 1 Then
    Print #TaxRpt, qcq$; Using$("####0", RealRec.OptRev1Chrg);
  End If
  If chkOpt2.Enabled = True And chkOpt2.Value = 1 Then
    Print #TaxRpt, qcq$; Using$("####0", RealRec.OptRev2Chrg);
  End If
  If chkOpt3.Enabled = True And chkOpt3.Value = 1 Then
    Print #TaxRpt, qcq$; Using$("####0", RealRec.OptRev3Chrg);
  End If
  If chkBalance.Value = 1 Then Print #TaxRpt, qcq$; QPTrim$(Using$("$##,###.##", RealBal));
  Print #TaxRpt, q$
    
  Return
  
End Sub

Private Sub cmdTagAll_Click()
  chkAcctNum.Value = 1
  chkCurrOwner.Value = 1
  chkAddress.Value = 1
  chkLandRec.Value = 1
  chkRealVal.Value = 1
  chkMap.Value = 1
  chkBlock.Value = 1
  chkLot.Value = 1
  chkLotAcre.Value = 1
  chkSize.Value = 1
  chkSnrDisc.Value = 1
  chkOthDisc.Value = 1
  chkDiscovery.Value = 1
  chkLateList.Value = 1
  chkLien.Value = 1
  chkLienDesc.Value = 1
  chkMortCode.Value = 1
  If chkOpt1.Enabled = True Then
    chkOpt1.Value = 1
  End If
  If chkOpt2.Enabled = True Then
    chkOpt2.Value = 1
  End If
  If chkOpt3.Enabled = True Then
    chkOpt3.Value = 1
  End If
  chkBalance.Value = 1
End Sub

Private Sub cmdUntagAll_Click()
  chkAcctNum.Value = 0
  chkCurrOwner.Value = 0
  chkAddress.Value = 0
  chkLandRec.Value = 0
  chkRealVal.Value = 0
  chkMap.Value = 0
  chkBlock.Value = 0
  chkLot.Value = 0
  chkLotAcre.Value = 0
  chkSize.Value = 0
  chkSnrDisc.Value = 0
  chkOthDisc.Value = 0
  chkDiscovery.Value = 0
  chkLateList.Value = 0
  chkLien.Value = 0
  chkLienDesc.Value = 0
  chkMortCode.Value = 0
  chkOpt1.Value = 0
  chkOpt2.Value = 0
  chkOpt3.Value = 0
  chkBalance.Value = 0
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpRealProperty
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxExpRealInfo.")
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub cmdExit_Click()
  frmTaxReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  OptDesc1 = QPTrim$(TaxMasterRec.OptRev1)
  If OptDesc1 = "" Then
    chkOpt1.Caption = "NOT IN USE"
    chkOpt1.Enabled = False
  Else
    chkOpt1.Caption = OptDesc1
  End If
  OptDesc2 = QPTrim$(TaxMasterRec.OptRev2)
  If OptDesc2 = "" Then
    chkOpt2.Caption = "NOT IN USE"
    chkOpt2.Enabled = False
  Else
    chkOpt2.Caption = OptDesc2
  End If
  OptDesc3 = QPTrim$(TaxMasterRec.OptRev3)
  If OptDesc3 = "" Then
    chkOpt3.Caption = "NOT IN USE"
    chkOpt3.Enabled = False
  Else
    chkOpt3.Caption = OptDesc3
  End If
  
  OptDelimiter1.Value = True
End Sub

