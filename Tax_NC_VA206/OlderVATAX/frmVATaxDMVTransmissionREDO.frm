VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmVATaxDMVTransmissionREDO 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DMV Transmission Redo"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxDMVTransmissionREDO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EditLib.fpLongInteger fpLIBatchNum 
      Height          =   372
      Left            =   6480
      TabIndex        =   4
      Top             =   4680
      Width           =   732
      _Version        =   196608
      _ExtentX        =   1291
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
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
      MarginTop       =   3
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
      MinValue        =   "0"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
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
   Begin EditLib.fpDoubleSingle fpDSPersRate 
      Height          =   372
      Left            =   6480
      TabIndex        =   3
      Top             =   4080
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
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
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "0.0000"
      DecimalPlaces   =   4
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
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   3480
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6600
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
      ButtonDesigner  =   "frmVATaxDMVTransmissionREDO.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   492
      Left            =   6120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6600
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
      ButtonDesigner  =   "frmVATaxDMVTransmissionREDO.frx":0AA6
   End
   Begin EditLib.fpDateTime fptxtPayThruDate 
      Height          =   372
      Left            =   6480
      TabIndex        =   0
      Top             =   2280
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
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
      InvalidColor    =   12648447
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
      Text            =   "02/24/2005"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
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
   Begin EditLib.fpDateTime fptxtSubDate 
      Height          =   372
      Left            =   6480
      TabIndex        =   1
      Top             =   2880
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
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
      InvalidColor    =   12648447
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
      Text            =   "02/24/2005"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
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
   Begin EditLib.fpDateTime fptxtTaxYear 
      Height          =   348
      Left            =   6480
      TabIndex        =   2
      Top             =   3480
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
      _ExtentY        =   614
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
   Begin EditLib.fpText fptxtJuris 
      Height          =   396
      Left            =   6480
      TabIndex        =   5
      Top             =   5280
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
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
      AlignTextH      =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   10
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Submission Date:"
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
      Left            =   3960
      TabIndex        =   14
      Top             =   3000
      Width           =   2292
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payments Through Date:"
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
      Left            =   3360
      TabIndex        =   13
      Top             =   2400
      Width           =   2892
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4092
      Left            =   2400
      Top             =   1920
      Width           =   7092
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   660
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DMV File Transmission REDO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2940
      TabIndex        =   12
      Top             =   792
      Width           =   6012
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Year To Process:"
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
      Height          =   252
      Left            =   3960
      TabIndex        =   11
      Top             =   3552
      Width           =   2292
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Property Rate:"
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
      Left            =   3240
      TabIndex        =   10
      Top             =   4200
      Width           =   3012
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Number:"
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
      Left            =   4440
      TabIndex        =   9
      Top             =   4800
      Width           =   1812
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jurisdiction ID:"
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
      Left            =   4440
      TabIndex        =   8
      Top             =   5400
      Width           =   1812
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   600
      Width           =   8652
   End
End
Attribute VB_Name = "frmVATaxDMVTransmissionREDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmVATaxDMVMenu.Show
  DoEvents
  Unload Me
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
  Me.HelpContextID = hlpReprocessDMV
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxDMVTransmissionREDO.")
      Call Terminate
      End
    End If
  End If

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

Private Sub LoadMe()
  Dim DMVLiveIF As DMVInformationType
  Dim DMVInfoHandle As Integer
  Dim NumOfDMVLiveRecs As Long
  
  fptxtPayThruDate = Date
  fptxtSubDate = Date
  If Exist(DMVInfoFile) Then
    OpenDMVInfoFile DMVInfoHandle, NumOfDMVLiveRecs
    Get DMVInfoHandle, 1, DMVLiveIF
    Close
    fpDSPersRate = DMVLiveIF.PerRate
    fpLIBatchNum = DMVLiveIF.Batch + 1
    fptxtJuris = DMVLiveIF.JCode
  Else
    fpDSPersRate = 0
    fpLIBatchNum = 0
    fptxtJuris = ""
  End If
  
End Sub

Private Sub cmdProcess_Click()
  Dim DMVHeader As DMVHeader
  Dim DMVRecord As DMVRecord
  Dim TaxSetUp As TaxMasterType
  Dim TaxCustRec As TaxCustType
  Dim TransRec As TaxTransactionType
  Dim PersRec As PersonalRecType
  Dim DMVLiveIF As DMVInformationType
  Dim DMVInfoHandle As Integer
  Dim NumOfDMVLiveRecs As Long
  Dim SSN1$(175), LastName1$(175), FirstName1$(175), Addr1$(175), Addr2$(175), City$(175), State$(175), Zip$(175), Vin$(175), VehValue#(175), PPTaxPd$(175), PPTaxReimb$(175), EMth$(175), SMth$(175)
  Dim Batch$, Jury$
  Dim ProcessDate As Integer
  Dim TaxYear$
  Dim ThisYear$
  Dim TAXRATE As Double
  Dim ThisDateI As Integer
  Dim ThisDateS As String
  Dim CalcDate As Integer
  Dim JulianDate As Integer
  Dim JulianDateS$
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim RptHandle10 As Integer
  Dim RptHandle11 As Integer
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim FF$, x As Long, y As Integer
  Dim PersHandle As Integer
  Dim NumOfPersRecs As Long
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxCust As TaxCustType
  Dim TransFile As TaxTransactionType
  Dim TransHandle As Integer
  Dim NumOfTransRec As Long
  Dim PayJourName$
  Dim PayJourNameOld$
  Dim Header$, DF$
  Dim ThisName$
  Dim V As Integer
  Dim kk1 As Integer
  Dim kk2 As Integer
  Dim FirstName$
  Dim LastName$
  Dim TaxPaidD@
  Dim TaxPaid$, veh As Integer
  Dim ReimbursementD@
  Dim Reimbursement$
  Dim TotalReimb#, Early As Integer
  Dim PersValue#
  Dim TransRecord&, Records As Long
  Dim ProcessThisCustomer$
  Dim Balance#, UpdateRecord&
  Dim LastPaidDate As Integer
  Dim PropertyRecord!
  Dim TLen As Integer
  Dim TotalCrCnt$
  Dim TotalCrAmt$
  Dim VehiclesS$, VehValueD#, VehValueS$
  Dim SMth1$, TotalAmt$
  Dim EMth1$, PERC!
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim OldYN As Boolean
  Dim OldNum As Integer
  Dim ThisLen As Integer
  Dim MaxVehVal As Double
  Dim MinVehVal As Double
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close
  If TaxMasterRec.PPTRADisc = 0 Then
    Call TaxMsg(750, "The current setting for the PPTRA Discount is zero. This setting (found on the System Setup screen) needs to be changed to a value above zero.")
    Exit Sub
  End If
  
  PERC! = TaxMasterRec.PPTRADisc * 0.01
  MaxVehVal = TaxMasterRec.MaxVehTaxVal
  MinVehVal = TaxMasterRec.MinVehTaxVal
  MaxLines = 58
  FF$ = Chr(12)
  Batch$ = CStr(fpLIBatchNum)
  If Len(Batch$) < 3 Then
   Batch$ = String$(3 - Len(Batch$), "0") + Batch$
  End If
  
  ProcessDate = Date2Num%(fptxtPayThruDate.Text)
  Jury$ = QPTrim$(fptxtJuris.Text)
  TaxYear$ = fptxtTaxYear.Text
  
  ThisLen = Len(Date)
  ThisYear = Mid(Date, ThisLen - 3, 4)
  
  If Abs(CInt(ThisYear) - CInt(TaxYear)) > 5 Then
    If TaxMsgWOpts(700, "The tax year entered is more than five years from the current year. If this is accurate then press F10 to continue. Otherwise, press ESC to escape and review.", "F10 Continue", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtTaxYear.SetFocus
      Exit Sub
    End If
  End If
  TAXRATE = fpDSPersRate
  
  ThisDateI = Date2Num(fptxtSubDate.Text)
  ThisDateS = fptxtSubDate.Text
  CalcDate = Date2Num%("12-31-2003")
  JulianDate = ThisDateI - CalcDate
  If JulianDate > 365 Then JulianDate = JulianDate - 365
  JulianDateS$ = LTrim$(Str$(JulianDate))
  If Len(JulianDateS$) < 3 Then JulianDateS$ = String$(3 - Len(JulianDateS$), "0") + JulianDateS$
  RptHandle = FreeFile
  ReportFile$ = "DMVFILE.RPT"
  Open ReportFile$ For Output As #RptHandle
  
  GoSub ReportHeading
  
  OldYN = False
  OldNum = 1
  PayJourName$ = "T" + Jury$ + JulianDateS$ + "." + Batch$
  If Exist("T" + Jury$ + JulianDateS$ + "." + Batch$) Then
    OldYN = True
    If Exist("T" + Jury$ + JulianDateS$ + "old." + Batch$) Then
      If TaxMsgWOpts(700, "The file " + "T" + Jury$ + JulianDateS$ + "old." + Batch$ + " has already been saved as a backup for the file being redone. If you wish to overwrite this file then press F10. To keep this backup file untouched press ESC.", "F10 Overwrite", "ESC Keep As Is") <> "abort" Then
        Unload frmVATaxMsgWOpts
        KillFile "T" + Jury$ + JulianDateS$ + "old." + Batch$
        Name "T" + Jury$ + JulianDateS$ + "." + Batch$ As "T" + Jury$ + JulianDateS$ + "old." + Batch$
      End If
    Else
      Name "T" + Jury$ + JulianDateS$ + "." + Batch$ As "T" + Jury$ + JulianDateS$ + "old." + Batch$
    End If
  End If
  
  Header$ = "Creating DMV Data File"
  
'  KillFile PayJourName$
  RptHandle11 = FreeFile
  Open PayJourName$ For Output As RptHandle11
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TransHandle, NumOfTransRec
  OpenTaxPersFile PersHandle, NumOfPersRecs
  
  frmVATaxShowPctComp.Label1 = "Creating DMV Transmission File"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    'If x = 1680 Then Stop
    If TaxCust.Deleted Then GoTo NextPlease
    ThisName = QPTrim$(TaxCust.CustName)
    LastPaidDate = 0
    PersValue# = 0
    ProcessThisCustomer$ = "N"
    TransRecord& = TaxCust.LastTrans
    Do While TransRecord& > 0
      Get TransHandle, TransRecord&, TransRec
      If TransRec.BillType <> "P" Then GoTo NextLoop
      If TransRec.TranType <> 1 Then GoTo NextLoop
      If TransRec.TaxYear <> CInt(TaxYear$) Then GoTo NextLoop      'Add Processing to Pull by Proper Tax Year
      If TransRec.DMVSubmitted = "Y" And TransRec.DMVBatch = Val(Batch$) Then
        Balance# = 0
        Balance# = OldRound#(TransRec.Revenue.Principle1 + TransRec.Revenue.Principle2 + TransRec.Revenue.Principle3 + TransRec.Revenue.Principle4 + TransRec.Revenue.Principle5)
        Balance# = OldRound#(Balance# - (TransRec.DiscAmt + TransRec.PPTRADisc + _
                             TransRec.Revenue.Principle1Pd + TransRec.Revenue.Principle2Pd + TransRec.Revenue.Principle3Pd + TransRec.Revenue.Principle4Pd + TransRec.Revenue.Principle5Pd))
        If Balance# <= 0 Then
          ProcessThisCustomer$ = "Y"
          UpdateRecord& = TransRecord&
          GoTo Processme
        End If
      End If
NextLoop:
      If TransRec.TranType = 2 Then
        LastPaidDate = TransRec.TransDate
      End If
      TransRecord& = TransRec.LastTrans
    Loop
Processme:
    If ProcessThisCustomer$ = "Y" And LastPaidDate <= ProcessDate Then
      If TaxCust.FirstPersRec > 0 Then
        PropertyRecord! = TaxCust.FirstPersRec
        Do While PropertyRecord! <> 0
          Get PersHandle, PropertyRecord!, PersRec
'             If Left$(PersRec.Desc5, 1) = "Y" And PersRec(1).PersVal > 0.01 And Mid$(PersRec(1).Desc5, 3, 2) <> "BP" And PersRec(1).DMVSubmitted <> "Y" And (Val(Right$(PersRec(1).Desc5, 5)) = Val(TaxYear$)) Then
          If PersRec.PPTRAYN = "Y" And PersRec.PersVal > 0.01 And Mid$(PersRec.Desc5, 3, 2) <> "BP" And PersRec.DMVSubmitted = "Y" And PersRec.TaxBillYear = CInt(TaxYear$) Then
'            PersRec.DMVSubmitted = "Y"
'            Put PersHandle, PropertyRecord, PersRec

        'Update Batch Here When you have a good vehicle
'            Get TransHandle, UpdateRecord&, TransRec
'            TransRec.DMVSubmitted = "Y"
'            TransRec.DMVBatch = Val(Batch$)
'            Put TransHandle, UpdateRecord&, TransRec
            V = V + 1
            SSN1$(V) = TaxCust.CSSN
            kk1 = InStr(TaxCust.CustName, " ")
            Dim ThisCh$
            LastName$ = QPTrim$(TaxCust.CustName)
            For y = 1 To Len(LastName$)
              ThisCh = Mid(LastName$, y, 1)
              If ThisCh = " " Then
                If y > 1 Then
                  If Mid(LastName$, y - 1, 1) <> "," Then
                    kk2 = y
                  End If
                Else
                  kk2 = y
                End If
              End If
            Next y
            If kk2 > 0 Then
              LastName$ = Right(LastName$, Len(LastName$) - (kk2))
            ElseIf kk1 > 0 Then
              LastName$ = Right(LastName$, Len(LastName$) - (kk1))
            End If
            
            FirstName$ = QPTrim$(TaxCust.CustName)
            If kk1 > 0 Then
              FirstName$ = Left(FirstName$, kk1)
            Else
              FirstName$ = ""
            End If
            LastName1$(V) = QPTrim$(LastName$)
            FirstName1$(V) = FirstName$
              'Normal
            Addr1$(V) = QPTrim$(TaxCust.Addr1)
            Addr2$(V) = QPTrim$(TaxCust.Addr2)
            City$(V) = QPTrim$(TaxCust.City)
            State$(V) = QPTrim$(TaxCust.State$)
            Zip$(V) = QPTrim$(TaxCust.Zip)
            Vin$(V) = QPTrim$(PersRec.Vin)
            VehValue#(V) = PersRec.PersVal
            SMth1$ = Mid$(PersRec.Desc5, 5, 2)
            EMth1$ = Mid$(PersRec.Desc5, 12, 2)

            If Val(SMth1$) < 1 Or Val(SMth1$) > 12 Then SMth1$ = "01"
            If Val(EMth1$) < 1 Or Val(EMth1$) > 12 Then EMth1$ = "12"
            SMth$(V) = SMth1$
            EMth$(V) = EMth1$
             
'            If VehValue#(V) > 20000 Then VehValue#(V) = 20000 'Maximum of 20,000
            If VehValue#(V) > MaxVehVal Then VehValue#(V) = MaxVehVal 'Maximum of 20,000

            ' Calculate Tax Paid
            TaxPaidD@ = (VehValue#(V) / 100) * TAXRATE
            TaxPaidD@ = Int((TaxPaidD@ * 100) + 0.5) / 100
            TaxPaid$ = LTrim$(Str$(TaxPaidD@ * 100))
            If Len(TaxPaid$) = 1 Then
              TaxPaid$ = TaxPaid$ + ".00"
            Else
              TaxPaid$ = Left$(TaxPaid$, Len(TaxPaid$) - 2) + "." + Right$(TaxPaid$, 2)
            End If
            PPTaxPd$(V) = TaxPaid$

            ' Calculate Reimbursement
'            If VehValue#(V) <= 1000 Then
            If VehValue#(V) <= MinVehVal Then
              ReimbursementD@ = Val(TaxPaid$)
              Reimbursement$ = TaxPaid$
            Else
'              If VehValue#(V) <= 20000 Then
              If VehValue#(V) <= MaxVehVal Then
                ReimbursementD@ = TaxPaidD@ * PERC!
                ReimbursementD@ = Int((ReimbursementD@ * 100) + 0.5) / 100
                Reimbursement$ = LTrim$(Str$(ReimbursementD@))
              End If
            End If
            Reimbursement$ = LTrim$(Str$(ReimbursementD@ * 100))
            Reimbursement$ = Left$(Reimbursement$, Len(Reimbursement$) - 2) + "." + Right$(Reimbursement$, 2)
            PPTaxReimb$(V) = Reimbursement$
            TotalReimb# = TotalReimb# + ReimbursementD@
            TotalReimb# = Int((TotalReimb# * 100) + 0.5) / 100
            veh = V
          End If
          PropertyRecord! = PersRec.NextRec
        Loop
      End If
    End If
  
    If veh > 170 Then
      Early = 1
      GoTo EndDMVProcess
    End If

NextPlease:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
  Next x
  
EndDMVProcess:
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  'Write Header
   GoSub DMVHeaderWrite
  'Write Records
   GoSub DMVLines
  'Show FileName

  Print #RptHandle, String$(79, "-")
  Print #RptHandle, "File Name to Send is "; PayJourName$
  Print #RptHandle, "Total Vehicles: "; Using("######0", Val(VehiclesS$))
  Print #RptHandle, "Total of Reimbursement: "; Using("###,##0.00", Val(TotalAmt$))
  Print #RptHandle, FF$
  Close

  ViewPrint ReportFile$, "DMV Transmission File", True
  Kill ReportFile$

DMVExitJournal:
  
'  DMVLiveIF.PerRate = fpDSPersRate
'  DMVLiveIF.Batch = fpLIBatchNum
'  DMVLiveIF.JCode = fptxtJuris
'  OpenDMVInfoFile DMVInfoHandle, NumOfDMVLiveRecs
'  Put DMVInfoHandle, 1, DMVLiveIF
  Close
  
  If OldYN = False Then
    Call Savemsg(800, "The DMV transmission file can be found in the Citipak directory as " + PayJourName$ + ".")
  ElseIf OldYN = True Then
    Call Savemsg(700, "The DMV transmission file can be found in the Citipak directory as " + PayJourName$ + ". The replaced file has been saved as " + "T" + Jury$ + JulianDateS$ + "old." + Batch$ + ".")
  End If
  
  Exit Sub
  
ReportHeading:
  Print #RptHandle, "DMV Processing : Data File Contents"
  Print #RptHandle, "Submission Date: "; fptxtSubDate.Text; Tab(60); "Batch #"; Batch$
  Print #RptHandle, ""
  Print #RptHandle,
  Print #RptHandle, "Name"; Tab(40); "VIN #"; Tab(60); "Tax Amount"; Tab(71); "PPTRA Amt"
  Print #RptHandle, String$(79, "-")
  LineCnt = 5
  Return
  
DMVLines:
  For Records = 1 To veh
    Print #RptHandle11, "D@";
    Print #RptHandle11, RTrim$(LTrim$(Str$(Records))) + "@";
    Print #RptHandle11, RTrim$(SSN1$(Records)) + "@";
    Print #RptHandle11, RTrim$(LastName1$(Records)) + "@";
    Print #RptHandle11, RTrim$(FirstName1$(Records)) + "@";
    Print #RptHandle11, "@";
    Print #RptHandle11, "@";
    Print #RptHandle11, "@";
    Print #RptHandle11, "@";
    Print #RptHandle11, "@";
    Print #RptHandle11, RTrim$(Addr1$(Records)) + "@";
    Print #RptHandle11, RTrim$(Addr2$(Records)) + "@";
    Print #RptHandle11, RTrim$(City$(Records)) + "@";
    Print #RptHandle11, RTrim$(State$(Records)) + "@";
    Print #RptHandle11, RTrim$(Zip$(Records)) + "@";
    Print #RptHandle11, RTrim$(Vin$(Records)) + "@";
    VehValueS$ = LTrim$(Str$(VehValue#(Records)))
    Print #RptHandle11, RTrim$(VehValueS$) + "@";
    Print #RptHandle11, RTrim$(PPTaxPd$(Records)) + "@";
    Print #RptHandle11, RTrim$(PPTaxReimb$(Records)) + "@";
    Print #RptHandle11, LTrim$(TaxYear$) + SMth$(Records) + "@";
    Print #RptHandle11, LTrim$(TaxYear$) + EMth$(Records) + "@";
    Print #RptHandle11, Jury$ + "@";
    Print #RptHandle11, RTrim$(ReplaceString(ThisDateS, "/", "")) + "@"        'ADDED NEW CODE FOR 2001
    ThisName = RTrim$(FirstName1$(Records)) + " " + RTrim$(LastName1$(Records))
    Print #RptHandle, Left$(ThisName, 38);
'   NME$ = RTrim$(FirstName1$(Records)) + " " + RTrim$(LastName1$(Records))
'   Print #RptHandle, Left$(NME$, 38);
    Print #RptHandle, Tab(40); RTrim$(Vin$(Records));
    Print #RptHandle, Tab(61); Using("##,##0.00", Val(PPTaxPd$(Records)));
    Print #RptHandle, Tab(73); Using("###0.00", Val(PPTaxReimb$(Records)))
    LineCnt = LineCnt + 1
    If LineCnt >= 56 Then
     Print #RptHandle, Chr$(12);
     GoSub ReportHeading
    End If
  Next Records
  Return

DMVHeaderWrite:
  TotalAmt$ = LTrim$(Str$(TotalReimb# * 100))
  TLen = Len(TotalAmt$)

  TotalCrCnt$ = ""
  TotalCrAmt$ = ""

  If Val(TotalAmt$) = 0 Then
    TotalAmt$ = "0.00"
   Else
    TotalAmt$ = Left$(TotalAmt$, TLen - 2) + "." + Right$(TotalAmt$, 2)
  End If
  VehiclesS$ = LTrim$(Str$(veh))
  Print #RptHandle11, "H@";
  Print #RptHandle11, "1@";
  Print #RptHandle11, Jury$ + "@";
  Print #RptHandle11, Right$(Date$, 4) + "@";
  Print #RptHandle11, RTrim$(ThisDateS$) + "@";
  Print #RptHandle11, RTrim$(VehiclesS$) + "@";
  Print #RptHandle11, TotalAmt$ + "@";
  Print #RptHandle11, TotalCrCnt$ + "@";
  Print #RptHandle11, TotalCrAmt$
  Return
  
End Sub

