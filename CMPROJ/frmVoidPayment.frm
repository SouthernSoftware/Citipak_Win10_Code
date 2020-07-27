VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmVoidPayment 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Void Payment Transaction"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmVoidPayment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   350
      Left            =   0
      Top             =   0
   End
   Begin EditLib.fpDateTime txtPaymentDate 
      Height          =   324
      Left            =   7704
      TabIndex        =   0
      Top             =   1008
      Width           =   1548
      _Version        =   196608
      _ExtentX        =   2730
      _ExtentY        =   572
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
      AllowNull       =   -1  'True
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
      Text            =   ""
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
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
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   58
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "2:23 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdVoid 
      Height          =   390
      Left            =   8730
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1245
      _Version        =   131072
      _ExtentX        =   2196
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmVoidPayment.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn CmdExit 
      Height          =   384
      Left            =   10092
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1236
      _Version        =   131072
      _ExtentX        =   2180
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmVoidPayment.frx":0AA6
   End
   Begin fpBtnAtlLibCtl.fpBtn fpcmdDrawer 
      Height          =   390
      Left            =   7365
      TabIndex        =   68
      Top             =   7800
      Width           =   1245
      _Version        =   131072
      _ExtentX        =   2196
      _ExtentY        =   688
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
      ButtonDesigner  =   "frmVoidPayment.frx":0C82
   End
   Begin EditLib.fpText fpDesc 
      Height          =   324
      Left            =   7704
      TabIndex        =   1
      Top             =   1440
      Width           =   3228
      _Version        =   196608
      _ExtentX        =   5694
      _ExtentY        =   572
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
      AutoAdvance     =   0   'False
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   19
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
   Begin EditLib.fpText fpTransRecNo 
      Height          =   324
      Left            =   504
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   312
      Visible         =   0   'False
      Width           =   1764
      _Version        =   196608
      _ExtentX        =   3111
      _ExtentY        =   572
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      NoSpecialKeys   =   3
      AutoAdvance     =   0   'False
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
      Text            =   "fpText1"
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
   Begin VB.Label LblNoInterface 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   912
      TabIndex        =   79
      Top             =   7344
      Visible         =   0   'False
      Width           =   7716
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Number:"
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
      Height          =   264
      Left            =   1704
      TabIndex        =   78
      Top             =   3252
      Width           =   1824
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Source:"
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
      Height          =   264
      Left            =   1872
      TabIndex        =   76
      Top             =   2964
      Width           =   1656
   End
   Begin VB.Label PayOperator 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   3648
      TabIndex        =   75
      Top             =   3216
      Width           =   732
   End
   Begin VB.Label PaySource 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   3648
      TabIndex        =   74
      Top             =   2952
      Width           =   1860
   End
   Begin VB.Label payDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   3648
      TabIndex        =   73
      Top             =   2688
      Width           =   1860
   End
   Begin VB.Label fpReceiptNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   3648
      TabIndex        =   72
      Top             =   2400
      Width           =   1860
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Number:"
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
      Height          =   264
      Left            =   1704
      TabIndex        =   71
      Top             =   2400
      Width           =   1824
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Void Description:"
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
      Index           =   3
      Left            =   5856
      TabIndex        =   70
      Top             =   1488
      Width           =   1776
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      Height          =   5580
      Left            =   828
      Top             =   2064
      Width           =   10572
   End
   Begin VB.Label lblSource 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   3600
      TabIndex        =   65
      Top             =   1608
      Width           =   1872
   End
   Begin VB.Label lblOperator 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   3600
      TabIndex        =   64
      Top             =   936
      Width           =   732
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Void Source:"
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
      Height          =   312
      Left            =   1848
      TabIndex        =   63
      Top             =   1632
      Width           =   1656
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Void Trans Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   444
      Index           =   2
      Left            =   5928
      TabIndex        =   62
      Top             =   936
      Width           =   3672
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Number:"
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
      Height          =   312
      Left            =   1680
      TabIndex        =   61
      Top             =   984
      Width           =   1824
   End
   Begin VB.Label lblOperName 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   3600
      TabIndex        =   60
      Top             =   1272
      Width           =   1860
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Name:"
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
      Height          =   312
      Left            =   1680
      TabIndex        =   59
      Top             =   1320
      Width           =   1824
   End
   Begin VB.Label fpTotAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9360
      TabIndex        =   57
      Top             =   7308
      Width           =   1788
   End
   Begin VB.Label fpChkChgAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3792
      TabIndex        =   56
      Top             =   6060
      Width           =   1956
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
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
      Left            =   8376
      TabIndex        =   55
      Top             =   7332
      Width           =   900
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      Left            =   1176
      TabIndex        =   54
      Top             =   6912
      Width           =   1224
   End
   Begin VB.Label Lbl11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Check/Charge Amt Paid:"
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
      Height          =   228
      Left            =   1224
      TabIndex        =   53
      Top             =   6084
      Width           =   2472
   End
   Begin VB.Label lblchange 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Change Due:"
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
      Height          =   228
      Left            =   1824
      TabIndex        =   52
      Top             =   6624
      Width           =   1872
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tender Type:"
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
      Height          =   228
      Left            =   2112
      TabIndex        =   51
      Top             =   5544
      Width           =   1584
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Amount Paid:"
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
      Height          =   228
      Left            =   1428
      TabIndex        =   50
      Top             =   5820
      Width           =   2268
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Received:"
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
      Height          =   228
      Left            =   1884
      TabIndex        =   49
      Top             =   6348
      Width           =   1812
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Owed:"
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
      Height          =   228
      Index           =   0
      Left            =   1968
      TabIndex        =   48
      Top             =   5280
      Width           =   1728
   End
   Begin VB.Label lblPayDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2496
      TabIndex        =   47
      Top             =   6912
      Width           =   3252
   End
   Begin VB.Label fpChange 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3792
      TabIndex        =   46
      Top             =   6612
      Width           =   1956
   End
   Begin VB.Label fptotreceived 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3792
      TabIndex        =   45
      Top             =   6336
      Width           =   1956
   End
   Begin VB.Label fpCashAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3792
      TabIndex        =   44
      Top             =   5784
      Width           =   1956
   End
   Begin VB.Label fpTenderType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3792
      TabIndex        =   43
      Top             =   5508
      Width           =   1956
   End
   Begin VB.Label fpAmtOwed 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3792
      TabIndex        =   42
      Top             =   5232
      Width           =   1956
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Amount Paid"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   9360
      TabIndex        =   18
      Top             =   2160
      Width           =   1788
   End
   Begin VB.Label Revs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   14
      Left            =   9360
      TabIndex        =   41
      Top             =   6960
      Width           =   1788
   End
   Begin VB.Label Revs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   13
      Left            =   9360
      TabIndex        =   40
      Top             =   6636
      Width           =   1788
   End
   Begin VB.Label Revs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   12
      Left            =   9360
      TabIndex        =   39
      Top             =   6312
      Width           =   1788
   End
   Begin VB.Label Revs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   11
      Left            =   9360
      TabIndex        =   38
      Top             =   5988
      Width           =   1788
   End
   Begin VB.Label Revs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   10
      Left            =   9360
      TabIndex        =   37
      Top             =   5664
      Width           =   1788
   End
   Begin VB.Label Revs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   9
      Left            =   9360
      TabIndex        =   36
      Top             =   5340
      Width           =   1788
   End
   Begin VB.Label Revs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   8
      Left            =   9360
      TabIndex        =   35
      Top             =   5016
      Width           =   1788
   End
   Begin VB.Label Revs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   7
      Left            =   9360
      TabIndex        =   34
      Top             =   4692
      Width           =   1788
   End
   Begin VB.Label Revs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   6
      Left            =   9360
      TabIndex        =   33
      Top             =   4368
      Width           =   1788
   End
   Begin VB.Label Revs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   5
      Left            =   9360
      TabIndex        =   32
      Top             =   4044
      Width           =   1788
   End
   Begin VB.Label Revs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   4
      Left            =   9360
      TabIndex        =   31
      Top             =   3720
      Width           =   1788
   End
   Begin VB.Label Revs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   3
      Left            =   9360
      TabIndex        =   30
      Top             =   3396
      Width           =   1788
   End
   Begin VB.Label Revs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   2
      Left            =   9360
      TabIndex        =   29
      Top             =   3072
      Width           =   1788
   End
   Begin VB.Label Revs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   1
      Left            =   9360
      TabIndex        =   28
      Top             =   2748
      Width           =   1788
   End
   Begin VB.Label Revs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   0
      Left            =   9360
      TabIndex        =   27
      Top             =   2424
      Width           =   1788
   End
   Begin VB.Label fpAcct 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   3648
      TabIndex        =   26
      Top             =   3936
      Width           =   2100
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      Height          =   300
      Left            =   528
      TabIndex        =   25
      Top             =   4476
      Width           =   1248
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Height          =   300
      Left            =   804
      TabIndex        =   24
      Top             =   4200
      Width           =   972
   End
   Begin VB.Label Label2b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Account Number:"
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
      Index           =   1
      Left            =   672
      TabIndex        =   23
      Top             =   3936
      Width           =   2856
   End
   Begin VB.Label fptxtCity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   1824
      TabIndex        =   22
      Top             =   4800
      Width           =   3924
   End
   Begin VB.Label fptxtAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   1824
      TabIndex        =   21
      Top             =   4512
      Width           =   3924
   End
   Begin VB.Label fptxtName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   1824
      TabIndex        =   20
      Top             =   4224
      Width           =   3924
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Void "
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
      Left            =   4110
      TabIndex        =   19
      Top             =   360
      Width           =   4020
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   456
      Left            =   2598
      Top             =   312
      Width           =   7020
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Detail Distribution"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   5904
      TabIndex        =   17
      Top             =   2160
      Width           =   3444
   End
   Begin VB.Label fpDetDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   14
      Left            =   5904
      TabIndex        =   16
      Top             =   6960
      Width           =   3444
   End
   Begin VB.Label fpDetDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   13
      Left            =   5904
      TabIndex        =   15
      Top             =   6636
      Width           =   3444
   End
   Begin VB.Label fpDetDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   12
      Left            =   5904
      TabIndex        =   14
      Top             =   6312
      Width           =   3444
   End
   Begin VB.Label fpDetDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   11
      Left            =   5904
      TabIndex        =   13
      Top             =   5988
      Width           =   3444
   End
   Begin VB.Label fpDetDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   10
      Left            =   5904
      TabIndex        =   12
      Top             =   5664
      Width           =   3444
   End
   Begin VB.Label fpDetDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   9
      Left            =   5904
      TabIndex        =   11
      Top             =   5340
      Width           =   3444
   End
   Begin VB.Label fpDetDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   8
      Left            =   5904
      TabIndex        =   10
      Top             =   5016
      Width           =   3444
   End
   Begin VB.Label fpDetDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   7
      Left            =   5904
      TabIndex        =   9
      Top             =   4692
      Width           =   3444
   End
   Begin VB.Label fpDetDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   6
      Left            =   5904
      TabIndex        =   8
      Top             =   4368
      Width           =   3444
   End
   Begin VB.Label fpDetDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   5
      Left            =   5904
      TabIndex        =   7
      Top             =   4044
      Width           =   3444
   End
   Begin VB.Label fpDetDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   4
      Left            =   5904
      TabIndex        =   6
      Top             =   3720
      Width           =   3444
   End
   Begin VB.Label fpDetDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   3
      Left            =   5904
      TabIndex        =   5
      Top             =   3396
      Width           =   3444
   End
   Begin VB.Label fpDetDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   2
      Left            =   5904
      TabIndex        =   4
      Top             =   3072
      Width           =   3444
   End
   Begin VB.Label fpDetDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   1
      Left            =   5904
      TabIndex        =   3
      Top             =   2748
      Width           =   3444
   End
   Begin VB.Label fpDetDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   348
      Index           =   0
      Left            =   5904
      TabIndex        =   2
      Top             =   2424
      Width           =   3444
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   2586
      Top             =   192
      Width           =   7044
   End
   Begin VB.Shape Shape3 
      Height          =   612
      Left            =   816
      Top             =   7632
      Width           =   10596
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Transaction Date:"
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
      Height          =   264
      Index           =   1
      Left            =   1920
      TabIndex        =   77
      Top             =   2688
      Width           =   3864
   End
End
Attribute VB_Name = "frmVoidPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim CashFlag As Boolean, uselook As Boolean, CustAcct As Long
Dim EditFlag As Boolean, TempAmtRecv As Double, Answer As Integer
Dim ChkOKFlag As Boolean, BeenDone As Boolean, PayListCnt As Long
Dim DistArray() As DistArrayType
Dim PayList() As PayListType
Dim PayFileName As String, CMTrRecLen As Integer, OldBlTran As Long
Dim fromform As Form, toform As Form, codeopt As Integer, noreset As Boolean
Dim Oper As String, PayListRec As Long, RecpPort As String, CmNum As Long
Dim DefPayDate As String, TrTypeNum As Integer, NoDoModTrans As Boolean
Dim RevText$(1 To MaxRevsCnt)
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
  uselook = True
End Sub
Private Sub Form_Activate()
  If Val(fpTransRecNo) > 0 And Not BeenDone Then
    BeenDone = True
    fpReceiptNo = fpTransRecNo
    DispTrans
    DoEvents
  End If
End Sub

Private Sub cmdExit_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
  FntSize = frmMsgDialog.Label(1).FontSize
  frmMsgDialog.Label(1).FontSize = (FntSize + 2)
  frmMsgDialog.Label(2).FontSize = (FntSize + 2)
  frmMsgDialog.Label(3).FontSize = (FntSize + 2)
  MsgText(0) = "WARNING:Transaction Incomplete"
  MsgText(1) = ""
  MsgText(2) = "Are You Sure You Wish to EXIT?"
  MsgText(3) = "Ok to Continue,"
  MsgText(4) = "or Cancel to Review."
  MsgText(5) = ""
  If GetOKorNot(MsgText()) Then
   CMLog "USER WANTS TO Exit"
   Answer = 2 'ok to continue
  Else
   CMLog "USER Remain on screen"
   Answer = 1
  End If
  If Answer = 2 Then
    DoExitStuff
  End If
End Sub
Private Sub DoExitStuff()
  CustAcct = 0
  fpTransRecNo = 0
  BeenDone = False

  ActivateControls frmCMDispList
  CMLog "OUT: CMvoid" + " Oper:" + Oper$
  Unload Me
  DoEvents

End Sub
Private Sub fpcmdDrawer_Click()
  Dim Port As String, PortFile As Integer ', DPName As String, DefPrinter As String
  On Local Error Resume Next
  If RecpDef = 99 Then Exit Sub
  'RecPort = GetDEFPort%
  Port$ = QPTrim$(RecpPort)
  
  CMLog "Oper: " + Oper$ + "CMVoid-Open Drawer"
  PortFile = FreeFile
  Open Port$ For Output As #PortFile
  Print #PortFile, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
  Print #PortFile, Chr$(7)
  Close PortFile
End Sub
Private Sub fpCmdVoid_Click()
  Dim FntSize As Integer
  On Local Error GoTo ERRORSTUFF
  ReDim MsgText(0 To 5) As String
    FntSize = frmMsgDialog.Label(1).FontSize
    frmMsgDialog.Label(0).FontSize = (FntSize + 2)
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    frmMsgDialog.Label(4).FontSize = (FntSize + 2)
    MsgText(0) = "WARNING:VOID DATE NOTICE."
    MsgText(1) = "Void Trans Date Entered Above"
    MsgText(2) = "Will Be The Date This VOID"
    MsgText(3) = "Displays on Your Reports"
    MsgText(4) = "Ok to Continue,"
    MsgText(5) = "or Cancel to Review."
    If GetOKorNot(MsgText()) Then
     CMLog "USER WANTS TO Continue"
    Else
     CMLog "USER Canceled"
     Exit Sub
    End If
 
  FntSize = frmMsgDialog.Label(1).FontSize
  frmMsgDialog.Label(1).FontSize = (FntSize + 2)
  frmMsgDialog.Label(2).FontSize = (FntSize + 2)
  frmMsgDialog.Label(3).FontSize = (FntSize + 2)
  MsgText(0) = "WARNING:Last Chance."
  MsgText(1) = ""
  MsgText(2) = "Are You Sure You Wish to VOID this Payment?"
  MsgText(3) = "Ok to Continue,"
  MsgText(4) = "or Cancel to Review."
  MsgText(5) = ""
  If GetOKorNot(MsgText()) Then
   CMLog "USER WANTS TO Continue"
   Answer = 2 'ok to continue
  Else
   CMLog "USER Canceled"
   Answer = 1
  End If
  If Answer = 2 Then
    frmPrintReceipt.Label1.Caption = "Would you like a receipt?"
    frmPrintReceipt.Show 1
    If SavePay = True Then
  
      Select Case TrTypeNum
      Case 1    'Misc
        'do the misc void
        VoidMisc
      Case 27   'UB Dep
        'do the ub dep void
        VoidDeposit
      Case 24   'UB Pay
        'do the ubpayment adj
        VoidUBTrans
      Case 30 To 39, 131 'Taxes
        'only do the cm void for now
        VoidTax
      Case 161, 171
        VoidVATax
      Case 40 To 49, 141  'BL
        'do the BL void
        VoidBL
      Case 50 To 59, 151  'DC
        VoidDecal
      End Select
        'SaveTransaction
    
      If PrnRecp = True Then
        PrintReceipt
      End If
    
     MsgBox "Transaction Complete.", vbOKOnly, "Complete"
     'MsgBox "Transaction Complete.", vbOKOnly, "Complete"
      'ClearScn
      DoEvents
      DoExitStuff
    End If
  End If
  Exit Sub
ERRORSTUFF:
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "CMVoidPayment", "cmdVoid", Erl)
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
   Unload Me

End Sub
Private Sub txtPaymentDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
    KeyCode = 0
    
  ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then
    
  End If
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
        CMLog "Closed via cmVoidPayment by " + PWUser$ + " operator-" + Oper$
        CitiTerminate
      End If
    End If
  End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      KeyCode = 0
      DoEvents
      If cmdExit.Enabled Then
        Call cmdExit_Click
      End If
    Case vbKeyF2:
      KeyCode = 0
      DoEvents
      fpcmdDrawer_Click
'    Case vbKeyF4:
'      KeyCode = 0
'      DoEvents
'      If fpCmdInfo.Enabled Then
'        Call fpCmdInfo_Click
'      End If
'    Case vbKeyF5:
'      KeyCode = 0
'      DoEvents
'      If fpCmdCash.Enabled Then
'        Call fpcmdCash_Click
'      End If
'    Case vbKeyF6:
'      KeyCode = 0
'      DoEvents
'      If fpcmdCheck.Enabled Then
'        Call fpcmdCheck_Click
'      End If
'    Case vbKeyF7:
'      KeyCode = 0
'      DoEvents
'      If fpcmdFind.Enabled Then
'        Call fpcmdFind_Click
'      End If
'    Case vbKeyF8:
'      KeyCode = 0
'      DoEvents
'      If fpCmdCharge.Enabled Then
'        Call fpCmdCharge_Click
'      End If
'    Case vbKeyF9:
'      KeyCode = 0
'      DoEvents
'      If fpCmdDist.Enabled Then
'        Call fpCmdDist_Click
'      End If
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      If fpCmdVoid.Enabled Then
        Call fpCmdVoid_Click
      End If
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TownName$
  'txtPaymentDate.Text = DefPayDate
  
  noreset = False
  'LoadRevs
  lblOperator = OperNum
  lblOperName.Caption = PWUser
  'lblSource.Caption = "CM Utility"
  Oper$ = QPTrim(lblOperator.Caption)
  CMLog " IN Oper " + Oper$ + "Void Payment"
  
  GetRcpInfo
End Sub
Private Sub GetRcpInfo()
  Dim RP As Integer, lenRP As Integer, RP1 As Integer
  Dim RcptPrnFile As ReceiptPRNType
  RP1 = FreeFile
  lenRP = Len(RcptPrnFile)
  If Exist(RcptFileName$) Then
    Open RcptFileName$ For Random Shared As RP1 Len = lenRP
    Get RP1, 1, RcptPrnFile
    RecpPort = QPTrim(RcptPrnFile.RcpPort)
    If RcptPrnFile.PrnDefYN = 0 Then
      RecpDef = 0
    Else
      On Local Error GoTo nofound
      RP = FreeFile
      Open RecpPort For Output As RP
      Close RP
      RecpDef = 1
    End If
    If RcptPrnFile.CtlDefYN = 0 Then
      CntrlDef = 0
    Else
      CntrlDef = 1
    End If
  Close RP1
  Else
    RecpDef = 99
  End If
Exit Sub
nofound:
  RecpDef = 99
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ''' Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
'
'  If Me.Visible Then
'    Temp_Class.ResizeControls Me
'    DoEvents
'  End If
End Sub
Public Sub OpenMiscCodeFile(NumOfMiscRecs, MCFile)
  Dim MiscCodeRecLen As Integer
  ReDim MiscCodeRec(1) As MiscCodeRecType
  MiscCodeRecLen = Len(MiscCodeRec(1))
  MCFile = FreeFile
  Open UBPath$ + "CMMISCCD.DAT" For Random Shared As MCFile Len = MiscCodeRecLen
  NumOfMiscRecs = LOF(MCFile) \ MiscCodeRecLen

End Sub
Private Sub DispTrans()
  Dim RecSource As String, OperatorNumber As Integer
  Dim Fmt1 As String, Fmt3 As String, Fmt4 As String
  Dim CMTrRecLen As Integer, TRHandle As Integer, TrNumRecs As Long
  Dim NumOfMiscRecs As Long, cnt As Long, TRType As String
  Dim TxRev As Double, TRev As Integer, FDate As String
  Dim TotalAmount As Double, Change As Double
  Dim PrintMiscFlag As Integer, MCnt As Integer
  Dim MiscRevAmt As Double, NumofRevs As Integer, RCnt As Integer
  Dim PrintUtilFlag As Integer, PrintTaxFlag As Integer, Header As String
  Dim BegRecNo As Long, TransDate As Integer, TempTot As Double
  Dim UBSetupLen As Integer, RecAmt As Double
  Dim RevCnt As Integer, OutOfOrder As Boolean, x As Integer
  Dim Temp2 As Integer, uCnt As Integer, dcnt As Integer
  Dim TCnt As Long, PrnOpr As String
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long
  Dim CustFile As Integer, TMHandle As Integer
  Dim TxOpt1 As String, TxOpt2 As String, TxOpt3 As String
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  CmNum = fpTransRecNo
  OldBlTran = 0
  ReDim CMTrRec(1) As CMTransRecType            ' open transaction file
  CMTrRecLen = Len(CMTrRec(1))
  TRHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Access Read Write Shared As TRHandle Len = CMTrRecLen
  TrNumRecs& = LOF(TRHandle) \ CMTrRecLen
      Get TRHandle, CmNum, CMTrRec(1)
        TRType$ = ""
        TrTypeNum = CMTrRec(1).TransSource
        Select Case CMTrRec(1).TransSource
        Case 1
          TRType$ = "Miscellaneous"
        Case 27
          TRType$ = "Utility Deposit"
        Case 24
          TRType$ = "Utility Billing"
        Case 30 To 39, 131
          TRType$ = "Tax Billing"
          NoDoModTrans = True
        Case 161
          TRType$ = "Real Tax Billing"
          NoDoModTrans = True
        Case 171
          TRType$ = "Pers Tax Billing"
          NoDoModTrans = True
        Case 40 To 49, 141
          TRType$ = "Business License"
          If CMTrRec(1).TransSource <> 141 Then
            NoDoModTrans = True
          End If
        Case 50 To 59, 151
          TRType$ = "Vehicle Decal"
          If CMTrRec(1).TransSource <> 151 Then
            NoDoModTrans = True
          End If
        Case Else
          NoDoModTrans = True
        End Select
  If TrTypeNum = 24 Or TrTypeNum = 27 Then
    If IsDeleted(CMTrRec(1).TransAcctNum) Then
      CMLog "ERROR cmvoidub: Deleted Account:" + Str$(CMTrRec(1).TransAcctNum) + " Oper:" + Oper$
      NoDoModTrans = True
    End If
  End If
  If TrTypeNum = 151 Then
    If IsDCDeleted(CMTrRec(1).TransAcctNum) Then
      NoDoModTrans = True
      CMLog "ERROR cmvoiddc: Deleted Account:" + Str$(CMTrRec(1).TransAcctNum) + " Oper:" + Oper$
    Else
      If CMTrRec(1).TransDetNum > 0 Then
        ReDim DCTranRec(1) As DCTransRecType
        Dim DCTranRecLen As Integer
        Dim DCFile As Integer
        DCTranRecLen = Len(DCTranRec(1))
        DCFile = FreeFile
        Open UBPath + "DCTRANS.DAT" For Random Shared As DCFile Len = DCTranRecLen
        Get DCFile, CMTrRec(1).TransDetNum, DCTranRec(1)
        Close DCFile
        If DCTranRec(1).TransType <> 2 Or DCTranRec(1).VoidFlag = "Y" Then
          NoDoModTrans = True
          CMLog "ERROR cmvoiddc: Invalid Trans:" + Str$(CMTrRec(1).TransAcctNum) + " Oper:" + Oper$
        End If
      Else
        NoDoModTrans = True
        CMLog "ERROR cmvoiddc: Invalid Trans:" + Str$(CMTrRec(1).TransAcctNum) + " Oper:" + Oper$
      End If
    End If
  End If
  ReDim UBSetUpRec(1) As UBSetupRecType
'  ReDim DistArray(1 To 1) As DistArrayType
  Dim MCFile As Integer
  MCFile = FreeFile
  OpenMiscCodeFile NumOfMiscRecs, MCFile     ' opens misc code file
  ReDim MiscCodeRec(1) As MiscCodeRecType

  Fmt1$ = "###,###.##"
  Fmt3$ = "$#,###,###.##"
  Fmt4$ = "$###,###,###.##"
  'ReDim Array1(1 To Size) As Struct
     '####################################
      FDate$ = Num2Date(CMTrRec(1).TransDate)
      lblOperator.Caption = OperNum
      lblOperName.Caption = PWUser
      lblSource.Caption = "VOID " + TRType$
      txtPaymentDate = Format(Now, "mm/dd/yyyy")
      fpReceiptNo.Caption = Str(CmNum)
      payDate.Caption = FDate$
      PaySource.Caption = TRType$
      PayOperator.Caption = CMTrRec(1).TransOperNum
      fpAcct.Caption = CMTrRec(1).TransAcctNum
      fptxtName.Caption = Left$(CMTrRec(1).TransName, 18)
      lblPayDesc.Caption = QPTrim$(CMTrRec(1).TransDesc)
      fpCashAmt.Caption = Using(Fmt1$, CMTrRec(1).TransCash)
        If CMTrRec(1).TransTender = 4 Then
          fpChkChgAmt.Caption = Using(Fmt1$, CMTrRec(1).TransCheck)
          fpTenderType.Caption = "Charge"
        ElseIf CMTrRec(1).TransTender = 2 Then
          fpChkChgAmt.Caption = Using(Fmt1$, CMTrRec(1).TransCheck)
          fpTenderType.Caption = "Check"
        ElseIf CMTrRec(1).TransTender = 3 Then
          fpChkChgAmt.Caption = Using(Fmt1$, CMTrRec(1).TransCheck)
          fpTenderType.Caption = "Cash & Check"
        Else
          fpChkChgAmt.Caption = Using(Fmt1$, CMTrRec(1).TransCheck)
          fpTenderType.Caption = "Cash"
        End If
        RecAmt# = Round#(CMTrRec(1).TransCash + CMTrRec(1).TransCheck)
        fpAmtOwed.Caption = Using(Fmt1$, CMTrRec(1).TransAmtOwed)
        fpTotReceived.Caption = Using(Fmt1$, RecAmt#)
        fpTotAmt.Caption = Using(Fmt1$, CMTrRec(1).TransAmount)
'?????????HEY LOOK HERE MADE THIS CHANGE WHY WOULDN'T WORK??????????????
'        CHANGE# = Round#(RecAmt# - CMTrRec(1).TransAmount)
'        If CHANGE# < 0 Then CHANGE# = 0
'        fpChange.Caption = Using(Fmt1$, CHANGE#)
'Do change after get total of revs calc
        If CMTrRec(1).TransSource = 1 Then
          ' Misc Code Transaction ****************
          fptxtAddress.Visible = False
          fptxtCity.Visible = False
          Label7.Visible = False
          TempTot# = 0
         For MCnt = 1 To 5
            MiscRevAmt# = (CMTrRec(1).TransRevAmt(MCnt))
            MiscRevAmt# = Round#(MiscRevAmt#)
            TempTot# = Round#(TempTot# + MiscRevAmt#)
            If MiscRevAmt# <> 0 Then
              ' If There Is an Amount in Misc Rev 1-5 then get code record number from 6-10
              If CMTrRec(1).TransRevAmt(MCnt + 5) >= 1 Then
                Get MCFile, CMTrRec(1).TransRevAmt(MCnt + 5), MiscCodeRec(1)
                fpDetDesc(MCnt - 1).Caption = MiscCodeRec(1).MiscCode + "  " + QPTrim$(MiscCodeRec(1).Description)
                Revs(MCnt - 1).Caption = Using(Fmt1$, MiscRevAmt#)
              End If
            End If
          Next MCnt
          fpTotAmt.Caption = Using(Fmt1$, TempTot#)
          'End Misc Code Print ********************************
          For MCnt = 6 To 15
            fpDetDesc(MCnt - 1).Visible = False
            Revs(MCnt - 1).Visible = False
          Next
        End If

        If CMTrRec(1).TransSource >= 20 And CMTrRec(1).TransSource <= 29 Then
          'If CMTrRec(1).TransSource <> 27 Then
            'Utility Transaction *****************
            GoSub GetUBCust
            GoSub GetRevenueSources
            If NumofRevs > 0 Then
              For RCnt = 1 To NumofRevs
                fpDetDesc(RCnt - 1).Caption = RevText$(RCnt)
                Revs(RCnt - 1).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(RCnt))
                TempTot# = Round#(TempTot# + CMTrRec(1).TransRevAmt(RCnt))
              Next RCnt
            End If
            For RCnt = NumofRevs + 1 To 15
              fpDetDesc(RCnt - 1).Visible = False
              Revs(RCnt - 1).Visible = False
            Next
            fpTotAmt.Caption = Using(Fmt1$, TempTot#)
           If NoDoModTrans = True Then
             LblNoInterface.Caption = "A reversing Utility Transaction will NOT be created, only a reversing Cash Management Transaction.  Call Software Support if questions."
             LblNoInterface.Visible = True
           End If
        End If

        If CMTrRec(1).TransSource >= 30 And CMTrRec(1).TransSource <= 39 Then
          'Tax Transaction     *****************
          fptxtAddress.Visible = False
          fptxtCity.Visible = False
          Label7.Visible = False

          fpDetDesc(0).Caption = "Tax:"
          Revs(0).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(1))
          fpDetDesc(1).Caption = "Interest:"
          Revs(1).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(2))
          fpDetDesc(2).Caption = "Penalty:"
          Revs(2).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(3))
          fpDetDesc(3).Caption = "Storm:"
          Revs(3).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(4))
          fpDetDesc(4).Caption = "Past Tax:"
          Revs(4).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(6))
          fpDetDesc(5).Caption = "Interest:"
          Revs(5).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(7))
          fpDetDesc(6).Caption = "Penalty:"
          Revs(6).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(8))
          fpDetDesc(7).Caption = "Storm:"
          Revs(7).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(9))
          For cnt = 1 To 9
            TempTot# = Round#(TempTot# + CMTrRec(1).TransRevAmt(cnt))
          Next
          For cnt = 8 To 15
            fpDetDesc(cnt - 1).Visible = False
            Revs(cnt - 1).Visible = False
          Next
          fpTotAmt.Caption = Using(Fmt1$, TempTot#)
          LblNoInterface.Caption = "A reversing Tax Transaction will NOT be created, only a reversing Cash Management Transaction.  Call Software Support if questions."
          LblNoInterface.Visible = True
        End If
        If CMTrRec(1).TransSource = 131 Then
          'New Tax Transaction     *****************
          If Exist("TAXSETUP.DAT") Then
            ReDim TaxMasterRec(1) As TaxMasterType
            OpenTaxSetUpFile TMHandle
            Get TMHandle, 1, TaxMasterRec(1)
            Close TMHandle
            TxOpt1 = Mid$(TaxMasterRec(1).OptRev1, 1, 5)
            TxOpt2 = Mid$(TaxMasterRec(1).OptRev2, 1, 5)
            TxOpt3 = Mid$(TaxMasterRec(1).OptRev3, 1, 5)
          End If

          fptxtAddress.Visible = False
          fptxtCity.Visible = False
          Label7.Visible = False

          fpDetDesc(0).Caption = "Principle:"
          Revs(0).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(1))
          fpDetDesc(1).Caption = "Interest:"
          Revs(1).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(2))
          fpDetDesc(2).Caption = "Collection:"
          Revs(2).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(3))
          fpDetDesc(3).Caption = "Late List:"
          Revs(3).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(4))
          fpDetDesc(4).Caption = TxOpt1 + ":"
          Revs(4).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(5))
          fpDetDesc(5).Caption = TxOpt2 + ":"
          Revs(5).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(6))
          fpDetDesc(6).Caption = TxOpt3 + ":"
          Revs(6).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(7))
          fpDetDesc(7).Caption = "Discount:"
          Revs(7).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(8))
          fpDetDesc(8).Caption = "PrePay:"
          Revs(8).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(9))
          fpDetDesc(9).Caption = "#ofBills:"
          Revs(9).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(10))
          For cnt = 1 To 7
            TempTot# = Round#(TempTot# + CMTrRec(1).TransRevAmt(cnt))
          Next
          TempTot# = Round#(TempTot# + CMTrRec(1).TransRevAmt(9))
          TempTot# = Round#(TempTot# - CMTrRec(1).TransRevAmt(8))
          For cnt = 11 To 15
            fpDetDesc(cnt - 1).Visible = False
            Revs(cnt - 1).Visible = False
          Next
          fpTotAmt.Caption = Using(Fmt1$, TempTot#)
          LblNoInterface.Caption = "A reversing Tax Transaction will NOT be created, only a reversing Cash Management Transaction.  Call Software Support if questions."
          LblNoInterface.Visible = True
        ElseIf CMTrRec(1).TransSource = 161 Then
          'New VA Tax Transaction     *****************
          If Exist("TAXSETUP.DAT") Then
            ReDim TaxMaster(1) As VATaxMasterType
            OpenVATaxSetUpFile TMHandle
            Get TMHandle, 1, TaxMaster(1)
            Close TMHandle
            TxOpt1 = Mid$(TaxMaster(1).OptRev1, 1, 5)
            TxOpt2 = Mid$(TaxMaster(1).OptRev2, 1, 5)
            TxOpt3 = Mid$(TaxMaster(1).OptRev3, 1, 5)
          End If

          fptxtAddress.Visible = False
          fptxtCity.Visible = False
          Label7.Visible = False

          fpDetDesc(0).Caption = "Principle:"
          Revs(0).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(1))
          fpDetDesc(1).Caption = "Interest:"
          Revs(1).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(2))
          fpDetDesc(2).Caption = "Collection:"
          Revs(2).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(3))
          fpDetDesc(3).Caption = "Late List:"
          Revs(3).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(4))
          fpDetDesc(4).Caption = "Penalty:"
          Revs(4).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(5))
          fpDetDesc(5).Caption = TxOpt1 + ":"
          Revs(5).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(6))
          fpDetDesc(6).Caption = TxOpt2 + ":"
          Revs(6).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(7))
          fpDetDesc(7).Caption = TxOpt3 + ":"
          Revs(7).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(8))
          fpDetDesc(7).Caption = "Discount:"
          Revs(7).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(9))
          fpDetDesc(8).Caption = "PrePay:"
          Revs(8).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(10))
          fpDetDesc(9).Caption = "#ofBills:"
          Revs(9).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(11))
          For cnt = 1 To 8
            TempTot# = Round#(TempTot# + CMTrRec(1).TransRevAmt(cnt))
          Next
          TempTot# = Round#(TempTot# + CMTrRec(1).TransRevAmt(10))
          TempTot# = Round#(TempTot# - CMTrRec(1).TransRevAmt(9))
          For cnt = 12 To 15
            fpDetDesc(cnt - 1).Visible = False
            Revs(cnt - 1).Visible = False
          Next
          fpTotAmt.Caption = Using(Fmt1$, TempTot#)
          LblNoInterface.Caption = "A reversing Tax Transaction will NOT be created, only a reversing Cash Management Transaction.  Call Software Support if questions."
          LblNoInterface.Visible = True
        ElseIf CMTrRec(1).TransSource = 171 Then
          'New VA Tax Transaction  Pers   *****************
          If Exist("TAXSETUP.DAT") Then
            ReDim TaxMaster(1) As VATaxMasterType
            OpenVATaxSetUpFile TMHandle
            Get TMHandle, 1, TaxMaster(1)
            Close TMHandle
            TxOpt1 = Mid$(TaxMaster(1).OptRev1, 1, 5)
            TxOpt2 = Mid$(TaxMaster(1).OptRev2, 1, 5)
            TxOpt3 = Mid$(TaxMaster(1).OptRev3, 1, 5)
          End If

          fptxtAddress.Visible = False
          fptxtCity.Visible = False
          Label7.Visible = False

          fpDetDesc(0).Caption = "Principle1:"
          Revs(0).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(1))
          fpDetDesc(1).Caption = "Principle2:"
          Revs(1).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(2))
          fpDetDesc(2).Caption = "Principle3:"
          Revs(2).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(3))
          fpDetDesc(3).Caption = "Principle4:"
          Revs(3).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(4))
          fpDetDesc(4).Caption = "Principle5:"
          Revs(4).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(5))
          fpDetDesc(5).Caption = "Interest:"
          Revs(5).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(6))
          fpDetDesc(6).Caption = "Penalty:"
          Revs(6).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(7))
          fpDetDesc(7).Caption = TxOpt1 + ":"
          Revs(7).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(8))
          fpDetDesc(8).Caption = TxOpt2 + ":"
          Revs(8).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(9))
          fpDetDesc(9).Caption = TxOpt3 + ":"
          Revs(9).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(10))
          fpDetDesc(10).Caption = "Discount:"
          Revs(10).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(11))
          fpDetDesc(11).Caption = "PrePay:"
          Revs(11).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(12))
          fpDetDesc(12).Caption = "#ofBills:"
          Revs(12).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(13))
          For cnt = 1 To 10
            TempTot# = Round#(TempTot# + CMTrRec(1).TransRevAmt(cnt))
          Next
          TempTot# = Round#(TempTot# + CMTrRec(1).TransRevAmt(12))
          TempTot# = Round#(TempTot# - CMTrRec(1).TransRevAmt(11))
          For cnt = 14 To 15
            fpDetDesc(cnt - 1).Visible = False
            Revs(cnt - 1).Visible = False
          Next
          fpTotAmt.Caption = Using(Fmt1$, TempTot#)
          LblNoInterface.Caption = "A reversing Tax Transaction will NOT be created, only a reversing Cash Management Transaction.  Call Software Support if questions."
          LblNoInterface.Visible = True
        End If

        If CMTrRec(1).TransSource >= 40 And CMTrRec(1).TransSource <= 49 Then
          LblNoInterface.Caption = "A reversing Business License Transaction will NOT be created, only a reversing Cash Management Transaction.  Call Software Support if questions."
          LblNoInterface.Visible = True
        End If
        If CMTrRec(1).TransSource = 141 Then
          OldBlTran = CMTrRec(1).TransDetNum
          If CMTrRec(1).TransRevAmt(1) <> 0 Then
            fpDetDesc(0).Caption = GetCatRec(CMTrRec(1).TransRevAmt(1))
            Revs(0).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(6))
          End If
          If CMTrRec(1).TransRevAmt(2) <> 0 Then
            fpDetDesc(1).Caption = GetCatRec(CMTrRec(1).TransRevAmt(2))
            Revs(1).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(7))
          End If
          If CMTrRec(1).TransRevAmt(3) <> 0 Then
            fpDetDesc(2).Caption = GetCatRec(CMTrRec(1).TransRevAmt(3))
            Revs(2).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(8))
          End If
          If CMTrRec(1).TransRevAmt(4) <> 0 Then
            fpDetDesc(3).Caption = GetCatRec(CMTrRec(1).TransRevAmt(4))
            Revs(3).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(9))
          End If
          If CMTrRec(1).TransRevAmt(5) <> 0 Then
            fpDetDesc(4).Caption = GetCatRec(CMTrRec(1).TransRevAmt(5))
            Revs(4).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(10))
          End If
          If CMTrRec(1).TransRevAmt(11) <> 0 Then
            fpDetDesc(4).Caption = "Penalty:"
            Revs(4).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(11))
          End If
          If CMTrRec(1).TransRevAmt(12) <> 0 Then
            fpDetDesc(4).Caption = "Issue Fee:"
            Revs(4).Caption = Using(Fmt1$, CMTrRec(1).TransRevAmt(12))
          End If
      
        End If
        If CMTrRec(1).TransSource >= 50 And CMTrRec(1).TransSource <= 59 Then
          LblNoInterface.Caption = "A reversing Decal Transaction will NOT be created, only a reversing Cash Management Transaction.  Call Software Support if questions."
          LblNoInterface.Visible = True
        End If
        If CMTrRec(1).TransSource = 151 Then
        'do what you need to do here for void of new decals
          Revs(0).Caption = Using(Fmt1$, CMTrRec(1).TransAmount)
          For cnt = 1 To 14
            Revs(cnt).Visible = False
          Next
          getdcCat (CMTrRec(1).TransRevAmt(2))
          GetVeh (CMTrRec(1).TransRevAmt(1))
          GetDCCust (CMTrRec(1).TransAcctNum)
        
          If NoDoModTrans = True Then
            LblNoInterface.Caption = "A reversing Decal Transaction will NOT be created, only a reversing Cash Management Transaction.  Call Software Support if questions."
            LblNoInterface.Visible = True
          End If
        End If
        If Not TempTot# <> 0 Then
          TempTot# = CMTrRec(1).TransAmount
        End If
        Change# = Round#(RecAmt# - TempTot#)
        If Change# < 0 Then Change# = 0
        fpChange.Caption = Using(Fmt1$, Change#)
        
  Close         'Close all open files now

  'End If
  Exit Sub

GetRevenueSources:

  NumofRevs = MaxRevsCnt
  ReDim UBSetUpRec(1) As UBSetupRecType
  ReDim DistArray(1 To MaxRevsCnt) As DistArrayType
  ReDim UBSetUpRec(1) As UBSetupRecType
  UBSetupLen = Len(UBSetUpRec(1))
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  For RevCnt = 1 To MaxRevsCnt
    RevText$(RevCnt) = Left$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName), 14)
    DistArray(RevCnt).DistOrder = UBSetUpRec(1).Revenues(RevCnt).DistOr
    DistArray(RevCnt).DistCnt = RevCnt
    If Len(RevText$(RevCnt)) = 0 Then
      NumofRevs = RevCnt - 1
      Exit For
    End If
  Next

  ReDim Preserve DistArray(1 To NumofRevs) As DistArrayType

  Do
    OutOfOrder = False          'assume it's sorted
    For x = 1 To NumofRevs - 1
      If DistArray(x).DistOrder > DistArray(x + 1).DistOrder Then
        Temp2 = DistArray(x).DistOrder
        DistArray(x).DistOrder = DistArray(x + 1).DistOrder
        DistArray(x + 1).DistOrder = Temp2
        'SWAP DistArray(x), DistArray(x + 1)     'if we had to swap
        OutOfOrder = True       'we're not done yet
      End If
    Next
  Loop While OutOfOrder
Return
GetUBCust:
  CustAcct = Val(fpAcct)
  NumOfCustRecs& = FileSize(UBPath$ + "UBCUST.DAT") \ UBCustRecLen
  If CustAcct& > NumOfCustRecs& Or CustAcct& <= 0 Then
    LblNoInterface.Caption = "Invalid Account! No UB Transaction will be created."
    LblNoInterface.Visible = True
    'GoTo SkiptoHere
  End If
  If IsDeleted(CustAcct&) Then
    CMLog "ERROR: Deleted Account:" + Str$(CustAcct&) + " Oper:" + Oper$
    CustAcct& = 0
    LblNoInterface.Caption = "Deleted Account! No UB Transaction will be created."
    LblNoInterface.Visible = True
    'GoTo SkiptoHere
  End If
  CustFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  Get CustFile, CustAcct&, UBCustRec(1)
  Close CustFile

    UBLog "Oper:" + Oper$ + " CM Entering Void for Account:" + Str$(CustAcct&)
    CMLog "Oper:" + Oper$ + " Entering Void for UBAccount:" + Str$(CustAcct&)
    fptxtAddress.Caption = UBCustRec(1).Addr1
    fptxtCity.Caption = UBCustRec(1).City
'  fpDeposit.Caption = Using$("$###,###.##", UBCustRec(1).DepositAmt)
'SkiptoHere:
'  BeenDone = True
'  frmLookupError.Label.Caption = "Invalid Account Number"
'  frmLookupError.Label1.Caption = "Please Enter A Valid Account Number."
'  frmLookupError.Show 1
'  ClearScn
Return
End Sub
Private Sub GetVeh(Vehrecnum)
  Dim NumOfVRecs As Long, DCvFile As Integer, DCVehReclen As Integer, cnt As Integer
  ReDim DCVRec(1) As DCVehType
  If Vehrecnum > 0 Then
    DCVehReclen = Len(DCVRec(1))
    DCvFile = FreeFile
    Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
    NumOfVRecs = LOF(DCvFile) \ DCVehReclen
    Get DCvFile, Vehrecnum, DCVRec(1)
    Close DCvFile
    fpDetDesc(1).Caption = QPTrim$(DCVRec(1).Desc)
    fpDetDesc(2).Caption = QPTrim$(DCVRec(1).makemodel)
    fpDetDesc(3).Caption = QPTrim$(DCVRec(1).StateTag)
    fpDetDesc(4).Caption = "Expire " + Num2Date$(DCVRec(1).ExpireDate)
    fpDetDesc(5).Caption = "Sticker# " + DCVRec(1).Sticker
    For cnt = 6 To 14
      fpDetDesc(cnt).Visible = False
    Next
  End If
End Sub
Private Sub getdcCat(x As Long)
  Dim DCCatCodeRec As DCCatCodeRecType
  Dim DCCatCodeRecLen As Integer, ghandle As Integer
  DCCatCodeRecLen = Len(DCCatCodeRec)
  ghandle = FreeFile
  Open "DCCODE.DAT" For Random Access Read Write Shared As ghandle Len = DCCatCodeRecLen
    Get #ghandle, x, DCCatCodeRec
    fpDetDesc(0).Caption = QPTrim$(DCCatCodeRec.CATCODE) & " " & QPTrim$(DCCatCodeRec.CODEDESC)
  Close ghandle
End Sub

Private Sub GetDCCust(CNum As Long)
  Dim DCCustRecLen As Integer
  Dim CustFile As Integer
  ReDim DCCustREc(1) As DCCustRecType
  DCCustRecLen = Len(DCCustREc(1))
  If CNum& > 0 Then
    If Not IsDCDeleted(CNum) Then
    CustFile = FreeFile
    Open UBPath$ + "DCCUST.DAT" For Random Shared As CustFile Len = DCCustRecLen
    Get CustFile, CNum&, DCCustREc(1)
    Close CustFile
    fptxtAddress.Caption = DCCustREc(1).ADDRESS1
    fptxtCity.Caption = DCCustREc(1).City
    End If
  End If
End Sub

Private Sub VoidMisc()
  Dim ListFile As Integer, CHandle As Integer, MiscRevAmt As Double
  Dim PayFileName As String, UBPayRecLen As Integer
  Dim NumOfRecs As Long, CMTrRecLen As Integer
  Dim cnt As Integer, TRHandle As Integer, TrNumRecs As Long
  ReDim UBPaymentRec(1) As UBPaymentRecType
  Oper$ = QPTrim$(lblOperator.Caption)
  PayFileName$ = "C:\CPWork\CMPAY" + Oper$ + ".DAT"
  UBPayRecLen = Len(UBPaymentRec(1))
  ReDim CMTrRec(1) As CMTransRecType            ' open transaction file
  CMTrRecLen = Len(CMTrRec(1))
  TRHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Access Read Write Shared As TRHandle Len = CMTrRecLen
  TrNumRecs& = LOF(TRHandle) \ CMTrRecLen
  Get TRHandle, CmNum, CMTrRec(1)
  Close TRHandle
    For cnt = 1 To 5
      MiscRevAmt# = (CMTrRec(1).TransRevAmt(cnt))
      MiscRevAmt# = Round#(MiscRevAmt#)
      If MiscRevAmt# <> 0 Then
        ' If There Is an Amount in Misc Rev 1-5 then get code record number from 6-10
        If CMTrRec(1).TransRevAmt(cnt) <> 0 Then
          UBPaymentRec(1).PaidOwed(cnt + 5).AMTPD1 = CMTrRec(1).TransRevAmt(cnt + 5)
          UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = MiscRevAmt#
        End If
      Else
        UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = 0
        UBPaymentRec(1).PaidOwed(cnt + 5).AMTPD1 = 0
      End If
    Next cnt
    For cnt = 1 To 5
      UBPaymentRec(1).PaidOwed(cnt).AMTOWE1 = 0
    Next
  UBPaymentRec(1).OperNum = QPTrim(lblOperator.Caption)
  UBPaymentRec(1).payDate = Date2Num(txtPaymentDate)
  UBPaymentRec(1).CustAcct = 99999
  UBPaymentRec(1).CustName = QPTrim(fptxtName)
  UBPaymentRec(1).CustAddr = QPTrim(fptxtAddress)
  'UBPaymentRec(1).CUSTCMNT = 'QPTrim(Label4.Caption)
  'UBPaymentRec(1).TaxExempt = QPTrim(fptaxexmpt)
  UBPaymentRec(1).AmtOwed = fpAmtOwed
  UBPaymentRec(1).TenderTY = QPTrim$(fpTenderType.Caption)
  UBPaymentRec(1).CashAmt = fpCashAmt
  UBPaymentRec(1).ChkAmt = fpChkChgAmt
  UBPaymentRec(1).AmtRecd = fpTotReceived
  UBPaymentRec(1).Change = fpChange
  UBPaymentRec(1).Desc = QPTrim(fpDesc)
  UBPaymentRec(1).TotOwed = fpAmtOwed
  UBPaymentRec(1).AmtPaid = fpTotAmt
  'UBPaymentRec(1).Status = QPTrim(fpstatus)
  ListFile = FreeFile
  Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
    Put #ListFile, 1, UBPaymentRec(1)
    EditFlag = False
  CMLog "Oper:" + Oper$ + " Updated TempFile for Void Misc Pay"
  

  ReDim CMTrRec(1) As CMTransRecType
  CMTrRecLen = Len(CMTrRec(1))
  CHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Access Read Write Shared As CHandle Len = CMTrRecLen
  'CmNum = (LOF(CHandle) \ CMTrRecLen) + 1
  CMTrRec(1).TransDate = UBPaymentRec(1).payDate
  CMTrRec(1).TransAmount = -(Val(UBPaymentRec(1).AmtPaid))
  CMTrRec(1).TransCash = -(Val(UBPaymentRec(1).CashAmt))
  CMTrRec(1).TransCheck = -(Val(UBPaymentRec(1).ChkAmt))
  CMTrRec(1).TransAmtOwed = -(Val(UBPaymentRec(1).TotOwed))
  If Len(QPTrim$(UBPaymentRec(1).Desc)) = 0 Then
    CMTrRec(1).TransDesc = "V-Miscellaneous Payment"
  Else
    CMTrRec(1).TransDesc = (QPTrim$(UBPaymentRec(1).Desc))
  End If
  CMTrRec(1).TransSource = 201
  CMTrRec(1).TransName = UBPaymentRec(1).CustName
  CMTrRec(1).TransAcctNum = 99999
  CMTrRec(1).TransDetNum = CmNum
  CMTrRec(1).TransOperNum = OperNum
  Select Case QPTrim(UBPaymentRec(1).TenderTY)
    Case "Cash":
      CMTrRec(1).TransTender = 1
    Case "Check":
      CMTrRec(1).TransTender = 2
    Case "Cash & Check":
      CMTrRec(1).TransTender = 3
    Case "Charge":
      CMTrRec(1).TransTender = 4
    Case Else:
      '
  End Select
  CMTrRec(1).ChkByte = Chr$(1)
  CMTrRec(1).TransPad = ""
  CMTrRec(1).TransVoidNum = CmNum
  For cnt = 1 To 5
    CMTrRec(1).TransRevAmt(cnt) = -(UBPaymentRec(1).PaidOwed(cnt).AMTPD1)
  Next cnt

  For cnt = 1 To 5              ' Store the Misc Code Record Number in Rev Amt 6-10
    CMTrRec(1).TransRevAmt(cnt + 5) = UBPaymentRec(1).PaidOwed(cnt + 5).AMTPD1
  Next cnt
  Put CHandle, (LOF(CHandle) / CMTrRecLen) + 1, CMTrRec(1)
  Get CHandle, CmNum, CMTrRec(1)
  CMTrRec(1).TransVoidNum = (LOF(CHandle) / CMTrRecLen)
  Put CHandle, CmNum, CMTrRec(1)
  
  CmNum = (LOF(CHandle) / CMTrRecLen)

  Close
  CMLog "CMVoid-Misc Posted" + "  TRANS:" + Str$(CmNum&)
  'ClearScn
End Sub


Private Sub VoidUBTrans()
  Dim NumofRevs As Integer, RevCnt As Integer, ListFile As Integer
  Dim UBPayRecLen As Integer, CHandle As Integer
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long, NumOfRecs As Long
  Dim CustFile As Integer, cnt As Integer, UBTransRecLen As Integer
  Dim THandle As Integer, OldTotBalance As Double, RevAmts As Integer
  Dim TotalCustBalance As Double, CustChCnt As Integer, NextTransRec As Long
  Dim TAmtPaid As Double, AdjDesc As String
  Dim UBTran As Integer, NextTranRecs As Long, PrevLastTrans As Long
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBPaymentRec(1) As UBPaymentRecType
  Oper$ = QPTrim$(lblOperator.Caption)
  PayFileName$ = "C:\CPWork\CMPAY" + Oper$ + ".DAT"

  UBPayRecLen = Len(UBPaymentRec(1))
  UBCustRecLen = Len(UBCustRec(1))
  NumofRevs = MaxRevsCnt
  For cnt = 1 To 15
    If Revs(cnt - 1).Visible = True Then
    If Revs(cnt - 1) < -100000# Then
      UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = 0
    Else
      UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = Revs(cnt - 1)
    End If
    Else
      UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = 0
    End If
      UBPaymentRec(1).PaidOwed(cnt).AMTOWE1 = 0
  Next
  UBPaymentRec(1).OperNum = QPTrim(lblOperator.Caption)
  UBPaymentRec(1).payDate = Date2Num(txtPaymentDate)
  UBPaymentRec(1).CustAcct = fpAcct
  UBPaymentRec(1).CustName = QPTrim(fptxtName.Caption)
  UBPaymentRec(1).CustAddr = QPTrim(fptxtAddress.Caption)
  'UBPaymentRec(1).CUSTCMNT = 'QPTrim(Label4.Caption)
  UBPaymentRec(1).AmtOwed = fpAmtOwed
  UBPaymentRec(1).TenderTY = QPTrim$(fpTenderType.Caption)
  UBPaymentRec(1).CashAmt = fpCashAmt
  UBPaymentRec(1).ChkAmt = fpChkChgAmt
  UBPaymentRec(1).AmtRecd = fpTotReceived
  UBPaymentRec(1).Change = fpChange
  UBPaymentRec(1).Desc = QPTrim(fpDesc)
  UBPaymentRec(1).TotOwed = fpAmtOwed
  UBPaymentRec(1).AmtPaid = fpTotAmt
  'UBPaymentRec(1).Status = QPTrim(fpstatus)
  ListFile = FreeFile
  Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
  Put #ListFile, 1, UBPaymentRec(1)

  CMLog "CMOper:" + Oper$ + " Updated CMTempfile VoidUtilPay for Account:" + Str$(UBPaymentRec(1).CustAcct)
  UBLog "CMOper:" + Oper$ + " Updated CMTempfile VoidUtilPay for Account:" + Str$(UBPaymentRec(1).CustAcct)
  ReDim UBTransRec(1) As UBTransRecType
  If NoDoModTrans = False Then
  UBCustRecLen = Len(UBCustRec(1))
  UBTransRecLen = Len(UBTransRec(1))
'oooeeeooohaha
  CustFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  Get CustFile, UBPaymentRec(1).CustAcct, UBCustRec(1)
  Close CustFile

  UBTransRec(1).TransDate = UBPaymentRec(1).payDate
  UBTransRec(1).TransType = TranOverPayAdjustment
  UBTransRec(1).TransDesc = "CM-Payment Adjustment"
  UBTransRec(1).TransAmt = UBPaymentRec(1).AmtPaid
  For cnt = 1 To 15
    UBTransRec(1).RevAmt(cnt) = UBPaymentRec(1).PaidOwed(cnt).AMTPD1
  Next
  UBTransRec(1).CustStatus = UBCustRec(1).Status
  UBTransRec(1).CustAcctNo = UBPaymentRec(1).CustAcct
  
  UBTransRec(1).CheckAmount = UBPaymentRec(1).ChkAmt
  UBTransRec(1).CashAmount = UBPaymentRec(1).CashAmt

  If UBTransRec(1).CheckAmount > 0 And UBTransRec(1).CashAmount > 0 Then
    UBTransRec(1).PayTypeCode = 3
  ElseIf UBTransRec(1).CashAmount > 0 Then
    UBTransRec(1).PayTypeCode = 1
  ElseIf UBTransRec(1).CheckAmount > 0 Then
     If QPTrim(UBPaymentRec(1).TenderTY) = "Charge" Then
       UBTransRec(1).PayTypeCode = 4
     Else
       UBTransRec(1).PayTypeCode = 2
     End If
  End If
  UBTransRec(1).OperatorNumber = OperNum
    For RevCnt = 1 To MaxRevsCnt
      UBCustRec(1).CurrRevAmts(RevCnt) = Round#(UBCustRec(1).CurrRevAmts(RevCnt) + UBTransRec(1).RevAmt(RevCnt))
      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + UBTransRec(1).RevAmt(RevCnt))
    Next
  UBTransRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  UBTransRec(1).FromCMFlag = True
  AdjDesc$ = QPTrim$(fpDesc)
  If Len(AdjDesc$) > 0 Then
    UBTransRec(1).BillMsg = AdjDesc$
  End If

  CustFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  UBTran = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTransRecLen
  NextTranRecs& = (LOF(UBTran) \ UBTransRecLen) + 1
  PrevLastTrans& = UBCustRec(1).LastTrans
  UBTransRec(1).PrevTrans = PrevLastTrans&
  UBCustRec(1).LastTrans = NextTranRecs&

  If Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) = 0 Then
    If UBCustRec(1).Status = "B" Then
      CustChCnt = CustChCnt + 1
      UBLog "CM-ADJUST: SET CUST STATUS to I. Acct:" + Str$(UBTransRec(1).CustAcctNo)
      CMLog "CM-ADJUST: SET CUST STATUS to I. Acct:" + Str$(UBTransRec(1).CustAcctNo)
      UBCustRec(1).Status = "I"
    End If
  End If
  Put CustFile, UBTransRec(1).CustAcctNo, UBCustRec(1)
  Put UBTran, NextTranRecs&, UBTransRec(1)
  'Get #CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1)
  Close UBTran, CustFile

  UBLog "CMVoid Posted Util-PayADJ CUST:" + Str$(UBPaymentRec(1).CustAcct) + "  TRANS:" + Str$(NextTranRecs&)
  CMLog "CMVoid-Posted Util-PayADJ CUST:" + Str$(UBPaymentRec(1).CustAcct) + "  TRANS:" + Str$(NextTranRecs&)

'  MsgBox "Save procedure complete.", vbOKOnly, "Completed"
 End If
  ReDim CMTrRec(1) As CMTransRecType
  CMTrRecLen = Len(CMTrRec(1))
  CMTrRec(1).TransDate = UBPaymentRec(1).payDate
  CMTrRec(1).TransAmount = -(Val(UBPaymentRec(1).AmtPaid)) 'UBTransRec(1).CashAmount + UBTransRec(1).CheckAmount
  CMTrRec(1).TransCash = -(Val(UBPaymentRec(1).CashAmt))
  CMTrRec(1).TransAmtOwed = -(Val(UBPaymentRec(1).AmtOwed))
  CMTrRec(1).TransCheck = -(Val(UBPaymentRec(1).ChkAmt))
  CMTrRec(1).TransDesc = UBPaymentRec(1).Desc
  CMTrRec(1).TransSource = 224
  CMTrRec(1).TransName = UBPaymentRec(1).CustName
  CMTrRec(1).TransAcctNum = UBPaymentRec(1).CustAcct
  If CMTrRec(1).TransCheck <> 0 And CMTrRec(1).TransCash <> 0 Then
    CMTrRec(1).TransTender = 3
  ElseIf CMTrRec(1).TransCash <> 0 Then
    CMTrRec(1).TransTender = 1
  ElseIf CMTrRec(1).TransCheck <> 0 Then
     If QPTrim(UBPaymentRec(1).TenderTY) = "Charge" Then
       CMTrRec(1).TransTender = 4
     Else
       CMTrRec(1).TransTender = 2
     End If
  End If
  CMTrRec(1).ChkByte = Chr$(1)
  CMTrRec(1).TransDetNum = 0
  CMTrRec(1).TransOperNum = OperNum
  CMTrRec(1).TransPad = ""
  CMTrRec(1).TransVoidNum = CmNum
  For cnt = 1 To 15
    CMTrRec(1).TransRevAmt(cnt) = -(Val(UBPaymentRec(1).PaidOwed(cnt).AMTPD1))
  Next cnt
  Close ListFile
  CHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Shared As CHandle Len = CMTrRecLen
  Put CHandle, (LOF(CHandle) / CMTrRecLen) + 1, CMTrRec(1)
  'CmNum = (LOF(CHandle) / CMTrRecLen) + 1
  Get CHandle, CmNum, CMTrRec(1)
    CMTrRec(1).TransVoidNum = (LOF(CHandle) / CMTrRecLen)
  Put CHandle, CmNum, CMTrRec(1)
  
  CmNum = (LOF(CHandle) / CMTrRecLen)
  Close CHandle
  UBLog "CMVoid Posted CM-CUST:" + Str$(UBPaymentRec(1).CustAcct) + "  TRANS:" + Str$(NextTranRecs&)
  CMLog "CMVoid-Posted CM-CUST:" + Str$(UBPaymentRec(1).CustAcct) + "  TRANS:" + Str$(NextTranRecs&)
 'ClearScn
End Sub
Private Sub VoidDeposit()
  Dim UBTransRecLen As Integer, NextTranRecs As Long
  Dim TransDate As Integer, TransAmt As Double, CustChCnt As Integer
  Dim UBCustRecLen As Integer, NumOfCustRecs As Long
  Dim UBPayRecLen As Integer, CHandle As Integer, NumofRevs As Integer
  Dim CustFile  As Integer, cnt As Integer, RevCnt As Integer
  Dim UBTran As Integer, NumOfTranRecs As Long, PrevLastTrans As Long
  Dim TotalDepAmt As Double, LastTran As Long, ListFile As Integer
  ReDim RevAmts(1 To 15) As Double
  ReDim UBCustRec(1 To 2) As NewUBCustRecType
  CustFile = FreeFile
  UBCustRecLen = Len(UBCustRec(1))
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  Get CustFile, CustAcct&, UBCustRec(1)
  Close CustFile
  ReDim UBPaymentRec(1) As UBPaymentRecType
  Oper$ = QPTrim$(lblOperator.Caption)
  PayFileName$ = "C:\CPWork\CMPAY" + Oper$ + ".DAT"
  UBPayRecLen = Len(UBPaymentRec(1))

  GoSub GetDepRevAmts
  NumofRevs = MaxRevsCnt
  For cnt = 1 To 15
    If Revs(cnt - 1).Visible = True Then
    If Revs(cnt - 1) < -100000# Then
      UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = 0
    Else
      UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = Revs(cnt - 1)
    End If
    Else
      UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = 0
    End If
      UBPaymentRec(1).PaidOwed(cnt).AMTOWE1 = 0
  Next
  UBPaymentRec(1).OperNum = QPTrim(lblOperator.Caption)
  UBPaymentRec(1).payDate = Date2Num(txtPaymentDate)
  UBPaymentRec(1).CustAcct = fpAcct
  UBPaymentRec(1).CustName = QPTrim(fptxtName.Caption)
  UBPaymentRec(1).CustAddr = QPTrim(fptxtAddress.Caption)
  'UBPaymentRec(1).CUSTCMNT = 'QPTrim(Label4.Caption)
  UBPaymentRec(1).AmtOwed = fpAmtOwed
  UBPaymentRec(1).TenderTY = QPTrim$(fpTenderType.Caption)
  UBPaymentRec(1).CashAmt = fpCashAmt
  UBPaymentRec(1).ChkAmt = fpChkChgAmt
  UBPaymentRec(1).AmtRecd = fpTotReceived
  UBPaymentRec(1).Change = fpChange
  UBPaymentRec(1).Desc = QPTrim(fpDesc)
  UBPaymentRec(1).TotOwed = fpAmtOwed
  UBPaymentRec(1).AmtPaid = fpTotAmt
  'UBPaymentRec(1).Status = QPTrim(fpstatus)
  ListFile = FreeFile
  Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
  Put #ListFile, 1, UBPaymentRec(1)

  CMLog "CMOper:" + Oper$ + " Updated CMTempfile VoidUDep for Account:" + Str$(UBPaymentRec(1).CustAcct)
  UBLog "CMOper:" + Oper$ + " Updated CMTempfile VoidUDep for Account:" + Str$(UBPaymentRec(1).CustAcct)
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  CustFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  Get CustFile, UBPaymentRec(1).CustAcct, UBCustRec(1)
  Close CustFile
  TransDate = Date2Num(txtPaymentDate)
  'Transamt# = Val(Label4.Caption)
  TransAmt# = UBPaymentRec(1).AmtPaid
  If NoDoModTrans = False Then
  UBTransRec(1).TransDate = TransDate
  'UBTransRec(1)CustLocation = RecNo&
  UBTransRec(1).CustStatus = UBCustRec(1).Status
  UBTransRec(1).CustAcctNo = UBPaymentRec(1).CustAcct
  UBTransRec(1).TransAmt = UBPaymentRec(1).AmtPaid
  UBTransRec(1).TransDesc = "CM-Void Deposit Payment"

  UBTransRec(1).TransType = TranDepPaymentVoid
  UBTransRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  If Not Round#(TotalDepAmt# - TransAmt#) < 0 Then
    UBCustRec(1).DepositAmt = Round#(TotalDepAmt# - TransAmt#)
  Else
    UBCustRec(1).DepositAmt = 0
  End If
  For RevCnt = 1 To 15
    If Revs(RevCnt - 1).Visible Then
      UBTransRec(1).RevAmt(RevCnt) = Revs(RevCnt - 1)
    Else
      UBTransRec(1).RevAmt(RevCnt) = 0
    End If
  Next
  
  CustFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen

  UBTran = FreeFile
  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTransRecLen

  NextTranRecs& = (LOF(UBTran) \ UBTransRecLen) + 1
  PrevLastTrans& = UBCustRec(1).LastTrans
  UBTransRec(1).PrevTrans = PrevLastTrans&
  UBCustRec(1).LastTrans = NextTranRecs&
  UBTransRec(1).OperatorNumber = OperNum
  UBTransRec(1).VoidFlag = False
  UBTransRec(1).FromCMFlag = True
'remark these
'*******************************
'''  UBTransRec(1).TransDate = Date2Num("01-31-2001")
'''  UBTransRec(1).TransDesc = "Applied Deposit"
'''  UBTransRec(1).TransType = TranAppliedDeposit
'''  UBTransRec(1).Transamt = TotalDepAmt#
'*******************************

  Put CustFile, CustAcct&, UBCustRec(1)
  Put UBTran, NextTranRecs&, UBTransRec(1)
'  'write the original trans with voidflag set
'Can't do this in cm no trans from ub
'      Get UBTran, fpTranNum, UBTransRec(1)
'        UBTransRec(1).VoidFlag = True
'        Put UBTran, fpTranNum, UBTransRec(1)
  End If
  UBLog "CMVoid-VoidDep Post Util CUST:" + Str$(UBPaymentRec(1).CustAcct) + "  TRANS:" + Str$(NextTranRecs&)
  CMLog "CMVoid-VoidDep Post Util CUST:" + Str$(UBPaymentRec(1).CustAcct) + "  TRANS:" + Str$(NextTranRecs&)
  Close UBTran, CustFile
  ReDim CMTrRec(1) As CMTransRecType
  CMTrRecLen = Len(CMTrRec(1))
  CMTrRec(1).TransDate = UBPaymentRec(1).payDate
  CMTrRec(1).TransAmount = -(Val(UBPaymentRec(1).AmtPaid)) 'UBTransRec(1).CashAmount + UBTransRec(1).CheckAmount
  CMTrRec(1).TransCash = -(Val(UBPaymentRec(1).CashAmt))
  CMTrRec(1).TransAmtOwed = -(Val(UBPaymentRec(1).AmtOwed))
  CMTrRec(1).TransCheck = -(Val(UBPaymentRec(1).ChkAmt))
  CMTrRec(1).TransDesc = UBPaymentRec(1).Desc
  CMTrRec(1).TransSource = 227
  CMTrRec(1).TransName = UBPaymentRec(1).CustName
  CMTrRec(1).TransAcctNum = UBPaymentRec(1).CustAcct
  If CMTrRec(1).TransCheck <> 0 And CMTrRec(1).TransCash <> 0 Then
    CMTrRec(1).TransTender = 3
  ElseIf CMTrRec(1).TransCash <> 0 Then
    CMTrRec(1).TransTender = 1
  ElseIf CMTrRec(1).TransCheck <> 0 Then
     If QPTrim(UBPaymentRec(1).TenderTY) = "Charge" Then
       CMTrRec(1).TransTender = 4
     Else
       CMTrRec(1).TransTender = 2
     End If
  End If
  CMTrRec(1).TransDetNum = 0
  CMTrRec(1).TransOperNum = OperNum
  CMTrRec(1).TransPad = ""
  CMTrRec(1).TransVoidNum = CmNum
  CMTrRec(1).ChkByte = Chr$(1)
  For cnt = 1 To 15
    CMTrRec(1).TransRevAmt(cnt) = -(Val(UBPaymentRec(1).PaidOwed(cnt).AMTPD1))
  Next cnt
  Close ListFile
  CHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Shared As CHandle Len = CMTrRecLen
  Put CHandle, (LOF(CHandle) / CMTrRecLen) + 1, CMTrRec(1)
  'CmNum = (LOF(CHandle) / CMTrRecLen) + 1
  Get CHandle, CmNum, CMTrRec(1)
    CMTrRec(1).TransVoidNum = (LOF(CHandle) / CMTrRecLen)
  Put CHandle, CmNum, CMTrRec(1)
  
  CmNum = (LOF(CHandle) / CMTrRecLen)
  Close CHandle
  UBLog "CMVoid-VoidDep Post CM CUST:" + Str$(UBPaymentRec(1).CustAcct) + "  TRANS:" + Str$(NextTranRecs&)
  CMLog "CMVoid-VoidDep Post CM CUST:" + Str$(UBPaymentRec(1).CustAcct) + "  TRANS:" + Str$(NextTranRecs&)
Exit Sub

GetDepRevAmts:
  TotalDepAmt# = 0
  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  LastTran& = UBCustRec(1).LastTrans
  If LastTran& > 0 Then
    UBTran = FreeFile
    Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTransRecLen
    Do
      Get #UBTran, LastTran&, UBTransRec(1)
      If UBTransRec(1).TransType = TranDepositPayment Then
        For RevCnt = 1 To 15
          If UBTransRec(1).RevAmt(RevCnt) <> 0 Then
            RevAmts(RevCnt) = Round#(RevAmts(RevCnt) + UBTransRec(1).RevAmt(RevCnt))
            TotalDepAmt# = Round#(TotalDepAmt# + UBTransRec(1).RevAmt(RevCnt))
          End If
        Next
      ElseIf (UBTransRec(1).TransType = TranAppliedDeposit) Or (UBTransRec(1).TransType = TranRefundDeposit) Or (UBTransRec(1).TransType = TranDepPaymentVoid) Then
        For RevCnt = 1 To 15
          If UBTransRec(1).RevAmt(RevCnt) <> 0 Then
            RevAmts(RevCnt) = Round#(RevAmts(RevCnt) - UBTransRec(1).RevAmt(RevCnt))
            TotalDepAmt# = Round#(TotalDepAmt# - UBTransRec(1).RevAmt(RevCnt))
          End If
        Next
      End If
      LastTran& = UBTransRec(1).PrevTrans
    Loop While LastTran& > 0
    Close UBTran
  End If

Return
End Sub
Private Sub VoidBL()
  Dim THandle As Integer
  Dim ARTransRec(1 To 2) As ARTransRecType
  Dim NumOfTransRecs As Long
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim cnt As Long
  Dim NextTransRec As Long
  Dim CustRecNum As Integer
  Dim Adj$, TBal#, Prev&
  Dim ThisType As Integer, x As Integer
  Dim PrintType$
  PayFileName$ = "C:\CPWork\CMPAY" + Oper$ + ".DAT"
  If OldBlTran > 0 Then
  'On Error GoTo ERRORSTUFF
  OpenBLTransFile THandle
  OpenBLCustFile CHandle

  If NoDoModTrans = False Then
  NumOfTransRecs = LOF(THandle) \ Len(ARTransRec(1))
  NextTransRec = NumOfTransRecs + 1
  Get THandle, OldBlTran, ARTransRec(1)
  Close THandle
  

  'If OldRound(fpcurrTotAmt.DoubleValue) <> 0 Then
    CustRecNum = ARTransRec(1).CustomerNumber
    'retrieve the customer file on the customer being adjusted
    Get CHandle, CustRecNum, CustRec
    ARTransRec(2).CustomerNumber = CustRecNum
    ARTransRec(2).TransDate = Date2Num(txtPaymentDate)
    ARTransRec(2).Posted2GL = "N"
    'SavedTotAmt = fpcurrTotAmt.DoubleValue
        ARTransRec(2).TransType = 13 'adjust up Payment
        Adj$ = "ADJUPA "
        If (ARTransRec(1).LicAmt > 0 Or ARTransRec(1).IssAmt > 0) And ARTransRec(1).PenAmt > 0 Then
          ARTransRec(2).DetailTransType = 411
        ElseIf (ARTransRec(1).LicAmt > 0 Or ARTransRec(1).IssAmt > 0) And ARTransRec(1).PenAmt = 0 Then
          ARTransRec(2).DetailTransType = 410
        ElseIf ARTransRec(1).LicAmt = 0 And ARTransRec(1).PenAmt > 0 Then
          ARTransRec(2).DetailTransType = 401
        Else
          ARTransRec(2).DetailTransType = 0
        End If

        ARTransRec(2).TransAmount = ARTransRec(1).TransAmount
        ARTransRec(2).PenAmt = ARTransRec(1).PenAmt
        ARTransRec(2).LicAmt = ARTransRec(1).LicAmt
        ARTransRec(2).IssAmt = ARTransRec(1).IssAmt


        ARTransRec(2).CatLicAmt1 = ARTransRec(1).CatLicAmt1
        ARTransRec(2).CatLicAmt2 = ARTransRec(1).CatLicAmt2
        ARTransRec(2).CatLicAmt3 = ARTransRec(1).CatLicAmt3
        ARTransRec(2).CatLicAmt4 = ARTransRec(1).CatLicAmt4
        ARTransRec(2).CatLicAmt5 = ARTransRec(1).CatLicAmt5
        'adjusts payment transactions, not billing
          CustRec.FeeLicPay1 = CustRec.FeeLicPay1 + ARTransRec(1).CatLicAmt1
          CustRec.FeeLicPay2 = CustRec.FeeLicPay2 + ARTransRec(1).CatLicAmt2
          CustRec.FeeLicPay3 = CustRec.FeeLicPay3 + ARTransRec(1).CatLicAmt3
          CustRec.FeeLicPay4 = CustRec.FeeLicPay4 + ARTransRec(1).CatLicAmt4
          CustRec.FeeLicPay5 = CustRec.FeeLicPay5 + ARTransRec(1).CatLicAmt5

        ARTransRec(2).CatLicBal1 = CustRec.FeeLicBal1 + ARTransRec(1).CatLicAmt1 'CatLicBal1
        ARTransRec(2).CatLicBal2 = CustRec.FeeLicBal2 + ARTransRec(1).CatLicAmt2
        ARTransRec(2).CatLicBal3 = CustRec.FeeLicBal3 + ARTransRec(1).CatLicAmt3
        ARTransRec(2).CatLicBal4 = CustRec.FeeLicBal4 + ARTransRec(1).CatLicAmt4
        ARTransRec(2).CatLicBal5 = CustRec.FeeLicBal5 + ARTransRec(1).CatLicAmt5

        CustRec.FeeLicBal1 = CustRec.FeeLicBal1 + ARTransRec(1).CatLicAmt1
        CustRec.FeeLicBal2 = CustRec.FeeLicBal2 + ARTransRec(1).CatLicAmt2
        CustRec.FeeLicBal3 = CustRec.FeeLicBal3 + ARTransRec(1).CatLicAmt3
        CustRec.FeeLicBal4 = CustRec.FeeLicBal4 + ARTransRec(1).CatLicAmt4
        CustRec.FeeLicBal5 = CustRec.FeeLicBal5 + ARTransRec(1).CatLicAmt5


        CustRec.LicBal = OldRound#(CustRec.LicBal + ARTransRec(1).LicAmt)
        CustRec.PenBal = OldRound#(CustRec.PenBal + ARTransRec(1).PenAmt)
        CustRec.IssuanceBal = OldRound#(CustRec.IssuanceBal + ARTransRec(1).IssAmt)

        TBal# = OldRound#(CustRec.LicBal + CustRec.PenBal + CustRec.IssuanceBal)

    '-----------------------------------------------------------

    CustRec.AcctBal = TBal#
    ARTransRec(2).BalanceAfterTrans = TBal#

    Adj$ = Adj$ + QPTrim$(fpDesc.Text)
    ARTransRec(2).TransDesc = "CM- " + QPTrim$(fpDesc.Text) 'Adj$
    ARTransRec(2).FeeAmt = 0
    ARTransRec(2).CashAmount = 0                'EditBegBalRec(1).Amount
    ARTransRec(2).ChkAmount = 0
    ARTransRec(2).ExtraRoom = ""
    ARTransRec(2).NextTrans = 0 ' CustRec.LastTrans
    ARTransRec(2).CatCodeRec1 = ARTransRec(1).CatCodeRec1
    ARTransRec(2).CatCodeRec2 = ARTransRec(1).CatCodeRec2
    ARTransRec(2).CatCodeRec3 = ARTransRec(1).CatCodeRec3
    ARTransRec(2).CatCodeRec4 = ARTransRec(1).CatCodeRec4
    ARTransRec(2).CatCodeRec5 = ARTransRec(1).CatCodeRec5
    ARTransRec(2).PenBal = CustRec.PenBal
    ARTransRec(2).LicBal = CustRec.LicBal
    ARTransRec(2).IssBal = CustRec.IssuanceBal
      OpenBLTransFile THandle
 ' OpenBLCustFile CHandle

    Put THandle, NextTransRec, ARTransRec(2)

    If CustRec.FirstTrans = 0 Then
      CustRec.FirstTrans = NextTransRec
      CustRec.LastTrans = NextTransRec
      Put CHandle, CustRecNum, CustRec
    Else
      Prev& = CustRec.LastTrans
      CustRec.LastTrans = NextTransRec
      Put CHandle, CustRecNum, CustRec
      Get THandle, Prev&, ARTransRec(2)
      ARTransRec(2).NextTrans = NextTransRec
      Put THandle, Prev&, ARTransRec(2)
    End If
  End If
  End If
  'now record the activity that took place here in the arlog.dat file
  BLLog "Void-BLOverPay PostedBL CUST:" + QPTrim$(fpAcct.Caption) + "  TRANS:" + Str$(NextTransRec&)
  CMLog "CMVoid-BLPay PostedBL CUST:" + QPTrim$(fpAcct.Caption) + "  TRANS:" + Str$(CmNum)

  'Call LogSaves
  
  Close
  ReDim CMTrRec(1 To 2) As CMTransRecType
  CHandle = FreeFile
  CMTrRecLen = Len(CMTrRec(1))
  Open UBPath$ + "CMTRANS.DAT" For Random Shared As CHandle Len = CMTrRecLen
    'CmNum = Val(fpReceiptNo.Caption)
    Get CHandle, CmNum, CMTrRec(1)
  Close CHandle
  
  
  CMTrRec(2).TransDate = Date2Num(txtPaymentDate)
  CMTrRec(2).TransAmount = -(Val(CMTrRec(1).TransAmount)) 'UBTransRec(1).CashAmount + UBTransRec(1).CheckAmount
  CMTrRec(2).TransCash = -(Val(CMTrRec(1).TransCash))
  CMTrRec(2).TransAmtOwed = -(Val(CMTrRec(1).TransAmtOwed))
  CMTrRec(2).TransCheck = -(Val(CMTrRec(1).TransCheck))
  CMTrRec(2).TransDesc = QPTrim$(fpDesc.Text)
  CMTrRec(2).TransSource = 241
  CMTrRec(2).TransName = CMTrRec(1).TransName
  CMTrRec(2).TransAcctNum = CMTrRec(1).TransAcctNum
  CMTrRec(2).TransTender = CMTrRec(1).TransTender
  CMTrRec(2).TransDetNum = NextTransRec
  CMTrRec(2).TransOperNum = OperNum
  CMTrRec(2).TransPad = ""
  CMTrRec(2).TransVoidNum = CmNum
  CMTrRec(1).ChkByte = Chr$(1)
  For cnt = 1 To 5  'recnum for categories
    CMTrRec(2).TransRevAmt(cnt) = CMTrRec(1).TransRevAmt(cnt)
  Next cnt
  For cnt = 6 To 10  'amount for categories
    CMTrRec(2).TransRevAmt(cnt) = -(Val(CMTrRec(1).TransRevAmt(cnt)))
  Next cnt
  For cnt = 11 To 12  'Penalty is 11 and Issue Fee is 12
    CMTrRec(2).TransRevAmt(cnt) = -(Val(CMTrRec(1).TransRevAmt(cnt)))
  Next cnt
  CHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Shared As CHandle Len = CMTrRecLen
  Put CHandle, (LOF(CHandle) / CMTrRecLen) + 1, CMTrRec(2)
  'CmNum = (LOF(CHandle) / CMTrRecLen) + 1
  Get CHandle, CmNum, CMTrRec(2)
    CMTrRec(2).TransVoidNum = (LOF(CHandle) / CMTrRecLen)
  Put CHandle, CmNum, CMTrRec(2)

  CmNum = (LOF(CHandle) / CMTrRecLen)
 ' Close CHandle


  GCustNum = 0
  BLLog "Void-BLOverPay Posted CM CUST:" + QPTrim$(fpAcct.Caption) + "  TRANS:" + Str$(NextTransRec&)
  CMLog "CMVoid-BLPay Posted CM CUST:" + QPTrim$(fpAcct.Caption) + "  TRANS:" + Str$(CmNum)
  Close
  'Call LoadMe
  'End If
  Exit Sub

ERRORSTUFF:
   Unload FrmShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLAdjustBal", "cmdPost_Click", Erl)
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
   ' ClearInUse PWcnt
   ' CitiTerminate

End Sub

Private Sub VoidTax()
  Dim ListFile As Integer, CHandle As Integer
  Dim PayFileName As String, UBPayRecLen As Integer
  Dim NumOfRecs As Long, CMTrRecLen As Integer
  Dim cnt As Integer, TRHandle As Integer, TrNumRecs As Long
  ReDim UBPaymentRec(1) As UBPaymentRecType
  Oper$ = QPTrim$(lblOperator.Caption)
  PayFileName$ = "C:\CPWork\CMPAY" + Oper$ + ".DAT"
  UBPayRecLen = Len(UBPaymentRec(1))
  ReDim CMTrRec(1) As CMTransRecType            ' open transaction file
  CMTrRecLen = Len(CMTrRec(1))
  TRHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Access Read Write Shared As TRHandle Len = CMTrRecLen
  TrNumRecs& = LOF(TRHandle) \ CMTrRecLen
  Get TRHandle, CmNum, CMTrRec(1)
  Close TRHandle
    For cnt = 1 To 15
        ' If There Is an Amount in Misc Rev 1-5 then get code record number from 6-10
      If CMTrRec(1).TransRevAmt(cnt) <> 0 Then
        UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = CMTrRec(1).TransRevAmt(cnt)
      Else
        UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = 0
      End If
    Next cnt
    For cnt = 1 To 15
      UBPaymentRec(1).PaidOwed(cnt).AMTOWE1 = 0
    Next
  UBPaymentRec(1).OperNum = QPTrim(lblOperator.Caption)
  UBPaymentRec(1).payDate = Date2Num(txtPaymentDate)
  UBPaymentRec(1).CustAcct = QPTrim(fpAcct)
  UBPaymentRec(1).CustName = QPTrim(fptxtName)
  UBPaymentRec(1).CustAddr = QPTrim(fptxtAddress)
  'UBPaymentRec(1).CUSTCMNT = 'QPTrim(Label4.Caption)
  'UBPaymentRec(1).TaxExempt = QPTrim(fptaxexmpt)
  UBPaymentRec(1).AmtOwed = fpAmtOwed
  UBPaymentRec(1).TenderTY = QPTrim$(fpTenderType.Caption)
  UBPaymentRec(1).CashAmt = fpCashAmt
  UBPaymentRec(1).ChkAmt = fpChkChgAmt
  UBPaymentRec(1).AmtRecd = fpTotReceived
  UBPaymentRec(1).Change = fpChange
  UBPaymentRec(1).Desc = QPTrim(fpDesc)
  UBPaymentRec(1).TotOwed = fpAmtOwed
  UBPaymentRec(1).AmtPaid = fpTotAmt
  'UBPaymentRec(1).Status = QPTrim(fpstatus)
  ListFile = FreeFile
  Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
    Put #ListFile, 1, UBPaymentRec(1)
    EditFlag = False
  CMLog "Oper:" + Oper$ + " Updated TempFile for Void Tax Pay"
  

  ReDim CMTrRec(1) As CMTransRecType
  CMTrRecLen = Len(CMTrRec(1))
  CHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Access Read Write Shared As CHandle Len = CMTrRecLen
  'CmNum = (LOF(CHandle) \ CMTrRecLen) + 1
  CMTrRec(1).TransDate = UBPaymentRec(1).payDate
  CMTrRec(1).TransAmount = -(Val(UBPaymentRec(1).AmtPaid))
  CMTrRec(1).TransCash = -(Val(UBPaymentRec(1).CashAmt))
  CMTrRec(1).TransCheck = -(Val(UBPaymentRec(1).ChkAmt))
  CMTrRec(1).TransAmtOwed = -(Val(UBPaymentRec(1).TotOwed))
  If Len(QPTrim$(UBPaymentRec(1).Desc)) = 0 Then
    CMTrRec(1).TransDesc = "Void Tax Payment"
  Else
    CMTrRec(1).TransDesc = (QPTrim$(UBPaymentRec(1).Desc))
  End If
  CMTrRec(1).TransSource = 231
  CMTrRec(1).TransName = UBPaymentRec(1).CustName
  CMTrRec(1).TransAcctNum = UBPaymentRec(1).CustAcct
  CMTrRec(1).TransDetNum = CmNum
  CMTrRec(1).TransOperNum = OperNum
  Select Case QPTrim(UBPaymentRec(1).TenderTY)
    Case "Cash":
      CMTrRec(1).TransTender = 1
    Case "Check":
      CMTrRec(1).TransTender = 2
    Case "Cash & Check":
      CMTrRec(1).TransTender = 3
    Case "Charge":
      CMTrRec(1).TransTender = 4
    Case Else:
      '
  End Select
  CMTrRec(1).ChkByte = Chr$(1)
  CMTrRec(1).TransPad = ""
  CMTrRec(1).TransVoidNum = CmNum
  For cnt = 1 To 15
    CMTrRec(1).TransRevAmt(cnt) = -(Val(UBPaymentRec(1).PaidOwed(cnt).AMTPD1))
  Next cnt


  Put CHandle, (LOF(CHandle) / CMTrRecLen) + 1, CMTrRec(1)
  Get CHandle, CmNum, CMTrRec(1)
  CMTrRec(1).TransVoidNum = (LOF(CHandle) / CMTrRecLen)
  Put CHandle, CmNum, CMTrRec(1)
  
  CmNum = (LOF(CHandle) / CMTrRecLen)

  Close
  CMLog "CMVoid-Tax Posted" + "  TRANS:" + Str$(CmNum&)
  'ClearScn
End Sub
Private Sub VoidVATax()
  Dim ListFile As Integer, CHandle As Integer
  Dim PayFileName As String, UBPayRecLen As Integer
  Dim NumOfRecs As Long, CMTrRecLen As Integer
  Dim cnt As Integer, TRHandle As Integer, TrNumRecs As Long
  ReDim UBPaymentRec(1) As UBPaymentRecType
  Oper$ = QPTrim$(lblOperator.Caption)
  PayFileName$ = "C:\CPWork\CMPAY" + Oper$ + ".DAT"
  UBPayRecLen = Len(UBPaymentRec(1))
  ReDim CMTrRec(1) As CMTransRecType            ' open transaction file
  CMTrRecLen = Len(CMTrRec(1))
  TRHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Access Read Write Shared As TRHandle Len = CMTrRecLen
  TrNumRecs& = LOF(TRHandle) \ CMTrRecLen
  Get TRHandle, CmNum, CMTrRec(1)
  Close TRHandle
    For cnt = 1 To 15
        ' If There Is an Amount in Misc Rev 1-5 then get code record number from 6-10
      If CMTrRec(1).TransRevAmt(cnt) <> 0 Then
        UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = CMTrRec(1).TransRevAmt(cnt)
      Else
        UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = 0
      End If
    Next cnt
    For cnt = 1 To 15
      UBPaymentRec(1).PaidOwed(cnt).AMTOWE1 = 0
    Next
  UBPaymentRec(1).OperNum = QPTrim(lblOperator.Caption)
  UBPaymentRec(1).payDate = Date2Num(txtPaymentDate)
  UBPaymentRec(1).CustAcct = QPTrim(fpAcct)
  UBPaymentRec(1).CustName = QPTrim(fptxtName)
  UBPaymentRec(1).CustAddr = QPTrim(fptxtAddress)
  'UBPaymentRec(1).CUSTCMNT = 'QPTrim(Label4.Caption)
  'UBPaymentRec(1).TaxExempt = QPTrim(fptaxexmpt)
  UBPaymentRec(1).AmtOwed = fpAmtOwed
  UBPaymentRec(1).TenderTY = QPTrim$(fpTenderType.Caption)
  UBPaymentRec(1).CashAmt = fpCashAmt
  UBPaymentRec(1).ChkAmt = fpChkChgAmt
  UBPaymentRec(1).AmtRecd = fpTotReceived
  UBPaymentRec(1).Change = fpChange
  UBPaymentRec(1).Desc = QPTrim(fpDesc)
  UBPaymentRec(1).TotOwed = fpAmtOwed
  UBPaymentRec(1).AmtPaid = fpTotAmt
  'UBPaymentRec(1).Status = QPTrim(fpstatus)
  ListFile = FreeFile
  Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
    Put #ListFile, 1, UBPaymentRec(1)
    EditFlag = False
  CMLog "Oper:" + Oper$ + " Updated TempFile for Void Tax Pay"
  

  ReDim CMTrRec(1) As CMTransRecType
  CMTrRecLen = Len(CMTrRec(1))
  CHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Access Read Write Shared As CHandle Len = CMTrRecLen
  'CmNum = (LOF(CHandle) \ CMTrRecLen) + 1
  CMTrRec(1).TransDate = UBPaymentRec(1).payDate
  CMTrRec(1).TransAmount = -(Val(UBPaymentRec(1).AmtPaid))
  CMTrRec(1).TransCash = -(Val(UBPaymentRec(1).CashAmt))
  CMTrRec(1).TransCheck = -(Val(UBPaymentRec(1).ChkAmt))
  CMTrRec(1).TransAmtOwed = -(Val(UBPaymentRec(1).TotOwed))
  If Len(QPTrim$(UBPaymentRec(1).Desc)) = 0 Then
    CMTrRec(1).TransDesc = "Void Tax Payment"
  Else
    CMTrRec(1).TransDesc = (QPTrim$(UBPaymentRec(1).Desc))
  End If
  If TrTypeNum = 161 Then
    CMTrRec(1).TransSource = 261
  ElseIf TrTypeNum = 171 Then
    CMTrRec(1).TransSource = 271
  End If
  CMTrRec(1).TransName = UBPaymentRec(1).CustName
  CMTrRec(1).TransAcctNum = UBPaymentRec(1).CustAcct
  CMTrRec(1).TransDetNum = CmNum
  CMTrRec(1).TransOperNum = OperNum
  Select Case QPTrim(UBPaymentRec(1).TenderTY)
    Case "Cash":
      CMTrRec(1).TransTender = 1
    Case "Check":
      CMTrRec(1).TransTender = 2
    Case "Cash & Check":
      CMTrRec(1).TransTender = 3
    Case "Charge":
      CMTrRec(1).TransTender = 4
    Case Else:
      '
  End Select
  CMTrRec(1).ChkByte = Chr$(1)
  CMTrRec(1).TransPad = ""
  CMTrRec(1).TransVoidNum = CmNum
  For cnt = 1 To 15
    CMTrRec(1).TransRevAmt(cnt) = -(Val(UBPaymentRec(1).PaidOwed(cnt).AMTPD1))
  Next cnt


  Put CHandle, (LOF(CHandle) / CMTrRecLen) + 1, CMTrRec(1)
  Get CHandle, CmNum, CMTrRec(1)
  CMTrRec(1).TransVoidNum = (LOF(CHandle) / CMTrRecLen)
  Put CHandle, CmNum, CMTrRec(1)
  
  CmNum = (LOF(CHandle) / CMTrRecLen)

  Close
  CMLog "CMVoid-Tax Posted" + "  TRANS:" + Str$(CmNum&)
  'ClearScn
End Sub

Private Sub VoidDecal()
  Dim DCFile As Integer, NumOfDCRecs As Long
  Dim DCTransRecLen As Integer, TRHandle As Integer, CHandle As Integer
  Dim DCTransFile As Integer, NumOfTransRecs As Long, NextTransRec As Long
  Dim Prev As Long, DCVehReclen As Integer, DCvFile As Integer, VehRecord As Long
  Dim NumOfVRecs As Long, UseDate As Integer
  ReDim DCCustREc(1) As DCCustRecType
  OpenDCCustFile NumOfDCRecs, DCFile
  ReDim DCVRec(1) As DCVehType
  ReDim DCTransRec(1 To 2) As DCTransRecType
  ReDim CMTrRec(1) As CMTransRecType            ' open transaction file
  CMTrRecLen = Len(CMTrRec(1))
  TRHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Access Read Write Shared As TRHandle Len = CMTrRecLen
  Get TRHandle, CmNum, CMTrRec(1)
  Close TRHandle
  If NoDoModTrans = False Then
    DCTransRecLen = Len(DCTransRec(1))
    DCTransFile = FreeFile
    Open "DCTrans.DAT" For Random Access Read Write Shared As DCTransFile Len = DCTransRecLen
    NumOfTransRecs = LOF(DCTransFile) \ DCTransRecLen
    NextTransRec = NumOfTransRecs + 1

    Get DCTransFile, CMTrRec(1).TransDetNum, DCTransRec(1)
    If DCTransRec(1).TransAmount >= 0 And Val(DCTransRec(1).CustomerNumber) > 0 Then
      DCTransRec(1).VoidFlag = "Y"
      VehRecord = DCTransRec(1).VehRecord
      Put DCTransFile, CMTrRec(1).TransDetNum, DCTransRec(1)
      If VehRecord > 0 Then GoSub UpdateVehRecord
      Get DCFile, Val(DCTransRec(1).CustomerNumber), DCCustREc(1)
      ' Post Void Charge First to Offset Void Payment of Decal

      DCTransRec(2).CustomerNumber = DCTransRec(1).CustomerNumber
      DCTransRec(2).TransDate = Date2Num(txtPaymentDate)
      DCTransRec(2).TransAmount = DCTransRec(1).TransAmount
      DCTransRec(2).TransType = 3               ' Type 3 = Void Charge
      DCTransRec(2).TRVinDesc = DCTransRec(1).TRVinDesc
      DCTransRec(2).TransTender = DCTransRec(1).TransTender
      DCTransRec(2).CashAmount = DCTransRec(1).CashAmount
      DCTransRec(2).ChkAmount = DCTransRec(1).ChkAmount
      DCTransRec(2).BalanceAfterTrans = DCCustREc(1).AcctBal - DCTransRec(1).TransAmount
      DCTransRec(2).makemodel = DCTransRec(1).makemodel
      DCTransRec(2).StateTag = DCTransRec(1).StateTag
      DCTransRec(2).Sticker = DCTransRec(1).Sticker
      DCTransRec(2).ExpireDate = DCTransRec(1).ExpireDate
      DCTransRec(2).ExtraDesc = "CMVoid " + QPTrim(fpDesc)
      DCTransRec(2).ExtraRoom = ""
      DCTransRec(2).NextTrans = 0
      DCTransRec(2).OperNum = OperNum
      DCTransRec(2).GLInterfaced = "Y"
      DCTransRec(2).DecalCat = DCTransRec(1).DecalCat    'Dale Need This in his stuff
      DCTransRec(2).ChkByte = Chr$(1)
      DCTransRec(2).VoidFlag = "N"
      DCTransRec(2).VehRecord = DCTransRec(1).VehRecord
'      DCTransFile = FreeFile
'      Open "DCTrans.DAT" For Random Access Read Write Shared As DCTransFile Len = DCTransRecLen
'      NumOfTransRecs = LOF(DCTransFile) \ DCTransRecLen
'      NextTransRec = NumOfTransRecs + 1

      Put DCTransFile, NextTransRec, DCTransRec(2)
      
      Get DCFile, Val(DCTransRec(1).CustomerNumber), DCCustREc(1)
      DCCustREc(1).AcctBal = DCCustREc(1).AcctBal - DCTransRec(1).TransAmount
      Put DCFile, Val(DCTransRec(1).CustomerNumber), DCCustREc(1)
      If DCCustREc(1).FirstTrans = 0 Then
        DCCustREc(1).FirstTrans = NextTransRec
        DCCustREc(1).LastTrans = NextTransRec
        Put DCFile, Val(DCTransRec(1).CustomerNumber), DCCustREc(1)
      Else
        Prev = DCCustREc(1).LastTrans
        DCCustREc(1).LastTrans = NextTransRec
        Put DCFile, Val(DCTransRec(1).CustomerNumber), DCCustREc(1)
        Get DCTransFile, Prev, DCTransRec(1)
        DCTransRec(1).NextTrans = NextTransRec
        Put DCTransFile, Prev, DCTransRec(1)
      End If
      Close DCTransFile
      DCTransFile = FreeFile
      Open "DCTrans.DAT" For Random Access Read Write Shared As DCTransFile Len = DCTransRecLen
      NumOfTransRecs = LOF(DCTransFile) \ DCTransRecLen
      NextTransRec = NumOfTransRecs + 1
      Get DCTransFile, NumOfTransRecs, DCTransRec(1)
      ' Post Transaction Record First
      DCTransRec(2).CustomerNumber = DCTransRec(1).CustomerNumber
      DCTransRec(2).TransDate = Date2Num(txtPaymentDate)
      DCTransRec(2).TransAmount = DCTransRec(1).TransAmount
      DCTransRec(2).TransType = 4               ' Type 4 = Void Payment
      DCTransRec(2).TRVinDesc = DCTransRec(1).TRVinDesc
      DCTransRec(2).TransTender = DCTransRec(1).TransTender
      DCTransRec(2).CashAmount = DCTransRec(1).CashAmount
      DCTransRec(2).ChkAmount = DCTransRec(1).ChkAmount
      DCTransRec(2).BalanceAfterTrans = DCCustREc(1).AcctBal + DCTransRec(1).TransAmount
      DCTransRec(2).ExtraDesc = "CMVoid " + QPTrim(fpDesc)
      DCTransRec(2).ExtraRoom = ""
      DCTransRec(2).NextTrans = 0
      DCTransRec(2).GLInterfaced = "N"
      DCTransRec(2).OperNum = OperNum
      DCTransRec(2).DecalCat = DCTransRec(1).DecalCat
      DCTransRec(2).ChkByte = Chr$(1)
      DCTransRec(2).VoidFlag = "N"
      
      Put DCTransFile, NextTransRec, DCTransRec(2)
      
      Get DCFile, Val(DCTransRec(1).CustomerNumber), DCCustREc(1)
      DCCustREc(1).AcctBal = DCCustREc(1).AcctBal + DCTransRec(1).TransAmount
      DCCustREc(1).LICENSE = ""
      Put DCFile, Val(DCTransRec(1).CustomerNumber), DCCustREc(1)

      If DCCustREc(1).FirstTrans = 0 Then
        DCCustREc(1).FirstTrans = NextTransRec
        DCCustREc(1).LastTrans = NextTransRec
        Put DCFile, Val(DCTransRec(1).CustomerNumber), DCCustREc(1)
      Else
        Prev = DCCustREc(1).LastTrans
        DCCustREc(1).LastTrans = NextTransRec
        Put DCFile, Val(DCTransRec(1).CustomerNumber), DCCustREc(1)
        Get DCTransFile, Prev, DCTransRec(1)
        DCTransRec(1).NextTrans = NextTransRec
        Put DCTransFile, Prev, DCTransRec(1)
      End If
      Close DCTransFile
   
    End If
  ' Show All Posted
  DCLog "Voided thru CM:" + Str$(CMTrRec(1).TransDetNum) + " by-" + Str(OperNum)
 ' MsgBox "Void Complete", vbOKOnly, "Complete"
  End If
  ReDim CMTrRec(1 To 2) As CMTransRecType
  CMTrRecLen = Len(CMTrRec(1))
  CHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Access Read Write Shared As CHandle Len = CMTrRecLen
      Get CHandle, CmNum, CMTrRec(1)
  Close CHandle

  'CmNum = (LOF(CHandle) \ CMTrRecLen) + 1
  CMTrRec(2).TransDate = Date2Num(txtPaymentDate)
  CMTrRec(2).TransAmount = -(Val(CMTrRec(1).TransAmount))
  CMTrRec(2).TransCash = -(Val(CMTrRec(1).TransCash))
  CMTrRec(2).TransCheck = -(Val(CMTrRec(1).TransCheck))
  CMTrRec(2).TransAmtOwed = -(Val(CMTrRec(1).TransAmtOwed))
  CMTrRec(2).TransDesc = "Void DC" + QPTrim(fpDesc)
  CMTrRec(2).TransSource = 251
  CMTrRec(2).TransName = QPTrim$(CMTrRec(1).TransName)
  CMTrRec(2).TransAcctNum = CMTrRec(1).TransAcctNum
  CMTrRec(2).TransDetNum = NextTransRec
  CMTrRec(2).TransOperNum = OperNum
  CMTrRec(2).TransTender = CMTrRec(1).TransTender
  CMTrRec(2).ChkByte = Chr$(1)
  CMTrRec(2).TransPad = ""
  CMTrRec(2).TransVoidNum = CmNum
  CMTrRec(2).TransRevAmt(1) = CMTrRec(1).TransRevAmt(1)
  CMTrRec(2).TransRevAmt(2) = CMTrRec(1).TransRevAmt(2)
  
  CHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Shared As CHandle Len = CMTrRecLen

  Put CHandle, (LOF(CHandle) / CMTrRecLen) + 1, CMTrRec(2)
  Get CHandle, CmNum, CMTrRec(1)
  CMTrRec(1).TransVoidNum = (LOF(CHandle) / CMTrRecLen)
  Put CHandle, CmNum, CMTrRec(1)
  
  CmNum = (LOF(CHandle) / CMTrRecLen)

  Close
  CMLog "CMVoid-Decal Posted" + "  TRANS:" + Str$(CmNum&)
  Exit Sub
UpdateVehRecord:
  DCVehReclen = Len(DCVRec(1))
  DCvFile = FreeFile
  Open "DCVEH.DAT" For Random Access Read Write Shared As DCvFile Len = DCVehReclen
  NumOfVRecs = LOF(DCvFile) \ DCVehReclen
  If VehRecord <= 0 Or VehRecord > NumOfVRecs Then Close DCvFile: Return
  Get DCvFile, VehRecord, DCVRec(1)
  DCVRec(1).ExpireDate = Date2Num("01/01/1980")
  DCVRec(1).Sticker = "VOID"
  Put DCvFile, VehRecord, DCVRec(1)
  Close DCvFile
Return
End Sub


Private Sub PrintReceipt()
  Dim ListFile As Integer, PayFileName As String, UBPayRecLen As Integer
  Dim RecptNum As Long, RHandle As Integer, PayRecpName As String
  Dim CutPaper As String, PostDate As String, RevCnt As Integer
  Dim NumofRevs As Integer, RecpRev As String
'  ReDim UBPaymentRec(1) As UBPaymentRecType
'  ReDim Preserve RevText$(1 To MaxRevsCnt)
  RecpRev$ = Space$(15)
  CutPaper$ = Chr$(29) + Chr$(86) + Chr$(66) + Chr$(64)
   If InStr(TownName$, "Dobson") > 0 Then CutPaper$ = Chr$(27) + Chr$(100)
'  UBPayRecLen = Len(UBPaymentRec(1))
'  PayFileName$ = "C:\CMPAY" + Oper$ + ".DAT"
  PayRecpName$ = "c:\CPWORK\CMRCP" + Oper$ + ".RPT"
  PostDate$ = txtPaymentDate
  ListFile = FreeFile
'  Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
  'RecptNum& = LOF(ListFile) / UBPayRecLen
'  Get #ListFile, 1, UBPaymentRec(1)
' Close
  NumofRevs = MaxRevsCnt
  RHandle = FreeFile
  Open PayRecpName$ For Output As RHandle
  If CntrlDef = 1 Then
    Print #RHandle, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
    Print #RHandle, Chr$(7)
  End If
  Print #RHandle, TownName$
  Print #RHandle, "CM- VOID FOR RECEIPT# " + fpReceiptNo.Caption
  Print #RHandle, "Date: "; PostDate$
  Print #RHandle, "Time: "; Time
  Print #RHandle,
  Print #RHandle, "CUSTOMER NAME & DESC. OF PAYMENT"
  Print #RHandle, QPTrim(fptxtName)
  Print #RHandle, QPTrim(fptxtAddress)
  Print #RHandle, QPTrim(fpDesc)
  Print #RHandle, "Acct. No. "; QPTrim(fpAcct)
  Print #RHandle,
  Print #RHandle, QPTrim$(fpTenderType.Caption)
  Print #RHandle,
  Print #RHandle, "Total Owed: "; Using("$##,###,###.##", -(fpAmtOwed))
  Print #RHandle, "Total Paid: "; Using("$##,###,###.##", -(fpTotReceived))
  Print #RHandle, "Change Due: "; Using("$##,###,###.##", -(fpChange))
  Print #RHandle, '"   Balance: "; Using("$##,###.##", (-(Val(UBPaymentRec(1).TOTOWED))) - (-(Val(UBPaymentRec(1).AMTPAID))))
  If TrTypeNum <> 1 Then
  For RevCnt = 1 To 15    ' NumOfRevs
    If Revs(RevCnt - 1).Visible = True Then
      If Val(Revs(RevCnt - 1)) <> 0 Then
        Print #RHandle, QPTrim$(fpDetDesc(RevCnt - 1).Caption) + "  "; Using("$########.##", (-(Revs(RevCnt - 1))))
      End If
    End If
  Next
   If TrTypeNum = 151 Then
     Print #RHandle, fpDetDesc(1).Caption
     Print #RHandle, fpDetDesc(2).Caption
     Print #RHandle, fpDetDesc(3).Caption
     Print #RHandle, fpDetDesc(4).Caption
     Print #RHandle, fpDetDesc(5).Caption
    End If
  Else
    For RevCnt = 1 To 10    ' NumOfRevs
    If Revs(RevCnt - 1).Visible = True Then
      If Val(Revs(RevCnt - 1)) <> 0 Then
        Print #RHandle, Mid$(fpDetDesc(RevCnt - 1).Caption, 11, 20) + "  "; Using("$########.##", (-(Revs(RevCnt - 1))))
      End If
    End If
    Next
    For RevCnt = 11 To 15    ' NumOfRevs
    If Revs(RevCnt - 1).Visible = True Then
      If Val(Revs(RevCnt - 1)) <> 0 Then
        Print #RHandle, QPTrim$(fpDetDesc(RevCnt - 1).Caption) + "  "; Using("$########.##", (-(Revs(RevCnt - 1))))
      End If
    End If
    Next
  End If
  Print #RHandle,
  Print #RHandle, lblSource.Caption
  Print #RHandle, "Operator: "; OperNum
  Print #RHandle, "Receipt#: "; Using("######", CmNum&)
  Print #RHandle,
  Print #RHandle, "        V O I D     V O I D  "
  Print #RHandle,
  If NoDoModTrans = True Then
    Print #RHandle, "Only The CM Transaction Created."
  Else
    Print #RHandle,
  End If
  Print #RHandle,
'  Print #RHandle,
'  Print #RHandle,
  If CntrlDef = 1 Then
    Print #RHandle, CutPaper$
  Else
    Print #RHandle,
    Print #RHandle,
    Print #RHandle,
  End If
  Close RHandle

  'Shell$ = "type " + PayRecpName$ + " > com2:"
  'SHELL Shell$
  If CntrlDef = 1 Then
    fpcmdDrawer_Click
  End If
  'PrintRptFile Header$, PayRecpName$, RecpPort, RetCode%, 5
  Dim RptHandle As Integer, LPTHandle As Integer
  Dim RptA As Integer, LPTA As Integer, ToPrintA As String
  Dim ToPrint As String, CopyLoop As Integer, DefPrinter As String
  On Error GoTo Cancel
  'Printer.Print
'''  to strReportFile DefPrinter'[ADDITIVE] | PortName]
10:
  DefPrinter = RecpPort '"LPT" + QPTrim$(Str$(RecpPort)) + ":"
20:
 ' MsgBox "Printer -" + DefPrinter, vbOKOnly
  
  For CopyLoop = 1 To 1 'Copies
    LPTHandle = FreeFile
    Open DefPrinter For Output As LPTHandle
    RptHandle = FreeFile
30:
    Open PayRecpName$ For Input As RptHandle
40:
    Do
      If frmPrint.cmdCancel = False Then
45:
        Line Input #RptHandle, ToPrint$
        
        ToPrint$ = RTrim$(ToPrint$)
        Print #LPTHandle, ToPrint$
      Else
50:
        Exit Do
        'Printer.EndDoc
      End If
    Loop Until eof(RptHandle)
60:
    Close RptHandle
62:
    Close LPTHandle
65:
    Next CopyLoop
68:
 Printer.EndDoc
70:
 CMLog "Oper: " + Oper$ + " Print receipt Acct:" + QPTrim$(fpAcct.Caption)
 KillFile PayRecpName$
' KillFile PayFileName$
80:
  Exit Sub
Cancel:
  If Err > 0 Then
    CMLog "CMRecp-Error Code Was " + DefPrinter + Err.Description + Str$(Err) + " (PrintWSet - Line:" & Erl & ")"
    MsgBox "Error Code Was " + DefPrinter + Err.Description + Str$(Err) + " (PrintWSet - Line:" & Erl & ")"
  End If
  Close
  Exit Sub

  
End Sub

Private Sub Timer1_Timer()
  Dim BkColor As Integer
  Static tog As Boolean
  tog = Not tog
  If tog Then
    Label2(1).BackColor = &HFFFF&
    Label2(2).BackColor = &HFFFF&
    Label2(2).ForeColor = &HFF&
  Else
    Label2(1).BackColor = &HFF&
    Label2(2).BackColor = &HFF&
    Label2(2).ForeColor = &HFFFF&
  End If
End Sub

