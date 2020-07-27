VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmBLAdjustBal 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Adjust Customer Balance"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLAdjustBal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H008F8265&
      BorderStyle     =   0  'None
      Height          =   2124
      Left            =   7776
      TabIndex        =   52
      Top             =   1968
      Width           =   3420
      Begin VB.OptionButton optOverPay 
         BackColor       =   &H008F8265&
         Caption         =   "Payment Adjustment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   192
         MaskColor       =   &H008F8265&
         TabIndex        =   3
         Tag             =   $"frmBLAdjustBal.frx":08CA
         Top             =   1584
         Width           =   2604
      End
      Begin VB.OptionButton optOverBill 
         BackColor       =   &H008F8265&
         Caption         =   "Billing Downward Adjustment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   192
         MaskColor       =   &H008F8265&
         TabIndex        =   1
         Tag             =   "If a customer has been over-billed then use this option to DECREASE a balance based on the amounts entered below."
         Top             =   240
         Width           =   3180
      End
      Begin VB.OptionButton optUnderBill 
         BackColor       =   &H008F8265&
         Caption         =   "Billing Upward Adjustment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   192
         MaskColor       =   &H008F8265&
         TabIndex        =   2
         Tag             =   $"frmBLAdjustBal.frx":09D7
         Top             =   912
         Width           =   3180
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCustList 
      Height          =   405
      Left            =   5280
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   $"frmBLAdjustBal.frx":0A65
      Top             =   1965
      Width           =   2025
      _Version        =   131072
      _ExtentX        =   3572
      _ExtentY        =   714
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
      ButtonDesigner  =   "frmBLAdjustBal.frx":0B6B
   End
   Begin EditLib.fpText fptxtName 
      Height          =   396
      Left            =   1884
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "This field contains the customer's business name. It cannot be edited."
      Top             =   2388
      Width           =   5388
      _Version        =   196608
      _ExtentX        =   9504
      _ExtentY        =   698
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
      ButtonStyle     =   0
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
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtAddress 
      Height          =   396
      Left            =   1884
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "This field contains the primary address of this business. This field cannot be edited."
      Top             =   2820
      Width           =   5388
      _Version        =   196608
      _ExtentX        =   9504
      _ExtentY        =   698
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
      ButtonStyle     =   0
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
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtCity 
      Height          =   396
      Left            =   1884
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "This field contains the name of the city where this business receives mail. This field cannot be edited."
      Top             =   3252
      Width           =   5388
      _Version        =   196608
      _ExtentX        =   9504
      _ExtentY        =   698
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
      ButtonStyle     =   0
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
      MaxLength       =   20
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtAccount 
      Height          =   396
      Left            =   1872
      TabIndex        =   0
      Tag             =   $"frmBLAdjustBal.frx":0D4F
      Top             =   1956
      Width           =   1308
      _Version        =   196608
      _ExtentX        =   2307
      _ExtentY        =   698
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
      ThreeDInsideShadowColor=   -2147483646
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
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
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtState 
      Height          =   396
      Left            =   1884
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "This field contains the state where this business receives mail. This field cannot be edited."
      Top             =   3684
      Width           =   540
      _Version        =   196608
      _ExtentX        =   952
      _ExtentY        =   698
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      CharValidationText=   "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z"
      MaxLength       =   2
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpMask fptxtZip 
      Height          =   396
      Left            =   5484
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "This field contains the postal code for this business. This field cannot be edited."
      Top             =   3684
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
      _ExtentY        =   698
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
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   "#####-####"
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   0   'False
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime fptxtToday 
      Height          =   348
      Left            =   5568
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "This field defaults to today's date and is not editable."
      ToolTipText     =   "This date, themost current depreciation date, is automatically calculated and cannot be edited."
      Top             =   1200
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
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
      ControlType     =   1
      Text            =   "11/20/2002"
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
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpcurrLicBal 
      Height          =   348
      Left            =   7692
      TabIndex        =   25
      TabStop         =   0   'False
      Tag             =   " This field displays the current total outstanding license balance for all license categories. It cannot be edited."
      Top             =   7200
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      ForeColor       =   -2147483645
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpcurrLicAmt 
      Height          =   348
      Left            =   9504
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   $"frmBLAdjustBal.frx":0EE3
      Top             =   7200
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      ForeColor       =   -2147483645
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   0
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
      InvalidColor    =   -2147483633
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpcurrPenBal 
      Height          =   348
      Left            =   912
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "This field displays the total outstanding penalty balance for this customer. This field is not editable."
      Top             =   5712
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      BackColor       =   16777215
      ForeColor       =   -2147483645
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpcurrPenAmt 
      Height          =   348
      Left            =   2880
      TabIndex        =   4
      Tag             =   "Enter the amount here that will either increase or decrease the current outstanding penalty balance. "
      Top             =   5712
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpcurrTotAmt 
      Height          =   348
      Left            =   2880
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   $"frmBLAdjustBal.frx":0F71
      Top             =   4704
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      ForeColor       =   -2147483645
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpcurrTotBal 
      Height          =   348
      Left            =   912
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "This field displays the total outstanding balance for this customer. This field cannot be edited."
      Top             =   4704
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      ForeColor       =   -2147483645
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   348
      Left            =   1104
      TabIndex        =   6
      Tag             =   $"frmBLAdjustBal.frx":100E
      Top             =   7200
      Width           =   3756
      _Version        =   196608
      _ExtentX        =   6625
      _ExtentY        =   614
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
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
      MaxLength       =   20
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   540
      Left            =   9264
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   $"frmBLAdjustBal.frx":10EA
      Top             =   7824
      Width           =   1956
      _Version        =   131072
      _ExtentX        =   3450
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmBLAdjustBal.frx":1172
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   540
      Left            =   7248
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   $"frmBLAdjustBal.frx":1350
      Top             =   7824
      Width           =   1860
      _Version        =   131072
      _ExtentX        =   3281
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmBLAdjustBal.frx":1445
   End
   Begin EditLib.fpCurrency fpcurrLicBalDet 
      Height          =   348
      Index           =   0
      Left            =   7680
      TabIndex        =   47
      TabStop         =   0   'False
      Tag             =   "This field displays the total outstanding balance for license category 1. It cannot be edited."
      Top             =   4752
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      ForeColor       =   -2147483645
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpcurrLicBalDet 
      Height          =   348
      Index           =   1
      Left            =   7680
      TabIndex        =   48
      TabStop         =   0   'False
      Tag             =   "This field displays the total outstanding balance for license category 2. It cannot be edited."
      Top             =   5184
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      ForeColor       =   -2147483645
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpcurrLicBalDet 
      Height          =   348
      Index           =   2
      Left            =   7680
      TabIndex        =   49
      TabStop         =   0   'False
      Tag             =   "This field displays the total outstanding balance for license category 3. It cannot be edited."
      Top             =   5616
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      ForeColor       =   -2147483645
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpcurrLicBalDet 
      Height          =   348
      Index           =   3
      Left            =   7680
      TabIndex        =   50
      TabStop         =   0   'False
      Tag             =   "This field displays the total outstanding balance for license category 4. It cannot be edited."
      Top             =   6048
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      ForeColor       =   -2147483645
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpcurrLicBalDet 
      Height          =   348
      Index           =   4
      Left            =   7680
      TabIndex        =   51
      TabStop         =   0   'False
      Tag             =   "This field displays the total outstanding balance for license category 5. It cannot be edited."
      Top             =   6480
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      ForeColor       =   -2147483645
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpcurrLicAmtDet 
      Height          =   348
      Index           =   0
      Left            =   9504
      TabIndex        =   7
      Tag             =   "Enter the desired adjustment amount for license category 1 here. "
      Top             =   4752
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpcurrLicAmtDet 
      Height          =   348
      Index           =   1
      Left            =   9504
      TabIndex        =   8
      Tag             =   "Enter the desired adjustment amount for license category 2 here. "
      Top             =   5184
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpcurrLicAmtDet 
      Height          =   348
      Index           =   2
      Left            =   9504
      TabIndex        =   9
      Tag             =   "Enter the desired adjustment amount for license category 3 here. "
      Top             =   5616
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpcurrLicAmtDet 
      Height          =   348
      Index           =   3
      Left            =   9504
      TabIndex        =   10
      Tag             =   "Enter the desired adjustment amount for license category 4 here. "
      Top             =   6048
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpcurrLicAmtDet 
      Height          =   348
      Index           =   4
      Left            =   9504
      TabIndex        =   11
      Tag             =   "Enter the desired adjustment amount for license category 5 here. "
      Top             =   6480
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpcurrIssFeeBal 
      Height          =   348
      Left            =   912
      TabIndex        =   55
      TabStop         =   0   'False
      Tag             =   "This field displays the total outstanding issuance fee balance. This field is not editable."
      Top             =   6480
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      BackColor       =   16777215
      ForeColor       =   -2147483645
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
      ControlType     =   1
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpcurrIssFeeAmt 
      Height          =   348
      Left            =   2880
      TabIndex        =   5
      Tag             =   "The amount entered here will affect the issuance balance either increasing or decreasing it depending on the adjustment type."
      Top             =   6480
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   614
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   540
      Left            =   4992
      TabIndex        =   57
      TabStop         =   0   'False
      Tag             =   $"frmBLAdjustBal.frx":1621
      Top             =   7824
      Width           =   2100
      _Version        =   131072
      _ExtentX        =   3704
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmBLAdjustBal.frx":16F1
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   348
      Left            =   8928
      TabIndex        =   58
      Top             =   8400
      Width           =   540
      _Version        =   131072
      _ExtentX        =   952
      _ExtentY        =   614
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
      ForeColor       =   8454143
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
      MaxWidth        =   6000
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
   Begin fpBtnAtlLibCtl.fpBtn cmdTransHist 
      Height          =   405
      Left            =   3210
      TabIndex        =   62
      TabStop         =   0   'False
      Tag             =   $"frmBLAdjustBal.frx":18D4
      Top             =   1965
      Width           =   2040
      _Version        =   131072
      _ExtentX        =   3598
      _ExtentY        =   714
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
      ButtonDesigner  =   "frmBLAdjustBal.frx":19DA
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   288
      X2              =   11376
      Y1              =   4272
      Y2              =   4272
   End
   Begin VB.Line Line10 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   288
      X2              =   288
      Y1              =   1776
      Y2              =   7680
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Category Descriptions"
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
      Height          =   252
      Left            =   5424
      TabIndex        =   61
      Top             =   4368
      Width           =   2028
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total License Amounts"
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
      Height          =   252
      Left            =   5328
      TabIndex        =   60
      Top             =   7248
      Width           =   2076
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblBalloon 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "HELP BALLOONS ON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   4992
      TabIndex        =   59
      Top             =   8400
      Width           =   2100
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Where Am I"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   732
      Left            =   528
      TabIndex        =   56
      Top             =   7824
      Width           =   4332
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Iss Fee Adj Amount"
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
      Height          =   300
      Left            =   2688
      TabIndex        =   54
      Top             =   6192
      Width           =   1932
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Issuance Fee Balance"
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
      Height          =   300
      Left            =   672
      TabIndex        =   53
      Top             =   6192
      Width           =   1932
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   288
      X2              =   11376
      Y1              =   1776
      Y2              =   1776
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   7500
      X2              =   7500
      Y1              =   1776
      Y2              =   4272
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   5052
      X2              =   5052
      Y1              =   4272
      Y2              =   7680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   288
      X2              =   5280
      Y1              =   7056
      Y2              =   7056
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   11376
      X2              =   288
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   5280
      X2              =   11376
      Y1              =   7056
      Y2              =   7056
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   288
      X2              =   5040
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   11376
      X2              =   11376
      Y1              =   1776
      Y2              =   7680
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CatDesc5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   252
      Index           =   4
      Left            =   5184
      TabIndex        =   46
      Top             =   6576
      UseMnemonic     =   0   'False
      Width           =   2460
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CatDesc4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   252
      Index           =   3
      Left            =   5184
      TabIndex        =   45
      Top             =   6144
      UseMnemonic     =   0   'False
      Width           =   2460
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CatDesc3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   252
      Index           =   2
      Left            =   5184
      TabIndex        =   44
      Top             =   5712
      UseMnemonic     =   0   'False
      Width           =   2460
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CatDesc2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   252
      Index           =   1
      Left            =   5184
      TabIndex        =   43
      Top             =   5280
      UseMnemonic     =   0   'False
      Width           =   2460
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CatDesc1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   252
      Index           =   0
      Left            =   5184
      TabIndex        =   42
      Top             =   4848
      UseMnemonic     =   0   'False
      Width           =   2460
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty Adj Amount"
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
      Height          =   300
      Left            =   2736
      TabIndex        =   41
      Top             =   5424
      Width           =   1836
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Adj Amount"
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
      Height          =   252
      Left            =   2736
      TabIndex        =   40
      Top             =   4416
      Width           =   1788
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty Balance"
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
      Height          =   300
      Left            =   912
      TabIndex        =   39
      Top             =   5424
      Width           =   1548
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "License Balance"
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
      Height          =   252
      Left            =   7632
      TabIndex        =   38
      Top             =   4368
      Width           =   1644
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "License Adj Amount"
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
      Height          =   300
      Left            =   9360
      TabIndex        =   37
      Top             =   4368
      Width           =   1788
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Adjust Customer Balance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2940
      TabIndex        =   36
      Top             =   336
      Width           =   6012
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   300
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Type: Adjustment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   300
      Left            =   4320
      TabIndex        =   35
      Top             =   672
      Width           =   3276
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Account #:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   492
      TabIndex        =   34
      Top             =   2052
      Width           =   1212
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   576
      TabIndex        =   33
      Top             =   2916
      Width           =   1116
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   972
      TabIndex        =   32
      Top             =   3348
      Width           =   780
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   684
      TabIndex        =   31
      Top             =   2484
      Width           =   1020
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   924
      TabIndex        =   30
      Top             =   3780
      Width           =   828
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   4140
      TabIndex        =   29
      Top             =   3780
      Width           =   1212
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Balance"
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
      Height          =   252
      Left            =   912
      TabIndex        =   28
      Top             =   4416
      Width           =   1500
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Note:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   432
      TabIndex        =   27
      Top             =   7248
      Width           =   684
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   4596
      TabIndex        =   26
      Top             =   1236
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   240
      Width           =   8652
   End
End
Attribute VB_Name = "frmBLAdjustBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Dim TempLicBal As Double
  Dim TempPenBal As Double
  Dim TempTotBal As Double
  Dim TempLicAmt As Double
  Dim TempLicDet(0 To 4) As Double
  Dim TempPenAmt As Double
  Dim TempTotAmt As Double
  Dim SavedLicAmt As Double
  Dim SavedPenAmt As Double
  Dim SavedIssAmt As Double
  Dim SavedTotAmt As Double
  Dim NumOfCodes As Integer
  Dim TempIssAmt As Double
  Dim TempIssBal As Double
  Dim TempAcctNum As String
  Dim TempThisCaption$
  Dim CatDesc(0 To 4) As String
  Dim DeletedFlag As Boolean
  Dim NoMatchFoundFlag As Boolean
  Private Temp_Class As Resize_Class

Private Sub cmdCustList_Click()
  frmBLCustomerList.Wheretogo frmBLAdjustBal, frmBLAdjustBal, 2
  frmBLCustomerList.Show vbModal
  DoEvents
End Sub

Private Sub cmdExit_Click()
  ThisCustXNum = 0
  GCustNum = 0
  KillFile "adjustbalance.dat"
  Load frmCMPaySource
  DoEvents
  frmCMPaySource.Show
  BLLog ("Adjust CM customer balance screen exited.")
  CMLog "BL Adj screen exited."
  Unload frmBLAdjustBal
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    fptxtToday.ToolTipText = ""
    fptxtAccount.ToolTipText = ""
    cmdCustList.ToolTipText = ""
    fptxtName.ToolTipText = ""
    fptxtAddress.ToolTipText = ""
    fptxtCity.ToolTipText = ""
    fptxtState.ToolTipText = ""
    fptxtZip.ToolTipText = ""
    optUnderBill.ToolTipText = ""
    optOverBill.ToolTipText = ""
    optOverPay.ToolTipText = ""
    fpcurrTotBal.ToolTipText = ""
    fpcurrTotAmt.ToolTipText = ""
    fpCurrPenBal.ToolTipText = ""
    fpcurrPenAmt.ToolTipText = ""
    fpcurrIssFeeBal.ToolTipText = ""
    fpcurrIssFeeAmt.ToolTipText = ""
    fptxtDesc.ToolTipText = ""
    fpcurrLicBal.ToolTipText = ""
    fpcurrLicAmt.ToolTipText = ""
    fpcurrLicBalDet(0).ToolTipText = ""
    fpcurrLicBalDet(1).ToolTipText = ""
    fpcurrLicBalDet(2).ToolTipText = ""
    fpcurrLicBalDet(3).ToolTipText = ""
    fpcurrLicBalDet(4).ToolTipText = ""
    fpcurrLicAmtDet(0).ToolTipText = ""
    fpcurrLicAmtDet(1).ToolTipText = ""
    fpcurrLicAmtDet(2).ToolTipText = ""
    fpcurrLicAmtDet(3).ToolTipText = ""
    fpcurrLicAmtDet(4).ToolTipText = ""
    cmdHelp.ToolTipText = ""
    cmdPost.ToolTipText = ""
    cmdExit.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    fptxtToday.ToolTipText = "Today's date."
'    fptxtAccount.ToolTipText = "Enter the account nunber of the customer whose data you wish to bring up and then press F4 to populate this screen."
'    cmdGetCust.ToolTipText = "After entering the desired customer number press this button to populate this screen with their data."
'    cmdCustList.ToolTipText = "Press this button to bring up a complete list of customers and then double click one of them to populate this screen with their data."
'    fptxtName.ToolTipText = "This is a Read Only field containing the customer's business name."
'    fptxtAddress.ToolTipText = "This is a Read Only field containing the customer's business address."
'    fptxtCity.ToolTipText = "This is a Read Only field containing the name of the city where the customer's business is located."
'    fptxtState.ToolTipText = "This is a Read Only field that contains the name of the state where this customer's business is located."
'    fptxtZip.ToolTipText = "This is a Read Only field containing the customer's zip code."
'    optUnderBill.ToolTipText = "Use this option to increase a customer's balance due to under billing."
'    optOverBill.ToolTipText = "Use this option to decrease a customer's balance due to an over billing."
'    optOverPay.ToolTipText = "Use this option to decrease a customer's balance due to an over payment."
'    fpcurrTotBal.ToolTipText = "This is a Read Only field containing the current total license balance for this customer."
'    fpcurrTotAmt.ToolTipText = "This is a Read Only field that contains a running total of the adjustment amounts entered on this screen."
'    fpcurrPenBal.ToolTipText = "This is a Read Only field that contains the outstanding penalty fee balance for this customer."
'    fpcurrPenAmt.ToolTipText = "The amount entered here will either increase or decrease the penalty balance for this customer."
'    fpcurrIssFeeBal.ToolTipText = "This field displays the total outstanding issuance fee balance."
'    fpcurrIssFeeAmt.ToolTipText = "The amount entered here will either increase or decrease the issuance fee balance for this customer."
'    fptxtDesc.ToolTipText = "Enter a more specific description of this transaction here."
'    fpcurrLicBal.ToolTipText = "This is a Read Only field that contains the outstanding license fee balance for this customer."
'    fpcurrLicAmt.ToolTipText = "This is a Read Only field that contains the running total of all license adjustments entered below."
'    fpcurrLicBalDet(0).ToolTipText = "This is a Read Only field that contains the outstanding balance for License Category 1 for this customer."
'    fpcurrLicBalDet(1).ToolTipText = "This is a Read Only field that contains the outstanding balance for License Category 2 for this customer."
'    fpcurrLicBalDet(2).ToolTipText = "This is a Read Only field that contains the outstanding balance for License Category 3 for this customer."
'    fpcurrLicBalDet(3).ToolTipText = "This is a Read Only field that contains the outstanding balance for License Category 4 for this customer."
'    fpcurrLicBalDet(4).ToolTipText = "This is a Read Only field that contains the outstanding balance for License Category 5 for this customer."
'    fpcurrLicAmtDet(0).ToolTipText = "Enter the adjustment amount desired for License Category 1 in this field."
'    fpcurrLicAmtDet(1).ToolTipText = "Enter the adjustment amount desired for License Category 2 in this field."
'    fpcurrLicAmtDet(2).ToolTipText = "Enter the adjustment amount desired for License Category 3 in this field."
'    fpcurrLicAmtDet(3).ToolTipText = "Enter the adjustment amount desired for License Category 4 in this field."
'    fpcurrLicAmtDet(4).ToolTipText = "Enter the adjustment amount desired for License Category 5 in this field."
'    cmdHelp.ToolTipText = "Press the 'Turn Help On' button to start the help feature. Press the 'Turn Help Off' button to disable the help feature."
'    cmdPost.ToolTipText = "Press F10 to permanently commit the data on the screen to memory."
'    cmdExit.ToolTipText = "Press Escape to leave this screen and return to the Customer Maintenance menu."
  End If
End Sub

Private Sub cmdPost_Click()
  Dim THandle As Integer
  Dim ARTransRec(1) As ARTransRecType
  Dim NumOfTransRecs As Long
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim cnt As Long
  Dim NextTransRec As Long
  Dim CustRecNum As Integer
  Dim Adj As String, TBal As Double, Prev As Long
  Dim ThisType As Integer, x As Integer
  Dim PrintType As String
  
  On Error GoTo ERRORSTUFF
10:
  If fpcurrTotAmt.Value = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No adjustment entries have been made. Post attempt aborted."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
20:
  If GCustNum > 0 Then
    If EmpInLicProcess(CStr(GCustNum)) = True And Exist("artmppst.dat") Then
      frmBLMessageBoxJrWOpts.Label1.Caption = "This customer is currently being processed for a business license renewal. If you wish to post this adjustment all temporary business license files will be deleted. You will be required to re-process the business license fee operation. Do you wish to continue to post anyway?"
      frmBLMessageBoxJrWOpts.Label1.Top = 500
      frmBLMessageBoxJrWOpts.Label1.Height = 1300
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
        Unload frmBLMessageBoxJrWOpts
        Close
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
        KillFile "artmppst.dat"
        KillFile "artmplic.dat"
        BLLog ("User warned that continuing to adjust the balance of customer # " + CStr(GCustNum) + " would delete the 'attmppst.dat' and the 'artmplic.dat' files and the user elected to continue to post the balance adjustment anyway.")
      End If
    End If
  End If
30:
  If GCustNum > 0 Then
    If EmpInPenProcess(CStr(GCustNum)) = True And Exist("artmppst.dat") Then
      frmBLMessageBoxJrWOpts.Label1.Caption = "This customer is currently being processed for a penalty fee. If you wish to continue to post this adjustment the temporary penalty fee file will be deleted. You will be required to re-process penalty fees. Do you wish to continue to post anyway?"
      frmBLMessageBoxJrWOpts.Label1.Top = 500
      frmBLMessageBoxJrWOpts.Label1.Height = 1300
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
        Unload frmBLMessageBoxJrWOpts
        Close
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
        KillFile "artmppen.dat"
        BLLog ("User warned that continuing to edit customer # " + CStr(GCustNum) + " would delete the 'attmppen.dat' file and the user elected to post the balance adjustment anyway.")
      End If
    End If
  End If
40:
  If QPTrim$(fptxtName.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "There must be valid customer data on the screen before posting can take place. Posting aborted."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    BLLog ("Adjustment post aborted because the business name was not valid.")
    Close
    Exit Sub
  End If
50:
  frmBLMessageBoxJrWOpts.Label1.Caption = "Are you sure you want to post these adjustment entries?"
  frmBLMessageBoxJrWOpts.Label1.Top = 900
  frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Post"
  frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Exit"
  frmBLMessageBoxJrWOpts.Show vbModal
60:
  If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
    Unload frmBLMessageBoxJrWOpts
    Close
    Exit Sub
  Else
    Unload frmBLMessageBoxJrWOpts
    BLLog ("User warned they were posting adj entries for customer # " + QPTrim$(fptxtAccount.Text) + " - " + QPTrim$(fptxtName.Text) + " and elected to continue the post.")
  End If
70:
  If fpcurrIssFeeAmt.DoubleValue < 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter positive amounts only."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fpcurrIssFeeAmt.SetFocus
    BLLog ("Adj post aborted because of a negative issuance fee entry.")
    CMLog "BLAdj post aborted-neg fee."
    Close
    Exit Sub
  End If
80:
  If fpcurrPenAmt.DoubleValue < 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter positive amounts only."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fpcurrPenAmt.SetFocus
    BLLog ("Adj post aborted because of a negative penalty entry.")
    CMLog "BLAdj post abort-neg pen"
    Close
    Exit Sub
  End If
90:
  For x = 0 To 4
    If fpcurrLicAmtDet(x).DoubleValue < 0 Then
      frmBLMessageBoxJr.Label1.Caption = "Please enter positive amounts only."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      fpcurrLicAmtDet(x).SetFocus
      BLLog ("Adj post aborted because of a negative entry for category # " + CStr(x + 1) + ".")
      CMLog "BLAdj post abort-neg"
      Close
      Exit Sub
    End If
  Next x
93:
  If CompareAcctNumWData = False Then
    Exit Sub
  End If
  
  'if a user brings up data for a customer and then changes the account number
  'but does not bring up the data for that new account number then this filter
  'is designed to catch it
'  If QPTrim$(TempAcctNum) <> QPTrim$(fptxtAccount.Text) Then
'    fptxtAccount.BackColor = &H80FFFF
'    frmBLMessageBoxJr.Label1.Caption = "Please check the account number to make sure it is the account number that is assigned to the business listed."
'    frmBLMessageBoxJr.Label1.Top = 700
'    frmBLMessageBoxJr.Show vbModal
'    fptxtAccount.BackColor = &H80000005
'    fptxtAccount.SetFocus
'    Exit Sub
'  End If
  
  'filter keeps the user from posting a negative amount for penalty
100:
  If optOverBill.Value = True Then
    If fpCurrPenBal.DoubleValue <= 0 And fpcurrPenAmt.DoubleValue > 0 Then
      fpcurrPenAmt.BackColor = &H80FFFF
      frmBLMessageBoxJr.Label1.Caption = "The amount entered for Penalty Adjustment Amount would result in a negative penalty balance for this customer. Please revise the penalty amount entered so that no negative penalty balance results."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      fpcurrPenAmt = 0
      fpcurrPenAmt.BackColor = &H80000005
      fpcurrPenAmt.SetFocus
      BLLog ("Adj post aborted because the penalty adjustment amount would cause the penalty balance to be negative.")
      CMLog "BLAdj post abort"
      Close
      Exit Sub
    ElseIf fpCurrPenBal.DoubleValue - fpcurrPenAmt.DoubleValue < 0 Then
      fpcurrPenAmt.BackColor = &H80FFFF
      frmBLMessageBoxJr.Label1.Caption = "The amount entered for Penalty Adjustment Amount would result in a negative penalty balance for this customer. Please revise the penalty amount entered so that no negative penalty balance results."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      fpcurrPenAmt = 0
      fpcurrPenAmt.BackColor = &H80000005
      fpcurrPenAmt.SetFocus
      BLLog ("Adj post aborted because the issuance fee adjustment amount would cause the issuance fee balance to be negative.")
      CMLog "BLAdj post aborted"
      Close
      Exit Sub
    End If
  End If
110:
  If TotalsOK = False Then
    Exit Sub
  End If
120:
  OpenBLTransFile THandle
  OpenBLCustFile CHandle
121:
  If optUnderBill.Value = True Or optOverPay.Value = True Then
    ThisType = 2 'billed too little...increase balance
  Else
    ThisType = 1 'payed or billed too much...decrease balance
  End If
122:
  NumOfTransRecs = LOF(THandle) \ Len(ARTransRec(1))
  NextTransRec = NumOfTransRecs + 1
123:
  If OldRound(fpcurrTotAmt.DoubleValue) <> 0 Then
    CustRecNum = Val(fptxtAccount.Text)
    'retrieve the customer file on the customer being adjusted
    Get CHandle, CustRecNum, CustRec
    ARTransRec(1).CustomerNumber = QPTrim$(fptxtAccount.Text)
    ARTransRec(1).TransDate = Date2Num(fptxtToday)
    ARTransRec(1).Posted2GL = "N"
    SavedTotAmt = fpcurrTotAmt.DoubleValue
    Select Case ThisType
      Case 1      'downward
      'transaction types are assigned a number according to the
      'adjustment that takes place...these numbers are used by the
      'General Ledger to know how to post the amounts there...they
      'are also used by business license reports to know what transactions
      'to display
        If optOverPay.Value = True Then
          ARTransRec(1).TransType = 13 'adjust down Payment
        ElseIf optOverBill.Value = True Then
          ARTransRec(1).TransType = 23 'adjust down Billing
        End If
        'Adj$ = "CM-ADJDNA "
        If (fpcurrLicAmt.DoubleValue > 0 Or fpcurrIssFeeAmt.DoubleValue > 0) And fpcurrPenAmt.DoubleValue > 0 Then
          ARTransRec(1).DetailTransType = 311
        ElseIf (fpcurrLicAmt.DoubleValue > 0 Or fpcurrIssFeeAmt.DoubleValue > 0) And fpcurrPenAmt.DoubleValue = 0 Then
          ARTransRec(1).DetailTransType = 310
        ElseIf fpcurrLicAmt.DoubleValue = 0 And fpcurrPenAmt.DoubleValue > 0 Then
          ARTransRec(1).DetailTransType = 301
        Else
          ARTransRec(1).DetailTransType = 0
        End If
125:
        ARTransRec(1).TransAmount = fpcurrTotAmt.DoubleValue
        ARTransRec(1).PenAmt = fpcurrPenAmt.DoubleValue
        ARTransRec(1).LicAmt = fpcurrLicAmt.DoubleValue
        ARTransRec(1).IssAmt = fpcurrIssFeeAmt.DoubleValue
        
        ARTransRec(1).CatLicAmt1 = fpcurrLicAmtDet(0).DoubleValue
        ARTransRec(1).CatLicAmt2 = fpcurrLicAmtDet(1).DoubleValue
        ARTransRec(1).CatLicAmt3 = fpcurrLicAmtDet(2).DoubleValue
        ARTransRec(1).CatLicAmt4 = fpcurrLicAmtDet(3).DoubleValue
        ARTransRec(1).CatLicAmt5 = fpcurrLicAmtDet(4).DoubleValue
        'calculations for down adjustments are not the same as
        'upward adjustments
        'will decrease the existing license fees because you are reducing
        'a billing amount (which is where fees are calculated)
        CustRec.Fee1 = CustRec.Fee1 - fpcurrLicAmtDet(0).DoubleValue
        CustRec.Fee2 = CustRec.Fee2 - fpcurrLicAmtDet(1).DoubleValue
        CustRec.Fee3 = CustRec.Fee3 - fpcurrLicAmtDet(2).DoubleValue
        CustRec.Fee4 = CustRec.Fee4 - fpcurrLicAmtDet(3).DoubleValue
        CustRec.Fee5 = CustRec.Fee5 - fpcurrLicAmtDet(4).DoubleValue
        CustRec.FeeAmt = CustRec.FeeAmt - fpcurrLicAmt.DoubleValue
        'now adjust each license balance transaction accordingly
        ARTransRec(1).CatLicBal1 = fpcurrLicBalDet(0).DoubleValue - fpcurrLicAmtDet(0).DoubleValue
        ARTransRec(1).CatLicBal2 = fpcurrLicBalDet(1).DoubleValue - fpcurrLicAmtDet(1).DoubleValue
        ARTransRec(1).CatLicBal3 = fpcurrLicBalDet(2).DoubleValue - fpcurrLicAmtDet(2).DoubleValue
        ARTransRec(1).CatLicBal4 = fpcurrLicBalDet(3).DoubleValue - fpcurrLicAmtDet(3).DoubleValue
        ARTransRec(1).CatLicBal5 = fpcurrLicBalDet(4).DoubleValue - fpcurrLicAmtDet(4).DoubleValue
        'then adjust the license balances saved in this customer's file
        CustRec.FeeLicBal1 = fpcurrLicBalDet(0).DoubleValue - fpcurrLicAmtDet(0).DoubleValue
        CustRec.FeeLicBal2 = fpcurrLicBalDet(1).DoubleValue - fpcurrLicAmtDet(1).DoubleValue
        CustRec.FeeLicBal3 = fpcurrLicBalDet(2).DoubleValue - fpcurrLicAmtDet(2).DoubleValue
        CustRec.FeeLicBal4 = fpcurrLicBalDet(3).DoubleValue - fpcurrLicAmtDet(3).DoubleValue
        CustRec.FeeLicBal5 = fpcurrLicBalDet(4).DoubleValue - fpcurrLicAmtDet(4).DoubleValue
128:
        SavedPenAmt = -fpcurrPenAmt.DoubleValue
        SavedLicAmt = -fpcurrLicAmt.DoubleValue
        SavedIssAmt = -fpcurrIssFeeAmt.DoubleValue
        
        CustRec.LicBal = OldRound#(CustRec.LicBal - fpcurrLicAmt.DoubleValue)
        CustRec.PenBal = OldRound#(CustRec.PenBal - fpcurrPenAmt.DoubleValue)
        CustRec.IssuanceBal = OldRound#(CustRec.IssuanceBal - fpcurrIssFeeAmt.DoubleValue)
        
        TBal# = OldRound#(CustRec.LicBal + CustRec.PenBal + CustRec.IssuanceBal)
130:
   Case 2 'adjust up works the same as adjust down but in reverse
        If optOverPay.Value = True Then
          ARTransRec(1).TransType = 13 'adjust up Payment
        Else
          ARTransRec(1).TransType = 24 '24 = adjust up Billing
        End If
        'Adj$ = "CM-ADJUPA "
        If (fpcurrLicAmt.DoubleValue > 0 Or fpcurrIssFeeAmt.DoubleValue > 0) And fpcurrPenAmt.DoubleValue > 0 Then
          ARTransRec(1).DetailTransType = 411
        ElseIf (fpcurrLicAmt.DoubleValue > 0 Or fpcurrIssFeeAmt.DoubleValue > 0) And fpcurrPenAmt.DoubleValue = 0 Then
          ARTransRec(1).DetailTransType = 410
        ElseIf fpcurrLicAmt.DoubleValue = 0 And fpcurrPenAmt.DoubleValue > 0 Then
          ARTransRec(1).DetailTransType = 401
        Else
          ARTransRec(1).DetailTransType = 0
        End If
        
        ARTransRec(1).TransAmount = fpcurrTotAmt.DoubleValue
        ARTransRec(1).PenAmt = fpcurrPenAmt.DoubleValue
        ARTransRec(1).LicAmt = fpcurrLicAmt.DoubleValue
        ARTransRec(1).IssAmt = fpcurrIssFeeAmt.DoubleValue
        
139:
        ARTransRec(1).CatLicAmt1 = fpcurrLicAmtDet(0).DoubleValue
        ARTransRec(1).CatLicAmt2 = fpcurrLicAmtDet(1).DoubleValue
        ARTransRec(1).CatLicAmt3 = fpcurrLicAmtDet(2).DoubleValue
        ARTransRec(1).CatLicAmt4 = fpcurrLicAmtDet(3).DoubleValue
        ARTransRec(1).CatLicAmt5 = fpcurrLicAmtDet(4).DoubleValue
140:
        If ARTransRec(1).TransType = 24 Then '24 = adjust billing up
          CustRec.Fee1 = CustRec.Fee1 + fpcurrLicAmtDet(0).DoubleValue
          CustRec.Fee2 = CustRec.Fee2 + fpcurrLicAmtDet(1).DoubleValue
          CustRec.Fee3 = CustRec.Fee3 + fpcurrLicAmtDet(2).DoubleValue
          CustRec.Fee4 = CustRec.Fee4 + fpcurrLicAmtDet(3).DoubleValue
          CustRec.Fee5 = CustRec.Fee5 + fpcurrLicAmtDet(4).DoubleValue
          CustRec.FeeAmt = CustRec.FeeAmt + fpcurrLicAmt.DoubleValue
        ElseIf ARTransRec(1).TransType = 13 Then '13 = adjust payment up
        'adjusts payment transactions, not billing
          CustRec.FeeLicPay1 = CustRec.FeeLicPay1 + fpcurrLicAmtDet(0).DoubleValue
          CustRec.FeeLicPay2 = CustRec.FeeLicPay2 + fpcurrLicAmtDet(1).DoubleValue
          CustRec.FeeLicPay3 = CustRec.FeeLicPay3 + fpcurrLicAmtDet(2).DoubleValue
          CustRec.FeeLicPay4 = CustRec.FeeLicPay4 + fpcurrLicAmtDet(3).DoubleValue
          CustRec.FeeLicPay5 = CustRec.FeeLicPay5 + fpcurrLicAmtDet(4).DoubleValue
        End If
145:
        ARTransRec(1).CatLicBal1 = fpcurrLicBalDet(0).DoubleValue + fpcurrLicAmtDet(0).DoubleValue
        ARTransRec(1).CatLicBal2 = fpcurrLicBalDet(1).DoubleValue + fpcurrLicAmtDet(1).DoubleValue
        ARTransRec(1).CatLicBal3 = fpcurrLicBalDet(2).DoubleValue + fpcurrLicAmtDet(2).DoubleValue
        ARTransRec(1).CatLicBal4 = fpcurrLicBalDet(3).DoubleValue + fpcurrLicAmtDet(3).DoubleValue
        ARTransRec(1).CatLicBal5 = fpcurrLicBalDet(4).DoubleValue + fpcurrLicAmtDet(4).DoubleValue
150:
        CustRec.FeeLicBal1 = fpcurrLicBalDet(0).DoubleValue + fpcurrLicAmtDet(0).DoubleValue
        CustRec.FeeLicBal2 = fpcurrLicBalDet(1).DoubleValue + fpcurrLicAmtDet(1).DoubleValue
        CustRec.FeeLicBal3 = fpcurrLicBalDet(2).DoubleValue + fpcurrLicAmtDet(2).DoubleValue
        CustRec.FeeLicBal4 = fpcurrLicBalDet(3).DoubleValue + fpcurrLicAmtDet(3).DoubleValue
        CustRec.FeeLicBal5 = fpcurrLicBalDet(4).DoubleValue + fpcurrLicAmtDet(4).DoubleValue
160:
        SavedPenAmt = fpcurrPenAmt.DoubleValue
        SavedLicAmt = fpcurrLicAmt.DoubleValue
        SavedIssAmt = fpcurrIssFeeAmt.DoubleValue
170:
        CustRec.LicBal = OldRound#(CustRec.LicBal + fpcurrLicAmt.DoubleValue)
        CustRec.PenBal = OldRound#(CustRec.PenBal + fpcurrPenAmt.DoubleValue)
        CustRec.IssuanceBal = OldRound#(CustRec.IssuanceBal + fpcurrIssFeeAmt.DoubleValue)
180:
        TBal# = OldRound#(CustRec.LicBal + CustRec.PenBal + CustRec.IssuanceBal)
    End Select
        
    '-----------------------------------------------------------
190:
    CustRec.AcctBal = TBal#
    ARTransRec(1).BalanceAfterTrans = TBal#
198:
    Adj$ = "CM-" + QPTrim$(fptxtDesc.Text)
    ARTransRec(1).TransDesc = Adj$
    ARTransRec(1).FeeAmt = 0
    ARTransRec(1).CashAmount = 0                'EditBegBalRec(1).Amount
    ARTransRec(1).ChkAmount = 0
    ARTransRec(1).ExtraRoom = ""
    ARTransRec(1).NextTrans = 0 ' CustRec.LastTrans
    ARTransRec(1).CatCodeRec1 = GetCatRecNum(CustRec.BILLCAT1)
    ARTransRec(1).CatCodeRec2 = GetCatRecNum(CustRec.BILLCAT2)
    ARTransRec(1).CatCodeRec3 = GetCatRecNum(CustRec.BILLCAT3)
    ARTransRec(1).CatCodeRec4 = GetCatRecNum(CustRec.BILLCAT4)
    ARTransRec(1).CatCodeRec5 = GetCatRecNum(CustRec.BILLCAT5)
    ARTransRec(1).PenBal = CustRec.PenBal
    ARTransRec(1).LicBal = CustRec.LicBal
    ARTransRec(1).IssBal = CustRec.IssuanceBal
200:
    Put THandle, NextTransRec, ARTransRec(1)
205:
    If CustRec.FirstTrans = 0 Then
      CustRec.FirstTrans = NextTransRec
      CustRec.LastTrans = NextTransRec
      Put CHandle, CustRecNum, CustRec
    Else
      Prev& = CustRec.LastTrans
      CustRec.LastTrans = NextTransRec
      Put CHandle, CustRecNum, CustRec
      Get THandle, Prev&, ARTransRec(1)
      ARTransRec(1).NextTrans = NextTransRec
      Put THandle, Prev&, ARTransRec(1)
    End If
210:
  End If
  'now record the activity that took place here in the arlog.dat file
  CMLog "BLAdj Save Adj Acct- " + Str(CustRecNum)
  Call LogSaves
220:
  Close
225:
  frmBLSucSave.Label1.Caption = "Transaction adjustment data for " + QPTrim$(fptxtName.Text) + " has been posted successfully."
  frmBLSucSave.Label1.Top = 700
  frmBLSucSave.Show vbModal
230:
  If optUnderBill.Value = True Then
    TempThisCaption = "Billing Upward Adjustment"
  ElseIf optOverBill.Value = True Then
    TempThisCaption = "Billing Downward Adjustment"
  ElseIf optOverPay.Value = True Then
    TempThisCaption = "Over Payment Adjustment"
  End If
240:
  frmBLReportOpt.Label2.Caption = "Do you wish to print a report?"
  frmBLReportOpt.Show vbModal 'opens small screen from which the
 ' user selects the printing method
  PrintType$ = frmBLReportOpt.fptxtPrintType
  Unload frmBLReportOpt
250:
  If fptxtAccount.Enabled = True Then
    fptxtAccount.SetFocus
  End If
260:
  Select Case PrintType$
    Case "Graphical"
      Call PrintGraphics
    Case "Text"
      frmBLMessageBoxJr.Label1.Caption = "Pitch 10 is recommended for this report."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Call PrintText
    Case "Exit"
  End Select
270:
  GCustNum = 0
273:
  Call LoadMe
275:
  Exit Sub
280:
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
   '   MsgBox "Error Code Was " + Err.Description + Str$(Err) + " (PrintWSet - Line:" & Erl & ")"

Proc_Exit:
  '--- Cleanup code goes here...
    Close
  '  ClearInUse PWcnt
  '  CitiTerminate
  
End Sub

Private Sub cmdTransHist_Click()
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim ThisCustNum As Integer
  
  If QPTrim$(fptxtAccount.Text) = "" Then
    Exit Sub
  End If
  
  ThisCustXNum = CInt(fptxtAccount.Text)
  
  If Check4ValidCustNum(QPTrim$(fptxtAccount.Text)) = False Then
    frmBLMessageBoxJr.Label1.Caption = "The customer number entered is not valid. Please enter a valid customer number."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Call Clearscreen
    If optOverBill.Enabled = True Then
      optOverBill.SetFocus
    End If
    Exit Sub
  End If
  
  If Exist("transhistjr.dat") Then Exit Sub
  
  If ThisCustXNum > 0 Then
    OpenBLCustFile CustHandle
    Get CustHandle, ThisCustXNum, CustRec
    Close CustHandle
    If CustRec.LastTrans = 0 Then
      frmBLMessageBoxJr.Label1.Caption = "This customer has no transaction activity."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      If fptxtAccount.Enabled = True Then
        fptxtAccount.SetFocus
      End If
      Exit Sub
    Else
      Load frmBLTransHistJr
      DoEvents
      frmBLTransHistJr.Show vbModal
      DoEvents
      Me.Hide
    End If
  End If
  
End Sub

Private Sub cmdTransHist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call cmdTransHist_Click
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  BLLog ("Adjust customer balance screen opened.")
  CMLog "BLAdj screen open"
  DeletedFlag = False
  NoMatchFoundFlag = False
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ''' Me.Visible = False
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
      Call cmdExit_Click
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF10:
      Call cmdPost_Click
      SendKeys "%P"
      KeyCode = 0
    Case vbKeyF7:
      Call cmdCustList_Click
      SendKeys "%L"
      KeyCode = 0
    Case vbKeyF4:
      Call cmdTransHist_Click
      SendKeys "%H"
      KeyCode = 0
    Case vbKeyF1:
      Call cmdHelp_Click
      SendKeys "%T"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      KillFile "adjustbalance.dat"
      ClearInUse PWcnt
      BLLog ("terminated via menu bar on frmBLAdjustBal.")
      CMLog ("terminated via menu bar on frmBLAdjustBal.")
      CitiTerminate
      End
    End If
  End If
End Sub

Public Sub LoadMe()
  
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim LicBal As Double
  Dim PenBal As Double
  Dim IssBal As Double
  Dim LicPenBal As Double
  Dim One As Integer
  Dim DHandle As Integer
  Dim ThisCap As String * 22
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  
  lblBalloon.Visible = False
  ThisCustXNum = 0
'  fptxtToday.ToolTipText = "Today's date."
'  fptxtAccount.ToolTipText = "Enter the account nunber of the customer whose data you wish to bring up and then press F4 to populate this screen."
'  cmdGetCust.ToolTipText = "After entering the desired customer number press this button to populate this screen with their data."
'  cmdCustList.ToolTipText = "Press this button to bring up a complete list of customers and then double click one of them to populate this screen with their data."
'  fptxtName.ToolTipText = "This is a Read Only field containing the customer's business name."
'  fptxtAddress.ToolTipText = "This is a Read Only field containing the customer's business address."
'  fptxtCity.ToolTipText = "This is a Read Only field containing the name of the city where the customer's business is located."
'  fptxtState.ToolTipText = "This is a Read Only field that contains the name of the state where this customer's business is located."
'  fptxtZip.ToolTipText = "This is a Read Only field containing the customer's zip code."
'  optUnderBill.ToolTipText = "Use this option to increase a customer's balance due to under billing."
'  optOverBill.ToolTipText = "Use this option to decrease a customer's balance due to an over billing."
'  optOverPay.ToolTipText = "Use this option to decrease a customer's balance due to an over payment."
'  fpcurrTotBal.ToolTipText = "This is a Read Only field containing the current total license balance for this customer."
'  fpcurrTotAmt.ToolTipText = "This is a Read Only field that contains a running total of the adjustment amounts entered on this screen."
'  fpcurrPenBal.ToolTipText = "This is a Read Only field that contains the outstanding penalty fee balance for this customer."
'  fpcurrPenAmt.ToolTipText = "The amount entered here will either increase or decrease the penalty balance for this customer."
'  fpcurrIssFeeBal.ToolTipText = "This field displays the total outstanding issuance fee balance."
'  fpcurrIssFeeAmt.ToolTipText = "The amount entered here will either increase or decrease the issuance fee balance for this customer."
'  fptxtDesc.ToolTipText = "Enter a more specific description of this transaction here."
'  fpcurrLicBal.ToolTipText = "This is a Read Only field that contains the outstanding license fee balance for this customer."
'  fpcurrLicAmt.ToolTipText = "This is a Read Only field that contains the running total of all license adjustments entered below."
'  fpcurrLicBalDet(0).ToolTipText = "This is a Read Only field that contains the outstanding balance for License Category 1 for this customer."
'  fpcurrLicBalDet(1).ToolTipText = "This is a Read Only field that contains the outstanding balance for License Category 2 for this customer."
'  fpcurrLicBalDet(2).ToolTipText = "This is a Read Only field that contains the outstanding balance for License Category 3 for this customer."
'  fpcurrLicBalDet(3).ToolTipText = "This is a Read Only field that contains the outstanding balance for License Category 4 for this customer."
'  fpcurrLicBalDet(4).ToolTipText = "This is a Read Only field that contains the outstanding balance for License Category 5 for this customer."
'  fpcurrLicAmtDet(0).ToolTipText = "Enter the adjustment amount desired for License Category 1 in this field."
'  fpcurrLicAmtDet(1).ToolTipText = "Enter the adjustment amount desired for License Category 2 in this field."
'  fpcurrLicAmtDet(2).ToolTipText = "Enter the adjustment amount desired for License Category 3 in this field."
'  fpcurrLicAmtDet(3).ToolTipText = "Enter the adjustment amount desired for License Category 4 in this field."
'  fpcurrLicAmtDet(4).ToolTipText = "Enter the adjustment amount desired for License Category 5 in this field."
'  cmdHelp.ToolTipText = "Press the 'Turn Help On' button to start the help feature. Press the 'Turn Help Off' button to disable the help feature."
'  cmdPost.ToolTipText = "Press F10 to permanently commit the data on the screen to memory."
'  cmdExit.ToolTipText = "Press Escape to leave this screen and return to the Customer Maintenance menu."
  
  One = 1
  DHandle = FreeFile
  Open "adjustbalance.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
  NumOfCodes = 0
  fptxtToday = Date$
  
  'opt buttons are highlighted in yellow when they
  'are selected (actually the background)...the buttons
  'are small and it can be difficult to see them
  optOverBill.ForeColor = &H80000012  'black
  optOverBill.Value = True
  optOverBill.BackColor = &H80FFFF
  optUnderBill.ForeColor = &H80000005 'white
  optUnderBill.Value = False
  optUnderBill.BackColor = &H8F8265
  optOverPay.ForeColor = &H80000005 'white
  optOverPay.Value = False
  optOverPay.BackColor = &H8F8265
  
  If GCustNum = 0 Then
    GoSub Clearscreen
    Exit Sub
  End If
  
  OpenBLCustFile CHandle
  NumOfCustRecs = LOF(CHandle) / Len(CustRec)
  
  If GCustNum > 0 And GCustNum <= NumOfCustRecs Then
    If EmpInPenProcess(CStr(GCustNum)) = True Then
      Get CHandle, GCustNum, CustRec
      frmBLMessageBoxJr.Label1.Caption = QPTrim$(CustRec.CustName) + " is currently involved in an unposted penalty assessment file. If you choose to post a balance adjustment for this customer then temporary penalty files will be deleted and you will be required to restart your penalty calculations."
      frmBLMessageBoxJr.Label1.Top = 500
      frmBLMessageBoxJr.Label1.Height = 1300
      frmBLMessageBoxJr.Show vbModal
    End If

    If EmpInLicProcess(CStr(GCustNum)) = True And Exist("artmppst.dat") Then
      Get CHandle, GCustNum, CustRec
      frmBLMessageBoxJr.Label1.Caption = QPTrim$(CustRec.CustName) + " is currently involved in an unposted business license renewal file. If you choose to post a balance adjustment for this customer then all temporary business license fee files will be deleted and you will be required to re-process the business license fee operation."
      frmBLMessageBoxJr.Label1.Height = 1300
      frmBLMessageBoxJr.Label1.Top = 500
      frmBLMessageBoxJr.Show vbModal
    End If
  End If
  ThisCap = ""
  If GCustNum > 0 And GCustNum <= NumOfCustRecs Then
    Get CHandle, GCustNum, CustRec
    If CustRec.Deleted = "Y" Or QPTrim$(CustRec.SORTNAME) = "DELETED" Then
      frmBLMessageBoxJr.Label1.Caption = "This customer has been deleted."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      If fptxtAccount.Enabled = True Then
        DeletedFlag = True
        fptxtAccount.SetFocus
      End If
      Close
      Exit Sub
    End If
  End If
  
  Close CHandle
  For x = 0 To 4
    Label3(x).Caption = ""
  Next x
  
  
  If QPTrim$(CustRec.BILLCAT1) <> "" Then
    ThisCap = GetCatDesc(CustRec.BILLCAT1)
    Label3(0).Caption = ThisCap 'QPTrim$(CustRec.DESC1)
    CatDesc(0) = GetCatDesc(CustRec.BILLCAT1)
    TempLicDet(0) = CustRec.FeeLicBal1
    NumOfCodes = NumOfCodes + 1
    fpcurrLicAmtDet(0).Enabled = True
  Else
    CatDesc(0) = ""
    If GCustNum > 0 And GCustNum <= NumOfCustRecs Then
      fpcurrLicAmtDet(0).Enabled = False
    End If
  End If
  If QPTrim$(CustRec.BILLCAT2) <> "" Then
    ThisCap = GetCatDesc(CustRec.BILLCAT2)
    Label3(1).Caption = ThisCap 'QPTrim$(CustRec.DESC2)
    CatDesc(1) = GetCatDesc(CustRec.BILLCAT2)
    TempLicDet(1) = CustRec.FeeLicBal2
    NumOfCodes = NumOfCodes + 1
    fpcurrLicAmtDet(1).Enabled = True
  Else
    CatDesc(1) = ""
    If GCustNum > 0 And GCustNum <= NumOfCustRecs Then
      fpcurrLicAmtDet(1).Enabled = False
    End If
  End If
  If QPTrim$(CustRec.BILLCAT3) <> "" Then
    ThisCap = GetCatDesc(CustRec.BILLCAT3)
    Label3(2).Caption = ThisCap 'QPTrim$(CustRec.DESC3)
    CatDesc(2) = GetCatDesc(CustRec.BILLCAT3)
    TempLicDet(2) = CustRec.FeeLicBal3
    NumOfCodes = NumOfCodes + 1
    fpcurrLicAmtDet(2).Enabled = True
  Else
    CatDesc(2) = ""
    If GCustNum > 0 And GCustNum <= NumOfCustRecs Then
      fpcurrLicAmtDet(2).Enabled = False
    End If
  End If
  If QPTrim$(CustRec.BILLCAT4) <> "" Then
    ThisCap = GetCatDesc(CustRec.BILLCAT4)
    Label3(3).Caption = ThisCap 'QPTrim$(CustRec.DESC4)
    CatDesc(3) = GetCatDesc(CustRec.BILLCAT4)
    TempLicDet(3) = CustRec.FeeLicBal4
    NumOfCodes = NumOfCodes + 1
    fpcurrLicAmtDet(3).Enabled = True
  Else
    CatDesc(3) = ""
    If GCustNum > 0 And GCustNum <= NumOfCustRecs Then
      fpcurrLicAmtDet(3).Enabled = False
    End If
  End If
  If QPTrim$(CustRec.BILLCAT5) <> "" Then
    ThisCap = GetCatDesc(CustRec.BILLCAT5)
    Label3(4).Caption = ThisCap 'QPTrim$(CustRec.DESC5)
    CatDesc(4) = GetCatDesc(CustRec.BILLCAT5)
    TempLicDet(4) = CustRec.FeeLicBal5
    NumOfCodes = NumOfCodes + 1
    fpcurrLicAmtDet(4).Enabled = True
  Else
    CatDesc(4) = ""
    If GCustNum > 0 And GCustNum <= NumOfCustRecs Then
      fpcurrLicAmtDet(4).Enabled = False
    End If
  End If
  
  'lock out any lic amount fields that don't contain a code
  For x = 0 To 4
    fpcurrLicAmtDet(x).ControlType = ControlTypeNormal
  Next x
  
  LicBal = OldRound(CustRec.LicBal)
  fpcurrLicBal = LicBal
  fpcurrLicBalDet(0) = OldRound(CustRec.FeeLicBal1)
  fpcurrLicBalDet(1) = OldRound(CustRec.FeeLicBal2)
  fpcurrLicBalDet(2) = OldRound(CustRec.FeeLicBal3)
  fpcurrLicBalDet(3) = OldRound(CustRec.FeeLicBal4)
  fpcurrLicBalDet(4) = OldRound(CustRec.FeeLicBal5)
  TempLicBal = LicBal
  PenBal = OldRound(CustRec.PenBal)
  fpCurrPenBal = PenBal
  TempPenBal = PenBal
  LicPenBal = OldRound(CustRec.PenBal + CustRec.LicBal + CustRec.IssuanceBal)
  IssBal = OldRound(CustRec.IssuanceBal)
  fpcurrIssFeeBal = IssBal
  TempIssBal = IssBal
  
  fpcurrTotBal = CustRec.AcctBal 'LicPenBal
  TempTotBal = CustRec.AcctBal ' LicPenBal
  TempLicAmt = 0
  TempPenAmt = 0
  TempIssAmt = 0
  TempTotAmt = 0
  fpcurrLicAmt = 0
  fpcurrTotAmt = 0
  fpcurrPenAmt = 0
  fptxtAccount.Text = QPTrim$(CustRec.CUSTNUMB)
  TempAcctNum = QPTrim$(CustRec.CUSTNUMB)
  fptxtName.Text = QPTrim$(CustRec.CustName)
  fptxtAddress.Text = QPTrim$(CustRec.ADDRESS1)
  fptxtCity.Text = QPTrim$(CustRec.City)
  fptxtState.Text = QPTrim$(CustRec.State)
  fptxtZip.Text = QPTrim$(CustRec.ZIPCODE)
  For x = 0 To 4
    fpcurrLicAmtDet(x) = 0
  Next x
  
  Exit Sub
    
Clearscreen:
    fptxtAccount.Text = ""
    fptxtName.Text = ""
    fptxtAddress.Text = ""
    fptxtCity.Text = ""
    fptxtState.Text = ""
    fptxtZip.Text = ""
    fpcurrLicBal = ""
    TempLicBal = 0
    fpCurrPenBal = ""
    TempPenBal = 0
    fpcurrTotBal = ""
    TempTotBal = 0
    fptxtDesc.Text = ""
    fpcurrLicAmt = 0
    fpcurrPenAmt = 0
    fpcurrTotAmt = 0
    fpcurrIssFeeAmt = 0
    fpcurrIssFeeBal = 0
    For x = 0 To 4
      fpcurrLicAmtDet(x).Enabled = False
      fpcurrLicBalDet(x) = 0
      fpcurrLicAmtDet(x) = 0
      Label3(x).Caption = ""
    Next x
    optOverBill.Value = True
    
    Return
    
ERRORSTUFF:
   Unload FrmShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLAdjustBal", "LoadMe", Erl)
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
  '  CitiTerminate
    Unload Me
  
End Sub

Private Sub fpcurrIssFeeAmt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If fptxtDesc.Enabled = True Then
      fptxtDesc.SetFocus
    End If
  ElseIf KeyCode = vbKeyUp Then
    If fpcurrPenAmt.Enabled = True Then
      fpcurrPenAmt.SetFocus
    End If
  End If
End Sub

Private Sub fpcurrIssFeeAmt_LostFocus()
  On Error Resume Next
  
  If optOverBill.Value = True Then
    If fpcurrIssFeeBal.DoubleValue <= 0 And fpcurrIssFeeAmt.DoubleValue > 0 Then
      fpcurrIssFeeAmt.BackColor = &H80FFFF
      frmBLMessageBoxJr.Label1.Caption = "The amount entered for Issuance Fee Adjustment Amount would result in a negative Issuance Fee balance for this customer. Please revise the Issuance Fee amount entered so that no negative Issuance Fee balance results."
      frmBLMessageBoxJr.Label1.Top = 600
      frmBLMessageBoxJr.Show vbModal
      fpcurrIssFeeAmt = 0
      fpcurrIssFeeAmt.BackColor = &H80000005
      fpcurrIssFeeAmt.SetFocus
      Exit Sub
    ElseIf fpcurrIssFeeBal.DoubleValue - fpcurrIssFeeAmt.DoubleValue < 0 Then
      fpcurrIssFeeAmt.BackColor = &H80FFFF
      frmBLMessageBoxJr.Label1.Caption = "The amount entered for Issuance Fee Adjustment Amount would result in a negative Issuance Fee balance for this customer. Please revise the Issuance Fee amount entered so that no negative Issuance Fee balance results."
      frmBLMessageBoxJr.Label1.Top = 600
      frmBLMessageBoxJr.Show vbModal
      fpcurrIssFeeAmt = 0
      fpcurrIssFeeAmt.BackColor = &H80000005
      fpcurrIssFeeAmt.SetFocus
      Exit Sub
    End If
  End If
  fpcurrTotAmt = OldRound(fpcurrLicAmt.DoubleValue + fpcurrPenAmt.DoubleValue + fpcurrIssFeeAmt.DoubleValue)

End Sub

Private Sub fpcurrLicAmt_GotFocus()
  frmBLMessageBoxJr.Label1.Caption = "This field shows a running total of individual License Balances and is not editable."
  frmBLMessageBoxJr.Label1.Top = 900
  frmBLMessageBoxJr.Show vbModal
  fpcurrLicAmtDet(0).SetFocus
End Sub

Private Sub fpcurrLicAmtDet_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If Index <> 4 Then
      If fpcurrLicAmtDet(Index + 1).Enabled = False Then
        fpcurrPenAmt.SetFocus
      Else
        fpcurrLicAmtDet(Index + 1).SetFocus
      End If
    ElseIf Index = 4 Then
      If fpcurrPenAmt.Enabled = True Then
        fpcurrPenAmt.SetFocus
      End If
    End If
  ElseIf KeyCode = vbKeyUp Then
    If Index <> 0 Then
      If fpcurrLicAmtDet(Index - 1).Enabled = True Then
        fpcurrLicAmtDet(Index - 1).SetFocus
      Else
        If fptxtDesc.Enabled = True Then
          fptxtDesc.SetFocus
        End If
      End If
    ElseIf Index = 0 Then
      If fptxtDesc.Enabled = True Then
        fptxtDesc.SetFocus
      End If
    End If
  End If
        
End Sub

Private Sub fpcurrLicAmtDet_LostFocus(Index As Integer)
  Dim x As Integer
  Dim ThisTotal As Double
  
  On Error Resume Next
  For x = 0 To 4 'NumOfCodes - 1
    ThisTotal = ThisTotal + fpcurrLicAmtDet(x).DoubleValue
  Next x
  
  fpcurrLicAmt = ThisTotal
  fpcurrTotAmt = ThisTotal + fpcurrPenAmt.DoubleValue + fpcurrIssFeeAmt.DoubleValue

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub fpcurrPenAmt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpcurrIssFeeAmt.SetFocus
  End If
  
  If KeyCode = vbKeyUp Then
    If fpcurrLicAmtDet(4).Enabled = True Then
      fpcurrLicAmtDet(4).SetFocus
    ElseIf fpcurrLicAmtDet(3).Enabled = True Then
      fpcurrLicAmtDet(3).SetFocus
    ElseIf fpcurrLicAmtDet(2).Enabled = True Then
      fpcurrLicAmtDet(2).SetFocus
    ElseIf fpcurrLicAmtDet(1).Enabled = True Then
      fpcurrLicAmtDet(1).SetFocus
    ElseIf fpcurrLicAmtDet(0).Enabled = True Then
      fpcurrLicAmtDet(0).SetFocus
    Else
      fpcurrIssFeeAmt.SetFocus
    End If
  End If
End Sub

Private Sub fpcurrPenAmt_LostFocus()
  fpcurrTotAmt = OldRound(fpcurrLicAmt.DoubleValue + fpcurrPenAmt.DoubleValue + fpcurrIssFeeAmt.DoubleValue)
End Sub

'Private Sub cmdGetCust_Click()
'  Dim CustRec As ARCustRecType
'  Dim CHandle As Integer
'  Dim TotalAccts As Integer
'  Dim x As Integer
'  Dim Number$
'  Dim Name$
'  Dim Found As Boolean
'
'  On Error GoTo ERRORSTUFF
'
'  If QPTrim$(fptxtAccount.Text) = "" Then
'    frmBLMessageBoxJr.Label1.Caption = "Please enter a customer number."
'    frmBLMessageBoxJr.Label1.Top = 900
'    frmBLMessageBoxJr.Show vbModal
'    Exit Sub
'  End If
'
'  Number = QPTrim$(fptxtAccount.Text)
'
'  OpenCustFile CHandle
'  TotalAccts = LOF(CHandle) \ Len(CustRec)
'
'  If TotalAccts = 0 Then
'    frmBLMessageBoxJr.Label1.Caption = "There are no business customers saved."
'    frmBLMessageBoxJr.Label1.Top = 900
'    frmBLMessageBoxJr.Show vbModal
'    Close
'    Exit Sub
'  End If
'
'  For x = 1 To TotalAccts
'    Get CHandle, x, CustRec
'    If Number$ = QPTrim$(CustRec.CustNumb) Then 'match the selected
'    'row with the right code
'      Found = True
'      GCustNum = x 'now you can assign the correct global
'      Exit For
'    Else
'      Found = False
'      GoTo NotAMatch
'    End If
'
'NotAMatch:
'   Next x
'  Close CHandle
'
'  If Found = False Then
'    frmBLMessageBoxJrWOpts.Label1.Caption = "The customer number entered does not match any of those saved. Would you like to see the customer list?"
'    frmBLMessageBoxJrWOpts.Label1.Top = 700
'    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Show List"
'    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
'    frmBLMessageBoxJrWOpts.Show vbModal
'    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
'      Unload frmBLMessageBoxJrWOpts
'      Call cmdCustList_Click
'    Else
'      NoMatchFoundFlag = True
'      Unload frmBLMessageBoxJrWOpts
'    End If
'  Else
'    Call ClearScreen
'    Call LoadMe
'    If DeletedFlag = False Then
'      optOverBill.SetFocus
'    Else
'      DeletedFlag = True
'    End If
'  End If
'
'  Exit Sub
'
'ERRORSTUFF:
'   Unload frmBLShowPctComp
'   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLAdjustBal", "cmdGetCust_Click", Erl)
'     Case emrExitProc:
'       Resume Proc_Exit
'     Case emrResume:
'       Resume
'     Case emrResumeNext:
'       Resume Next
'     Case Else
'      '--- Technically, this should never happen.
'       Resume Proc_Exit
'   End Select
'
'Proc_Exit:
'  '--- Cleanup code goes here...
'    Close
'    ClearInUse PWcnt
'    Terminate
'
'End Sub

Public Sub Clearscreen()
  Dim x
  
  For x = 0 To 4
    fpcurrLicAmtDet(x) = 0
    fpcurrLicBalDet(x) = 0
  Next x
  fpcurrLicAmt.Text = ""
  fpcurrLicBal.Text = ""
  fpcurrPenAmt.Text = ""
  fpCurrPenBal.Text = ""
  fpcurrIssFeeAmt.Text = ""
  fpcurrIssFeeBal.Text = ""
  fpcurrTotAmt.Text = ""
  fpcurrTotBal.Text = ""
  fptxtAccount.Text = ""
  fptxtAddress.Text = ""
  fptxtCity.Text = ""
  fptxtDesc.Text = ""
  fptxtName.Text = ""
  fptxtState.Text = ""
  fptxtToday.Text = ""
  fptxtZip.Text = ""
End Sub

Private Sub LogSaves()
  Dim CHandle As Integer
  Dim CustRec As ARCustRecType
  Dim x As Integer
  Dim ThisFee(0 To 4) As Double
  
  On Error GoTo ERRORSTUFF
  'all activity on this screen that is saved gets recorded in arlog.dat
  OpenBLCustFile CHandle
  Get CHandle, Val(fptxtAccount.Text), CustRec
  Close CHandle
  
  If TempLicBal <> CustRec.LicBal Then
    BLLog ("For " + QPTrim$(CustRec.CustName) + ": " + "License balance has changed from " + QPTrim$(Using("$##,###,##0.00", TempLicBal)) + " to " + QPTrim$(Using("$##,###,##0.00", CustRec.LicBal)) + " in frmBLAdjustBal.")
  End If

  If TempPenBal <> CustRec.PenBal Then
    BLLog ("For " + QPTrim$(CustRec.CustName) + ": " + "Penalty balance has changed from " + QPTrim$(Using("$##,###,##0.00", TempPenBal)) + " to " + QPTrim$(Using("$##,###,##0.00", CustRec.PenBal)) + " in frmBLAdjustBal.")
  End If

  If TempIssBal <> CustRec.IssuanceBal Then
    BLLog ("For " + QPTrim$(CustRec.CustName) + ": " + "Issuance Fee balance has changed from " + QPTrim$(Using("$##,###,##0.00", TempIssBal)) + " to " + QPTrim$(Using("$##,###,##0.00", CustRec.IssuanceBal)) + " in frmBLAdjustBal.")
  End If

  If TempTotBal <> CustRec.AcctBal Then
    BLLog ("For " + QPTrim$(CustRec.CustName) + ": " + "Total balance has changed from " + QPTrim$(Using("$##,###,##0.00", TempTotBal)) + " to " + QPTrim$(Using("$##,###,##0.00", CustRec.AcctBal)) + " in frmBLAdjustBal.")
  End If

  If TempLicAmt <> SavedLicAmt Then
    ThisFee(0) = CustRec.FeeLicBal1
    ThisFee(1) = CustRec.FeeLicBal2
    ThisFee(2) = CustRec.FeeLicBal3
    ThisFee(3) = CustRec.FeeLicBal4
    ThisFee(4) = CustRec.FeeLicBal5
    For x = 0 To 4
      If ThisFee(x) <> TempLicDet(x) Then
        BLLog ("For " + QPTrim$(CustRec.CustName) + ": " + "Total balance for license #" + CStr(x) + " has changed from " + QPTrim$(Using("$##,###,##0.00", TempLicDet(x))) + " to " + QPTrim$(Using("$##,###,##0.00", ThisFee(x))) + " in frmBLAdjustBal.")
      End If
    Next x
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Unload FrmShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLAdjustBal", "LogSaves", Erl)
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
  '  ClearInUse PWcnt
  '  CitiTerminate
    
End Sub

Private Function TotalsOK() As Boolean
  Dim x As Integer
  Dim ThisTotal As Double
  Dim CodeCnt As Integer
  Dim AmtOver(0 To 4) As Double
  Dim OverTot As Double
  Dim AmtUnder(0 To 4) As Double
  Dim UnderTot As Double
  
  On Error GoTo ERRORSTUFF
  
  CodeCnt = 4 'NumOfCodes - 1
  TotalsOK = True
  
  If optUnderBill.Value = True Or optOverPay.Value = True Then Exit Function 'if the balance
  'is being increased then there is no need to make sure that some
  'balances are not reduced at the expense of others
  
  For x = 0 To CodeCnt 'look for credit vs debit balances
    If fpcurrLicBalDet(x).DoubleValue - fpcurrLicAmtDet(x).DoubleValue > 0 Then 'balance not paid in full
      AmtUnder(x) = fpcurrLicBalDet(x).DoubleValue - fpcurrLicAmtDet(x).DoubleValue
      UnderTot = UnderTot + AmtUnder(x)
    ElseIf fpcurrLicAmtDet(x).DoubleValue - fpcurrLicBalDet(x).DoubleValue > 0 Then 'balance overpaid
      AmtOver(x) = fpcurrLicAmtDet(x).DoubleValue - fpcurrLicBalDet(x).DoubleValue
      OverTot = OverTot + AmtOver(x)
    End If
  Next x
    
  If OverTot > 0 Then 'found credit balances which is OK unless other balances
  'are debit...we don't want negative balances
    If fpCurrPenBal.DoubleValue - fpcurrPenAmt.DoubleValue > 0 Then 'still balance left for penalty
      fpcurrPenAmt.BackColor = &H8080FF
      For x = 0 To CodeCnt
        If AmtOver(x) > 0 Then 'highlight categories with credit balances
          fpcurrLicAmtDet(x).BackColor = &H80FFFF
        End If
      Next x
      frmBLMessageBoxJr.Label1.Caption = "The values entered would cause a debit balance for Penalty (red) while other License balances would have credit balances (yellow). Please reduce the Penalty balance before allowing License balance credits."
      frmBLMessageBoxJr.Label1.Top = 600
      frmBLMessageBoxJr.Show vbModal
      TotalsOK = False 'send it back as false
      fpcurrPenAmt.SetFocus
      fpcurrPenAmt.BackColor = &H80000014
      For x = 0 To CodeCnt
        fpcurrLicAmtDet(x).BackColor = &H80000014
      Next x
      Exit Function
    ElseIf fpcurrIssFeeBal.DoubleValue - fpcurrIssFeeAmt.DoubleValue > 0 Then 'still balance left for penalty
      fpcurrIssFeeAmt.BackColor = &H8080FF
      For x = 0 To CodeCnt
        If AmtOver(x) > 0 Then 'highlight categories with credit balances
          fpcurrLicAmtDet(x).BackColor = &H80FFFF
        End If
      Next x
      frmBLMessageBoxJr.Label1.Caption = "The values entered would cause a debit balance for Issuance Fee (red) while other License balances would have credit balances (yellow). Please reduce the Issuance Fee balance before allowing License balance credits."
      frmBLMessageBoxJr.Label1.Top = 600
      frmBLMessageBoxJr.Show vbModal
      TotalsOK = False 'send it back as false
      fpcurrIssFeeAmt.SetFocus
      fpcurrIssFeeAmt.BackColor = &H80000014
      For x = 0 To CodeCnt
        fpcurrLicAmtDet(x).BackColor = &H80000014
      Next x
      Exit Function
    ElseIf UnderTot > 0 Then 'found debit balances which is OK as long
    'as there are no credit balances
      For x = 0 To CodeCnt
        If AmtUnder(x) > 0 Then 'highlight debits
          fpcurrLicAmtDet(x).BackColor = &H8080FF 'red
        ElseIf AmtOver(x) > 0 Then 'highlight credits
          fpcurrLicAmtDet(x).BackColor = &H80FFFF 'yellow
        End If
      Next x
      frmBLMessageBoxJr.Label1.Caption = "The values entered would cause debit balances in some categories (red) while other categories would have credit balances (yellow). Please reduce any debit balances before allowing any credit balances."
      frmBLMessageBoxJr.Label1.Top = 600
      frmBLMessageBoxJr.Show vbModal
      TotalsOK = False 'send back as false
      For x = 0 To CodeCnt
        fpcurrLicAmtDet(x).BackColor = &H80000014
      Next x
      Exit Function
    End If
  End If
  
  If fpcurrPenAmt.DoubleValue - fpCurrPenBal.DoubleValue > 0 Then
  'if penalty has a credit balance while some license balances
  'still are debit then fix it
    If UnderTot > 0 Then
      fpcurrPenAmt.BackColor = &H80FFFF
      For x = 0 To CodeCnt
        If AmtUnder(x) > 0 Then
          fpcurrLicAmtDet(x).BackColor = &H8080FF 'red
        End If
      Next x
      frmBLMessageBoxJr.Label1.Caption = "The values entered would cause a credit balance for Penalty (yellow) while other License balances would have debit balances (red). Please reduce any debit balances before allowing a Penalty credit balance."
      frmBLMessageBoxJr.Label1.Top = 600
      frmBLMessageBoxJr.Show vbModal
      TotalsOK = False
      For x = 0 To CodeCnt
        fpcurrLicAmtDet(x).BackColor = &H80000014
      Next x
      fpcurrPenAmt.BackColor = &H80000014
      Exit Function
    End If
  End If
    
  If fpcurrIssFeeAmt.DoubleValue - fpcurrIssFeeBal.DoubleValue > 0 Then
  'if issue fee has a credit balance while some license balances
  'still are debit then fix it
    If UnderTot > 0 Then
      fpcurrIssFeeAmt.BackColor = &H80FFFF
      For x = 0 To CodeCnt
        If AmtUnder(x) > 0 Then
          fpcurrLicAmtDet(x).BackColor = &H8080FF 'red
        End If
      Next x
      frmBLMessageBoxJr.Label1.Caption = "The values entered would cause a credit balance for Issuance Fee (yellow) while other License balances would have debit balances (red). Please reduce any debit balances before allowing a Issuance Fee credit balance."
      frmBLMessageBoxJr.Label1.Top = 600
      frmBLMessageBoxJr.Show vbModal
      TotalsOK = False
      For x = 0 To CodeCnt
        fpcurrLicAmtDet(x).BackColor = &H80000014
      Next x
      fpcurrIssFeeAmt.BackColor = &H80000014
      Exit Function
    End If
  End If
  
  Exit Function
  
ERRORSTUFF:
   Unload FrmShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLAdjustBal", "TotalsOK", Erl)
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
  '  ClearInUse PWcnt
  '  CitiTerminate
  
End Function

Private Sub fptxtAccount_LostFocus()
  If frmBLMessageBoxJr.Visible = True Then Exit Sub
  Call LostFocusCheck

End Sub

Private Sub fptxtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If fpcurrLicAmtDet(0).Enabled = True Then
      fpcurrLicAmtDet(0).SetFocus
    ElseIf fpcurrLicAmtDet(1).Enabled = True Then
      fpcurrLicAmtDet(1).SetFocus
    ElseIf fpcurrLicAmtDet(2).Enabled = True Then
      fpcurrLicAmtDet(2).SetFocus
    ElseIf fpcurrLicAmtDet(3).Enabled = True Then
      fpcurrLicAmtDet(3).SetFocus
    ElseIf fpcurrLicAmtDet(4).Enabled = True Then
      fpcurrLicAmtDet(4).SetFocus
    Else
      If fpcurrPenAmt.Enabled = True Then
        fpcurrPenAmt.SetFocus
      End If
    End If
  ElseIf KeyCode = vbKeyUp Then
    If fpcurrIssFeeAmt.Enabled = True Then
      fpcurrIssFeeAmt.SetFocus
    End If
  End If
End Sub

Private Sub fptxtName_Change()
  If QPTrim$(fptxtName.Text) <> "" Then
    fptxtAccount.TabStop = False
  Else
    fptxtAccount.TabStop = True
  End If
End Sub

Private Sub optOverBill_GotFocus()
  If QPTrim$(fptxtAccount.Text) = "" Then
    If fptxtAccount.Enabled = True Then
      fptxtAccount.SetFocus
      DoEvents
      Exit Sub
    End If
  End If

  If QPTrim$(fptxtAccount.Text) <> "" Then
    fptxtAccount.TabStop = False
  End If
  
  fpcurrPenAmt.TabIndex = 1
  fptxtAccount.TabIndex = 0

End Sub

Private Sub optOverPay_Click()
  If optOverPay.Value = True Then
    optOverPay.ForeColor = &H80000012 'black
    optOverPay.BackColor = &H80FFFF 'yellow
    optOverBill.ForeColor = &H80000005 'white
    optOverBill.BackColor = &H8F8265 'clear
    optUnderBill.ForeColor = &H80000005 'white
    optUnderBill.BackColor = &H8F8265 'clear
    Label20.Caption = "Adjustments entered will increase balances."
  End If
End Sub

Private Sub optOverPay_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpcurrPenAmt.SetFocus
  End If
End Sub

Private Sub optUnderBill_Click()
  If optUnderBill.Value = True Then
    optUnderBill.ForeColor = &H80000012 'black
    optUnderBill.BackColor = &H80FFFF 'yellow
    optOverBill.ForeColor = &H80000005 'white
    optOverBill.BackColor = &H8F8265 'clear
    optOverPay.ForeColor = &H80000005 'white
    optOverPay.BackColor = &H8F8265 'clear
    Label20.Caption = "Adjustments entered will increase balances."
  End If
End Sub

Private Sub optOverBill_Click()
  
  If optOverBill.Value = True Then
    optOverBill.ForeColor = &H80000012 'black
    optOverBill.BackColor = &H80FFFF 'yellow
    optUnderBill.ForeColor = &H80000005 'white
    optUnderBill.BackColor = &H8F8265 'clear
    optOverPay.ForeColor = &H80000005 'white
    optOverPay.BackColor = &H8F8265 'clear
    Label20.Caption = "Adjustments entered will decrease balances."
  End If

End Sub

Private Function CompareAcctNumWData() As Boolean
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim x As Integer
  
  'A user can bring up a business's data and then change the account number
  'without bringing up the new account number's data...this function is
  'designed to trap for this situation...the data would still be saved for the
  'business whose address appears and not for the business associated with
  'the new account number
  On Error Resume Next
  CompareAcctNumWData = True
  OpenBLCustFile CHandle
  NumOfCustRecs = LOF(CHandle) / Len(CustRec)
  For x = 1 To NumOfCustRecs
    Get CHandle, x, CustRec
    If QPTrim$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SORTNAME) = "DELETED" Then GoTo NotThisOne
    If QPTrim$(CustRec.CustName) = QPTrim$(fptxtName.Text) And QPTrim$(CustRec.ADDRESS1) = QPTrim$(fptxtAddress.Text) Then
      If QPTrim$(CustRec.CUSTNUMB) = QPTrim$(fptxtAccount.Text) Then
        Exit For
      Else
        CompareAcctNumWData = False
        frmBLMessageBoxJr.Label1.Caption = "The account number entered does not match the other data shown for this business. Please check the customer list for the correct data."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        Exit For
      End If
    End If
NotThisOne:
  Next x
  Close CHandle
End Function

Private Sub PrintGraphics()
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim ThisType$
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim x As Integer
  Dim NegFlag As Boolean
  Dim dlm$
  
  dlm$ = "~"
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close
  
  OpenBLCustFile CustHandle
  Get CustHandle, GCustNum, CustRec
  Close
  
  NegFlag = False
  
  If optOverBill.Value = True Then
    NegFlag = True
  End If
  
  ReportFile$ = "BLRPTS\ARADJRPT.RPT"
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  '                              0
  Print #RptHandle, QPTrim$(TownRec.TownName); dlm;
  If NegFlag = False Then
    '                            1
    Print #RptHandle, fpcurrTotAmt.DoubleValue; dlm;
  ElseIf NegFlag = True Then
    '                            1
    Print #RptHandle, -fpcurrTotAmt.DoubleValue; dlm;
  End If
  '                             2                               3
  Print #RptHandle, QPTrim$(fptxtName.Text); dlm; QPTrim$(fptxtAccount.Text); dlm;
  '                             4                               5
  Print #RptHandle, QPTrim$(fptxtAddress.Text); dlm; QPTrim$(fptxtDesc.Text); dlm;
  '                         6                        7                          8
  If NegFlag = False Then
    Print #RptHandle, TempThisCaption; dlm; fpcurrPenAmt.DoubleValue; dlm; CustRec.PenBal; dlm;
    '                             9                             10
    Print #RptHandle, fpcurrIssFeeAmt.DoubleValue; dlm; CustRec.IssuanceBal; dlm;
  Else
    Print #RptHandle, TempThisCaption; dlm; -fpcurrPenAmt.DoubleValue; dlm; CustRec.PenBal; dlm;
    '                             9                             10
    Print #RptHandle, -fpcurrIssFeeAmt.DoubleValue; dlm; CustRec.IssuanceBal; dlm;
  End If
  
  If NegFlag = False Then
    If CatDesc(0) <> "" Then
      '                     11                        12                             13
      Print #RptHandle, CatDesc(0); dlm; fpcurrLicAmtDet(0).DoubleValue; dlm; CustRec.FeeLicBal1; dlm;
    Else
      '
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
    End If
    If CatDesc(1) <> "" Then
      '                     14                        15                             16
      Print #RptHandle, CatDesc(1); dlm; fpcurrLicAmtDet(1).DoubleValue; dlm; CustRec.FeeLicBal2; dlm;
    Else
      '
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
    End If
    If CatDesc(2) <> "" Then
      '                     17                        18                             19
      Print #RptHandle, CatDesc(2); dlm; fpcurrLicAmtDet(2).DoubleValue; dlm; CustRec.FeeLicBal3; dlm;
    Else
      '
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
    End If
    If CatDesc(3) <> "" Then
      '                     20                        21                             22
      Print #RptHandle, CatDesc(3); dlm; fpcurrLicAmtDet(3).DoubleValue; dlm; CustRec.FeeLicBal4; dlm;
    Else
      '
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
    End If
    If CatDesc(4) <> "" Then
      '                     23                        24                             25
      Print #RptHandle, CatDesc(4); dlm; fpcurrLicAmtDet(4).DoubleValue; dlm; CustRec.FeeLicBal5; dlm;
    Else
      '
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
    End If
  Else
    If CatDesc(0) <> "" Then
      '                     11                        12                             13
      Print #RptHandle, CatDesc(0); dlm; -fpcurrLicAmtDet(0).DoubleValue; dlm; CustRec.FeeLicBal1; dlm;
    Else
      '
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
    End If
    If CatDesc(1) <> "" Then
      '                     14                        15                             16
      Print #RptHandle, CatDesc(1); dlm; -fpcurrLicAmtDet(1).DoubleValue; dlm; CustRec.FeeLicBal2; dlm;
    Else
      '
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
    End If
    If CatDesc(2) <> "" Then
      '                     17                        18                             19
      Print #RptHandle, CatDesc(2); dlm; -fpcurrLicAmtDet(2).DoubleValue; dlm; CustRec.FeeLicBal3; dlm;
    Else
      '
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
    End If
    If CatDesc(3) <> "" Then
      '                     20                        21                             22
      Print #RptHandle, CatDesc(3); dlm; -fpcurrLicAmtDet(3).DoubleValue; dlm; CustRec.FeeLicBal4; dlm;
    Else
      '
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
    End If
    If CatDesc(4) <> "" Then
      '                     23                        24                             25
      Print #RptHandle, CatDesc(4); dlm; -fpcurrLicAmtDet(4).DoubleValue; dlm; CustRec.FeeLicBal5; dlm;
    Else
      '
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
    End If
  End If
  '                    26
  Print #RptHandle, TempTotBal; dlm;
  
  If NegFlag = False Then
    '                              27
    Print #RptHandle, fpcurrTotAmt.DoubleValue; dlm;
  Else
    '                              27
    Print #RptHandle, -fpcurrTotAmt.DoubleValue; dlm;
  End If
  '                        28
  Print #RptHandle, CustRec.AcctBal
  
  Close
  
  arBLAdjRpt.Show
  
End Sub

Private Sub PrintText()
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim ThisType$
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim x As Integer
  Dim NegFlag As Boolean
  Dim FF$
  
  FF$ = Chr$(12)
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close
  
  OpenBLCustFile CustHandle
  Get CustHandle, GCustNum, CustRec
  Close
  
  NegFlag = False
  
  If optOverBill.Value = True Then
    NegFlag = True
  End If
  
  ReportFile$ = "ARAdjRpt.PRN"
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  Print #RptHandle, Tab(29); "Adjustment Transaction"
  Print #RptHandle,
  Print #RptHandle, Tab(2); Date; Tab(14); Time
  If NegFlag = False Then
    Print #RptHandle, Tab(2); QPTrim$(TownRec.TownName); Tab(45); "Transaction Amount: "; Tab(66); Using$("$###,##0.00", fpcurrTotAmt.DoubleValue)
  ElseIf NegFlag = True Then
    Print #RptHandle, Tab(2); QPTrim$(TownRec.TownName); Tab(45); "Transaction Amount: "; Tab(66); Using$("$###,##0.00", -fpcurrTotAmt.DoubleValue)
  End If
  Print #RptHandle, Tab(2); String$(75, "-")
  Print #RptHandle,
  Print #RptHandle, Tab(12); "Customer Name:  " + QPTrim$(fptxtName.Text)
  Print #RptHandle, Tab(16); "Account #:  " + QPTrim$(fptxtAccount.Text)
  Print #RptHandle, Tab(18); "Address:  " + QPTrim$(fptxtAddress.Text)
  Print #RptHandle, Tab(20); "Notes:  " + QPTrim$(fptxtDesc.Text)
  Print #RptHandle,
  Print #RptHandle, Tab(10); "Adjustment Type:  " + TempThisCaption
  Print #RptHandle,
  Print #RptHandle, Tab(10); "Revenue Type"; Tab(46); "Adj Amt"; Tab(59); "New Balance"
  Print #RptHandle, Tab(10); "------------"; Tab(43); "----------"; Tab(59); "-----------"
  If NegFlag = False Then
    If CatDesc(0) <> "" Then
      Print #RptHandle, Tab(10); CatDesc(0); Tab(42); Using("$###,##0.00", fpcurrLicAmtDet(0).DoubleValue); Tab(59); Using("$###,##0.00", CustRec.FeeLicBal1)
    End If
    If CatDesc(1) <> "" Then
      Print #RptHandle, Tab(10); CatDesc(1); Tab(42); Using("$###,##0.00", fpcurrLicAmtDet(1).DoubleValue); Tab(59); Using("$###,##0.00", CustRec.FeeLicBal2)
    End If
    If CatDesc(2) <> "" Then
      Print #RptHandle, Tab(10); CatDesc(2); Tab(42); Using("$###,##0.00", fpcurrLicAmtDet(2).DoubleValue); Tab(59); Using("$###,##0.00", CustRec.FeeLicBal3)
    End If
    If CatDesc(3) <> "" Then
      Print #RptHandle, Tab(10); CatDesc(3); Tab(42); Using("$###,##0.00", fpcurrLicAmtDet(3).DoubleValue); Tab(59); Using("$###,##0.00", CustRec.FeeLicBal4)
    End If
    If CatDesc(4) <> "" Then
      Print #RptHandle, Tab(10); CatDesc(4); Tab(42); Using("$###,##0.00", fpcurrLicAmtDet(4).DoubleValue); Tab(59); Using("$###,##0.00", CustRec.FeeLicBal5)
    End If
  Else
    If CatDesc(0) <> "" Then
      Print #RptHandle, Tab(10); CatDesc(0); Tab(42); Using("$###,##0.00", -fpcurrLicAmtDet(0).DoubleValue); Tab(59); Using("$###,##0.00", CustRec.FeeLicBal1)
    End If
    If CatDesc(1) <> "" Then
      Print #RptHandle, Tab(10); CatDesc(1); Tab(42); Using("$###,##0.00", -fpcurrLicAmtDet(1).DoubleValue); Tab(59); Using("$###,##0.00", CustRec.FeeLicBal2)
    End If
    If CatDesc(2) <> "" Then
      Print #RptHandle, Tab(10); CatDesc(2); Tab(42); Using("$###,##0.00", -fpcurrLicAmtDet(2).DoubleValue); Tab(59); Using("$###,##0.00", CustRec.FeeLicBal3)
    End If
    If CatDesc(3) <> "" Then
      Print #RptHandle, Tab(10); CatDesc(3); Tab(42); Using("$###,##0.00", -fpcurrLicAmtDet(3).DoubleValue); Tab(59); Using("$###,##0.00", CustRec.FeeLicBal4)
    End If
    If CatDesc(4) <> "" Then
      Print #RptHandle, Tab(10); CatDesc(4); Tab(42); Using("$###,##0.00", -fpcurrLicAmtDet(4).DoubleValue); Tab(59); Using("$###,##0.00", CustRec.FeeLicBal5)
    End If
  End If
  If NegFlag = False Then
    If fpcurrPenAmt.DoubleValue > 0 Or CustRec.PenBal > 0 Then
      Print #RptHandle, Tab(10); "PENALTY"; Tab(42); Using("$###,##0.00", fpcurrPenAmt.DoubleValue); Tab(59); Using("$###,##0.00", CustRec.PenBal)
    End If
    If fpcurrIssFeeAmt.DoubleValue > 0 Or CustRec.IssuanceBal > 0 Then
      Print #RptHandle, Tab(10); "ISSUANCE"; Tab(42); Using("$###,##0.00", fpcurrIssFeeAmt.DoubleValue); Tab(59); Using("$###,##0.00", CustRec.IssuanceBal)
    End If
  Else
    If fpcurrPenAmt.DoubleValue > 0 Or CustRec.PenBal > 0 Then
      Print #RptHandle, Tab(10); "PENALTY"; Tab(42); Using("$###,##0.00", -fpcurrPenAmt.DoubleValue); Tab(59); Using("$###,##0.00", CustRec.PenBal)
    End If
    If fpcurrIssFeeAmt.DoubleValue > 0 Or CustRec.IssuanceBal > 0 Then
      Print #RptHandle, Tab(10); "ISSUANCE"; Tab(42); Using("$###,##0.00", -fpcurrIssFeeAmt.DoubleValue); Tab(59); Using("$###,##0.00", CustRec.IssuanceBal)
    End If
  End If
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(10); "----Account Balance Information-----"
  Print #RptHandle,
  Print #RptHandle, Tab(10); "Previous Balance: "; Tab(33); Using("$#,###,##0.00", TempTotBal)
  If NegFlag = False Then
    Print #RptHandle, Tab(10); "Current Adjustment: "; Tab(33); Using("$#,###,##0.00", fpcurrTotAmt.DoubleValue)
  Else
    Print #RptHandle, Tab(10); "Current Adjustment: "; Tab(33); Using("$#,###,##0.00", -fpcurrTotAmt.DoubleValue)
  End If
  Print #RptHandle, Tab(10); "Account Balance: "; Tab(33); Using("$#,###,##0.00", CustRec.AcctBal)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(36); "Signature:_____________________________"
  Print #RptHandle, FF$
  
  Close
  ViewPrint ReportFile$, "Balance Adjustment Report", True

End Sub

Private Sub GetCust()
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim Number$
  Dim Name$
  Dim Found As Boolean
  
'  On Error GoTo ERRORSTUFF
  
  If QPTrim$(fptxtAccount.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter a customer number."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  Number = QPTrim$(fptxtAccount.Text)
  
  OpenBLCustFile CHandle
  TotalAccts = LOF(CHandle) \ Len(CustRec)
  
  If TotalAccts = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no business customers saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  For x = 1 To TotalAccts
    Get CHandle, x, CustRec
    If Number$ = QPTrim$(CustRec.CUSTNUMB) Then 'match the selected
    'row with the right code
      Found = True
      GCustNum = x 'now you can assign the correct global
      Exit For
    Else
      Found = False
      GoTo NotAMatch
    End If
      
NotAMatch:
   Next x
  Close CHandle
  
  If Found = False Then
    frmBLMessageBoxJr.Label1.Caption = "The customer number entered does not match any of those saved. Please enter a valid customer number."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    NoMatchFoundFlag = True
  Else
    Call Clearscreen
    Call LoadMe
    If DeletedFlag = False Then
      If optOverBill.Enabled = True Then
        optOverBill.SetFocus
      End If
    Else
      DeletedFlag = True
    End If
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Unload FrmShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLAdjustBal", "cmdGetCust_Click", Erl)
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
  '  ClearInUse PWcnt
  '  CitiTerminate
  
End Sub

Private Sub LostFocusCheck()
  Dim ThisNum$
  
  ThisNum$ = QPTrim$(fptxtAccount.Text)
  If QPTrim$(fptxtAccount.Text) = "" Then
    Call Clearscreen
    Exit Sub
  End If
  
  If Check4ValidCustNum(QPTrim$(fptxtAccount.Text)) = False Then
    frmBLMessageBoxJr.Label1.Caption = "The customer number entered, " + ThisNum + ", is not valid. Please enter a valid customer number."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Call Clearscreen
    If optOverBill.Enabled = True Then
      optOverBill.SetFocus
    End If
    Exit Sub
  End If
  
  Call GetCust
  
  If NoMatchFoundFlag = True Then
    Call Clearscreen
    NoMatchFoundFlag = False
    If fptxtAccount.Enabled = True Then
      fptxtAccount.SetFocus
      Exit Sub
    End If
  End If
  
  If optOverBill.Enabled = True Then
    optOverBill.SetFocus
  End If

End Sub

Private Function Check4ValidCustNum(ThisCust As String) As Boolean
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim Number$
  Dim Name$
  Dim Found As Boolean

  Check4ValidCustNum = True
  
  
  If QPTrim$(fptxtAccount.Text) = "" Then
    Check4ValidCustNum = False
    Exit Function
  End If
  
  
  OpenBLCustFile CHandle
  TotalAccts = LOF(CHandle) \ Len(CustRec)

  If TotalAccts = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no business customers saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close CHandle
    Exit Function
  End If
  
  Get CHandle, CInt(fptxtAccount.Text), CustRec
  If CustRec.Deleted = "Y" Or QPTrim$(CustRec.SORTNAME) = "DELETED" Then
    Check4ValidCustNum = False
    Close
    Exit Function
  End If
  
  For x = 1 To TotalAccts
    Get CHandle, x, CustRec
    If ThisCust$ = QPTrim$(CustRec.CUSTNUMB) Then 'match the selected
    'row with the right code
      Exit For
    End If
  Next x

  Close CHandle

  If x > TotalAccts Then
    Call Clearscreen
    Check4ValidCustNum = False
  End If
  
End Function

