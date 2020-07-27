VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmVATaxDataRepair 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Repair"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11715
   Icon            =   "frmVATaxDataRepair.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11715
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdFixBadAdjBillsDown 
      Height          =   495
      Left            =   10080
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFindFixBillsWWrongTotals 
      Height          =   615
      Left            =   10080
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
      _ExtentY        =   1085
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":0AB6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixCreditZeros 
      Height          =   615
      Left            =   10080
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
      _ExtentY        =   1085
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":0CA4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMisc 
      Height          =   495
      Left            =   10080
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":0E8E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdOptRevInTaxBilling 
      Height          =   360
      Left            =   600
      TabIndex        =   34
      Top             =   7200
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":1065
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdListZeroCRecs 
      Height          =   360
      Left            =   3840
      TabIndex        =   28
      ToolTipText     =   "Files exported are: Cust Pin, County #, Personal Pin #, Personal Value, Property Type, PPTRA flag"
      Top             =   6600
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":1253
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixRealCustPins 
      Height          =   375
      Left            =   1080
      TabIndex        =   25
      Tag             =   "If real customer pin does not equal customer pin then run this utility."
      Top             =   4680
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":143C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess7 
      Height          =   360
      Left            =   4320
      TabIndex        =   0
      Tag             =   $"frmVATaxDataRepair.frx":1616
      Top             =   4080
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":171D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess4 
      Height          =   360
      Left            =   7440
      TabIndex        =   1
      Tag             =   "If customer pin numbers are not the same as the customer records then run this procedure. It will match them back up."
      Top             =   4440
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":18FA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   10200
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1215
      _Version        =   131072
      _ExtentX        =   2143
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":1AD8
      Begin fpBtnAtlLibCtl.fpBtn fpBtn1 
         Height          =   360
         Left            =   0
         TabIndex        =   40
         Top             =   1200
         Width           =   2775
         _Version        =   131072
         _ExtentX        =   4895
         _ExtentY        =   635
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
         ButtonDesigner  =   "frmVATaxDataRepair.frx":1CB4
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess2 
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Tag             =   $"frmVATaxDataRepair.frx":1E97
      Top             =   1635
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":1F78
   End
   Begin EditLib.fpDateTime fptxtFiscalBeg 
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   2040
      Width           =   1140
      _Version        =   196608
      _ExtentX        =   2011
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
      Text            =   "02/24"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd"
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
   Begin EditLib.fpDateTime fptxtFiscalEnd 
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   2760
      Width           =   1140
      _Version        =   196608
      _ExtentX        =   2011
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
      Text            =   "02/24"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess3 
      Height          =   360
      Left            =   7440
      TabIndex        =   6
      ToolTipText     =   $"frmVATaxDataRepair.frx":2156
      Top             =   3240
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":2244
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   450
      Left            =   120
      TabIndex        =   7
      Top             =   4800
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
      MaxWidth        =   5000
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
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess6 
      Height          =   360
      Left            =   4320
      TabIndex        =   8
      Tag             =   $"frmVATaxDataRepair.frx":2422
      Top             =   2880
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":2519
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess8 
      Height          =   360
      Left            =   4320
      TabIndex        =   9
      Tag             =   $"frmVATaxDataRepair.frx":26F7
      Top             =   5040
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":27FE
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess5 
      Height          =   360
      Left            =   1080
      TabIndex        =   10
      Tag             =   $"frmVATaxDataRepair.frx":29DB
      Top             =   2580
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":2B67
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCnvtPstedBills 
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Tag             =   "Use this utility if the posted bill data appears garbled."
      Top             =   1440
      Width           =   2295
      _Version        =   131072
      _ExtentX        =   4048
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":2D44
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess11 
      Height          =   375
      Left            =   1080
      TabIndex        =   23
      Tag             =   $"frmVATaxDataRepair.frx":2F2B
      Top             =   3600
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":300C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixRealPinsWOverPay 
      Height          =   360
      Left            =   3480
      TabIndex        =   27
      Top             =   6120
      Width           =   3135
      _Version        =   131072
      _ExtentX        =   5530
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":31E6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdRemovePenalties 
      Height          =   360
      Left            =   3840
      TabIndex        =   29
      Top             =   7080
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":33D5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdRunRealvsAcctBal 
      Height          =   360
      Left            =   3840
      TabIndex        =   30
      Top             =   7560
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":35BB
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFindCountyPin 
      Height          =   360
      Left            =   3840
      TabIndex        =   31
      Top             =   8040
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":37A6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixLunenburgPins 
      Height          =   360
      Left            =   600
      TabIndex        =   32
      Top             =   5280
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":3988
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPenGapMCtoPers 
      Height          =   360
      Left            =   600
      TabIndex        =   33
      Top             =   5760
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":3B6F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdRemoveRealPinSpaces 
      Height          =   360
      Left            =   600
      TabIndex        =   35
      Top             =   7680
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":3D55
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixLunenburgZeroYears 
      Height          =   360
      Left            =   600
      TabIndex        =   36
      Top             =   8160
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":3F3E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixAdTrans 
      Height          =   375
      Left            =   7440
      TabIndex        =   37
      ToolTipText     =   "Inserts real pin numbers where applicable"
      Top             =   5520
      Width           =   1575
      _Version        =   131072
      _ExtentX        =   2778
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":4129
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdRunningBal 
      Height          =   375
      Left            =   4080
      TabIndex        =   39
      Tag             =   "If customer pin numbers are not the same as the customer records then run this procedure. It will match them back up."
      Top             =   5640
      Width           =   2175
      _Version        =   131072
      _ExtentX        =   3836
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":4303
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdUpdateAdd1AndAdd2 
      Height          =   360
      Left            =   6840
      TabIndex        =   41
      Top             =   6120
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":44E5
      Begin fpBtnAtlLibCtl.fpBtn cmdUpdateAdd1andAdd2Short 
         Height          =   360
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   2775
         _Version        =   131072
         _ExtentX        =   4895
         _ExtentY        =   635
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
         ButtonDesigner  =   "frmVATaxDataRepair.frx":46CE
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdUpdateAdd1andAdd2Long 
      Height          =   360
      Left            =   6840
      TabIndex        =   43
      Top             =   6600
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":48B9
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixPersLinkToTrans 
      Height          =   360
      Left            =   600
      TabIndex        =   44
      Top             =   6240
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":4AA3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixHillsville 
      Height          =   360
      Left            =   600
      TabIndex        =   45
      Top             =   6720
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":4C86
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdRelinkBelongTosToBill 
      Height          =   375
      Left            =   6840
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   7080
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":4E67
      Begin fpBtnAtlLibCtl.fpBtn fpBtn3 
         Height          =   360
         Left            =   0
         TabIndex        =   47
         Top             =   600
         Width           =   2775
         _Version        =   131072
         _ExtentX        =   4895
         _ExtentY        =   635
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
         ButtonDesigner  =   "frmVATaxDataRepair.frx":5053
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCompareMBCH 
      Height          =   375
      Left            =   6840
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   7560
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":5236
      Begin fpBtnAtlLibCtl.fpBtn fpBtn4 
         Height          =   360
         Left            =   0
         TabIndex        =   49
         Top             =   600
         Width           =   2775
         _Version        =   131072
         _ExtentX        =   4895
         _ExtentY        =   635
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
         ButtonDesigner  =   "frmVATaxDataRepair.frx":541E
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExamineTrans 
      Height          =   375
      Left            =   10200
      TabIndex        =   50
      ToolTipText     =   "Inserts real pin numbers where applicable"
      Top             =   240
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":5601
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixCustRecAndPins 
      Height          =   375
      Left            =   6840
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   8040
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":57E1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixCrossedTrans 
      Height          =   615
      Left            =   10080
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
      _ExtentY        =   1085
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":59CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixOPatBilling 
      Height          =   375
      Left            =   6840
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   8040
      Width           =   2775
      _Version        =   131072
      _ExtentX        =   4895
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":5BB5
      Begin fpBtnAtlLibCtl.fpBtn fpBtn5 
         Height          =   360
         Left            =   0
         TabIndex        =   54
         Top             =   600
         Width           =   2775
         _Version        =   131072
         _ExtentX        =   4895
         _ExtentY        =   635
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
         ButtonDesigner  =   "frmVATaxDataRepair.frx":5D95
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGetBelongTos 
      Height          =   495
      Left            =   10080
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":5F78
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixInitialized 
      Height          =   495
      Left            =   10080
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":6159
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDiagAndFix 
      Height          =   615
      Left            =   10080
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
      _ExtentY        =   1085
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":633B
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixZeroedIntAndPen 
      Height          =   495
      Left            =   10080
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":6524
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFixRestOfTrans 
      Height          =   495
      Left            =   10080
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   6360
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
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
      ButtonDesigner  =   "frmVATaxDataRepair.frx":670F
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Insert Real Pin To Adv Trans"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6720
      TabIndex        =   38
      ToolTipText     =   "Use this utility only if the transaction journal release report is not displaying revenues correctly."
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Shape Shape15 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   975
      Left            =   6720
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Shape Shape13 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   975
      Left            =   360
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Fix Real Cust Pin "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   26
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Make Bill Type ""R"", ""P"", or Blank"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   24
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   975
      Left            =   360
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax DOS Data Repair"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3120
      TabIndex        =   22
      Top             =   330
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1500
      Top             =   165
      Width           =   8655
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   1095
      Left            =   3720
      Top             =   1020
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Make Zero Value Years Equal To Bill Trans Tax Years"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3720
      TabIndex        =   21
      Top             =   1020
      Width           =   2775
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   2655
      Left            =   6720
      Top             =   1020
      Width           =   3135
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Make Zero Value Tax Years Correspond to Its Trans Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6720
      TabIndex        =   20
      Top             =   1020
      Width           =   3135
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Beginning Fiscal Year Date:"
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
      Left            =   6840
      TabIndex        =   19
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Fiscal Year Date:"
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
      Left            =   6960
      TabIndex        =   18
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Make Pin Numbers and Acct Numbers The Same"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   6720
      TabIndex        =   17
      Top             =   3780
      Width           =   3135
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   1215
      Left            =   6720
      Top             =   3780
      Width           =   3135
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   1095
      Left            =   3720
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Bill Trans Only: Make Paid Equal Charged If Paid Is More"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3720
      TabIndex        =   16
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Fix Accumulated BelongTo Trans More Than Bill Itself"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3720
      TabIndex        =   15
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   1095
      Left            =   3720
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   855
      Left            =   3720
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Relink Posting Errors"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3720
      TabIndex        =   14
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   1095
      Left            =   360
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Reconstruct history by eliminating faulty negative balances"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   360
      TabIndex        =   13
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Convert Posted Bills"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   12
      Top             =   1020
      Width           =   3135
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      FillColor       =   &H0080FFFF&
      Height          =   855
      Left            =   360
      Top             =   1020
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1500
      Top             =   60
      Width           =   8655
   End
End
Attribute VB_Name = "frmVATaxDataRepair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim CArr() As Long
  Dim CArrCnt As Long
  Dim CrossArr() As Long
  Dim CrossCnt As Long
  Dim CrossGoodArr() As Long
  Dim CrossBadArr() As Long

Private Sub cmdCompareMBCH_Click()
  Call CompareMBWithCustHistory
End Sub

Private Sub cmdDiagAndFix_Click()
  Call RelinkBelongTosWithBills
  Call FindCorrectQueue
  Call CompareMBWithCustHistory(True)
  Call FixInitializedTrans
  Call CompareMBWithCustHistory(True)
  Call FixZeroedOutCreditAtBilling
  Call CompareMBWithCustHistory(True)
  Call FindAndFixZeroInterestCausedIssues
  Call CompareMBWithCustHistory(True)
  Call FindAndFixBillsWithWrongTotals
  Call CompareMBWithCustHistory(True)
  Call FixBillsWithAdjustsMoreThanTotals
  Call CompareMBWithCustHistory(True)
  Call FixFinalFew
  Call CompareMBWithCustHistory
  MsgBox ("All fixes have been completed.")
End Sub
Private Sub MakeAllRecsAndPinsTheSameAsCurrentRec()
  Dim TaxTrans As TaxTransactionType
  Dim NumOfTTRecs As Long
  Dim TTHandle As Integer
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim NextRec As Long
  Dim x As Long
  Dim cnt As Integer
  Dim AHandle As Integer
  
  AHandle = FreeFile
  Open "correctedrecsandpins.txt" For Output As AHandle
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo Skip
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
      If TaxTrans.CustomerRec <> x Then
        TaxTrans.CustomerRec = x
        Put TTHandle, NextRec, TaxTrans
        Print #AHandle, CStr(x) + "  CustomerRec"
        cnt = cnt + 1
      End If
       If TaxTrans.CustPin <> x Then
        TaxTrans.CustPin = x
        Put TTHandle, NextRec, TaxTrans
        cnt = cnt + 1
         Print #AHandle, CStr(x) + "  CustomerRec"
     End If
     
      NextRec = TaxTrans.LastTrans
    Loop
Skip:
  Next x
  Close
  MsgBox ("A total of " + CStr(cnt) + " customer records were corrected. Look for 'correctedrecsandpins.txt' in the Citipak folder for results.")
End Sub
Private Sub cmdExamineTrans_Click()
  Call ExamineATrans
End Sub

Private Sub cmdExit_Click()
  frmVATaxMainMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdExportFiles_Click()
Call ExportFiles
End Sub
Private Sub BuildTextFile()
Dim AHandle As Integer
Dim PersRec As PersonalRecType
Dim PHandle As Integer
Dim NumOfPersRecs As Long
Dim x As Integer, y As Integer
Dim PTaxBill As VAPPTaxBillType
Dim TBHandle As Integer
Dim NumOfTBRecs As Long
Dim TaxCust As TaxCustType
Dim TCHandle As Integer
Dim NumOfTCRecs As Long
Dim NextRec As Long
Dim SName As String
AHandle = FreeFile
Open "Remington.txt" For Output As AHandle
Print #AHandle, "Search Name~Cust Acct #~Prop Pin #~Prop Date"
OpenPersTaxBillFile TBHandle, NumOfTBRecs
OpenPersPropFile PHandle, NumOfPersRecs
OpenTaxCustFile TCHandle, NumOfTCRecs
frmVATaxShowPctComp.Label1 = "Creating Text File"
frmVATaxShowPctComp.Show , Me
For x = 1 To NumOfTCRecs
 Get TCHandle, x, TaxCust
 NextRec = TaxCust.FirstPersRec
 SName = QPTrim$(TaxCust.OptSrchDesc)
 If SName = "" Then SName = "No Search Name"
 Do While NextRec > 0
   Get PHandle, NextRec, PersRec
    Print #AHandle, SName + "~" + CStr(TaxCust.Acct) + "~" + QPTrim$(PersRec.PropPin) + "~" + MakeRegDate(PersRec.PROPDATE)
   NextRec = PersRec.NextRec
 Loop
 frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
Next x
Unload frmVATaxShowPctComp
Close
MsgBox ("Done. Look for the file 'Remington.txt'.")

End Sub

Private Sub cmdFindBadCreditAtBill_Click()
  Call Check4CreditAtBillingWithNoPrepayment
End Sub

Private Sub cmdFindCountyPin_Click()
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Integer
  
  OpenTaxCustFile CHandle, NumOfCRecs
  For x = 1 To NumOfCRecs
    Get CHandle, x, TaxCust
'    If x = 1781 Then Stop
    TaxCust.Deleted = TaxCust.Deleted
    If QPTrim$(TaxCust.CountyAcctString) = "26" Then Stop
  Next x
  
  Close
  MsgBox ("Finished")
  
End Sub

Private Sub cmdFindFixBillsWWrongTotals_Click()
  Call FindAndFixBillsWithWrongTotals
End Sub

Private Sub cmdFixAdTrans_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim RealPin As String
  Dim cnt As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If TaxTrans.TranType = 6 Then
      If QPTrim$(TaxTrans.RealPin) = "" Then
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
          RealPin = QPTrim$(TaxTrans.RealPin)
          Get TTHandle, x, TaxTrans
          TaxTrans.RealPin = RealPin
          Put TTHandle, x, TaxTrans
          cnt = cnt + 1
        End If
      End If
    End If
  Next x
  Close
  Call Savemsg(900, "A total of " + CStr(cnt) + " transactions were modified successfully.")

End Sub

Private Sub cmdFixCedarBluff_Click()
  Dim TransRec As TaxTransactionType
  Dim NumOfTRecs As Long
  Dim THandle As Integer
  
  'fix for cust #72
  OpenTaxTransFile THandle, NumOfTRecs
  Get THandle, 13128, TransRec
  TransRec.Revenue.Principle1Pd = 0.44
  Put THandle, 13128, TransRec
  
  Get THandle, 14687, TransRec
  TransRec.Revenue.Principle1Pd = 0.44
  TransRec.Amount = 2.15
  Put THandle, 14687, TransRec
  
  'fix for cust #592
  Get THandle, 5494, TransRec
  TransRec.Revenue.Principle1Pd = 4.56
  Put THandle, 5494, TransRec
  
  Get THandle, 6487, TransRec
  TransRec.Revenue.Principle1Pd = 4.56
  TransRec.Amount = 17.67
  Put THandle, 6487, TransRec
  
  'fix for cust #95
  Get THandle, 4961, TransRec
  TransRec.Revenue.Principle1 = 0.01
  Put THandle, 4961, TransRec
  
  'fix for cust #320
  Get THandle, 10905, TransRec
  TransRec.Revenue.Principle1 = 0.01
  Put THandle, 10905, TransRec
  
  'fix for cust #530
  Get THandle, 1262, TransRec
  TransRec.Revenue.Principle1Pd = 13.19
  Put THandle, 1262, TransRec
  
  Get THandle, 1928, TransRec
  TransRec.Revenue.Principle1Pd = 13.19
  TransRec.Amount = 15.47
  Put THandle, 1928, TransRec
  
  Get THandle, 3555, TransRec
  TransRec.Revenue.Principle1Pd = 7.19
  Put THandle, 3555, TransRec
  
  Get THandle, 4441, TransRec
  TransRec.Revenue.Principle1Pd = 7.19
  TransRec.Amount = 9.47
  Put THandle, 4441, TransRec
  
  Get THandle, 13566, TransRec
  TransRec.Revenue.Principle1Pd = 12.41
  Put THandle, 13566, TransRec
  
  Get THandle, 14678, TransRec
  TransRec.Revenue.Principle1Pd = 12.41
  TransRec.Amount = 14.69
  Put THandle, 14678, TransRec
  
  'fix for cust #109
  Get THandle, 8136, TransRec
  TransRec.Revenue.Principle1 = 0.01
  Put THandle, 8136, TransRec
  
  'fix for cust #164
  Get THandle, 5036, TransRec
  TransRec.Revenue.Principle1 = 0.01
  Put THandle, 5036, TransRec
  
  'fix for cust #358
  Get THandle, 10942, TransRec
  TransRec.Revenue.Principle1 = 0.01
  Put THandle, 10942, TransRec
  
  'fix for cust #540
  Get THandle, 5427, TransRec
  TransRec.Revenue.Principle1Pd = 0
  Put THandle, 5427, TransRec
  
  Get THandle, 7265, TransRec
  TransRec.Revenue.Principle1Pd = 0
  TransRec.Amount = 72.8
  Put THandle, 7265, TransRec
  
  'fix for cust #1211
  Get THandle, 7407, TransRec
  TransRec.Revenue.Principle1Pd = 303.05
  Put THandle, 7407, TransRec
  
  Close
  
  Call TaxMsg(900, "Finished.")
  
  
End Sub

Private Sub cmdFixChilhowie_Click()
  Dim TransRec As TaxTransactionType
  Dim NumOfTRecs As Long
  Dim THandle As Integer
  
  OpenTaxTransFile THandle, NumOfTRecs
'fix for #2013
  Get THandle, 272, TransRec
  TransRec.Revenue.Principle1 = 3.46
  TransRec.Revenue.Penalty = 0.03
  Put THandle, 272, TransRec
  Close
  
  MsgBox ("Finished.")
End Sub

Private Sub cmdFixDumfries_Click()
  Dim TransRec As TaxTransactionType
  Dim NumOfTRecs As Long
  Dim THandle As Integer
  
  OpenTaxTransFile THandle, NumOfTRecs
  'fix for cust# 376
  Get THandle, 3866, TransRec
  TransRec.Revenue.Principle1Pd = 20.99
  Put THandle, 3866, TransRec

  'fix for cust# 225
  Get THandle, 3962, TransRec
  TransRec.Revenue.Principle1Pd = 75.14
  Put THandle, 3962, TransRec

  Get THandle, 7617, TransRec
  TransRec.Amount = 0
  TransRec.Revenue.Principle1Pd = 0
  Put THandle, 7617, TransRec

  'fix for cust# 409
  Get THandle, 3963, TransRec
  TransRec.Revenue.Principle1Pd = 90.96
  Put THandle, 3963, TransRec

  Get THandle, 7618, TransRec
  TransRec.Amount = 0
  TransRec.Revenue.Principle1Pd = 0
  Put THandle, 7618, TransRec

  'fix for cust# 967
  Get THandle, 3959, TransRec
  TransRec.Revenue.Principle1Pd = 74.33
  Put THandle, 3959, TransRec
  
  Get THandle, 7614, TransRec
  TransRec.Amount = 0
  TransRec.Revenue.Principle1Pd = 0
  Put THandle, 7614, TransRec

  'fix for cust# 1533
  Get THandle, 4676, TransRec
  TransRec.Revenue.Principle1Pd = 76.19
  Put THandle, 4676, TransRec

  Get THandle, 7620, TransRec
  TransRec.Amount = 0
  TransRec.Revenue.Principle1Pd = 0
  Put THandle, 7620, TransRec

  'fix for cust# 753
  Get THandle, 3960, TransRec
  TransRec.Revenue.Principle1Pd = 46.32
  Put THandle, 3960, TransRec

  Get THandle, 7615, TransRec
  TransRec.Amount = 0
  TransRec.Revenue.Principle1Pd = 0
  Put THandle, 7615, TransRec

  'fix for cust# 2461
  Get THandle, 3961, TransRec
  TransRec.Revenue.Principle1Pd = 29.19
  Put THandle, 3961, TransRec

  Get THandle, 7616, TransRec
  TransRec.Amount = 0
  TransRec.Revenue.Principle1Pd = 0
  Put THandle, 7616, TransRec

  'fix for cust# 2257
  Get THandle, 5178, TransRec
  TransRec.Revenue.Principle1Pd = 66.24
  Put THandle, 5178, TransRec
  
  'fix for cust# 693
  Get THandle, 3964, TransRec
  TransRec.Revenue.Principle1Pd = 80.29
  Put THandle, 3964, TransRec

  Get THandle, 6054, TransRec
  TransRec.Amount = 0
  TransRec.Revenue.Principle1Pd = 0
  Put THandle, 6054, TransRec

  Close
  Call TaxMsg(900, "Finished.")
End Sub


Private Sub cmdFixIndependence_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  'fix for #233
  Get TTHandle, 819, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.PenaltyPd = 0
  Put TTHandle, 819, TaxTrans
  Get TTHandle, 41, TaxTrans
  TaxTrans.Revenue.InterestPd = 0.43
  TaxTrans.Revenue.PenaltyPd = 2.04
  Put TTHandle, 41, TaxTrans
  
  'fix for #448
  Get TTHandle, 820, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.PenaltyPd = 0
  Put TTHandle, 820, TaxTrans
  Get TTHandle, 104, TaxTrans
  TaxTrans.Revenue.InterestPd = 0.6
  TaxTrans.Revenue.PenaltyPd = 2.86
  Put TTHandle, 104, TaxTrans
  
  'fix for #898
  Get TTHandle, 823, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.PenaltyPd = 0
  Put TTHandle, 823, TaxTrans
  Get TTHandle, 218, TaxTrans
  TaxTrans.Revenue.InterestPd = 0.08
  TaxTrans.Revenue.PenaltyPd = 1.8
  Put TTHandle, 218, TaxTrans
  
  'fix for #899
  Get TTHandle, 219, TaxTrans
  TaxTrans.Revenue.InterestPd = 0.04
  TaxTrans.Revenue.PenaltyPd = 0.9
  Put TTHandle, 219, TaxTrans
  
  'fix for #900
  Get TTHandle, 220, TaxTrans
  TaxTrans.Revenue.InterestPd = 0.08
  TaxTrans.Revenue.PenaltyPd = 1.5
  Put TTHandle, 220, TaxTrans
  
  'fix for #901
  Get TTHandle, 221, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 15
  TaxTrans.Revenue.InterestPd = 0.68
  TaxTrans.Revenue.PenaltyPd = 3.75
  Put TTHandle, 221, TaxTrans
  
  'fix for #902
  Get TTHandle, 222, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 12
  TaxTrans.Revenue.InterestPd = 0.52
  TaxTrans.Revenue.PenaltyPd = 3
  Put TTHandle, 222, TaxTrans
  
  Close TTHandle
  
  Call TaxMsg(900, "Finished.")
End Sub

Private Sub cmdFixKenbridge_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  
  'fix for cust# 596
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, 13760, TaxTrans
    TaxTrans.Revenue.Penalty = 0
    TaxTrans.Revenue.PenaltyPd = 11.45
    
  Put TTHandle, 13760, TaxTrans
  Close TTHandle
  Call TaxMsg(900, "Finished.")

End Sub

Private Sub cmdFixAppalachia_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, 19386, TaxTrans
  TaxTrans.CustomerRec = 2178
  TaxTrans.CustPin = 2178
  Put TTHandle, 19386, TaxTrans
  
  Close
  MsgBox ("Done")

End Sub

Private Sub cmdFixHonaker_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim ThisDate As Integer
  Dim cnt As Integer
  ThisDate = Date2Num("9/19/2008")
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If TaxTrans.TranType = 1 Then
      If TaxTrans.TransDate = ThisDate Then
        TaxTrans.Amount = 0
        TaxTrans.FromPrePay = 0
        TaxTrans.Revenue.Collection = 0
        TaxTrans.Revenue.Interest = 0
        TaxTrans.Revenue.LateList = 0
        TaxTrans.Revenue.Penalty = 0
        TaxTrans.Revenue.PrePaidAmt = 0
        TaxTrans.Revenue.PrePaidBal = TaxTrans.Revenue.PrePaidBal + TaxTrans.Revenue.PrePaidUsed
        TaxTrans.Revenue.PrePaidUsed = 0
        TaxTrans.Revenue.Principle1 = 0
        TaxTrans.Revenue.Principle2 = 0
        TaxTrans.Revenue.Principle3 = 0
        TaxTrans.Revenue.Principle4 = 0
        TaxTrans.Revenue.Principle5 = 0
        TaxTrans.Revenue.RevOpt1 = 0
        TaxTrans.Revenue.RevOpt2 = 0
        TaxTrans.Revenue.RevOpt3 = 0
        Put TTHandle, x, TaxTrans
        cnt = cnt + 1
      End If
    End If
  Next x
  Close
  MsgBox ("Completed zeroing out " & CStr(cnt) & " billing transactions posted on 9/19/2008.")
End Sub

Private Sub cmdFixLawrenceville_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  
  'fix for cust# 596
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, 1799, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 41.04
  TaxTrans.Revenue.InterestPd = 21.24
  TaxTrans.Revenue.PenaltyPd = 6.4
  Put TTHandle, 1799, TaxTrans
  
  'fix for cust #1284
  Get TTHandle, 299, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 1.18
  TaxTrans.Revenue.PenaltyPd = 0
  Put TTHandle, 299, TaxTrans
  
  Get TTHandle, 1525, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 1.18
  TaxTrans.Revenue.PenaltyPd = 0
  Put TTHandle, 1525, TaxTrans
  
  Get TTHandle, 1528, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 0.59
  TaxTrans.Revenue.PenaltyPd = 0
  Put TTHandle, 1528, TaxTrans
  
  Get TTHandle, 1531, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 0.59
  TaxTrans.Revenue.PenaltyPd = 0
  Put TTHandle, 1531, TaxTrans
  
  'fix for cust #726
  Get TTHandle, 298, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 0.95
  TaxTrans.Revenue.PenaltyPd = 0
  Put TTHandle, 298, TaxTrans
  
  Get TTHandle, 1527, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 0.95
  TaxTrans.Revenue.PenaltyPd = 0
  Put TTHandle, 1527, TaxTrans
  
  Get TTHandle, 1524, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 0.95
  TaxTrans.Revenue.PenaltyPd = 0
  Put TTHandle, 1524, TaxTrans
  
  Close TTHandle
  Call TaxMsg(900, "Finished.")
End Sub

Private Sub cmdFixLunenburg_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  
  
'  Call FixPayPlusOPThatShouldHaveBeenApplied '2/6/09
  Exit Sub
'  OpenTaxTransFile TTHandle, NumOfTTRecs
  
'  'fix for 12662
'  Get TTHandle, 52444, TaxTrans
'  TaxTrans.BillType = "P"
'  Put TTHandle, 52444, TaxTrans
'
'  'fix for 15978
'  Get TTHandle, 52064, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.Penalty = 0
'  TaxTrans.Revenue.PenaltyPd = 0
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  TaxTrans.Revenue.RevOpt1 = 0
'  Put TTHandle, 52064, TaxTrans
'
'  Get TTHandle, 52065, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  TaxTrans.Revenue.RevOpt1 = 0
'  TaxTrans.Revenue.PenaltyPd = 0
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  Put TTHandle, 52065, TaxTrans
'
'  Get TTHandle, 43654, TaxTrans
'  TaxTrans.Amount = 11
'  TaxTrans.Revenue.RevOpt3Pd = 11
'  TaxTrans.TaxYear = 2006
'  Put TTHandle, 43654, TaxTrans
'
'  Get TTHandle, 2148, TaxTrans
'  TaxTrans.PPTRADisc = 0
'  TaxTrans.Amount = 17.7
'  TaxTrans.Revenue.Principle1 = 17.7
'  TaxTrans.Revenue.Principle1Pd = 10
'  TaxTrans.Revenue.Penalty = 1
'  TaxTrans.Revenue.PenaltyPd = 1
'  Put TTHandle, 2148, TaxTrans
'
'  Get TTHandle, 52066, TaxTrans
'  TaxTrans.PPTRADisc = 0
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Penalty = 0
'  TaxTrans.Revenue.PenaltyPd = 0
'  Put TTHandle, 52066, TaxTrans
'
'  Get TTHandle, 43684, TaxTrans
'  TaxTrans.PPTRADisc = 0
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Penalty = 0
'  TaxTrans.Revenue.PenaltyPd = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  Put TTHandle, 43684, TaxTrans
  
  
  'fix for #11906
'  Get TTHandle, 95304, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  Put TTHandle, 95304, TaxTrans
'
'  Get TTHandle, 95302, TaxTrans
'  TaxTrans.TranType = 14
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  TaxTrans.Revenue.RevOpt1 = 0
'  TaxTrans.Amount = 0
'  Put TTHandle, 95302, TaxTrans
'
'  Get TTHandle, 21179, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  TaxTrans.Revenue.RevOpt1 = 0
'  Put TTHandle, 21179, TaxTrans
'
'  Get TTHandle, 21902, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  TaxTrans.Revenue.RevOpt1 = 0
'  Put TTHandle, 21902, TaxTrans
'
'  Get TTHandle, 21169, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  TaxTrans.Revenue.RevOpt1 = 0
'  Put TTHandle, 21169, TaxTrans
'
'  Get TTHandle, 21176, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  TaxTrans.Revenue.RevOpt1 = 0
'  Put TTHandle, 21176, TaxTrans
'
'  Get TTHandle, 21909, TaxTrans
'  TaxTrans.Amount = 7.5
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 7.5
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  TaxTrans.Revenue.RevOpt1 = 0
'  Put TTHandle, 21909, TaxTrans
'
'  Get TTHandle, 4169, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 185.19
'  TaxTrans.Revenue.Principle1 = 427.88
'  Put TTHandle, 4169, TaxTrans
'
'  Call StripOutTrans
  
  
  'fix for #11623
'  Get TTHandle, 72993, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  Put TTHandle, 72993, TaxTrans
'
'  Get TTHandle, 72992, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  Put TTHandle, 72992, TaxTrans
'
'  Get TTHandle, 72991, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  Put TTHandle, 72991, TaxTrans
'
'  Get TTHandle, 72990, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  Put TTHandle, 72990, TaxTrans
'
'  'fix for 15497
'  Get TTHandle, 72608, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  TaxTrans.Revenue.RevOpt1 = 0
'  TaxTrans.Revenue.RevOpt2Pd = 0
'  TaxTrans.Revenue.RevOpt2 = 0
'  TaxTrans.Revenue.RevOpt3Pd = 0
'  TaxTrans.Revenue.RevOpt3 = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  Put TTHandle, 72608, TaxTrans
'
'  Get TTHandle, 72601, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  TaxTrans.Revenue.RevOpt1 = 0
'  TaxTrans.Revenue.RevOpt2Pd = 0
'  TaxTrans.Revenue.RevOpt2 = 0
'  TaxTrans.Revenue.RevOpt3Pd = 0
'  TaxTrans.Revenue.RevOpt3 = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  Put TTHandle, 72601, TaxTrans
'
'  Get TTHandle, 72600, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  TaxTrans.Revenue.RevOpt1 = 0
'  TaxTrans.Revenue.RevOpt2Pd = 0
'  TaxTrans.Revenue.RevOpt2 = 0
'  TaxTrans.Revenue.RevOpt3Pd = 0
'  TaxTrans.Revenue.RevOpt3 = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  Put TTHandle, 72600, TaxTrans
'
'  Get TTHandle, 72599, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  TaxTrans.Revenue.RevOpt1 = 0
'  TaxTrans.Revenue.RevOpt2Pd = 0
'  TaxTrans.Revenue.RevOpt2 = 0
'  TaxTrans.Revenue.RevOpt3Pd = 0
'  TaxTrans.Revenue.RevOpt3 = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  Put TTHandle, 72599, TaxTrans
'
'  Get TTHandle, 20799, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 141.96
'  TaxTrans.Amount = 141.96
'  Put TTHandle, 20799, TaxTrans
'  'fix for cust #16155
'  Get TTHandle, 51357, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.PenaltyPd = 0
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  Put TTHandle, 51357, TaxTrans
'
'  Get TTHandle, 7039, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 142.84
'  TaxTrans.Revenue.PenaltyPd = 19.28
'  TaxTrans.Revenue.RevOpt1Pd = 50#
'  Put TTHandle, 7039, TaxTrans
  
  Close
  Call TaxMsg(900, "Finished.")

End Sub

Private Sub cmdFixHillsville_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim ThisDate As String
  ThisDate = Date
 
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Call InsertCreditAtBillingTrans(ThisDate, 15.49, 4372, 2010, 14.63, 66260, "P", 73608, 66260, "1405")
'  Get TTHandle, 66260, TaxTrans
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.PenaltyPd = 0
'  Put TTHandle, 66260, TaxTrans
  
  'fix for #1075
'  Get TTHandle, 17749, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put TTHandle, 17749, TaxTrans
  
  Close
  MsgBox ("Done.")
  
End Sub

Private Sub cmdFixBadAdjBillsDown_Click()
  Call FixBillsWithAdjustsMoreThanTotals
End Sub

Private Sub cmdFixCreditZeros_Click()
   Call FixZeroedOutCreditAtBilling
End Sub

Private Sub cmdFixCrossedTrans_Click()
'  Call FindAndFixTransWithCrossCusts
  Call FindCorrectQueue
End Sub

Private Sub cmdFixCustRecAndPins_Click()
  Call MakeAllRecsAndPinsTheSameAsCurrentRec
End Sub

Private Sub cmdFixGrundy_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim BelongTo As Long
  Dim cnt As Integer
  Dim AHandle As Integer
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim OldDesc As String
  Dim NextRec As Long
  Dim BillNum As String
  Dim IntTot As Double
  Dim PenTot As Double
  Dim SaveRec As Long
  Dim NextRecToo As Long
'  Dim ArrString As String
'  ArrString = "161421, 161422, 162512, 162513, 164428, 164429, 175198, "
'  ArrString = ArrString + "175199, 176398, 176399, 177567, 177568, 178736, 178737, "
'  ArrString = ArrString + "179905, 179906, 181074, 181075, 182243, 182244, 183412, "
'  ArrString = ArrString + "183413, 184581, 184582, 185750, 185751, 186919, 186920, "
'  Dim CArr() As Long
'  Call BuildArray(ArrString, CArr(), cnt)
'
'  For x = 1 To cnt
'    ClearTrans (CArr(x))
'  Next x
  
'  cnt = 0
  'fix for 1350 on 6/9/2010
  OpenTaxCustFile CHandle, NumOfCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  Get CHandle, 1350, TaxCust
  NextRec = TaxCust.LastTrans
  Do While NextRec > 0
    Get TTHandle, NextRec, TaxTrans
    If TaxTrans.BelongTo > 0 Then
      Get TTHandle, TaxTrans.BelongTo, TaxTrans
        BillNum = ParseBillNum(TaxTrans.Description)
      Get TTHandle, NextRec, TaxTrans
      
        TaxTrans.Description = TaxTrans.Description
        TaxTrans.Description = "Applied to: " + BillNum
        TaxTrans.Description = TaxTrans.Description
      Put TTHandle, NextRec, TaxTrans
    End If
    NextRec = TaxTrans.LastTrans
  Loop
  
'  NextRec = TaxCust.LastTrans
'  SaveRec = NextRec
'  Do While NextRec > 0
'    Get TTHandle, NextRec, TaxTrans
'    If TaxTrans.TranType = 1 Then
'      NextRecToo = SaveRec
'      IntTot = 0
'      PenTot = 0
'      Do While NextRecToo > 0
'        Get TTHandle, NextRecToo, TaxTrans
'          If TaxTrans.BelongTo = NextRec Then
'            If TaxTrans.TranType = 4 Then
'              IntTot = IntTot + TaxTrans.Revenue.Interest
'            ElseIf TaxTrans.TranType = 5 Then
'              PenTot = PenTot + TaxTrans.Revenue.Penalty
'            End If
'          End If
'        NextRecToo = TaxTrans.LastTrans
'      Loop
'    End If
'    Get TTHandle, NextRec, TaxTrans
'    NextRec = TaxTrans.LastTrans
'  Loop
  
  
  Get TTHandle, 477, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.Principle2Pd = 0
  TaxTrans.Revenue.Principle3Pd = 0
  TaxTrans.Revenue.Principle4Pd = 0
  TaxTrans.Revenue.Principle5Pd = 0
  TaxTrans.Revenue.RevOpt1Pd = 0
  TaxTrans.Revenue.RevOpt2Pd = 0
  TaxTrans.Revenue.RevOpt2Pd = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.Revenue.LateListPd = 0
  TaxTrans.Revenue.PenaltyPd = 0
  TaxTrans.Revenue.Interest = 39.77
  Put TTHandle, 477, TaxTrans
'  AHandle = FreeFile
'  Open "grundy.dat" For Output As AHandle
'  frmVATaxShowPctComp.Label1 = "Grundy Bill Zero Update"
'  frmVATaxShowPctComp.Show , Me
'  'changing all zero bills to allow for release 12/16/2009
'  Print #AHandle, "Trans Type ~ + Cust # ~ Trans # ~ Name ~ County # ~ Tax Year ~ Old Bill"
'  For x = 1 To NumOfTTRecs
'    Get TTHandle, x, TaxTrans
'    If InStr(1, TaxTrans.Description, "Tax Bill 0") > 0 Then
'      OldDesc = TaxTrans.Description
'      TaxTrans.Description = "Tax Bill " + ReplaceString(MakeRegDate(TaxTrans.TransDate), "/", "")
'      Put TTHandle, x, TaxTrans
'      Get CHandle, TaxTrans.CustomerRec, CustTrans
'      Print #AHandle, CStr(TaxTrans.TranType) + "~" + CStr(TaxTrans.CustomerRec) + "~" + CStr(x) + "~" + CustTrans.CustName + "~" + CStr(CustTrans.CountyAcct) + "~" + CStr(TaxTrans.TaxYear) + "~" + OldDesc
'      If TaxTrans.TranType = 1 Then cnt = cnt + 1
'    End If
'    frmVATaxShowPctComp.ShowPctComp x, NumOfTTRecs
'    If frmVATaxShowPctComp.Out = True Then
'      Close
'      frmVATaxShowPctComp.Out = False
'      Unload frmVATaxShowPctComp
'      Exit Sub
'    End If
'
'  Next x
'  Unload frmVATaxShowPctComp

  Close
  MsgBox ("Fix for Grundy is done.")
'  MsgBox ("A total of " + CStr(cnt) + " bills were changed from zero to the transaction date. See 'grundy.dat' in Citipak for spreadsheet data.")
  
'  ClearTrans (661)
'  For x = 1 To NumOfTTRecs
'    Get TTHandle, x, TaxTrans
'    If TaxTrans.BelongTo = 359 Then
'      TaxTrans.Description = "Bill #2"
'      Put TTHandle, x, TaxTrans
'    End If
'    If TaxTrans.BelongTo = 357 Then
'      TaxTrans.Description = "Bill #1"
'      Put TTHandle, x, TaxTrans
'    End If
'  Next x
'  Get TTHandle, 359, TaxTrans
'  TaxTrans.Description = "Bill #2"
'  Put TTHandle, 359, TaxTrans
'
'  Get TTHandle, 357, TaxTrans
'  TaxTrans.Description = "Bill #1"
'  Put TTHandle, 357, TaxTrans
  
  'fix for 1583 on 3/4/09
'  Get TTHandle, 824, TaxTrans
'  TaxTrans.Revenue.InterestPd = 8.63
'  Put TTHandle, 824, TaxTrans
'
'  'fix for 1532 on 3/4/09
'  For x = 1 To NumOfTTRecs
'  Get TTHandle, x, TaxTrans
'    If TaxTrans.CustomerRec = 1532 Then
'    If x = 133381 Then TaxTrans.TaxYear = 2005
'    If x = 129508 Then TaxTrans.TaxYear = 2004
'    If x = 129507 Then TaxTrans.TaxYear = 2003
'    If x = 129506 Then TaxTrans.TaxYear = 2002
'    If x = 129505 Then TaxTrans.TaxYear = 2001
'    If x = 129504 Then TaxTrans.TaxYear = 2000
'    If x = 129503 Then TaxTrans.TaxYear = 1999
'    Put TTHandle, x, TaxTrans
'    If TaxTrans.TaxYear < 1999 Then
'    TaxTrans.Amount = 0
'    TaxTrans.Revenue.Collection = 0
'    TaxTrans.Revenue.CollectionPd = 0
'    TaxTrans.Revenue.Interest = 0
'    TaxTrans.Revenue.InterestPd = 0
'    TaxTrans.Revenue.LateList = 0
'    TaxTrans.Revenue.LateListPd = 0
'    TaxTrans.Revenue.Penalty = 0
'    TaxTrans.Revenue.PenaltyPd = 0
'    TaxTrans.Revenue.PrePaidUsed = 0
'    TaxTrans.Revenue.Principle1 = 0
'    TaxTrans.Revenue.Principle1Pd = 0
'    TaxTrans.Revenue.Principle2 = 0
'    TaxTrans.Revenue.Principle2Pd = 0
'    TaxTrans.Revenue.Principle3 = 0
'    TaxTrans.Revenue.Principle3Pd = 0
'    TaxTrans.Revenue.Principle4 = 0
'    TaxTrans.Revenue.Principle4Pd = 0
'    TaxTrans.Revenue.Principle5 = 0
'    TaxTrans.Revenue.Principle5Pd = 0
'    TaxTrans.Revenue.RevOpt1 = 0
'    TaxTrans.Revenue.RevOpt1Pd = 0
'    TaxTrans.Revenue.RevOpt2 = 0
'    TaxTrans.Revenue.RevOpt2Pd = 0
'    TaxTrans.Revenue.RevOpt3 = 0
'    TaxTrans.Revenue.RevOpt3Pd = 0
'  Put TTHandle, x, TaxTrans
'   End If
'  End If
'  Next x
'
'  Get TTHandle, 157274, TaxTrans
'  TaxTrans.Amount = 84.48
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.TranType = 2
'  TaxTrans.Description = "Bill No# 296"
'  Put TTHandle, 157274, TaxTrans
'
'  'fix for 1880 on 3/4/09
'  Get TTHandle, 136942, TaxTrans
'  TaxTrans.RealPin = "2HH-18-22-13"
'  Put TTHandle, 136942, TaxTrans
  
  
  'fix for 1583
  
'  Get TTHandle, 824, TaxTrans
'  TaxTrans.Revenue.Interest = TaxTrans.Revenue.Interest + 1.18
'  Put TTHandle, 824, TaxTrans
  
'  Get TTHandle, 430, TaxTrans
'  TaxTrans.Revenue.Interest = 0
'  TaxTrans.Amount = 0
'  Put TTHandle, 430, TaxTrans

'  Close
'  MsgBox ("Done.")

End Sub

Private Sub cmdFixLunen_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, 11631, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  Put TTHandle, 11631, TaxTrans

  Close
  MsgBox ("Done.")
'FixPayPlusOPThatShouldHaveBeenApplied
End Sub

Private Sub cmdFixLunenburg12_10_08_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Integer
  Dim BelongTo As Long
  Dim CollectionPd As Double
  Dim InterestPd As Double
  Dim LateListPd As Double
  Dim PenaltyPd As Double
  Dim PrePaidUsed As Double
  Dim Principle1Pd As Double
  Dim Principle2Pd As Double
  Dim Principle3Pd As Double
  Dim Principle4Pd As Double
  Dim Principle5Pd As Double
  Dim RevOpt1Pd As Double
  Dim RevOpt2Pd As Double
  Dim RevOpt3Pd As Double
  
  Dim TransArray(1 To 13) As Long
  TransArray(1) = 134778
  TransArray(2) = 134776
  TransArray(3) = 134676
  TransArray(4) = 134785
  TransArray(5) = 134777
  TransArray(6) = 134681
  
  TransArray(7) = 134779
  TransArray(8) = 134775
  TransArray(9) = 134774
  TransArray(10) = 134675
  TransArray(11) = 134780
  TransArray(12) = 134773
  TransArray(13) = 134680
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To 13
    Get TTHandle, TransArray(x), TaxTrans
    BelongTo = TaxTrans.BelongTo
    If x < 7 Then
      GoSub ZeroBillPay
    Else
      GoSub ZeroBillAdjDown
    End If
    GoSub ZeroPay
    Put TTHandle, TransArray(x), TaxTrans
  Next x
  Close
  MsgBox ("Done.")
  
  Exit Sub
ZeroPay:
    TaxTrans.Amount = 0
    TaxTrans.Revenue.Collection = 0
    TaxTrans.Revenue.CollectionPd = 0
    TaxTrans.Revenue.Interest = 0
    TaxTrans.Revenue.InterestPd = 0
    TaxTrans.Revenue.LateList = 0
    TaxTrans.Revenue.LateListPd = 0
    TaxTrans.Revenue.Penalty = 0
    TaxTrans.Revenue.PenaltyPd = 0
    TaxTrans.Revenue.PrePaidUsed = 0
    TaxTrans.Revenue.Principle1 = 0
    TaxTrans.Revenue.Principle1Pd = 0
    TaxTrans.Revenue.Principle2 = 0
    TaxTrans.Revenue.Principle2Pd = 0
    TaxTrans.Revenue.Principle3 = 0
    TaxTrans.Revenue.Principle3Pd = 0
    TaxTrans.Revenue.Principle4 = 0
    TaxTrans.Revenue.Principle4Pd = 0
    TaxTrans.Revenue.Principle5 = 0
    TaxTrans.Revenue.Principle5Pd = 0
    TaxTrans.Revenue.RevOpt1 = 0
    TaxTrans.Revenue.RevOpt1Pd = 0
    TaxTrans.Revenue.RevOpt2 = 0
    TaxTrans.Revenue.RevOpt2Pd = 0
    TaxTrans.Revenue.RevOpt3 = 0
    TaxTrans.Revenue.RevOpt3Pd = 0
  Return
  
ZeroBillPay:
    CollectionPd = TaxTrans.Revenue.CollectionPd
    InterestPd = TaxTrans.Revenue.InterestPd
    LateListPd = TaxTrans.Revenue.LateListPd
    PenaltyPd = TaxTrans.Revenue.PenaltyPd
    PrePaidUsed = 0
    Principle1Pd = TaxTrans.Revenue.Principle1Pd
    Principle2Pd = TaxTrans.Revenue.Principle2Pd
    Principle3Pd = TaxTrans.Revenue.Principle3Pd
    Principle4Pd = TaxTrans.Revenue.Principle4Pd
    Principle5Pd = TaxTrans.Revenue.Principle5Pd
    RevOpt1Pd = TaxTrans.Revenue.RevOpt1Pd
    RevOpt2Pd = TaxTrans.Revenue.RevOpt2Pd
    RevOpt3Pd = TaxTrans.Revenue.RevOpt3Pd
    Get TTHandle, BelongTo, TaxTrans
    TaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.CollectionPd - CollectionPd
    TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - InterestPd
    TaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateListPd - LateListPd
    TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - PenaltyPd
    TaxTrans.Revenue.PrePaidUsed = 0
    TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - Principle1Pd
    TaxTrans.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2Pd - Principle2Pd
    TaxTrans.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3Pd - Principle3Pd
    TaxTrans.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4Pd - Principle4Pd
    TaxTrans.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5Pd - Principle5Pd
    TaxTrans.Revenue.RevOpt1Pd = TaxTrans.Revenue.RevOpt1Pd - RevOpt1Pd
    TaxTrans.Revenue.RevOpt2Pd = TaxTrans.Revenue.RevOpt2Pd - RevOpt2Pd
    TaxTrans.Revenue.RevOpt3Pd = TaxTrans.Revenue.RevOpt3Pd - RevOpt3Pd
    Put TTHandle, BelongTo, TaxTrans
    Get TTHandle, TransArray(x), TaxTrans
  Return
  
ZeroBillAdjDown:
    CollectionPd = TaxTrans.Revenue.CollectionPd
    InterestPd = TaxTrans.Revenue.InterestPd
    LateListPd = TaxTrans.Revenue.LateListPd
    PenaltyPd = TaxTrans.Revenue.PenaltyPd
    PrePaidUsed = 0
    Principle1Pd = TaxTrans.Revenue.Principle1Pd
    Principle2Pd = TaxTrans.Revenue.Principle2Pd
    Principle3Pd = TaxTrans.Revenue.Principle3Pd
    Principle4Pd = TaxTrans.Revenue.Principle4Pd
    Principle5Pd = TaxTrans.Revenue.Principle5Pd
    RevOpt1Pd = TaxTrans.Revenue.RevOpt1Pd
    RevOpt2Pd = TaxTrans.Revenue.RevOpt2Pd
    RevOpt3Pd = TaxTrans.Revenue.RevOpt3Pd
    Get TTHandle, BelongTo, TaxTrans
    TaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.CollectionPd + CollectionPd
    TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd + InterestPd
    TaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateListPd + LateListPd
    TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd + PenaltyPd
    TaxTrans.Revenue.PrePaidUsed = 0
    TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd + Principle1Pd
    TaxTrans.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2Pd + Principle2Pd
    TaxTrans.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3Pd + Principle3Pd
    TaxTrans.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4Pd + Principle4Pd
    TaxTrans.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5Pd + Principle5Pd
    TaxTrans.Revenue.RevOpt1Pd = TaxTrans.Revenue.RevOpt1Pd + RevOpt1Pd
    TaxTrans.Revenue.RevOpt2Pd = TaxTrans.Revenue.RevOpt2Pd + RevOpt2Pd
    TaxTrans.Revenue.RevOpt3Pd = TaxTrans.Revenue.RevOpt3Pd + RevOpt3Pd
    Put TTHandle, BelongTo, TaxTrans
    Get TTHandle, TransArray(x), TaxTrans
  Return
  
End Sub

Private Sub cmdFixInitialized_Click()
  Call FixInitializedTrans
End Sub

Private Sub cmdFixLunenburgPins_Click()
  Dim x As Long, y As Long
  Dim TextLine$
  Dim ThisFile$
  Dim LHandle As Integer
  Dim WordCnt As Integer
  Dim TextLen As Integer
  Dim ThisCh As String
  Dim ThisWord$
  Dim CntyNum As String
  Dim NewPin As String
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim OldPin As String
  Dim NextRec As Long
  Dim dlm As String
  Dim cnt As Integer
  
  dlm = "~"
  WordCnt = 0
  ReDim Words(1 To 1) As String
  OpenRealPropFile RHandle, NumOfRRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  frmVATaxShowPctComp.Label1 = "Lunenburg Pin Update"
  frmVATaxShowPctComp.Show , Me
  
  If Exist("lunny.csv") Then
    LHandle = FreeFile
    ThisFile = "lunny.csv"
    Open ThisFile For Input As #LHandle
    Do While ThisWord <> "End"
      cnt = cnt + 1
      Line Input #LHandle, TextLine
      TextLen = Len(TextLine)
      TextLine = TextLine + dlm
      For x = 1 To TextLen + 1
        ThisCh = Mid(TextLine, x, 1)
        If ThisCh = dlm Then
          WordCnt = WordCnt + 1
          ReDim Preserve Words(1 To WordCnt) As String
          If WordCnt = 1 Then
            CntyNum = ThisWord
            ThisWord = ""
            GoTo NewWord
          ElseIf WordCnt = 2 Then
            NewPin = ThisWord
            GoSub SaveNewPin
            NewPin = ""
            CntyNum = ""
            ThisWord = ""
            WordCnt = 0
            GoTo NewLoop
          End If
        End If
        ThisWord = ThisWord + ThisCh
        If ThisWord = "End" Then Exit Do
NewWord:
      Next x
NewLoop:
    frmVATaxShowPctComp.ShowPctComp cnt, 11332
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
    Loop
  Else
    MsgBox ("The file 'lunny.csv' cannot be found.")
    Exit Sub
  End If
  Unload frmVATaxShowPctComp
  
  MsgBox ("Lunenburg real pin update has completed successfully.")
  Close
  Exit Sub
  
SaveNewPin:
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If QPTrim$(TaxCust.CountyAcctString) = "" Then
      TaxCust.CountyAcctString = CStr(TaxCust.CountyAcct)
    End If
    If QPTrim$(TaxCust.CountyAcctString) = CntyNum Then
      NextRec = TaxCust.FirstPropRec
      If NextRec = 0 Then GoTo Skip
'      Do While NextRec > 0
       Get RHandle, NextRec, RealRec
'       If RealRec.RealPin = OldPin Then
         RealRec.RealPin = NewPin
         Put RHandle, NextRec, RealRec
         Return
'       End If
'       NextRec = RealRec.NextRec
'      Loop
    End If
Skip:
  Next x
  
  Return

End Sub

Private Sub cmdFixMontross_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
'  'fix for cust #302
'  Get TTHandle, 2782, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 44.58
'  TaxTrans.Revenue.InterestPd = 0.49
'  TaxTrans.Revenue.PenaltyPd = 4.46
'  Put TTHandle, 2782, TaxTrans
'
'  Get TTHandle, 2021, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 44.58
'  TaxTrans.Revenue.InterestPd = 5.88
'  TaxTrans.Revenue.PenaltyPd = 4.46
'  Put TTHandle, 2021, TaxTrans
'
'  Get TTHandle, 1680, TaxTrans
'  TaxTrans.Revenue.Interest = 0
'  TaxTrans.Amount = 0
'  Put TTHandle, 1680, TaxTrans
'
'  Get TTHandle, 1657, TaxTrans
'  TaxTrans.Revenue.Interest = 0
'  TaxTrans.Amount = 0
'  Put TTHandle, 1657, TaxTrans
'
'  'fix for cust #301
'  Get TTHandle, 2019, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0.06
'  TaxTrans.Revenue.PenaltyPd = 0.01
'  Put TTHandle, 2019, TaxTrans
'
'  Get TTHandle, 2780, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0.06
'  TaxTrans.Revenue.PenaltyPd = 0.01
'  Put TTHandle, 2780, TaxTrans
'
'  'fix for cust #303
'  Get TTHandle, 2020, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0.12
'  TaxTrans.Revenue.PenaltyPd = 0.01
'  Put TTHandle, 2020, TaxTrans
'
'  Get TTHandle, 2781, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0.12
'  TaxTrans.Revenue.PenaltyPd = 0.01
'  Put TTHandle, 2781, TaxTrans
 'fix done on 2/8/10
 'fix for #53
  Get TTHandle, 1272, TaxTrans
  TaxTrans.Revenue.Principle1 = 25.38
  TaxTrans.Revenue.Principle1Pd = 25.38
  TaxTrans.Amount = 25.38
  Put TTHandle, 1272, TaxTrans
  
  'fix for #52
  Get TTHandle, 2883, TaxTrans
  TaxTrans.Revenue.Principle1 = 61.38
  TaxTrans.Revenue.Principle1Pd = 61.38
  TaxTrans.Amount = 61.38
  Put TTHandle, 2883, TaxTrans

 'fix for #71
  Get TTHandle, 3860, TaxTrans
  TaxTrans.Revenue.Principle1 = 3.18
  TaxTrans.Revenue.Principle1Pd = 3.18
  TaxTrans.Amount = 3.18
  Put TTHandle, 3860, TaxTrans
  
 'fix for #32
  Get TTHandle, 513, TaxTrans
  TaxTrans.Revenue.Principle1 = 96.9
  TaxTrans.Revenue.Principle1Pd = 96.9
  TaxTrans.Amount = 96.9
  Put TTHandle, 513, TaxTrans
  
 'fix for #276
  Get TTHandle, 46, TaxTrans
  TaxTrans.Revenue.Principle1 = 66.59
  TaxTrans.Revenue.Principle1Pd = 66.59
  TaxTrans.Amount = 66.59
  Put TTHandle, 46, TaxTrans

  Close
  Call TaxMsg(900, "Finished.")
End Sub

Private Sub cmdFixPenGap_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim ThisRec As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRec As Long
  
  OpenPersPropFile PHandle, NumOfPersRec
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  'fix for #74/#3798 on 5/18/09
  Get TTHandle, 17346, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.LastTrans = 0
  Put TTHandle, 17346, TaxTrans
 
  Get TTHandle, 4017, TaxTrans
  TaxTrans.BelongTo = 1186
  TaxTrans.CustomerRec = 74
  TaxTrans.CustPin = 74
  TaxTrans.LastTrans = 1186
  Put TTHandle, 4017, TaxTrans
  
  Get TTHandle, 6680, TaxTrans
  TaxTrans.LastTrans = 4017
  Put TTHandle, 6680, TaxTrans
  
  'cust #288 5/18/09
  Get TTHandle, 186, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1 = 0
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Revenue.Penalty = 0
  Put TTHandle, 186, TaxTrans
  
  'cust #156 5/18/09
  Get TTHandle, 122, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1 = 0
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Revenue.Penalty = 0
  Put TTHandle, 122, TaxTrans
  
  
'  Get TCHandle, 74, TaxCust
'  ThisRec = TaxCust.FirstPropRec
'  TaxCust.FirstPropRec = 0
'  Put TCHandle, 74, TaxCust
'
'  Get TTHandle, TaxCust.LastTrans, TaxTrans
'  TaxTrans.LastTrans = 71
'  Put TTHandle, TaxCust.LastTrans, TaxTrans
'
'  Get TCHandle, 3798, TaxCust
'  TaxCust.LastTrans = 1186
'  Get TTHandle, 1186, TaxTrans
'  TaxTrans.LastTrans = 0
'  Put TTHandle, 1186, TaxTrans
'  TaxCust.FirstPersRec = 0
'  Put TCHandle, 3798, TaxCust
  
'  Get TCHandle, 1315, TaxCust
'  ThisRec = TaxCust.FirstPersRec
'  TaxCust.FirstPersRec = 0
'  TaxCust.LastTrans = 1156
'  Get TTHandle, 1156, TaxTrans
'  TaxTrans.LastTrans = 0
'  Put TTHandle, 1156, TaxTrans
'  Put TCHandle, 1315, TaxCust
  
'  Get TCHandle, 3050, TaxCust
'  TaxCust.FirstPersRec = ThisRec
'  TaxCust.LastTrans = 2208
'  Get TTHandle, 2208, TaxTrans
'  TaxTrans.LastTrans = 1154
'  Put TTHandle, 2208, TaxTrans
'
'  Put TCHandle, 3050, TaxCust

'  Get TCHandle, 3050, TaxCust
'  ThisRec = TaxCust.FirstPersRec
'  TaxCust.LastTrans = 0
'  Put TCHandle, 3050, TaxCust
'
'  Get TCHandle, 827, TaxCust
'  TaxCust.LastTrans = 2208
'  TaxCust.FirstPersRec = ThisRec
'  Get TTHandle, 2208, TaxTrans
'  TaxTrans.CustomerRec = 827
'  TaxTrans.LastTrans = 2610
'  Put TTHandle, 2208, TaxTrans
'
'  Put TCHandle, 827, TaxCust
  
  'fix for cust #529-> 1415
'  Get TCHandle, 529, TaxCust
'  TaxCust.FirstPersRec = 0
'  Put TCHandle, 529, TaxCust
'
'  Get TTHandle, 3007, TaxTrans
'  TaxTrans.BelongTo = 0
'  TaxTrans.TranType = 22
'  TaxTrans.Revenue.PrePaidAmt = 15.67
'  TaxTrans.Description = "Prepaid Amount"
'  TaxTrans.CustomerRec = 529
'  TaxTrans.CustPin = 529
'  Put TTHandle, 3007, TaxTrans
'
'  Get TTHandle, 3001, TaxTrans
'  TaxTrans.LastTrans = 1987
'  TaxTrans.CustomerRec = 529
'  TaxTrans.CustPin = 529
'  Put TTHandle, 3001, TaxTrans
'
'  Get TCHandle, 1415, TaxCust
'  Get TTHandle, TaxCust.FirstPersRec, TaxTrans
'  Get PHandle, 1084, PersRec
'  PersRec.CustPin = 1415
'  Put PHandle, 1084, PersRec
'  Get PHandle, 1085, PersRec
'  PersRec.CustPin = 1415
'  Put PHandle, 1085, PersRec
'  Put TTHandle, TaxCust.FirstPersRec, TaxTrans
'
'  Get TTHandle, 2299, TaxTrans
'  TaxTrans.LastTrans = 2860
'  Put TTHandle, 2299, TaxTrans
'
'  Get TTHandle, 2860, TaxTrans
'  TaxTrans.Revenue.CollectionPd = 0
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.LateListPd = 0
'  TaxTrans.Revenue.PenaltyPd = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle2Pd = 0
'  TaxTrans.Revenue.Principle3Pd = 0
'  TaxTrans.Revenue.Principle4Pd = 0
'  TaxTrans.Revenue.Principle5Pd = 0
'  TaxTrans.CustomerRec = 1415
'  TaxTrans.CustPin = 1415
'  TaxTrans.LastTrans = 0
'  Put TTHandle, 2860, TaxTrans
'
'  'fix for cust #767-> 2864
'  Get TCHandle, 767, TaxCust
'  Get PHandle, TaxCust.FirstPersRec, PersRec
'  PersRec.CustPin = 2864
'  Put PHandle, TaxCust.FirstPersRec, PersRec
'  TaxCust.FirstPersRec = 0
'  TaxCust.LastTrans = 1585
'  Put TCHandle, 767, TaxCust
'
'  Get TTHandle, 1548, TaxTrans
'  TaxTrans.LastTrans = 0
'  Put TTHandle, 1548, TaxTrans
'
'  Get TTHandle, 2573, TaxTrans
'  TaxTrans.LastTrans = 1548
'  TaxTrans.CustomerRec = 2864
'  TaxTrans.CustPin = 2864
'  Put TTHandle, 2573, TaxTrans
'
'  Get TCHandle, 2864, TaxCust
'  TaxCust.FirstPersRec = 1435
'  Put TCHandle, 2864, TaxCust
'
'  Get TTHandle, 3334, TaxTrans
'  TaxTrans.LastTrans = 2573
'  Put TTHandle, 3334, TaxTrans
'
'  Get PHandle, 1435, PersRec
'  PersRec.CustPin = 2864
'  Put PHandle, 1435, PersRec
'
'  Get PHandle, 1436, PersRec
'  PersRec.CustPin = 2864
'  Put PHandle, 1436, PersRec
'
'  'fix for 792
'  Get PHandle, 810, PersRec
'  PersRec.CustPin = 0
'  PersRec.Deleted = True
'  Put PHandle, 810, PersRec
'
'  Get PHandle, 811, PersRec
'  PersRec.CustPin = 0
'  PersRec.Deleted = True
'  Put PHandle, 811, PersRec
'
'  Get TCHandle, 792, TaxCust
'  TaxCust.FirstPersRec = 812
'  TaxCust.LastTrans = 3324 '1576
'  Put TCHandle, 792, TaxCust
'
'  Get TTHandle, 3324, TaxTrans
'  TaxTrans.LastTrans = 1576
'  Put TTHandle, 3324, TaxTrans
'
'  Get TTHandle, 2570, TaxTrans
'  TaxTrans.LastTrans = 0
'  TaxTrans.CustomerRec = 0
'  TaxTrans.CustPin = 0
'  Put TTHandle, 2570, TaxTrans
  
  Close
  Call TaxMsg(900, "Finished.")


'First fix is here
'  TaxCust.Addr1 = "C/O WSWV RADIO STATION"
'  TaxCust.Addr2 = "PO BOX 630"
'  TaxCust.City = "PENNINGTON GAP"
'  TaxCust.State = "VA"
'  TaxCust.Zip = "24277"
'  TaxCust.CustName = "IBS COMMUNICATIONS LLC"
'  TaxCust.SName = "IBS COMMUN"
'  TaxCust.CSSN = "384"
'  TaxCust.Active = "Y"
'  TaxCust.TaxExempt = "N"
'  TaxCust.Interest = "Y"
'  TaxCust.LateNotice = "Y"
'  TaxCust.Penalty = "Y"
'  TaxCust.Bankrupt = "N"
''  TaxCust.LastTrans = 53040
'  TaxCust.CountyAcctString = "6412"
End Sub

Private Sub cmdFixLunenburgZeroYears_Click()
 Call FixLunenburgZeroYears
End Sub

Private Sub cmdFixOPatBilling_Click()
  Call FixErrorInOPAtBilling
End Sub

Private Sub cmdFixPaidBillsWNoPay_Click()
  Call FixPaidBillsWithNoPayments
End Sub

Private Sub cmdFixPennyBalances_Click()
  Call FindPennyPlusInErrorAndFix
End Sub

Private Sub cmdFixPersLinkToTrans_Click()
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long, y As Long, z As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTaxCusts As Long
  Dim NewPersPin As String
  Dim PinIdx As Integer
  Dim PinIdxS As String
  Dim ThisPersPin As String
  Dim PinCnt As Integer
  Dim NextRec As Long
  Dim NextPersRec As Long
  
  OpenPersPropFile PHandle, NumOfPersRecs
  OpenTaxTransFile THandle, NumOfTRecs
  OpenTaxCustFile TCHandle, NumOfTaxCusts
  PinCnt = 0
  ReDim PinArr(1 To 1) As Long
  
  For z = 1 To NumOfTaxCusts
    Get TCHandle, z, TaxCust
'    If z = 206 Then Stop
'    TaxCust.CustName = TaxCust.CustName
    If TaxCust.FirstPersRec <> 0 Then 'see if cust has personal property
      NextPersRec = TaxCust.FirstPersRec
      Do While NextPersRec > 0 'retrieve pers property
        Get PHandle, NextPersRec, PersRec
        ThisPersPin = QPTrim$(PersRec.PropPin) 'establish pin #
        NextRec = TaxCust.LastTrans
        PinIdx = PinIdx + 1 'reset pin # to apply
        PinIdxS = CStr(PinIdx) & CStr(TaxCust.Acct)
        Do While NextRec > 0
         Get THandle, NextRec, TaxTrans
'         If TaxTrans.TranType = 5 Then Stop
         If ThisPersPin = QPTrim$(TaxTrans.PersPin) Then
           TaxTrans.PersPin = CStr(PinIdxS) 'save new pin # to this property
           'and this transaction linking to that property
           Put THandle, NextRec, TaxTrans
'           PersRec.PropPin = CStr(PinIdxS)
'           Put PHandle, NextPersRec, PersRec
         End If
           PersRec.PropPin = CStr(PinIdxS)
           Put PHandle, NextPersRec, PersRec
         NextRec = TaxTrans.LastTrans
        Loop
        NextPersRec = PersRec.NextRec
      Loop
    End If
  Next z
       
  Close
  MsgBox ("Finished.")

End Sub

Private Sub cmdFixPersReprintFiles_Click()
  Dim PPRec As VAPPTaxBillTypeOld1
  Dim PPRecNew As VAPPTaxBillType
  Dim x As Integer
  Dim PPHandle As Integer
  Dim PPHandleNew As Integer
  Dim NumOfPPRecs As Long
  Dim RmvlRec As TaxPPTRARemovalType
  Dim RHandle As Integer
  Dim NumOfRmvlRecs As Long
  Dim ThisFile$
  Dim y As Integer
  Dim PostRec As TaxBillPostDateType
  Dim PostHandle As Integer
  Dim NumOfPostRecs As Long
  Dim CurrYear As Integer
  Dim ThisType As String * 1
  Dim CompareYr As Integer
  ReDim Bill(1 To 1) As String
  Dim BillCnt As Integer

  If Exist(TaxBillPostDateFile) Then
    OpenBillPostDateFile PostHandle, NumOfPostRecs
    For x = 1 To NumOfPostRecs
      Get PostHandle, x, PostRec
      If PostRec.PPTRAPosted = "Y" Then GoTo SkipIt
      If QPTrim$(PostRec.BackUpName) = "" Then GoTo SkipIt
GoAhead:
      If PostRec.BillType = "P" Then
        BillCnt = BillCnt + 1
        ReDim Preserve Bill(1 To BillCnt) As String
        Bill(BillCnt) = PostRec.BackUpName
      End If
SkipIt:
    Next x
    Close PostHandle
  End If

  KillFile PPTRARemovalFile
  OpenPPTRARmvlFile RHandle, NumOfRmvlRecs

  For y = 1 To 3
    ThisFile = Bill(y)
    If y = 3 Then
     OpenPersPostedReprintFile PPHandle, NumOfPPRecs, ThisFile
     For x = 1 To NumOfPPRecs
       Get PPHandle, x, PPRecNew
       Put PPHandle, x, PPRecNew
     Next x
     GoTo Done
    End If
    OpenPersPostedReprintFileOld1 PPHandle, NumOfPPRecs, ThisFile
    ReDim CustRec(1 To NumOfPPRecs) As Long                            'Acct #
    ReDim CustName(1 To NumOfPPRecs) As String * 40
    ReDim CustAdd1(1 To NumOfPPRecs) As String * 35
    ReDim CustAdd2(1 To NumOfPPRecs) As String * 35
    ReDim CustAdd3(1 To NumOfPPRecs) As String * 35
    ReDim CustZip(1 To NumOfPPRecs) As String * 10
    ReDim RDesc1(1 To NumOfPPRecs) As String * 30
    ReDim RDesc2(1 To NumOfPPRecs) As String * 30
    ReDim RealPin(1 To NumOfPPRecs) As String * 16
    ReDim PersValue(1 To NumOfPPRecs) As Double
    ReDim MHValue(1 To NumOfPPRecs) As Double
    ReDim MCValue(1 To NumOfPPRecs) As Double
    ReDim FEValue(1 To NumOfPPRecs) As Double
    ReDim MTValue(1 To NumOfPPRecs) As Double
    ReDim ExptValue(1 To NumOfPPRecs) As Double
    ReDim PersTaxDue(1 To NumOfPPRecs) As Double
    ReDim MHTaxDue(1 To NumOfPPRecs) As Double
    ReDim MCTaxDue(1 To NumOfPPRecs) As Double
    ReDim FETaxDue(1 To NumOfPPRecs) As Double
    ReDim MTTaxDue(1 To NumOfPPRecs) As Double
    ReDim LateTaxDue(1 To NumOfPPRecs) As Double
    ReDim TotalBillDue(1 To NumOfPPRecs) As Double
    ReDim BillNumber(1 To NumOfPPRecs) As Long                    'Recpt #
    ReDim TaxYear(1 To NumOfPPRecs) As Integer
    ReDim BillPrinted(1 To NumOfPPRecs) As Integer                '-1 = printed
    ReDim PersPropRecord(1 To NumOfPPRecs) As Long
    ReDim PriorYrBalance(1 To NumOfPPRecs) As Double
    ReDim PersTaxRate(1 To NumOfPPRecs) As Double
    ReDim MTTaxRate(1 To NumOfPPRecs) As Double
    ReDim MCTaxRate(1 To NumOfPPRecs) As Double
    ReDim FETaxRate(1 To NumOfPPRecs) As Double
    ReDim MHTaxRate(1 To NumOfPPRecs) As Double
    ReDim CustPin(1 To NumOfPPRecs) As Long                       'Same as Record #
    ReDim ChillHowieFudge(1 To NumOfPPRecs) As Single
    ReDim PPTRAValue(1 To NumOfPPRecs) As Double
    ReDim PPTRADiscnt(1 To NumOfPPRecs) As Double
    ReDim InternalPin(1 To NumOfPPRecs) As Long           'added 5/12/05
    ReDim OptRevTax1(1 To NumOfPPRecs) As Double            'added 5/12/05
    ReDim OptRevTax2(1 To NumOfPPRecs) As Double            'added 5/12/05
    ReDim OptRevTax3(1 To NumOfPPRecs) As Double            'added 5/12/05
    ReDim OverPayAmt(1 To NumOfPPRecs) As Double            'added 5/24/05
    ReDim RDesc3(1 To NumOfPPRecs) As String * 30
    ReDim PersPin(1 To NumOfPPRecs) As String * 20
    ReDim Prorate(1 To NumOfPPRecs) As String * 1               'new for VA 2.05
    ReDim PersTaxNet(1 To NumOfPPRecs) As Double            'new for VA 2.05
    ReDim MultiYrVal(1 To NumOfPPRecs) As Integer            'new for VA 2.05
    ReDim DueDate(1 To NumOfPPRecs) As Integer
    ReDim OptRevDesc1(1 To NumOfPPRecs) As String * 20
    ReDim OptRevDesc2(1 To NumOfPPRecs) As String * 20
    ReDim OptRevDesc3(1 To NumOfPPRecs) As String * 20
    ReDim PostDate(1 To NumOfPPRecs) As Integer
    ReDim TransRec(1 To NumOfPPRecs) As Long
    ReDim Comment(1 To NumOfPPRecs) As String * 40
    ReDim Padding(1 To NumOfPPRecs) As String * 92
    For x = 1 To NumOfPPRecs
      Get PPHandle, x, PPRec
      If PPRec.BillNumber > 0 And PPRec.TransRec = 0 Then
       If PPRec.TotalBillDue = 0 Then
         PPRec.BillNumber = -1
       End If
      End If
      CustRec(x) = PPRec.CustRec
      CustName(x) = PPRec.CustName
      CustAdd1(x) = PPRec.CustAdd1
      CustAdd2(x) = PPRec.CustAdd2
      CustAdd3(x) = PPRec.CustAdd3
      CustZip(x) = PPRec.CustZip
      RDesc1(x) = PPRec.RDesc1
      RDesc2(x) = PPRec.RDesc2
      RealPin(x) = PPRec.RealPin
      PersValue(x) = PPRec.PersValue
      MHValue(x) = PPRec.MHValue
      MCValue(x) = PPRec.MCValue
      FEValue(x) = PPRec.FEValue
      MTValue(x) = PPRec.MTValue
      ExptValue(x) = PPRec.ExptValue
      PersTaxDue(x) = PPRec.PersTaxDue
      MHTaxDue(x) = PPRec.MHTaxDue
      MCTaxDue(x) = PPRec.MCTaxDue
      FETaxDue(x) = PPRec.FETaxDue
      MTTaxDue(x) = PPRec.MTTaxDue
      LateTaxDue(x) = PPRec.LateTaxDue
      TotalBillDue(x) = PPRec.TotalBillDue
      BillNumber(x) = PPRec.BillNumber
      TaxYear(x) = PPRec.TaxYear
      BillPrinted(x) = PPRec.BillPrinted
      PersPropRecord(x) = PPRec.PersPropRecord
      PriorYrBalance(x) = PPRec.PriorYrBalance
      PersTaxRate(x) = PPRec.PersTaxRate
      MTTaxRate(x) = PPRec.MTTaxRate
      MCTaxRate(x) = PPRec.MCTaxRate
      FETaxRate(x) = PPRec.FETaxRate
      MHTaxRate(x) = PPRec.MHTaxRate
      CustPin(x) = PPRec.CustPin
      ChillHowieFudge(x) = PPRec.ChillHowieFudge
      PPTRAValue(x) = PPRec.PPTRAValue
      PPTRADiscnt(x) = PPRec.PPTRADiscnt
      InternalPin(x) = PPRec.InternalPin
      OptRevTax1(x) = PPRec.OptRevTax1
      OptRevTax2(x) = PPRec.OptRevTax2
      OptRevTax3(x) = PPRec.OptRevTax3
      OverPayAmt(x) = PPRec.OverPayAmt
      RDesc3(x) = PPRec.RDesc3
      PersPin(x) = PPRec.PersPin
      Prorate(x) = PPRec.Prorate
      PersTaxNet(x) = PPRec.PersTaxNet
      MultiYrVal(x) = PPRec.MultiYrVal
      DueDate(x) = PPRec.DueDate
      OptRevDesc1(x) = PPRec.OptRevDesc1
      OptRevDesc2(x) = PPRec.OptRevDesc2
      OptRevDesc3(x) = PPRec.OptRevDesc3
      PostDate(x) = PPRec.PostDate
      TransRec(x) = PPRec.TransRec
      Comment(x) = PPRec.Comment
'      PPRec.Comment2 = PPRec.Comment2
'      PPRec.CommentPlace = PPRec.CommentPlace
      Padding(x) = PPRec.Padding
'      PPRec.PrintPrior = PPRec.PrintPrior
'      PPRec.SetDscvry2No = PPRec.SetDscvry2No
     Next x
     
     OpenPersPostedReprintFile PPHandle, NumOfPPRecs, ThisFile
     For x = 1 To NumOfPPRecs
       Get PPHandle, x, PPRecNew
       PPRecNew.CustRec = CustRec(x)
       PPRecNew.CustName = CustName(x)
       PPRecNew.CustAdd1 = CustAdd1(x)
       PPRecNew.CustAdd2 = CustAdd2(x)
       PPRecNew.CustAdd3 = CustAdd3(x)
       PPRecNew.CustZip = CustZip(x)
       PPRecNew.RDesc1 = RDesc1(x)
       PPRecNew.RDesc2 = RDesc2(x)
       PPRecNew.RealPin = RealPin(x)
       PPRecNew.PersValue = PersValue(x)
       PPRecNew.MHValue = MHValue(x)
       PPRecNew.MCValue = MCValue(x)
       PPRecNew.FEValue = FEValue(x)
       PPRecNew.MTValue = MTValue(x)
       PPRecNew.ExptValue = ExptValue(x)
       PPRecNew.PersTaxDue = PersTaxDue(x)
       PPRecNew.MHTaxDue = MHTaxDue(x)
       PPRecNew.MCTaxDue = MCTaxDue(x)
       PPRecNew.FETaxDue = FETaxDue(x)
       PPRecNew.MTTaxDue = MTTaxDue(x)
       PPRecNew.LateTaxDue = LateTaxDue(x)
       PPRecNew.TotalBillDue = TotalBillDue(x)
       PPRecNew.BillNumber = BillNumber(x)
       PPRecNew.TaxYear = TaxYear(x)
       PPRecNew.BillPrinted = BillPrinted(x)
       PPRecNew.PersPropRecord = PersPropRecord(x)
       PPRecNew.PriorYrBalance = PriorYrBalance(x)
       PPRecNew.PersTaxRate = PersTaxRate(x)
       PPRecNew.MTTaxRate = MTTaxRate(x)
       PPRecNew.MCTaxRate = MCTaxRate(x)
       PPRecNew.FETaxRate = FETaxRate(x)
       PPRecNew.MHTaxRate = MHTaxRate(x)
       PPRecNew.CustPin = CustPin(x)
       PPRecNew.ChillHowieFudge = ChillHowieFudge(x)
       PPRecNew.PPTRAValue = PPTRAValue(x)
       PPRecNew.PPTRADiscnt = PPTRADiscnt(x)
       PPRecNew.InternalPin = InternalPin(x)
       PPRecNew.OptRevTax1 = OptRevTax1(x)
       PPRecNew.OptRevTax2 = OptRevTax2(x)
       PPRecNew.OptRevTax3 = OptRevTax3(x)
       PPRecNew.OverPayAmt = OverPayAmt(x)
       PPRecNew.RDesc3 = RDesc3(x)
       PPRecNew.PersPin = PersPin(x)
       PPRecNew.Prorate = Prorate(x)
       PPRecNew.PersTaxNet = PersTaxNet(x)
       PPRecNew.MultiYrVal = MultiYrVal(x)
       PPRecNew.DueDate = DueDate(x)
       PPRecNew.OptRevDesc1 = OptRevDesc1(x)
       PPRecNew.OptRevDesc2 = OptRevDesc2(x)
       PPRecNew.OptRevDesc3 = OptRevDesc3(x)
       PPRecNew.PostDate = PostDate(x)
       PPRecNew.TransRec = TransRec(x)
       PPRecNew.Comment = Comment(x)
       PPRecNew.Comment2 = " "
       PPRecNew.CommentPlace = " "
       PPRecNew.Padding = Padding(x)
       PPRecNew.PrintPrior = False
       PPRecNew.SetDscvry2No = ""
       Put PPHandle, x, PPRecNew
     Next x
   Next y
Done:
   Close
   MsgBox ("Finished.")


End Sub

Private Sub cmdFixPound_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TT2Handle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim BelongTo As Long
  Dim Amount As Double
  Dim CollectionPd As Double
  Dim InterestPd As Double
  Dim LateListPd As Double
  Dim PenaltyPd As Double
  Dim PrePaidAmt As Double
  Dim PrePaidBal As Double
  Dim PrePaidUsed As Double
  Dim Principle1Pd As Double
  Dim Principle2Pd As Double
  Dim Principle3Pd As Double
  Dim Principle4Pd As Double
  Dim Principle5Pd As Double
  Dim RevOpt1Pd As Double
  Dim RevOpt2Pd As Double
  Dim RevOpt3Pd As Double
  Dim TheDate As Integer
  Dim cnt As Integer
  Dim TArr(1 To 7) As Long
  TArr(1) = 3095
  TArr(2) = 3073
  TArr(3) = 3072
  TArr(4) = 3071
  TArr(5) = 3070
  TArr(6) = 3069
  TArr(7) = 3068
  TheDate = Date2Num("12/07/2009")
  OpenTaxTransFile TT2Handle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TT2Handle, x, TaxTrans
    If TaxTrans.TranType = 4 Then
      If TaxTrans.TransDate = TheDate Then
        ClearTrans (x)
        cnt = cnt + 1
      End If
    End If
  Next x
  
'  For x = 2 To 7
'    Get TT2Handle, TArr(x), TaxTrans
'    GoSub ZeroOut
'    Put TT2Handle, TArr(x), TaxTrans
'  Next x
'
'  Get TT2Handle, TArr(1), TaxTrans
'    BelongTo = TaxTrans.BelongTo
'    Get TT2Handle, BelongTo, TaxTrans
'    TaxTrans.Revenue.InterestPd = 0
'    Put TT2Handle, BelongTo, TaxTrans
'    Get TT2Handle, TArr(1), TaxTrans
'    TaxTrans.Amount = 0
'    TaxTrans.Revenue.InterestPd = 0
'    Put TT2Handle, TArr(1), TaxTrans
'
'  Close TT2Handle
  
  MsgBox "A total of " + CStr(cnt) + " interest transactions have been zeroed out."
  Exit Sub
  
ZeroOut:
    BelongTo = TaxTrans.BelongTo
    Amount = TaxTrans.Amount
    CollectionPd = TaxTrans.Revenue.CollectionPd
    InterestPd = TaxTrans.Revenue.InterestPd
    LateListPd = TaxTrans.Revenue.LateListPd
    PenaltyPd = TaxTrans.Revenue.PenaltyPd
    PrePaidAmt = TaxTrans.Revenue.PrePaidAmt
    PrePaidBal = TaxTrans.Revenue.PrePaidBal
    PrePaidUsed = TaxTrans.Revenue.PrePaidUsed
    Principle1Pd = TaxTrans.Revenue.Principle1Pd
    Principle2Pd = TaxTrans.Revenue.Principle2Pd
    Principle3Pd = TaxTrans.Revenue.Principle3Pd
    Principle4Pd = TaxTrans.Revenue.Principle4Pd
    Principle5Pd = TaxTrans.Revenue.Principle5Pd
    RevOpt1Pd = TaxTrans.Revenue.RevOpt1Pd
    RevOpt2Pd = TaxTrans.Revenue.RevOpt2Pd
    RevOpt3Pd = TaxTrans.Revenue.RevOpt3Pd
    
    Get TT2Handle, BelongTo, TaxTrans
    TaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.CollectionPd - CollectionPd
    TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - InterestPd
    TaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateListPd - LateListPd
    TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - PenaltyPd
    TaxTrans.Revenue.PrePaidAmt = TaxTrans.Revenue.PrePaidAmt - PrePaidAmt
    TaxTrans.Revenue.PrePaidBal = TaxTrans.Revenue.PrePaidBal - PrePaidBal
    TaxTrans.Revenue.PrePaidUsed = TaxTrans.Revenue.PrePaidUsed - PrePaidUsed
    TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - Principle1Pd
    TaxTrans.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2Pd - Principle2Pd
    TaxTrans.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3Pd - Principle3Pd
    TaxTrans.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4Pd - Principle4Pd
    TaxTrans.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5Pd - Principle5Pd
    TaxTrans.Revenue.RevOpt1Pd = 0 'TaxTrans.Revenue.RevOpt1Pd = RevOpt1Pd
    TaxTrans.Revenue.RevOpt2Pd = 0 'TaxTrans.Revenue.RevOpt2Pd = RevOpt2Pd
    TaxTrans.Revenue.RevOpt3Pd = 0 'TaxTrans.Revenue.RevOpt3Pd = RevOpt3Pd
    Put TT2Handle, BelongTo, TaxTrans
    
    Get TT2Handle, TArr(x), TaxTrans
    TaxTrans.Amount = 0
    TaxTrans.Revenue.CollectionPd = 0
    TaxTrans.Revenue.InterestPd = 0
    TaxTrans.Revenue.LateListPd = 0
    TaxTrans.Revenue.PenaltyPd = 0
    TaxTrans.Revenue.PrePaidAmt = 0
    TaxTrans.Revenue.PrePaidBal = 0
    TaxTrans.Revenue.PrePaidUsed = 0
    TaxTrans.Revenue.Principle1Pd = 0
    TaxTrans.Revenue.Principle2Pd = 0
    TaxTrans.Revenue.Principle3Pd = 0
    TaxTrans.Revenue.Principle4Pd = 0
    TaxTrans.Revenue.Principle5Pd = 0
    TaxTrans.Revenue.RevOpt1Pd = 0
    TaxTrans.Revenue.RevOpt2Pd = 0
    TaxTrans.Revenue.RevOpt3Pd = 0

  Return
End Sub

Private Sub cmdFixPocohontas_Click()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim ThisDate As Integer
  ThisDate = Date2Num("1/6/2009")
  OpenTaxTransFile THandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get THandle, x, TaxTrans

    If TaxTrans.TranType = 30 And TaxTrans.TransDate = ThisDate Then
      TaxTrans.Amount = 0
      TaxTrans.PPTRARmvl = 0
      TaxTrans.PPTRAVal = 0
      Put THandle, x, TaxTrans
    End If
  Next x
  Close THandle
  
  MsgBox "Completed successfully."

End Sub

Private Sub cmdFixPocahontas_Click()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTTRecs As Long
  OpenTaxTransFile THandle, NumOfTTRecs
  Get THandle, 4864, TaxTrans
  TaxTrans.Amount = 150.22
  TaxTrans.Revenue.Principle1 = 150.22
  Put THandle, 4864, TaxTrans
  Close THandle
  
  MsgBox "Completed successfully."

End Sub

Private Sub cmdFixRealCustPins_Click()
  Dim RealPropRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim x As Long
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTaxCusts As Long
  Dim ThisRec As Long
  Dim cnt As Long
  
  frmVATaxShowPctComp.Label1 = "Fixing Real Prop Cust Pins"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenTaxCustFile TCHandle, NumOfTaxCusts
  For x = 1 To NumOfTaxCusts
    Get TCHandle, x, TaxCustRec
    If TaxCustRec.FirstPropRec > 0 Then ThisRec = TaxCustRec.FirstPropRec
    Do While ThisRec > 0
      Get RHandle, ThisRec, RealPropRec
      If RealPropRec.CustPin <> TaxCustRec.Acct Then
        RealPropRec.CustPin = TaxCustRec.Acct
        Put RHandle, ThisRec, RealPropRec
        cnt = cnt + 1
      End If
      ThisRec = RealPropRec.NextRec
    Loop
    frmVATaxShowPctComp.ShowPctComp x, NumOfTaxCusts
  Next x
  Close RHandle
  Close TCHandle
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  
  Call Savemsg(900, "A total of " + CStr(cnt) + " real property records were corrected successfully.")
  
End Sub

Private Sub cmdFixRealPinsWOverPay_Click()
  Dim TransRec As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  Dim cnt As Long
  Dim PHandle As Integer
  
  PHandle = FreeFile
  Open "RealPinsRemovedFromPrepay.txt" For Output As PHandle
  frmVATaxShowPctComp.Label1 = "Updating Real Pins"
  frmVATaxShowPctComp.Show , Me
  OpenTaxTransFile THandle, NumOfTRecs
  Print #PHandle, "Real Pin ~ Customer Pin"
  For x = 1 To NumOfTRecs
    Get THandle, x, TransRec
    If TransRec.TranType = 22 Or TransRec.TranType = 11 Then
      If QPTrim$(TransRec.RealPin) <> "" Then
       Print #PHandle, TransRec.RealPin + "~" + CStr(TransRec.CustPin)
       TransRec.RealPin = ""
       cnt = cnt + 1
       Put THandle, x, TransRec
      End If
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
    End If
    Next x
    Unload frmVATaxShowPctComp
   Close
   MsgBox ("A total of " + CStr(cnt) + " real pins were removed. Look for 'RealPinsRemovedFromPrepay.txt' for the results.")


End Sub
Private Sub FixStephensCity()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long, y As Long, z As Integer
  Dim BelongTo As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim Before As Long
  Dim Before2 As Long
  Dim BeforeOld As Long
  Dim NextRec As Long
  Dim NextRec2 As Long
  Dim NextRec3 As Long
  Dim MoveToCust As Long
  Dim SaveRec As Long
  Dim JumpRec As Long
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile THandle, NumOfTRecs
  
'  Get TCHandle, 1781, TaxCust
'  NextRec = TaxCust.LastTrans
'  SaveRec = TaxCust.LastTrans
'  BeforeOld = TaxCust.LastTrans
'  Do While NextRec > 0
'    Get THandle, NextRec, TaxTrans
''    SaveRec = TaxCust.LastTrans
''    If NextRec = 8438 Then Stop
'    If TaxTrans.TranType = 9 Then
'      BelongTo = TaxTrans.BelongTo
'      Get THandle, BelongTo, TaxTrans
'      If TaxTrans.CustomerRec = 1781 Then GoTo SameCust
'      Get TCHandle, TaxTrans.CustomerRec, TaxCust
'      MoveToCust = TaxTrans.CustomerRec
'      If TaxCust.Deleted <> 0 Then
'        Get THandle, NextRec, TaxTrans
'        GoTo SameCust
'      End If
''      TaxTrans.CustPin = TaxTrans.CustPin
'      NextRec2 = TaxCust.LastTrans
'      Before = TaxCust.LastTrans
'      Do While NextRec2 > 0
'        Get THandle, NextRec2, TaxTrans 'looking up new cust trans
'        If NextRec2 = BelongTo Then
'          Get THandle, Before, TaxTrans 'insert new trans
'            TaxTrans.LastTrans = NextRec
'          Put THandle, Before, TaxTrans
'          Get THandle, NextRec, TaxTrans 'switch next trans on the new insert to next in line
'            NextRec3 = SaveRec 'SaveRec is 1781's starting trans
'            Before2 = NextRec3
'            Do While NextRec3 > 0 'but first find the old cust trans to delete
'              Get THandle, NextRec3, TaxTrans
'              If NextRec3 < NextRec Then
'                Get THandle, NextRec, TaxTrans
'                NextRec3 = TaxTrans.LastTrans 'old trans gets new next trans
'                Get THandle, BeforeOld, TaxTrans
'                  TaxTrans.LastTrans = NextRec3
'                Put THandle, BeforeOld, TaxTrans
'                Get THandle, NextRec2, TaxTrans
'                NextRec = NextRec3
'                GoTo SavedJump
'              End If
'              Before2 = NextRec3
'              NextRec3 = TaxTrans.LastTrans
'            Loop
'SavedJump:
'
'            TaxTrans.LastTrans = NextRec2
'          Put THandle, NextRec, TaxTrans
'          Debug.Print CStr(MoveToCust) + " "
''          Get THandle, NextRec, TaxTrans
'          GoTo FoundIt
'        End If
'        Before = NextRec2
'        NextRec2 = TaxTrans.LastTrans
'      Loop
'
'SameCust:
'    End If
'FoundIt:
'    Get THandle, NextRec, TaxTrans
'    BeforeOld = NextRec
'    NextRec = TaxTrans.LastTrans
'  Loop
'  Close
'  MsgBox ("Done.")
'  Exit Sub
'  Dim TrueCust As Long
'  Dim Amount As Double
'  Dim cnt As Integer
'  Get TCHandle, 1781, TaxCust
'  NextRec = TaxCust.LastTrans
'  Do While NextRec > 0
'  Get THandle, NextRec, TaxTrans
''  If NextRec = 8296 Then Stop
'  If TaxTrans.TranType = 9 Then
'    BelongTo = TaxTrans.BelongTo
'    Amount = TaxTrans.Revenue.PrePaidUsed
'    Get THandle, BelongTo, TaxTrans
'    cnt = cnt + 1
'    Debug.Print "1781" + "~" + CStr(NextRec) + "~" + CStr(TaxTrans.CustomerRec) + "~" + CStr(BelongTo) + "~" + Using("#,###.##", Amount)
'  End If
'  Get THandle, NextRec, TaxTrans
'  NextRec = TaxTrans.LastTrans
'  Loop
'  Debug.Print "Count = " + CStr(cnt)
'  MsgBox ("Done")
'  Exit Sub
  
  
  Get THandle, 7894, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  Put THandle, 7894, TaxTrans
  
  Get THandle, 7893, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  Put THandle, 7893, TaxTrans
 
 
  Get THandle, 11572, TaxTrans
  TaxTrans.TranType = 22
  TaxTrans.Revenue.PrePaidAmt = 79.99
  TaxTrans.Revenue.PrePaidBal = 79.99
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 79.99
  Put THandle, 11572, TaxTrans
  
  Get THandle, 8734, TaxTrans
  TaxTrans.Revenue.Principle1 = 0
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 0
  Put THandle, 8734, TaxTrans
  
  Get THandle, 7213, TaxTrans
  TaxTrans.Revenue.Principle1 = 0
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 0
  Put THandle, 7213, TaxTrans


  Get THandle, 11573, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 0
  Put THandle, 11573, TaxTrans

  Get THandle, 11574, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 0
  Put THandle, 11574, TaxTrans

  Get THandle, 11575, TaxTrans
  TaxTrans.TranType = 22
  TaxTrans.Revenue.PrePaidAmt = 0.01
  TaxTrans.Revenue.PrePaidBal = 0.01
  TaxTrans.Amount = 0.01
  Put THandle, 11575, TaxTrans
  
  Get THandle, 14658, TaxTrans
  TaxTrans.Description = "Bill # 381"
  TaxTrans.Revenue.Principle1Pd = 0.01
  TaxTrans.Amount = 0.01
  Put THandle, 14658, TaxTrans
  
  Get THandle, 13437, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 79.99
  Put THandle, 13437, TaxTrans
  
  Close
  MsgBox ("Finished.")


End Sub
Private Sub ClearIntAndPenAndLLFromBills(ByVal RecNum As Integer)
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  OpenTaxTransFile THandle, NumOfTRecs
 
  Get THandle, RecNum, TaxTrans
    TaxTrans.Revenue.Interest = 0
    TaxTrans.Revenue.InterestPd = 0
    TaxTrans.Revenue.LateList = 0
    TaxTrans.Revenue.LateListPd = 0
    TaxTrans.Revenue.Penalty = 0
    TaxTrans.Revenue.PenaltyPd = 0
    Put THandle, RecNum, TaxTrans
  Close THandle

End Sub
Private Sub ClearTrans(ByVal RecNum As Long)
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  OpenTaxTransFile THandle, NumOfTRecs
  '#72 3/24/09
  Get THandle, RecNum, TaxTrans
    TaxTrans.Amount = 0
    TaxTrans.Revenue.Collection = 0
    TaxTrans.Revenue.CollectionPd = 0
    TaxTrans.Revenue.Interest = 0
    TaxTrans.Revenue.InterestPd = 0
    TaxTrans.Revenue.LateList = 0
    TaxTrans.Revenue.LateListPd = 0
    TaxTrans.Revenue.Penalty = 0
    TaxTrans.Revenue.PenaltyPd = 0
    TaxTrans.Revenue.PrePaidUsed = 0
    TaxTrans.Revenue.Principle1 = 0
    TaxTrans.Revenue.Principle1Pd = 0
    TaxTrans.Revenue.Principle2 = 0
    TaxTrans.Revenue.Principle2Pd = 0
    TaxTrans.Revenue.Principle3 = 0
    TaxTrans.Revenue.Principle3Pd = 0
    TaxTrans.Revenue.Principle4 = 0
    TaxTrans.Revenue.Principle4Pd = 0
    TaxTrans.Revenue.Principle5 = 0
    TaxTrans.Revenue.Principle5Pd = 0
    TaxTrans.Revenue.RevOpt1 = 0
    TaxTrans.Revenue.RevOpt1Pd = 0
    TaxTrans.Revenue.RevOpt2 = 0
    TaxTrans.Revenue.RevOpt2Pd = 0
    TaxTrans.Revenue.RevOpt3 = 0
    TaxTrans.Revenue.RevOpt3Pd = 0
    TaxTrans.DiscAmt = 0
    TaxTrans.PPTRADisc = 0
    TaxTrans.PPTRARmvl = 0
    TaxTrans.Revenue.PrePaidBal = TaxTrans.Revenue.PrePaidBal - TaxTrans.Revenue.PrePaidAmt
    TaxTrans.Revenue.PrePaidAmt = 0
    TaxTrans.Revenue.PrePaidUsed = 0
    Put THandle, RecNum, TaxTrans
  Close THandle
End Sub
Private Sub ChangePayToPrePay(ByVal RecNum As Long, ByVal Amount As Double)
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  
  OpenTaxTransFile THandle, NumOfTRecs
  Get THandle, RecNum, TaxTrans
  TaxTrans.Amount = Amount
  TaxTrans.TranType = 22
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.Revenue.LateListPd = 0
  TaxTrans.Revenue.RevOpt1Pd = 0
  TaxTrans.Revenue.RevOpt2Pd = 0
  TaxTrans.Revenue.RevOpt3Pd = 0
  TaxTrans.Revenue.Principle1 = 0
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Revenue.Collection = 0
  TaxTrans.Revenue.LateList = 0
  TaxTrans.Revenue.RevOpt1 = 0
  TaxTrans.Revenue.RevOpt2 = 0
  TaxTrans.Revenue.RevOpt3 = 0
  TaxTrans.FromPrePay = 0
  TaxTrans.BelongTo = 0
  TaxTrans.Description = "Prepay"
  TaxTrans.Revenue.PrePaidAmt = Amount
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidBal = Amount
'  TaxTrans.InternalPin = CustPin
  Put THandle, RecNum, TaxTrans
  Close THandle
  
End Sub
Private Sub ChangePayToPayPlusPrePay(ByVal RecNum As Long, ByVal Amount As Double, ByVal Princ1Paid As Double, ByVal PrePayAmt As Double, ByVal BillNum As String)
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  
  OpenTaxTransFile THandle, NumOfTRecs
  Get THandle, RecNum, TaxTrans
  TaxTrans.Amount = Amount
  TaxTrans.TranType = 21
  TaxTrans.Revenue.Principle1Pd = Princ1Paid
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.Revenue.LateListPd = 0
  TaxTrans.Revenue.RevOpt1Pd = 0
  TaxTrans.Revenue.RevOpt2Pd = 0
  TaxTrans.Revenue.RevOpt3Pd = 0
  TaxTrans.Revenue.Principle1 = 0
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Revenue.Collection = 0
  TaxTrans.Revenue.LateList = 0
  TaxTrans.Revenue.RevOpt1 = 0
  TaxTrans.Revenue.RevOpt2 = 0
  TaxTrans.Revenue.RevOpt3 = 0
  TaxTrans.FromPrePay = 0
  TaxTrans.BelongTo = 0
  TaxTrans.Description = BillNum
  TaxTrans.Revenue.PrePaidAmt = PrePayAmt
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidBal = TaxTrans.Revenue.PrePaidBal + PrePayAmt
'  TaxTrans.InternalPin = CustPin
  Put THandle, RecNum, TaxTrans
  Close THandle
  
End Sub


Private Sub ChangePrePayToPay(ByVal RecNum As Integer, ByVal Amount As Double, ByVal BillNum As String, ByVal BelongTo As Long)
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  
  OpenTaxTransFile THandle, NumOfTRecs
  Get THandle, RecNum, TaxTrans
  TaxTrans.Amount = Amount
  TaxTrans.TranType = 2
  TaxTrans.Revenue.Principle1Pd = Amount
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.Revenue.LateListPd = 0
  TaxTrans.Revenue.RevOpt1Pd = 0
  TaxTrans.Revenue.RevOpt2Pd = 0
  TaxTrans.Revenue.RevOpt3Pd = 0
  TaxTrans.FromPrePay = 0
  TaxTrans.Description = "Bill# " + BillNum
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.BelongTo = BelongTo
'  TaxTrans.InternalPin = CustPin
  Put THandle, RecNum, TaxTrans
  Close THandle
  
End Sub
'Private Sub InsertPayTrans(ByVal ThisDate As String, ByVal PrincPd As Double, ByVal CustPin As Integer, ByVal TaxYear As Integer, ByVal PrePay As Double, ByVal BelongTo As Integer, ByVal PropType As String, ByVal Trans1 As Long, ByVal Trans2 As Long, ByVal BillNum As String, ByVal IntPd As Double, ByVal PenPd As Double, ByRef SaveRec As Long)
'  Dim TaxTrans As TaxTransactionType
'  Dim TTHandle As Integer
'  Dim NumOfTTRecs As Long
'  Dim x As Long, y As Long
''  Dim SaveRec As Long
'  OpenTaxTransFile TTHandle, NumOfTTRecs
'  SaveRec = NumOfTTRecs + 1
'
'  Get TTHandle, SaveRec, TaxTrans
'  TaxTrans.Revenue.Interest = 0#
'  TaxTrans.Amount = PrePay
'  TaxTrans.TransDate = Date2Num%(ThisDate)
'  TaxTrans.TranType = 2
'  TaxTrans.Revenue.Principle1Pd = PrincPd
'  TaxTrans.Revenue.InterestPd = IntPd
'  TaxTrans.Revenue.CollectionPd = 0
'  TaxTrans.Revenue.PenaltyPd = PenPd
'  TaxTrans.Revenue.LateListPd = 0
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  TaxTrans.Revenue.RevOpt2Pd = 0
'  TaxTrans.Revenue.RevOpt3Pd = 0
'  TaxTrans.CustPin = CustPin
'  TaxTrans.DiscXDate = 0
'  TaxTrans.RealPin = ""
'  TaxTrans.PersPin = ""
'  TaxTrans.Posted2GL = "N"
'  TaxTrans.TaxYear = TaxYear
'  TaxTrans.DiscAmt = 0
'  TaxTrans.OperNum = 0
'  TaxTrans.FromPrePay = 0
'  TaxTrans.Description = "Bill# " + BillNum
'  TaxTrans.CustomerRec = CustPin
'  TaxTrans.BelongTo = BelongTo
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.InternalPin = CustPin
'  TaxTrans.CntyPara = ""
'  TaxTrans.CyclPara = ""
'  TaxTrans.TShpPara = ""
'  TaxTrans.BillType = PropType
'  Put TTHandle, SaveRec, TaxTrans
'
''  Get TTHandle, Trans1, TaxTrans
''  TaxTrans.LastTrans = SaveRec
''  Put TTHandle, Trans1, TaxTrans
'
'  Get TTHandle, SaveRec, TaxTrans
'  TaxTrans.LastTrans = Trans2
'  Put TTHandle, SaveRec, TaxTrans
'
'  Get TTHandle, BelongTo, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = PrincPd
'  TaxTrans.Revenue.InterestPd = IntPd
'  TaxTrans.Revenue.PenaltyPd = PenPd
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  TaxTrans.Revenue.RevOpt3Pd = 0
'  Put TTHandle, BelongTo, TaxTrans
'
'  Close
'
'End Sub

Private Sub InsertPayTrans(ByVal ThisDate As String, ByVal PrincPd As Double, ByVal CustPin As Integer, ByVal TaxYear As Integer, ByVal PrePay As Double, ByVal BelongTo As Integer, ByVal PropType As String, ByVal Trans1 As Integer, ByVal Trans2 As Long, ByVal BillNum As String, ByVal IntPd As Double, ByVal PenPd As Double, ByRef SaveRec As Long)
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
'  Dim SaveRec As Long
  OpenTaxTransFile TTHandle, NumOfTTRecs
  SaveRec = NumOfTTRecs + 1

  Get TTHandle, SaveRec, TaxTrans
  TaxTrans.Revenue.Interest = 0#
  TaxTrans.Amount = PrePay
  TaxTrans.TransDate = Date2Num%(ThisDate)
  TaxTrans.TranType = 2
  TaxTrans.Revenue.Principle1Pd = PrincPd
  TaxTrans.Revenue.InterestPd = IntPd
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.Revenue.PenaltyPd = PenPd
  TaxTrans.Revenue.LateListPd = 0
  TaxTrans.Revenue.RevOpt1Pd = 0
  TaxTrans.Revenue.RevOpt2Pd = 0
  TaxTrans.Revenue.RevOpt3Pd = 0
  TaxTrans.CustPin = CustPin
  TaxTrans.DiscXDate = 0
  TaxTrans.RealPin = ""
  TaxTrans.PersPin = ""
  TaxTrans.Posted2GL = "N"
  TaxTrans.TaxYear = TaxYear
  TaxTrans.DiscAmt = 0
  TaxTrans.OperNum = 0
  TaxTrans.FromPrePay = 0
  TaxTrans.Description = "Bill# " + BillNum
  TaxTrans.CustomerRec = CustPin
  TaxTrans.BelongTo = BelongTo
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.InternalPin = CustPin
  TaxTrans.CntyPara = ""
  TaxTrans.CyclPara = ""
  TaxTrans.TShpPara = ""
  TaxTrans.BillType = PropType
  Put TTHandle, SaveRec, TaxTrans
  
'  Get TTHandle, Trans1, TaxTrans
'  TaxTrans.LastTrans = SaveRec
'  Put TTHandle, Trans1, TaxTrans
  
  Get TTHandle, SaveRec, TaxTrans
  TaxTrans.LastTrans = Trans2
  Put TTHandle, SaveRec, TaxTrans

  Get TTHandle, BelongTo, TaxTrans
  TaxTrans.Revenue.Principle1Pd = PrincPd
  TaxTrans.Revenue.InterestPd = IntPd
  TaxTrans.Revenue.PenaltyPd = PenPd
  TaxTrans.Revenue.RevOpt1Pd = 0
  TaxTrans.Revenue.RevOpt3Pd = 0
  Put TTHandle, BelongTo, TaxTrans
  
  Close TTHandle

End Sub
Private Sub ChangeCreditAtBillingtoRegPayment(ByVal Trans2Change As Long)
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, Trans2Change, TaxTrans
  TaxTrans.Amount = TaxTrans.Revenue.PrePaidUsed
  TaxTrans.TranType = 2
  TaxTrans.BelongTo = TaxTrans.BelongTo
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0 'TaxTrans.Revenue.PrePaidBal - TaxTrans.Revenue.PrePaidUsed
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, Trans2Change, TaxTrans
  Close TTHandle
  MsgBox ("Done.")
End Sub

Private Sub ChangePayOverPayToJustPay(ByVal Trans As Long, ByVal NextTrans As Long, ByRef Amount As Double)
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Get TTHandle, Trans, TaxTrans
  TaxTrans.Amount = TaxTrans.Amount - TaxTrans.Revenue.PrePaidAmt
  Amount = TaxTrans.Amount
  TaxTrans.TranType = 2
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidBal = 0
  Put TTHandle, Trans, TaxTrans
  Close TTHandle
End Sub

Private Sub InsertPrepaidOnlyTrans(ByVal ThisDate As String, ByVal Amount As Double, ByVal CustPin As Long, ByVal TaxYear As Integer, ByVal Trans1 As Long, Optional ByVal BillType As String = "")
  Dim PayTranRec As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim SaveRec As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  SaveRec = NumOfTTRecs + 1
  PayTranRec.TransDate = Date2Num(ThisDate)
  PayTranRec.TranType = 22 'overpay only
  PayTranRec.Revenue.Principle1Pd = 0
  PayTranRec.Revenue.InterestPd = 0
  PayTranRec.Revenue.CollectionPd = 0
  PayTranRec.Revenue.LateListPd = 0
  PayTranRec.Revenue.PenaltyPd = 0
  PayTranRec.Revenue.RevOpt1Pd = 0
  PayTranRec.Revenue.RevOpt2Pd = 0
  PayTranRec.Revenue.RevOpt3Pd = 0
  PayTranRec.Revenue.Principle2Pd = 0
  PayTranRec.Revenue.Principle3Pd = 0
  PayTranRec.Revenue.Principle4Pd = 0
  PayTranRec.Revenue.Principle5Pd = 0
  PayTranRec.CustPin = CustPin
  PayTranRec.DiscXDate = Date2Num("12/31/1979")
  PayTranRec.RealPin = " "
  PayTranRec.PersPin = ""
  PayTranRec.Posted2GL = "N"
  PayTranRec.TaxYear = 0
  PayTranRec.DiscAmt = 0
  PayTranRec.OperNum = 0
  PayTranRec.Amount = Amount
  PayTranRec.Description = "Prepay"
  PayTranRec.CustomerRec = CustPin
  PayTranRec.LastTrans = Trans1
  PayTranRec.BelongTo = 0
  PayTranRec.Revenue.PrePaidAmt = Amount
  PayTranRec.Revenue.PrePaidUsed = 0
  PayTranRec.Revenue.PrePaidBal = OldRound(GetOverPayBalance(CustPin, "N") + Amount)
  PayTranRec.BillType = BillType
  Put TTHandle, SaveRec, PayTranRec
  
  Get TCHandle, CustPin, TaxCust
  TaxCust.LastTrans = SaveRec
  Put TCHandle, CustPin, TaxCust
  
  Close TCHandle
  Close TTHandle
  
End Sub

Private Sub InsertCreditAtBillingTrans(ByVal ThisDate As String, ByVal PrincPd As Double, ByVal CustPin As Integer, ByVal TaxYear As Integer, ByVal PrePayUsed As Double, ByVal BelongTo As Long, ByVal PropType As String, ByVal B4Trans As Long, ByVal BillRec As Long, ByVal BillNum As String)
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim SaveRec As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim Intr As Double
  Dim IntrPd As Double
  Dim Pen As Double
  Dim PenPd As Double
  Dim Adv As Double
  Dim AdvPd As Double
  Dim LL As Double
  Dim LLPd As Double
  Dim Pr1 As Double
  Dim Pr1Pd As Double
  Dim Pr2 As Double
  Dim Pr2Pd As Double
  Dim Pr3 As Double
  Dim Pr3Pd As Double
  Dim Pr4 As Double
  Dim Pr4Pd As Double
  Dim Pr5 As Double
  Dim Pr5Pd As Double
  Dim TotalPd As Double
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  SaveRec = NumOfTTRecs + 1
'  If BillRec = 12289 Then Stop
  Get TTHandle, BillRec, TaxTrans
  Pr1 = TaxTrans.Revenue.Principle1 - (TaxTrans.Revenue.Principle1Pd + TaxTrans.PPTRADisc)
  Pr2 = TaxTrans.Revenue.Principle2 - TaxTrans.Revenue.Principle2Pd
  Pr3 = TaxTrans.Revenue.Principle3 - TaxTrans.Revenue.Principle3Pd
  Pr4 = TaxTrans.Revenue.Principle4 - TaxTrans.Revenue.Principle4Pd
  Pr5 = TaxTrans.Revenue.Principle5 - TaxTrans.Revenue.Principle5Pd
  Intr = TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd
  Pen = TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd
  Adv = TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd
  LL = TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd
  TotalPd = OldRound(PrincPd)
  
  If TotalPd > Intr Then
    TotalPd = TotalPd - Intr
    IntrPd = Intr
  Else
    IntrPd = TotalPd
    TotalPd = 0
  End If
  
  If TotalPd > Pen Then
    TotalPd = TotalPd - Pen
    PenPd = Pen
  Else
    PenPd = TotalPd
    TotalPd = 0
  End If
  
  If TotalPd > Adv Then
    TotalPd = TotalPd - Adv
    AdvPd = Adv
  Else
    AdvPd = TotalPd
    TotalPd = 0
  End If
  
  If TotalPd > LL Then
    TotalPd = TotalPd - LL
    LLPd = LL
  Else
    LLPd = TotalPd
    TotalPd = 0
  End If
  
  If TotalPd > Pr1 Then
    TotalPd = TotalPd - Pr1
    Pr1Pd = Pr1
  Else
    Pr1Pd = TotalPd
    TotalPd = 0
  End If
  
  If TotalPd > Pr2 Then
    TotalPd = TotalPd - Pr2
    Pr2Pd = Pr2
  Else
    Pr2Pd = TotalPd
    TotalPd = 0
  End If
  
'  If TotalPd > Pr2 Then
'    TotalPd = TotalPd - Pr2
'    Pr2Pd = Pr2
'  Else
'    Pr2Pd = TotalPd
'    TotalPd = 0
'  End If
  
  If TotalPd > Pr3 Then
    TotalPd = TotalPd - Pr3
    Pr3Pd = Pr3
  Else
    Pr3Pd = TotalPd
    TotalPd = 0
  End If
  
  If TotalPd > Pr4 Then
    TotalPd = TotalPd - Pr4
    Pr4Pd = Pr4
  Else
    Pr4Pd = TotalPd
    TotalPd = 0
  End If
  
  If TotalPd > Pr5 Then
    TotalPd = TotalPd - Pr5
    Pr5Pd = Pr5
  Else
    Pr5Pd = TotalPd
    TotalPd = 0
  End If
  
  Get TTHandle, SaveRec, TaxTrans
  TaxTrans.Revenue.Interest = 0#
  TaxTrans.Amount = 0
  TaxTrans.TransDate = Date2Num%(ThisDate)
  TaxTrans.TranType = 9
  TaxTrans.Revenue.Principle1Pd = PrincPd
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.Revenue.LateListPd = 0
  TaxTrans.Revenue.RevOpt1Pd = 0
  TaxTrans.Revenue.RevOpt2Pd = 0
  TaxTrans.Revenue.RevOpt3Pd = 0
  TaxTrans.CustPin = CustPin
  TaxTrans.DiscXDate = 0
  TaxTrans.RealPin = ""
  TaxTrans.PersPin = ""
  TaxTrans.Posted2GL = "N"
  TaxTrans.TaxYear = TaxYear
  TaxTrans.DiscAmt = 0
  TaxTrans.OperNum = 0
  TaxTrans.FromPrePay = PrePayUsed
  TaxTrans.Description = "Bill# " + BillNum
  TaxTrans.CustomerRec = CustPin
  TaxTrans.BelongTo = BelongTo
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidUsed = PrePayUsed
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.InternalPin = CustPin
  TaxTrans.CntyPara = ""
  TaxTrans.CyclPara = ""
  TaxTrans.TShpPara = ""
  TaxTrans.BillType = PropType
  Put TTHandle, SaveRec, TaxTrans
  
  Get TTHandle, B4Trans, TaxTrans
  TaxTrans.LastTrans = SaveRec
  Put TTHandle, B4Trans, TaxTrans
  
  Get TTHandle, SaveRec, TaxTrans
  TaxTrans.LastTrans = BelongTo
  Put TTHandle, SaveRec, TaxTrans

  Get TTHandle, BillRec, TaxTrans
  TaxTrans.Revenue.Principle1Pd = Pr1Pd
  TaxTrans.Revenue.Principle2Pd = Pr2Pd
  TaxTrans.Revenue.Principle3Pd = Pr3Pd
  TaxTrans.Revenue.Principle4Pd = Pr4Pd
  TaxTrans.Revenue.Principle5Pd = Pr5Pd
  TaxTrans.Revenue.InterestPd = IntrPd
  TaxTrans.Revenue.PenaltyPd = PenPd
  TaxTrans.Revenue.CollectionPd = AdvPd
  TaxTrans.Revenue.LateListPd = LLPd
'  TaxTrans.Revenue.RevOpt1Pd = 0
'  TaxTrans.Revenue.RevOpt3Pd = 0
  Put TTHandle, BillRec, TaxTrans
  
'  Get TCHandle, CustPin, TaxCust
'  TaxCust.LastTrans = SaveRec
'  Put TCHandle, CustPin, TaxCust
  Close TCHandle
  Close TTHandle


End Sub

Private Sub BuildArray(ByVal ArrString As String, ByRef CArr() As Long, ByRef cnt As Integer)
  Dim x As Integer, y As Integer
  Dim ch As String
  Dim NewWord As String
  Dim Lgt As Integer
  ReDim arr(1 To 1) As Long
  Lgt = Len(ArrString)
  
  For x = 1 To Lgt
    ch = Mid(ArrString, x, 1)
    If QPTrim$(ch) = "" Then GoTo NextOne
    If ch <> "," Then
      NewWord = NewWord + ch
    Else
      cnt = cnt + 1
      ReDim Preserve arr(1 To cnt) As Long
      arr(cnt) = CLng(NewWord)
      NewWord = ""
    End If
NextOne:
  Next x
  CArr() = arr()
End Sub
Private Sub cmdFixStephensCityOld_Click()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim NextRec As Long
  Dim LastRec As Long
  Dim Amt As Double
  Dim UseThisOne As Boolean
  Dim ArrString As String
  Dim PCnt As Integer
  Dim ICnt As Integer
  Dim TransDate As Integer
  'cnt = 0
'  ArrString = "670, 1164, 1126, 360, 1217, 1374, 1295, 1315,"
'  ArrString = ArrString + "1368, 1302, 729, 1371, 1360, 1119, 494, "
'  ArrString = ArrString + "1054, 597, 1215, 1219, 1356, 1174, 358, "
'  ArrString = ArrString + "346, 1361, 359, 973, 565, 633, 349, 285, "
'  ArrString = ArrString + "442, 942, 622, 729, 775, 443, 606, 630, "
'  ArrString = ArrString + "888, 516, 1023, 349, 1040,"
'  Dim CArr() As Integer
'  Call BuildArray(ArrString, CArr(), cnt)
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile THandle, NumOfTRecs
  TransDate = Date2Num("02/01/2010")
'  For x = 1 To NumOfTRecs
'    Get THandle, x, TaxTrans
'    If TaxTrans.TransDate = TransDate Then
'      ClearTrans (x)
'    End If
'  Next x
  
'  For x = 1 To NumOfTRecs
'    Get THandle, x, TaxTrans
'    If TaxTrans.TransDate = TransDate And TaxTrans.TranType = 5 Then
'      ClearTrans (x)
'      If TaxTrans.TranType = 5 Then PCnt = PCnt + 1
'    End If
'  Next x
 
'  For x = 1 To NumOfTRecs
'    Get THandle, x, TaxTrans
'    If TaxTrans.TransDate = TransDate And TaxTrans.TranType = 4 Then
'      ClearTrans (x)
'      If TaxTrans.TranType = 4 Then Icnt = Icnt + 1
'    End If
'  Next x

  'many fixes on 10/30/2009
'  For x = 1 To cnt
'    Get TCHandle, CArr(x), TaxCust
'    NextRec = TaxCust.LastTrans
'    Do While NextRec > 0
'    Get THandle, NextRec, TaxTrans
'    ClearTrans (NextRec)
'    NextRec = TaxTrans.LastTrans
'    Loop
'  Next x
  
  '#194 9/3/2009
'  Get THandle, 6043, TaxTrans 'took an existing pay trans of .01 and increased to match bill amount
'  TaxTrans.Amount = 25.75
'  TaxTrans.Revenue.Principle1Pd = 25.75
'  Put THandle, 6043, TaxTrans
'
'  '#4431 9/3/2009
'  Call InsertCreditAtBillingTrans("05/23/2007", 5.8, 4431, 2007, 5.8, 2537, "P", 7948, 7862, "900")
'  Get THandle, 7948, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put THandle, 7948, TaxTrans
'
'  '#173 9/3/2009
'  Call InsertCreditAtBillingTrans("05/23/2007", 66.11, 173, 2007, 66.11, 7873, "P", 7925, 7873, "107")
'  Call ClearTrans(7925)
'
'  '#3723 9/3/2009
'  Call InsertCreditAtBillingTrans("05/14/2007", 54.53, 3723, 2007, 54.53, 7401, "R", 8890, 7401, "231")
'
'  '#3908 9/3/2009
'  Call InsertCreditAtBillingTrans("05/14/2007", 30.63, 3908, 2007, 30.63, 7470, "R", 9075, 7470, "412")
'
'  '#3492 9/3/2009
'  Call InsertCreditAtBillingTrans("05/14/2007", 170.66, 3492, 2007, 170.66, 7056, "R", 8662, 7456, "9")
'
'  '#3720 9/3/2009
'  Call ClearTrans(6549)
'
'  '#3790 9/3/2009
'  Call ClearTrans(20492)
'  Call ClearTrans(17981)
'  Call ClearTrans(20880)
'  Call ClearTrans(11156)
'  Call ClearTrans(11155)
'  Get THandle, 17980, TaxTrans
'  TaxTrans.Revenue.Principle1 = 41.35
'  TaxTrans.Revenue.Principle1Pd = 41.35
'  Put THandle, 17980, TaxTrans
'
'  Get THandle, 8958, TaxTrans
'  TaxTrans.Amount = 41.35
'  Put THandle, 8958, TaxTrans
'
'  '#1688 9/3/2009
'  Call ClearTrans(7746)
'  Call ClearTrans(7744)
'  Call ClearTrans(7754)
'  Call ClearTrans(7752)
'  Call ClearTrans(7750)
'  Call ClearTrans(7748)
'  Call ClearTrans(7952)
'
'  '#3892 9/3/2009
'  Call InsertCreditAtBillingTrans("05/14/2007", 31.12, 3892, 2007, 31.12, 7468, "R", 9059, 7468, "396")
  
  
  '#72 3/24/09
'  Get THandle, 7488, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 144.12
'  TaxTrans.Amount = 144.12
'  Put THandle, 7488, TaxTrans
'  Close
'  '#1946 3/24/09
'  Call InsertPayTrans("06/09/2006", 27.84, 1946, 2006, 27.84, 7891, "P", 10049, 7891, "5222206")
'  Call ClearTrans(7719)
'  Call ChangePrePayToPay(2784, 27.84, "622")
'
'  '#1615 3/24/09
'  OpenTaxTransFile THandle, NumOfTRecs
'  Get THandle, 2722, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 39.44
'  TaxTrans.Revenue.Principle5Pd = 4.39
'  TaxTrans.Amount = 43.83
'  Put THandle, 2722, TaxTrans
'
'  Call ChangePrePayToPay(2723, 43.83, "450")
'  Close
'  '#1806 3/24/09
'
'  OpenTaxTransFile THandle, NumOfTRecs
'  Call ChangePrePayToPay(3329, 32.67, "553")
'  Call InsertPayTrans("05/22/2007", 32.66, 1806, 2007, 0, 7937, "P", 10011, 7937, "52222007")
'  Call ClearTrans(7780)
'  Close
  
  '#3610 3/24/09
'  OpenTaxTransFile THandle, NumOfTRecs
'  Get THandle, 8781, TaxTrans
'  TaxTrans.Amount = 72.13
'  Put THandle, 8781, TaxTrans
'
'  Get THandle, 7405, TaxTrans
'  TaxTrans.Amount = 72.13
'  TaxTrans.Revenue.Principle1 = 72.13
'  TaxTrans.Revenue.Principle1Pd = 72.13
'  Put THandle, 7405, TaxTrans
'
'  Get THandle, 20912, TaxTrans
'  TaxTrans.BelongTo = 7405
'  Put THandle, 20912, TaxTrans
'
'  Get THandle, 5956, TaxTrans
'  TaxTrans.Amount = 72.13
'  TaxTrans.Revenue.Principle1 = 72.13
'  Put THandle, 5956, TaxTrans
'
'  Call ClearTrans(24540)
'  Call ClearTrans(6877)
'  Call ClearTrans(11300)
'  Call ClearTrans(11299)
'  Call ClearTrans(11298)
'  Call ClearTrans(4069)
'  Call ClearTrans(19736)
'  Call ClearTrans(19738)
'  Call ClearTrans(19737)
'  Call ChangePrePayToPay(17868, 72.13, "351")
'  Call ChangePrePayToPay(3366, 72.12, "124")
'
'  'cust 329 2/9/09
'
''  Get THandle, 7935, TaxTrans
''  TaxTrans.Revenue.Principle1Pd = 0
''  Put THandle, 7935, TaxTrans
''
''  'cust 4023 2/9/09
''  Get THandle, 6076, TaxTrans
''  TaxTrans.RealPin = ""
''  Put THandle, 6076, TaxTrans
''
''  Get THandle, 6075, TaxTrans
''  TaxTrans.RealPin = ""
''  Put THandle, 6075, TaxTrans
'
'  'cust 264 2/9/09
'  Get THandle, 6582, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put THandle, 6582, TaxTrans
'
'  FixPayPlusOPThatShouldHaveBeenApplied '2/9/09
  
'  For x = 1 To NumOfTCRecs
'    Get TCHandle, x, TaxCust
'    UseThisOne = False
'    NextRec = TaxCust.LastTrans
'    LastRec = 0
'    Amt = 0
'    Do While NextRec > 0
'      Get THandle, NextRec, TaxTrans
'      If MakeRegDate(TaxTrans.TransDate) = "05/14/2007" Then
'        Amt = TaxTrans.Amount
'        If TaxTrans.OperNum = 3 And TaxTrans.TranType = 1 Then
'          Get THandle, LastRec, TaxTrans
'          TaxTrans.Revenue.PrePaidAmt = 0
'          TaxTrans.Revenue.PrePaidUsed = Amt
'          TaxTrans.FromPrePay = Amt
'          TaxTrans.Revenue.PrePaidBal = 0
'          TaxTrans.Revenue.Principle1Pd = 0
'          TaxTrans.Amount = 0
'          Put THandle, LastRec, TaxTrans
'          Get THandle, NextRec, TaxTrans
'          UseThisOne = True
'        End If
'      End If
'      If UseThisOne = True And TaxTrans.OperNum = 1 And MakeRegDate(TaxTrans.TransDate) = "05/14/2007" Then
'        TaxTrans.Amount = 0
'        TaxTrans.Revenue.Principle1 = 0
'        TaxTrans.Revenue.PrePaidAmt = 0
'        TaxTrans.Revenue.PrePaidUsed = 0
'        TaxTrans.FromPrePay = 0
'        TaxTrans.Revenue.PrePaidBal = 0
'        TaxTrans.Revenue.Principle1Pd = 0
'        Put THandle, NextRec, TaxTrans
'      End If
'    LastRec = NextRec
'    NextRec = TaxTrans.LastTrans
'    Loop
'  Next x
  
  'fix fo 4149
'  Get THandle, 8451, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put THandle, 8451, TaxTrans
'
'  Get THandle, 8450, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put THandle, 8450, TaxTrans
'
'  Get THandle, 7245, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.PrePaidAmt = 3.04
'  TaxTrans.Revenue.PrePaidUsed = 3.04
'  TaxTrans.FromPrePay = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.Principle1Pd = 3.04
'  Put THandle, 7245, TaxTrans
  
  
  'fix for 3640
'   Get THandle, 7252, TaxTrans
'   TaxTrans.Amount = 0
'   TaxTrans.Revenue.PrePaidUsed = 0
'   Put THandle, 7252, TaxTrans
'
'   Get THandle, 7251, TaxTrans
'   TaxTrans.Amount = 0
'   TaxTrans.Revenue.Principle1 = 0
'   TaxTrans.Revenue.Principle1Pd = 0
'   Put THandle, 7251, TaxTrans
'
'   Get THandle, 6952, TaxTrans
'   TaxTrans.Amount = 0
'   TaxTrans.Revenue.PrePaidUsed = 0
'   Put THandle, 6952, TaxTrans
'
'   Get THandle, 6951, TaxTrans
'   TaxTrans.Amount = 0
'   TaxTrans.Revenue.Principle1 = 0
'   TaxTrans.Revenue.Principle1Pd = 0
'   Put THandle, 6951, TaxTrans
  
'   'fix for 3503
'   Get THandle, 7161, TaxTrans
'   TaxTrans.Amount = 138.7
'   TaxTrans.Revenue.PrePaidUsed = 0
'   Put THandle, 7161, TaxTrans
'
'   'fix for 3492
'   Get THandle, 7457, TaxTrans
'   TaxTrans.Amount = 170.66
'   TaxTrans.Revenue.PrePaidUsed = 0
'   Put THandle, 7457, TaxTrans
'
'   'fix for 3561
'   Get THandle, 7407, TaxTrans
'   TaxTrans.Amount = 107.35
'   TaxTrans.Revenue.PrePaidUsed = 0
'   Put THandle, 7407, TaxTrans
'
'   Get THandle, 7107, TaxTrans
'   TaxTrans.Amount = 107.35
'   TaxTrans.Revenue.PrePaidUsed = 0
'   Put THandle, 7107, TaxTrans


'  'fix for 164
'  Get THandle, 7202, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put THandle, 7202, TaxTrans
'
'  Get THandle, 7201, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  Put THandle, 7201, TaxTrans
'
'  Get THandle, 6902, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put THandle, 6902, TaxTrans
'
'  Get THandle, 6901, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  Put THandle, 6901, TaxTrans
'
'  'fix for 186
'  Get THandle, 7208, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put THandle, 7208, TaxTrans
'
'  Get THandle, 7207, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  Put THandle, 7207, TaxTrans
'
'  Get THandle, 6908, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put THandle, 6908, TaxTrans
'
'  Get THandle, 6907, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  Put THandle, 6907, TaxTrans
'
'  'fix for 142
'  Get THandle, 7212, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put THandle, 7212, TaxTrans
'
'  Get THandle, 7211, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  Put THandle, 7211, TaxTrans
'
'  Get THandle, 6912, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put THandle, 6912, TaxTrans
'
'  Get THandle, 6911, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  Put THandle, 6911, TaxTrans
'
'  'fix for 1962
'  Get THandle, 7232, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put THandle, 7232, TaxTrans
'
'  Get THandle, 7231, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  Put THandle, 7231, TaxTrans
'
'  Get THandle, 6932, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put THandle, 6932, TaxTrans
'
'  Get THandle, 6931, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  Put THandle, 6931, TaxTrans
'
'  'fix for 45
'  Get THandle, 7215, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put THandle, 7215, TaxTrans
'
'  Get THandle, 7214, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  Put THandle, 7214, TaxTrans
'
'  Get THandle, 6915, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put THandle, 6915, TaxTrans
'
'  Get THandle, 6914, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  Put THandle, 6914, TaxTrans
'
'  'fix for 1155
'  Get THandle, 7228, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put THandle, 7228, TaxTrans
'
'  Get THandle, 7227, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  Put THandle, 7227, TaxTrans
'
'  Get THandle, 6928, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put THandle, 6928, TaxTrans
'
'  Get THandle, 6927, TaxTrans
'  TaxTrans.Amount = 0
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle1 = 0
'  Put THandle, 6927, TaxTrans
  
  
  Close
  Call TaxMsg(900, "A total of " + CStr(PCnt) + " penalty transactions and " + CStr(ICnt) + " interest transactions have been cleared.")

End Sub

Private Sub cmdFixRemington_Click()
  Call BuildTextFile
  Exit Sub
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  
  OpenTaxTransFile THandle, NumOfTRecs
  'fix for 2105
   Get THandle, 66132, TaxTrans
   TaxTrans.LastTrans = 66076
   Put THandle, 66132, TaxTrans
  
  'fix for 2097
   Get THandle, 66133, TaxTrans
   TaxTrans.LastTrans = 66131
   Put THandle, 66133, TaxTrans
   
   Get THandle, 66131, TaxTrans
   TaxTrans.LastTrans = 66130
   TaxTrans.BelongTo = TaxTrans.BelongTo
   Put THandle, 66131, TaxTrans
  
  Close THandle
  MsgBox ("Finished")

End Sub

Private Sub cmdFixStPauls_Click()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  
  OpenTaxTransFile THandle, NumOfTRecs
  'fix for 1212
  Get THandle, 7200, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  Put THandle, 7200, TaxTrans
  
  'fix for 1213
  Get THandle, 7199, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  Put THandle, 7199, TaxTrans
  
  'fix for 434
  Get THandle, 5676, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.PenaltyPd = 0
  Put THandle, 5676, TaxTrans
  
  Get THandle, 5675, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.PenaltyPd = 0
  Put THandle, 5675, TaxTrans
  
  'fix for 435
  Get THandle, 5674, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.PenaltyPd = 0
  Put THandle, 5674, TaxTrans
  
  Get THandle, 5673, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.PenaltyPd = 0
  Put THandle, 5673, TaxTrans
  
  Close
  Call TaxMsg(900, "Finished")
  
End Sub

Private Sub cmdFixStuart_Click()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  
  OpenTaxTransFile THandle, NumOfTRecs
  'fix cust #102
  Get THandle, 519, TaxTrans
  TaxTrans.RealPin = QPTrim$(TaxTrans.RealPin) + "A"
  Put THandle, 519, TaxTrans
  
  'fix cust #253
  Get THandle, 87, TaxTrans
  TaxTrans.RealPin = QPTrim$(TaxTrans.RealPin) + "A"
  Put THandle, 87, TaxTrans
  
  'fix cust #345
  Get THandle, 607, TaxTrans
  TaxTrans.RealPin = QPTrim$(TaxTrans.RealPin) + "A"
  Put THandle, 607, TaxTrans
  
  'fix cust# 2130
'  Get THandle, 415, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 2.85
'  Put THandle, 415, TaxTrans
'
'  Get THandle, 416, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 8.93
'  Put THandle, 416, TaxTrans
'
'  Get THandle, 417, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 11.21
'  Put THandle, 417, TaxTrans
'
'  Get THandle, 418, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 245.29
'  Put THandle, 418, TaxTrans
'
'  'fix cust# 1538
'  Get THandle, 8313, TaxTrans
'  TaxTrans.Revenue.Principle5Pd = 2.85
'  Put THandle, 8313, TaxTrans
'
'  Get THandle, 9905, TaxTrans
'  TaxTrans.Revenue.Principle5Pd = 2.55
'  Put THandle, 9905, TaxTrans
'
'  Get THandle, 12807, TaxTrans
'  TaxTrans.Revenue.Principle5Pd = 2.55
'  Put THandle, 12807, TaxTrans
  
  Close
  
  Call TaxMsg(900, "Finished.")
  
End Sub

Private Sub cmdFixWarsaw_Click()
  Dim RealRec As PropertyRecType
  Dim x As Long
  Dim NumOfRRecs As Long
  Dim RHandle As Integer
  
  OpenRealPropFile RHandle, NumOfRRecs
  For x = 1 To NumOfRRecs
    Get RHandle, x, RealRec
    RealRec.EXMPOTHR = 0
    Put RHandle, x, RealRec
  Next x
  
  Close
  
  Call TaxMsg(600, "Almost finished. Go to cust #1622 (real) and enter $43,680.00 as other exemption. Then go to cust #1756 (real) and enter $67,910.00 as other exemption.")
End Sub

Private Sub cmdMiddletown_Click()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  
  OpenTaxTransFile THandle, NumOfTRecs
  
  'fix for #1330
  Get THandle, 11134, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 22.78
  Put THandle, 11134, TaxTrans
  
  Get THandle, 11980, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 0
  Put THandle, 11980, TaxTrans

  'fix for #1850
  Get THandle, 8002, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 61.36
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 6.14
  Put THandle, 8002, TaxTrans
  
  Get THandle, 10247, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.PenaltyPd = 0
  TaxTrans.Amount = 0
  Put THandle, 10247, TaxTrans
  
  'fix for #781
  Get THandle, 11194, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 107.19
  Put THandle, 11194, TaxTrans
  
  Get THandle, 11982, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 107.19
  TaxTrans.Amount = 0
  Put THandle, 11982, TaxTrans
  
  'fix for #2072
  Get THandle, 293, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 46.67
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 8.22
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 4.67
  Put THandle, 293, TaxTrans
  
  Get THandle, 1295, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 46.67
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 4.11
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 4.67
  Put THandle, 1295, TaxTrans
  
  'fix for #225
  Get THandle, 11297, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 34.78
  Put THandle, 11297, TaxTrans
  
  Get THandle, 11985, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 0
  Put THandle, 11985, TaxTrans
  
  'fix for #1247
  Get THandle, 10811, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 26.63
  Put THandle, 10811, TaxTrans
  
  Get THandle, 11976, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 0
  Put THandle, 11976, TaxTrans
  
  'fix for #1331
  Get THandle, 11133, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 27#
  Put THandle, 11133, TaxTrans
  
  Get THandle, 11981, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 0
  Put THandle, 11981, TaxTrans
  
  'fix for #1356
  Get THandle, 11259, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 8.16
  Put THandle, 11259, TaxTrans
  
  Get THandle, 11983, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 0
  Put THandle, 11983, TaxTrans
  
  'fix for #1940
  Get THandle, 163, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 0.07
  Put THandle, 163, TaxTrans
  
  Get THandle, 2137, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 0
  Put THandle, 2137, TaxTrans
  
  'fix for #1007
  Get THandle, 11355, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 8.16
  Put THandle, 11355, TaxTrans
  
  Get THandle, 11986, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 0
  Put THandle, 11986, TaxTrans
  
  'fix for #1821
  Get THandle, 9880, TaxTrans
  TaxTrans.Revenue.Principle1 = 0
  TaxTrans.Amount = 0
  Put THandle, 9880, TaxTrans
  
  'fix for #688
  Get THandle, 10825, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 11.25
  Put THandle, 10825, TaxTrans
  
  Get THandle, 11977, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 0
  Put THandle, 11977, TaxTrans
  
  'fix for #68
  Get THandle, 10876, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 97.89
  Put THandle, 10876, TaxTrans
  
  Get THandle, 11978, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 0
  Put THandle, 11978, TaxTrans
  
  'fix for #883
  Get THandle, 6213, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 34.12
  Put THandle, 6213, TaxTrans
  
  Get THandle, 6703, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 0
  Put THandle, 6703, TaxTrans
  
  'fix for #671
  Get THandle, 8248, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 7.88
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 0.78
  Put THandle, 8248, TaxTrans
  
  Get THandle, 10203, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.PenaltyPd = 0
  TaxTrans.Amount = 0
  Put THandle, 10203, TaxTrans
  
  Get THandle, 10202, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.PenaltyPd = 0
  TaxTrans.Amount = 0
  Put THandle, 10202, TaxTrans
  
  'fix for #2168
  Get THandle, 3126, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 45.11
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 3.97
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 4.51
  Put THandle, 3126, TaxTrans
  
  Get THandle, 389, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 45.11
  Put THandle, 389, TaxTrans
  
  'fix for #927
  Get THandle, 1159, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 21.88
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 5.79
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 2.19
  Put THandle, 1159, TaxTrans
  
  Get THandle, 1160, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 15.63
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 5.52
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 1.56
  Put THandle, 1160, TaxTrans
  
  'fix for #1525
  Get THandle, 16410, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 2.86
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 4.91
  Put THandle, 16410, TaxTrans
  
  'fix for #509
  Get THandle, 1270, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 26.56
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 7.02
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 2.66
  Put THandle, 1270, TaxTrans
  
  'fix for #1269
  Get THandle, 10906, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 34.96
  Put THandle, 10906, TaxTrans
  
  Get THandle, 11979, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 0
  Put THandle, 11979, TaxTrans
  
  'fix for #470
  Get THandle, 744, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 164.96
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 28.57
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 16.5
  Put THandle, 744, TaxTrans
  
  Get THandle, 1268, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 157.64
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 18.75
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 15.77
  Put THandle, 1268, TaxTrans
  
  'fix for #470
  Get THandle, 1269, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 125.77
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 17.72
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 12.58
  Put THandle, 1269, TaxTrans
  
  Get THandle, 3679, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 181.12
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 32.52
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 18.12
  Put THandle, 3679, TaxTrans
  
  'fix for #2208
  Get THandle, 2936, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 2.26
  Put THandle, 2936, TaxTrans
  
  Get THandle, 3987, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 205.04
  Put THandle, 3987, TaxTrans
  
  'fix for #930
  Get THandle, 1163, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 15.63
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 2.89
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 1.56
  Put THandle, 1163, TaxTrans
  
  Get THandle, 1164, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 15.63
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 4.14
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 1.56
  Put THandle, 1164, TaxTrans
  
  Get THandle, 4286, TaxTrans
  TaxTrans.Revenue.Principle1 = 0.13
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 31.13
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 6.77
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 3.12
  Put THandle, 4286, TaxTrans
  
  'fix for #2097
  Get THandle, 318, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 44.07
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 4.41
  Put THandle, 318, TaxTrans
  
  Get THandle, 2931, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 44.07
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 3.88
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 4.41
  Put THandle, 2931, TaxTrans
  
  'fix for #472
  Get THandle, 1278, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 6.26
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 1.4
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 0.62
  Put THandle, 1278, TaxTrans
  
  Get THandle, 1279, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 6.26
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 1.68
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 0.62
  Put THandle, 1279, TaxTrans
  
  Get THandle, 1280, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 6.26
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 1.96
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 0.62
  Put THandle, 1280, TaxTrans
  
  'fix for #1363
  Get THandle, 11279, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 54.94
  Put THandle, 11279, TaxTrans
  
  Get THandle, 11984, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Amount = 0
  Put THandle, 11984, TaxTrans
  
  'fix for #2207
  Get THandle, 3210, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 11.24
  Put THandle, 3210, TaxTrans
  
  Get THandle, 3988, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 510.38
  Put THandle, 3988, TaxTrans
  
  'fix for #509
  Get THandle, 1271, TaxTrans
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 31.88
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 11.24
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 3.19
  Put THandle, 1271, TaxTrans
  Close
  
  Call TaxMsg(900, "Finished.")
End Sub

Private Sub cmdFixRestOfTrans_Click()
  Call FixFinalFew
End Sub
'Private Sub FixFinalFew2()
'  Dim TaxTrans As TaxTransactionType
'  Dim TTHandle As Integer
'  Dim NumOfTTRecs As Long
'  Dim TaxCust As TaxCustType
'  Dim TCHandle As Integer
'  Dim NumOfTCRecs As Long
'  Dim x As Long, y As Long
'  Dim NextRec As Long
'  Dim BottomRec As Long
'  Dim TopRec As Long
'  Dim Found As Boolean
'  Dim AHandle As Integer
'  Dim BelongTo As Long
'  Dim Bal As Double
'
'  OpenTaxTransFile TTHandle, NumOfTTRecs
'  OpenTaxCustFile TCHandle, NumOfTCRecs
'  AHandle = FreeFile
'  Open "Trans9Error.txt" For Output As AHandle
'   'fix for 1258
'  Call InsertCreditAtBillingTrans("11/25/2008", 2158.21, 1258, 2007, 2158.21, 24238, "Real", 25170, 24238, "901")
'  Get TTHandle, 24238, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 2158.21
'  TaxTrans.Revenue.InterestPd = 0
'  Put TTHandle, 24238, TaxTrans
'
'  'fix for 97
'  Get TTHandle, 7908, TaxTrans
'  TaxTrans.TranType = 2
'  TaxTrans.Revenue.Principle1Pd = 8.26
'  TaxTrans.Amount = 8.26
'  TaxTrans.Amount = TaxTrans.Revenue.PrePaidUsed
'  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.PrePaidUsed
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put TTHandle, 7908, TaxTrans
'
'  'fix for 3805
'  Call ClearTrans(8617)
'  Get TTHandle, 7325, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 38.83
'  Put TTHandle, 7325, TaxTrans
'
'  Get TTHandle, 18321, TaxTrans
'  TaxTrans.TranType = 2
'  TaxTrans.Amount = 38.83
'  TaxTrans.Revenue.Principle1Pd = 38.83
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put TTHandle, 18321, TaxTrans
'
'  Get TTHandle, 6797, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put TTHandle, 6797, TaxTrans
'
'   Get TTHandle, 7025, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put TTHandle, 7025, TaxTrans
'
'  Get TTHandle, 45510, TaxTrans
'  TaxTrans.Revenue.PrePaidBal = 0
'  Put TTHandle, 45510, TaxTrans
'
'  'fix for 1781
'  Call ClearTrans(7949)
'  'fix for 1781 -> 1015
'  Get TTHandle, 8307, TaxTrans
'  TaxTrans.LastTrans = 8185
'  Put TTHandle, 8307, TaxTrans
'
'  Get TTHandle, 8261, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 14.05
'  Put TTHandle, 8261, TaxTrans
'
'  Get TTHandle, 11611, TaxTrans
'  TaxTrans.LastTrans = 8262
'  Put TTHandle, 11611, TaxTrans
'
'  Get TTHandle, 8262, TaxTrans
'  TaxTrans.LastTrans = 8261
'  TaxTrans.CustomerRec = 1015
'  TaxTrans.CustPin = 1015
'  Put TTHandle, 8262, TaxTrans
' 'GoTo Skip
'  'fix for 4140
'  Get TTHandle, 7187, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 40.02
'  Put TTHandle, 7187, TaxTrans
'
'
'  Call BuildMBvsCustHistArr
'
'  For x = 1 To CArrCnt
'    Get TCHandle, CArr(x), TaxCust
'    NextRec = TaxCust.LastTrans
'    Do While NextRec > 0
'      Get TTHandle, NextRec, TaxTrans
''        If NextRec = 8262 Then Stop
'      '  Print #AHandle, CStr(CArr(x)) + "~" + CStr(NextRec)
'        If TaxTrans.TranType = 9 And TaxTrans.CustPin <> CArr(x) Then
'           BelongTo = TaxTrans.BelongTo
''          If TaxTrans.CustPin = 1806 Then Stop
'          Get TCHandle, TaxTrans.CustPin, TaxCust
'          If TaxCust.Deleted <> 0 Then
'            Call ClearTrans(NextRec)
'          Else
'            Get TTHandle, BelongTo, TaxTrans
'            Bal = 0
'            Bal# = OldRound#(Bal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
'            Bal# = OldRound#(Bal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
'            Bal# = OldRound#(Bal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
'            Bal# = OldRound#(Bal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
'            Bal# = OldRound#(Bal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
'            If Bal = 0 Then
'              Call ClearTrans(NextRec)
'              Print #AHandle, CStr(TaxTrans.CustPin) + "~" + CStr(NextRec)
'            Else
'              Print #AHandle, CStr(TaxTrans.CustPin) + "~" + CStr(NextRec) + "~" + Using("##,###.00", Bal)
'           End If
'          End If
'        End If
'          Get TTHandle, NextRec, TaxTrans
'         NextRec = TaxTrans.LastTrans
'    Loop
'   Next x
'  ' GoTo Skip
'
'  'fix for 4393
'  Get TTHandle, 7951, TaxTrans
'  TaxTrans.Amount = 13.49
'  TaxTrans.Revenue.Principle2Pd = 13.49
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.TranType = 2
'  Put TTHandle, 7951, TaxTrans
'
'  Call ClearTrans(7863)
'  Call ClearTrans(7853)
'
'
'  'fix for 4253
'  Call ClearTrans(8639)
'
'  'fix for 4158
'  Get TTHandle, 7356, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0.76
'  Put TTHandle, 7356, TaxTrans
'
'
'  'fix for 3886
'  Call ClearTrans(7834)
'  Call ClearTrans(7822)
'  Call ClearTrans(7820)
'
'  'fix for 1832
'  Call ClearTrans(5583)
'  Call ClearTrans(6579)
'
'  Get TTHandle, 5100, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 10.95
'  Put TTHandle, 5100, TaxTrans
'
'  Get TTHandle, 6578, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 10.94
'  Put TTHandle, 6578, TaxTrans
'
'  Get TTHandle, 7490, TaxTrans
'  TaxTrans.Amount = 10.94
'  TaxTrans.Revenue.Principle1Pd = 10.94
'  Put TTHandle, 7490, TaxTrans
'
'
'
'  'fix for 1735
'  Get TTHandle, 7972, TaxTrans
'  TaxTrans.Revenue.Principle2 = 0
'  Put TTHandle, 7972, TaxTrans
'
'
'
'  'fix for 1713
'  Call ClearTrans(26894)
'  Get TTHandle, 25475, TaxTrans
'  TaxTrans.Amount = 7.35
'  TaxTrans.Revenue.Interest = 0
'  TaxTrans.Revenue.Penalty = 0
'  Put TTHandle, 25475, TaxTrans
'
'  Get TTHandle, 26893, TaxTrans
'  TaxTrans.Revenue.Principle1 = 19.58
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put TTHandle, 26893, TaxTrans
'
'  Get TTHandle, 32753, TaxTrans
'  TaxTrans.Amount = 52.92
'  TaxTrans.Revenue.Principle1 = 52.92
'  Put TTHandle, 32753, TaxTrans
'
'  Get TTHandle, 12097, TaxTrans
'  TaxTrans.Revenue.Interest = 0
'  TaxTrans.Revenue.Penalty = 0
'  Put TTHandle, 12097, TaxTrans
'
'  Get TTHandle, 9975, TaxTrans
'  TaxTrans.Revenue.Interest = 0
'  TaxTrans.Revenue.Penalty = 0
'  Put TTHandle, 9975, TaxTrans
'
'    'fix for 1568
'  Call ClearTrans(7304)
'
'  Get TTHandle, 37432, TaxTrans
'  TaxTrans.Amount = 151.49
'  TaxTrans.Revenue.PenaltyPd = 0
'  TaxTrans.Revenue.Penalty = 151.49
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.TranType = 14
'  Put TTHandle, 37432, TaxTrans
'
'  Get TTHandle, 31510, TaxTrans
'  TaxTrans.Revenue.Principle1 = 64.06
'  TaxTrans.Revenue.Principle1Pd = 52.21
'  TaxTrans.Revenue.Penalty = 75.84
'  TaxTrans.Revenue.PenaltyPd = 50.94
'  Put TTHandle, 31510, TaxTrans
'
'  Get TTHandle, 4966, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 28.84
'  Put TTHandle, 4966, TaxTrans
'
'  Call ClearTrans(5537)
'
'  'fix for 933
'  Get TTHandle, 11377, TaxTrans
'  TaxTrans.Amount = 44.8
'  Put TTHandle, 11377, TaxTrans
'
'  Get TTHandle, 9812, TaxTrans
'  TaxTrans.Revenue.Principle1 = 0
'  TaxTrans.PPTRADisc = 0
'  Put TTHandle, 9812, TaxTrans
'  'fix for 427
'  Get TTHandle, 29614, TaxTrans
'  TaxTrans.Amount = 10.65
'  TaxTrans.Revenue.Principle1 = 6.96
'  TaxTrans.Revenue.Interest = 3.34
'  TaxTrans.Revenue.Penalty = 0.35
'  Put TTHandle, 29614, TaxTrans
'
'  Call ClearTrans(29615)
'
'  Get TTHandle, 63, TaxTrans
'  TaxTrans.Revenue.Interest = 0
'  TaxTrans.Revenue.Penalty = 0
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.PenaltyPd = 0
'  Put TTHandle, 63, TaxTrans
'
'  'fix for 880
'  For x = 1 To NumOfTCRecs
'    Get TCHandle, x, TaxCust
'    NextRec = TaxCust.LastTrans
'    Do While NextRec > 0
'      Get TTHandle, NextRec, TaxTrans
'      If NextRec = 7873 Then
'        Found = False
'        BottomRec = TaxTrans.LastTrans
'        If TaxCust.LastTrans = NextRec Then
'          TaxCust.LastTrans = BottomRec
'          Put TCHandle, x, TaxCust
'          Exit Do
'         Else
'          Get TTHandle, TopRec, TaxTrans
'          TaxTrans.LastTrans = BottomRec
'          Put TTHandle, TopRec, TaxTrans
'          Exit Do
'        End If
'      End If
'      TopRec = NextRec
'      NextRec = TaxTrans.LastTrans
'    Loop
'  Next x
'
'  Get TCHandle, 880, TaxCust
'  NextRec = TaxCust.LastTrans
'  Do While NextRec > 0
'     Get TTHandle, NextRec, TaxTrans
'     If NextRec = 7874 Then
'       BottomRec = TaxTrans.LastTrans
'       TaxTrans.LastTrans = 7873
'       Put TTHandle, 7874, TaxTrans
'       Get TTHandle, 7873, TaxTrans
'       TaxTrans.LastTrans = BottomRec
'       TaxTrans.Revenue.Principle1Pd = 0
'       Put TTHandle, 7873, TaxTrans
'       Exit Do
'     End If
'     TopRec = NextRec
'     NextRec = TaxTrans.LastTrans
'  Loop
'  Call ClearTrans(7874)
'
'  'fix for 173
'  Call ClearTrans(29591)
'
'Skip:
'  Close
'  MsgBox ("Done.")
'
'End Sub
Private Sub FixStephensCityNov30Ten()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  
  OpenTaxTransFile THandle, NumOfTRecs
  'fix for #3761
  Get THandle, 62550, TaxTrans
  TaxTrans.BelongTo = 60097
  TaxTrans.Description = "10020218"
  Put THandle, 62550, TaxTrans
  
  Get THandle, 60097, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 18.8
  Put THandle, 60097, TaxTrans
  
  Get THandle, 7294, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 50.23
  TaxTrans.Revenue.Interest = 0
  Put THandle, 7294, TaxTrans
  
  Get THandle, 7883, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 50.23
  TaxTrans.Amount = 50.23
  Put THandle, 7883, TaxTrans
  
  Call ClearTrans(6994)
  Call ClearTrans(6766)
  Call ClearTrans(7884)
  Call ClearTrans(7885)
  Call ClearTrans(52663)
  Call ClearTrans(55612)
  Call ClearTrans(55839)
  Call ChangePayToPayPlusPrePay(14106, 100.47, 50.24, 50.23, "601")
  Get THandle, 14106, TaxTrans
  TaxTrans.BelongTo = 13670
  Put THandle, 14106, TaxTrans
  
  'fix for #3780
  Get THandle, 62582, TaxTrans
  TaxTrans.BelongTo = 60095
  TaxTrans.Description = "10020216"
  TaxTrans.Revenue.Principle1Pd = 56.03
  TaxTrans.Revenue.InterestPd = 0
  Put THandle, 62582, TaxTrans
  
  Get THandle, 60095, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 57.46
  Put THandle, 60095, TaxTrans
  
  Get THandle, 7253, TaxTrans
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.Interest = 0
  Put THandle, 7253, TaxTrans
  
  Get THandle, 7866, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 48.71
  TaxTrans.Amount = 48.71
  Put THandle, 7866, TaxTrans
  
  Call ClearTrans(7868)
  Call ClearTrans(7867)
  Call ClearTrans(6953)
  Call ClearTrans(6725)
  Call ClearTrans(55843)
  Call ClearTrans(55616)
  Call ClearTrans(52667)
  Call ChangePayToPayPlusPrePay(14673, 97.43, 48.72, 48.71, "181")
  Get THandle, 14673, TaxTrans
  TaxTrans.BelongTo = 13226
  Put THandle, 14673, TaxTrans
 
  Close
  MsgBox ("Finished.")

End Sub


Private Sub FixStephensCityOct25Ten()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim x As Long
  
  OpenTaxPropFile RHandle, NumOfRRecs
  Get RHandle, 722, RealRec
  RealRec.NextRec = 0
  Put RHandle, 722, RealRec
  Call ChangePayToPrePay(8446, 35.63)
  
  OpenTaxTransFile THandle, NumOfTRecs
  Get THandle, 8447, TaxTrans
  TaxTrans.BelongTo = 7351
  Put THandle, 8447, TaxTrans
  
  Get THandle, 61897, TaxTrans
  TaxTrans.BelongTo = 60627
  TaxTrans.Description = "10020724"
  Put THandle, 61897, TaxTrans
  
  Get THandle, 60627, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 35.47
  Put THandle, 60627, TaxTrans
  Call ClearTrans(55845)
  Call ClearTrans(55618)
  Call ClearTrans(52669)
  
  Close
  MsgBox ("Done")
End Sub

Private Sub cmdFixStephensCity_Click()
  Call FixStephensCityNov30Ten
  Exit Sub
  Call FixStephensCityOct25Ten
  Exit Sub
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTRecs As Long
  Dim TransDate As Integer
  Dim x As Long
  Dim y As Long
  Dim IntTot As Double
  Dim ArrCnt As Integer
  Dim Bills() As Long
  Dim SaveRec As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long

  OpenTaxTransFile TTHandle, NumOfTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  'fix for 3720
   Get TCHandle, 3720, TaxCust
   Call InsertPrepaidOnlyTrans("8/02/2010", 2.98, 3720, 2010, TaxCust.LastTrans, "Real")
   
   Get TTHandle, 31336, TaxTrans
   TaxTrans.Revenue.Interest = 15.54
   TaxTrans.Revenue.InterestPd = 15.54
   Put TTHandle, 31336, TaxTrans
   
   Get TTHandle, 44496, TaxTrans
   TaxTrans.Revenue.Principle1Pd = 32.1
   TaxTrans.Revenue.InterestPd = 15.54
   TaxTrans.Revenue.PenaltyPd = 5.36
   TaxTrans.Amount = 53
   Put TTHandle, 44496, TaxTrans
   
   Get TTHandle, 50840, TaxTrans
   TaxTrans.Revenue.Principle1Pd = 21.44
   TaxTrans.Revenue.InterestPd = 0
   TaxTrans.Amount = 21.44
   Put TTHandle, 50840, TaxTrans
   
   Get TTHandle, 50839, TaxTrans
   TaxTrans.Amount = 57
   TaxTrans.Revenue.Principle1Pd = 53.54
   TaxTrans.Revenue.InterestPd = 3.46
   Put TTHandle, 50839, TaxTrans
   
   Get TTHandle, 45472, TaxTrans
   TaxTrans.Revenue.Principle1Pd = 53.54
   TaxTrans.Revenue.InterestPd = 3.46
   Put TTHandle, 45472, TaxTrans
   
   Get TTHandle, 44497, TaxTrans
   TaxTrans.Revenue.InterestPd = 13.4
   TaxTrans.Amount = 72.3
   Put TTHandle, 44497, TaxTrans
   
   Get TTHandle, 26661, TaxTrans
   TaxTrans.Revenue.InterestPd = 13.4
   Put TTHandle, 26661, TaxTrans
   
  'fix for #1137
  Get TTHandle, 4898, TaxTrans
  TaxTrans.Amount = 0.18
  TaxTrans.Revenue.Principle3Pd = 0.18
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.TranType = 2
  Put TTHandle, 4898, TaxTrans

  'fix for 4093
  Get TTHandle, 30937, TaxTrans
  TaxTrans.Revenue.Interest = 17.5
  Put TTHandle, 30937, TaxTrans

  Get TTHandle, 26269, TaxTrans
  TaxTrans.Revenue.Interest = 31.5
  Put TTHandle, 26269, TaxTrans

  Get TTHandle, 39585, TaxTrans
  TaxTrans.Amount = 3.5
  TaxTrans.Revenue.Interest = 3.5
  Put TTHandle, 39585, TaxTrans

  Get TTHandle, 39383, TaxTrans
  TaxTrans.Amount = 3.5
  TaxTrans.Revenue.Interest = 3.5
  Put TTHandle, 39383, TaxTrans

  Get TTHandle, 38954, TaxTrans
  TaxTrans.Amount = 3.5
  TaxTrans.Revenue.Interest = 3.5
  Put TTHandle, 38954, TaxTrans

   Get TTHandle, 38752, TaxTrans
  TaxTrans.Amount = 3.5
  TaxTrans.Revenue.Interest = 3.5
  Put TTHandle, 38752, TaxTrans

  Get TTHandle, 45071, TaxTrans
  TaxTrans.Amount = 28
  TaxTrans.Revenue.Principle1Pd = 28
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.TranType = 2
  Put TTHandle, 45071, TaxTrans

  Call ClearTrans(44439)

  'fix for 3621
'  Get TTHandle, 8792, TaxTrans
'  TaxTrans.Amount = 69.83
'  TaxTrans.Revenue.Principle1 = 69.83
'  TaxTrans.Revenue.Principle1Pd = 69.83
'  Put TTHandle, 8792, TaxTrans

  Call ClearTrans(17954)
  Call ClearTrans(15143)
  Call ClearTrans(10871)

  Get TTHandle, 4079, TaxTrans
  TaxTrans.Amount = 0.01
  TaxTrans.Revenue.Principle1Pd = 0.01
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.TranType = 2
  Put TTHandle, 4079, TaxTrans

  Get TTHandle, 7162, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  Put TTHandle, 7162, TaxTrans

  'fix for 502
  Get TTHandle, 12060, TaxTrans
  TaxTrans.TranType = 2
  TaxTrans.Amount = 5.34
  TaxTrans.Revenue.Principle3Pd = 5.34
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 12060, TaxTrans

  'fixes 7/15/2010
   'fix for 5277
  Get TTHandle, 45438, TaxTrans
  TaxTrans.TranType = 2
  TaxTrans.Amount = 2.68
  TaxTrans.Revenue.Principle1Pd = 2.68
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 45438, TaxTrans

  'fix for 4421
  Call InsertCreditAtBillingTrans("5/30/2007", 17.91, 4421, 2006, 17.91, 8531, "Personal", 8635, 8531, "4421")
  Get TTHandle, 8531, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 17.91
  TaxTrans.Revenue.InterestPd = 0
  Put TTHandle, 8531, TaxTrans
  
  'fix for 4093
  Get TTHandle, 45071, TaxTrans
  TaxTrans.TranType = 2
  TaxTrans.Amount = 28
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 45071, TaxTrans

  'fix for 3810
  Call InsertCreditAtBillingTrans("5/30/2007", 45.37, 3810, 2007, 45.37, 8593, "Personal", 8976, 8593, "5302007")

  Get TTHandle, 8593, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 45.37
  Put TTHandle, 8593, TaxTrans

  Call ClearTrans(21385)
  Call ClearTrans(20753)
  Get TTHandle, 18359, TaxTrans
  TaxTrans.Revenue.Principle1 = 45.37
  Put TTHandle, 18359, TaxTrans

  'fix for 3621
  Call ClearTrans(4079)
  Call ClearTrans(33522)
  Get TTHandle, 7162, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
   Put TTHandle, 7162, TaxTrans

  Get TTHandle, 6098, TaxTrans
  TaxTrans.Amount = 69.85
  TaxTrans.Revenue.Principle1 = 69.85
  TaxTrans.Revenue.Principle1Pd = 69.85
  Put TTHandle, 6098, TaxTrans

  Get TTHandle, 17954, TaxTrans
  TaxTrans.TranType = 2
  TaxTrans.Amount = 69.85
  TaxTrans.Revenue.Principle1Pd = 69.85
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 17954, TaxTrans

  Get TTHandle, 8792, TaxTrans
  TaxTrans.Revenue.Principle1 = 143.19
  TaxTrans.Revenue.Principle1Pd = 143.17
  TaxTrans.Revenue.Interest = 3.49
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Amount = 143.19
  Put TTHandle, 8792, TaxTrans
  
  Get TTHandle, 10872, TaxTrans
  TaxTrans.TranType = 9
  TaxTrans.Amount = 0.01
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0.01
  Put TTHandle, 10872, TaxTrans
  
  'fix for 1735
  Get TTHandle, 7972, TaxTrans
  TaxTrans.Revenue.Principle2 = 0
  Put TTHandle, 7972, TaxTrans

  Get TTHandle, 510, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 184.87
  Put TTHandle, 510, TaxTrans

  Get TTHandle, 3886, TaxTrans
  TaxTrans.TranType = 2
  TaxTrans.Amount = 184.87
  TaxTrans.Revenue.Principle1Pd = 184.87
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 3886, TaxTrans

    'fix for 933
  Call InsertCreditAtBillingTrans("11/03/2009", 9.82, 933, 2009, 13.3, 31644, "Personal", 32946, 31644, "187")
  Get TTHandle, 31644, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 34.04
  TaxTrans.Revenue.InterestPd = 2.32
  TaxTrans.Revenue.PenaltyPd = 1.16
  Put TTHandle, 31644, TaxTrans

  Get TTHandle, 5920, TaxTrans
  TaxTrans.BillType = "Personal"
  Put TTHandle, 5920, TaxTrans

  Get TTHandle, 3347, TaxTrans
  TaxTrans.Amount = 12.49
  TaxTrans.Revenue.Principle1Pd = 12.49
  Put TTHandle, 3347, TaxTrans

  Get TTHandle, 142, TaxTrans
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.PenaltyPd = 0
  Put TTHandle, 142, TaxTrans

  Get TTHandle, 18616, TaxTrans
  TaxTrans.TranType = 2
  TaxTrans.Amount = 24.81
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 18616, TaxTrans

  Get TTHandle, 12202, TaxTrans
  TaxTrans.TranType = 2
  TaxTrans.Amount = 44.8
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 12202, TaxTrans

  Get TTHandle, 9812, TaxTrans
  TaxTrans.Revenue.Principle1 = 0
  TaxTrans.PPTRADisc = 0
  Put TTHandle, 9812, TaxTrans

  Get TTHandle, 11377, TaxTrans
  TaxTrans.Amount = 44.8
  Put TTHandle, 11377, TaxTrans

  'fix for 502
'  Get TTHandle, 12059, TaxTrans
'  TaxTrans.Revenue.Principle3Pd = 63.22
'  Put TTHandle, 12059, TaxTrans
'
'  Get TTHandle, 12060, TaxTrans
'  TaxTrans.Revenue.PrePaidBal = 0
'  Put TTHandle, 12060, TaxTrans
'
'  Call ClearTrans(12060)
'  Get TTHandle, 21537, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle3Pd = 44.74
'  TaxTrans.Revenue.Interest = 0
'  TaxTrans.Revenue.Penalty = 0
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.PenaltyPd = 0
'  Put TTHandle, 21537, TaxTrans
'
'  Get TTHandle, 31474, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Interest = 16.1
'  TaxTrans.Revenue.Penalty = 8.05
'  TaxTrans.Revenue.Principle3Pd = 23.82
'  TaxTrans.Revenue.InterestPd = 2.38
'  TaxTrans.Revenue.PenaltyPd = 1.19
'  Put TTHandle, 31474, TaxTrans
'
'  Get TTHandle, 11863, TaxTrans
'  TaxTrans.Amount = 52.98
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.PenaltyPd = 0
'  Put TTHandle, 11863, TaxTrans
'
'  Get TTHandle, 74, TaxTrans
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.PenaltyPd = 0
'  Put TTHandle, 74, TaxTrans

  'fix for 712
  Call InsertCreditAtBillingTrans("5/30/2007", 19.16, 712, 2007, 7, 8511, "Personal", 9772, 8511, "5302007")
  Get TTHandle, 8511, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 19.16
  TaxTrans.Revenue.InterestPd = 0
  Put TTHandle, 8511, TaxTrans

  'fix for #280
  Get TTHandle, 24, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 34.24
  Put TTHandle, 24, TaxTrans

   Get TTHandle, 3517, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 34.24
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidAmt = 9.55
  TaxTrans.Revenue.PrePaidBal = 9.55
  Put TTHandle, 3517, TaxTrans

  'fix for #15
  Get TTHandle, 7945, TaxTrans
  TaxTrans.Amount = 7.55
  TaxTrans.Revenue.Principle1Pd = 7.55
  TaxTrans.TranType = 2
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 7945, TaxTrans

  'fix for 3100
  Call ClearTrans(44711)
  Call ClearTrans(40748)
  Call ClearTrans(16746)
  Call ClearTrans(25575)
  Call ClearTrans(6294)
  Call InsertCreditAtBillingTrans("12/27/2007", 2.02, 3100, 2007, 2.02, 16405, "Real", 16746, 16405, "5")

  Get TTHandle, 6277, TaxTrans
  TaxTrans.Amount = 1.3
  TaxTrans.Revenue.Principle1Pd = 1.3
  TaxTrans.Description = "6"
  TaxTrans.BelongTo = 5826
  Put TTHandle, 6277, TaxTrans

  Get TTHandle, 16405, TaxTrans
  TaxTrans.Revenue.Principle1 = 2.02
  TaxTrans.Revenue.Principle1Pd = 2.02
  TaxTrans.Revenue.Interest = 0
  Put TTHandle, 16405, TaxTrans

  'fix for 2712
  Get TTHandle, 3972, TaxTrans
  TaxTrans.Amount = 22.88
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.TranType = 2
  TaxTrans.Description = "816"
  TaxTrans.Revenue.Principle1Pd = 22.88
  TaxTrans.BillType = "Personal"
  Put TTHandle, 3972, TaxTrans

  Get TTHandle, 816, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 22.88
  Put TTHandle, 816, TaxTrans

    'fix for 2539
  Get TTHandle, 25322, TaxTrans
  TaxTrans.Amount = 56.32
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.TranType = 2
  TaxTrans.Description = "808"
  TaxTrans.Revenue.Principle1Pd = 56.32
  Put TTHandle, 25322, TaxTrans

  Get TTHandle, 808, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 56.32
  Put TTHandle, 808, TaxTrans

   'fix for 2251
  ClearTrans (8595)
  Call InsertPayTrans("03/19/2007", 13.63, 2251, 2007, 13.63, 780, "Personal", 0, 8595, "780", 0, 0, SaveRec)
  Call InsertPayTrans("03/19/2007", 13.63, 2251, 2007, 13.63, 781, "Personal", 0, SaveRec, "781", 0, 0, SaveRec)
  Get TCHandle, 2251, TaxCust
  TaxCust.LastTrans = SaveRec
  Put TCHandle, 2251, TaxCust


'  'fix 3552 6/3/2010
'  Get TTHandle, 4033, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 53.52
'  Put TTHandle, 4033, TaxTrans
'  Call ClearTrans(7442)
'  Call ClearTrans(14588)
'  Get TTHandle, 39227, TaxTrans
'  TaxTrans.Amount = 0.25
'  TaxTrans.Revenue.Penalty = 0.25
'  Put TTHandle, 39227, TaxTrans
'  Get TTHandle, 31114, TaxTrans
'  TaxTrans.Revenue.Interest = 72.66
'  Put TTHandle, 31114, TaxTrans
'
'  Get TTHandle, 26439, TaxTrans
'  TaxTrans.Revenue.Interest = 72.66
'  Put TTHandle, 26439, TaxTrans
'
  'fix 5/20/2010
'  Call ClearTrans(6548)
'  frmVATaxShowPctComp.Label1 = "Processing"
'  frmVATaxShowPctComp.Show , Me


'  TransDate = Date2Num("02/01/2010")
'  For x = 1 To NumOfTRecs
'    Get TTHandle, x, TaxTrans
'    If TaxTrans.TransDate = TransDate And TaxTrans.TranType = 4 Then
'      For y = 1 To ArrCnt
'        If Bills(y) = TaxTrans.BelongTo Then
'          Exit For
'        End If
'      Next y
'      If y > ArrCnt Then
'        ArrCnt = ArrCnt + 1
'        ReDim Preserve Bills(1 To ArrCnt) As Long
'        Bills(ArrCnt) = TaxTrans.BelongTo
'      End If
'    End If
'  Next x
'
'  For x = 1 To ArrCnt
'    IntTot = 0
'    Get TTHandle, Bills(x), TaxTrans
'    For y = 1 To NumOfTRecs
'      Get TTHandle, y, TaxTrans
'        If TaxTrans.TranType = 4 And TaxTrans.BelongTo = Bills(x) Then
'          IntTot = IntTot + TaxTrans.Revenue.Interest
'        End If
'    Next y
'    Get TTHandle, Bills(x), TaxTrans
'    TaxTrans.Revenue.Interest = IntTot
'    Put TTHandle, Bills(x), TaxTrans
'    frmVATaxShowPctComp.ShowPctComp x, ArrCnt
'    If frmVATaxShowPctComp.Out = True Then
'      Close
'      frmVATaxShowPctComp.Out = False
'      Unload frmVATaxShowPctComp
'      Exit Sub
'    End If
'  Next x
'
'  '#112 to corect adj down bill
'  Get TTHandle, 44432, TaxTrans
'  TaxTrans.Revenue.Interest = 17.85
'  TaxTrans.Amount = 24.99
'  Put TTHandle, 44432, TaxTrans
'
'  Get TTHandle, 30536, TaxTrans
'  TaxTrans.Revenue.Interest = 1.43
'  Put TTHandle, 30536, TaxTrans
'
'  Get TTHandle, 44435, TaxTrans
'  TaxTrans.Revenue.Interest = 17.85
'  TaxTrans.Amount = 24.99
'  Put TTHandle, 44435, TaxTrans
'
'  Get TTHandle, 25865, TaxTrans
'  TaxTrans.Revenue.Interest = 0
'  Put TTHandle, 25865, TaxTrans
'
'  Get TTHandle, 8715, TaxTrans
'  TaxTrans.Amount = 80.87
'  TaxTrans.Revenue.Principle1 = 80.87
'  TaxTrans.Revenue.Principle1Pd = 80.87
'  Put TTHandle, 8715, TaxTrans
'
'  Get TTHandle, 17713, TaxTrans
'  TaxTrans.Amount = 80.87
'  TaxTrans.Revenue.Principle1 = 80.87
'  TaxTrans.Revenue.Principle1Pd = 80.87
'  Put TTHandle, 17713, TaxTrans
'
'  Get TTHandle, 24934, TaxTrans
'  TaxTrans.BelongTo = 17713
'  TaxTrans.Description = "Bill #208"
'  Put TTHandle, 24934, TaxTrans
'
'  SaveRec = 0
'
'  Call InsertPayTrans("5/13/2010", 52.79, 3543, 2009, 80.87, 30996, "Real", 0, 45059, "438", 21.06, 7.02, SaveRec)
'  OpenTaxCustFile TCHandle, NumOfTCRecs
'  Get TCHandle, 3543, TaxCust
'  TaxCust.LastTrans = SaveRec
'  Put TCHandle, 3543, TaxCust
'  Close TCHandle
'
'
'  Call ClearTrans(20444)
'  Call ClearTrans(21329)
'  Call ClearTrans(17714)
'
'  Unload frmVATaxShowPctComp
'  Close
'  MsgBox ("All done.")


'  Dim TaxTrans As TaxTransactionType
'  Dim THandle As Integer
'  Dim NumOfTRecs As Long
'  Dim TransDate As Integer
'  Dim x As Long
'  Dim y As Long
'  Dim IntTot As Double
'  Dim ArrCnt As Integer
'  Dim Bills() As Long
'  Dim SaveRec As Long
'  Dim TaxCust As TaxCustType
'  Dim TCHandle As Integer
'  Dim NumOfTCRecs As Long
'
'  OpenTaxTransFile THandle, NumOfTRecs
'  'fix for #1137
'  Get THandle, 4898, TaxTrans
'  TaxTrans.Amount = 0.18
'  TaxTrans.Revenue.Principle3Pd = 0.18
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.TranType = 2
'  Put THandle, 4898, TaxTrans
'
'  'fix for 4093
'  Get THandle, 30937, TaxTrans
'  TaxTrans.Revenue.Interest = 17.5
'  Put THandle, 30937, TaxTrans
'
'  Get THandle, 26269, TaxTrans
'  TaxTrans.Revenue.Interest = 31.5
'  Put THandle, 26269, TaxTrans
'
'  Get THandle, 39585, TaxTrans
'  TaxTrans.Amount = 3.5
'  TaxTrans.Revenue.Interest = 3.5
'  Put THandle, 39585, TaxTrans
'
'  Get THandle, 39383, TaxTrans
'  TaxTrans.Amount = 3.5
'  TaxTrans.Revenue.Interest = 3.5
'  Put THandle, 39383, TaxTrans
'
'  Get THandle, 38954, TaxTrans
'  TaxTrans.Amount = 3.5
'  TaxTrans.Revenue.Interest = 3.5
'  Put THandle, 38954, TaxTrans
'
'   Get THandle, 38752, TaxTrans
'  TaxTrans.Amount = 3.5
'  TaxTrans.Revenue.Interest = 3.5
'  Put THandle, 38752, TaxTrans
'
'  Get THandle, 45071, TaxTrans
'  TaxTrans.Amount = 28
'  TaxTrans.Revenue.Principle1Pd = 28
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.TranType = 2
'  Put THandle, 45071, TaxTrans
'
'  Call ClearTrans(44439)
'
'  'fix for 3621
'  Get THandle, 8792, TaxTrans
'  TaxTrans.Amount = 69.83
'  TaxTrans.Revenue.Principle1 = 69.83
'  TaxTrans.Revenue.Principle1Pd = 69.83
'  Put THandle, 8792, TaxTrans
'
'  Call ClearTrans(17954)
'  Call ClearTrans(15143)
'  Call ClearTrans(10871)
'
'  Get THandle, 4079, TaxTrans
'  TaxTrans.Amount = 0.01
'  TaxTrans.Revenue.Principle1Pd = 0.01
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  TaxTrans.TranType = 2
'  Put THandle, 4079, TaxTrans
'
'  Get THandle, 7162, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0
'  Put THandle, 7162, TaxTrans
'
'  'fix for 502
''  Get THandle, 74, TaxTrans
''  TaxTrans.Revenue.Interest = 0
''  TaxTrans.Revenue.Penalty = 0
''  TaxTrans.Revenue.InterestPd = 0
''  TaxTrans.Revenue.PenaltyPd = 0
''  Put THandle, 74, TaxTrans
''
''  Get THandle, 11863, TaxTrans
''  TaxTrans.Amount = 52.98
''  TaxTrans.Revenue.InterestPd = 0
''  TaxTrans.Revenue.PenaltyPd = 0
''  Put THandle, 11863, TaxTrans
''
''  Get THandle, 43294, TaxTrans
''  TaxTrans.Revenue.PenaltyPd = 0
''  TaxTrans.Amount = 26.2
''  Put THandle, 43294, TaxTrans
'
'  Get THandle, 12060, TaxTrans
'  TaxTrans.TranType = 2
'  TaxTrans.Amount = 5.34
'  TaxTrans.Revenue.Principle3Pd = 5.34
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put THandle, 12060, TaxTrans
'
'  'fix 3552 6/3/2010
'  Get THandle, 4033, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 53.52
'  Put THandle, 4033, TaxTrans
'  Call ClearTrans(7442)
'  Call ClearTrans(14588)
'  Get THandle, 39227, TaxTrans
'  TaxTrans.Amount = 0.25
'  TaxTrans.Revenue.Penalty = 0.25
'  Put THandle, 39227, TaxTrans
''  Get THandle, 31114, TaxTrans
''  TaxTrans.Revenue.Interest = 72.66
''  Put THandle, 31114, TaxTrans
''
''  Get THandle, 26439, TaxTrans
''  TaxTrans.Revenue.Interest = 72.66
''  Put THandle, 26439, TaxTrans
''
'  'fix 5/20/2010
''  Call ClearTrans(6548)
''  frmVATaxShowPctComp.Label1 = "Processing"
''  frmVATaxShowPctComp.Show , Me
'
'
''  TransDate = Date2Num("02/01/2010")
''  For x = 1 To NumOfTRecs
''    Get THandle, x, TaxTrans
''    If TaxTrans.TransDate = TransDate And TaxTrans.TranType = 4 Then
''      For y = 1 To ArrCnt
''        If Bills(y) = TaxTrans.BelongTo Then
''          Exit For
''        End If
''      Next y
''      If y > ArrCnt Then
''        ArrCnt = ArrCnt + 1
''        ReDim Preserve Bills(1 To ArrCnt) As Long
''        Bills(ArrCnt) = TaxTrans.BelongTo
''      End If
''    End If
''  Next x
''
''  For x = 1 To ArrCnt
''    IntTot = 0
''    Get THandle, Bills(x), TaxTrans
''    For y = 1 To NumOfTRecs
''      Get THandle, y, TaxTrans
''        If TaxTrans.TranType = 4 And TaxTrans.BelongTo = Bills(x) Then
''          IntTot = IntTot + TaxTrans.Revenue.Interest
''        End If
''    Next y
''    Get THandle, Bills(x), TaxTrans
''    TaxTrans.Revenue.Interest = IntTot
''    Put THandle, Bills(x), TaxTrans
''    frmVATaxShowPctComp.ShowPctComp x, ArrCnt
''    If frmVATaxShowPctComp.Out = True Then
''      Close
''      frmVATaxShowPctComp.Out = False
''      Unload frmVATaxShowPctComp
''      Exit Sub
''    End If
''  Next x
''
''  '#112 to corect adj down bill
''  Get THandle, 44432, TaxTrans
''  TaxTrans.Revenue.Interest = 17.85
''  TaxTrans.Amount = 24.99
''  Put THandle, 44432, TaxTrans
''
''  Get THandle, 30536, TaxTrans
''  TaxTrans.Revenue.Interest = 1.43
''  Put THandle, 30536, TaxTrans
''
''  Get THandle, 44435, TaxTrans
''  TaxTrans.Revenue.Interest = 17.85
''  TaxTrans.Amount = 24.99
''  Put THandle, 44435, TaxTrans
''
''  Get THandle, 25865, TaxTrans
''  TaxTrans.Revenue.Interest = 0
''  Put THandle, 25865, TaxTrans
''
''  Get THandle, 8715, TaxTrans
''  TaxTrans.Amount = 80.87
''  TaxTrans.Revenue.Principle1 = 80.87
''  TaxTrans.Revenue.Principle1Pd = 80.87
''  Put THandle, 8715, TaxTrans
''
''  Get THandle, 17713, TaxTrans
''  TaxTrans.Amount = 80.87
''  TaxTrans.Revenue.Principle1 = 80.87
''  TaxTrans.Revenue.Principle1Pd = 80.87
''  Put THandle, 17713, TaxTrans
''
''  Get THandle, 24934, TaxTrans
''  TaxTrans.BelongTo = 17713
''  TaxTrans.Description = "Bill #208"
''  Put THandle, 24934, TaxTrans
''
''  SaveRec = 0
''
''  Call InsertPayTrans("5/13/2010", 52.79, 3543, 2009, 80.87, 30996, "Real", 0, 45059, "438", 21.06, 7.02, SaveRec)
''  OpenTaxCustFile TCHandle, NumOfTCRecs
''  Get TCHandle, 3543, TaxCust
''  TaxCust.LastTrans = SaveRec
''  Put TCHandle, 3543, TaxCust
''  Close TCHandle
''
''
''  Call ClearTrans(20444)
''  Call ClearTrans(21329)
''  Call ClearTrans(17714)
''
''  Unload frmVATaxShowPctComp
  Close
  MsgBox ("All done.")


End Sub

Private Sub cmdMBToBillBal_Click()
' MakeMasterBalEqualBillBal
Call FixErrorInOPAtBilling
End Sub

Private Sub cmdFixUnusedPrepay_Click()
  Call FixUnusedPrepayIn2006
End Sub

Private Sub cmdFixZeroedIntAndPen_Click()
  Call FindAndFixZeroInterestCausedIssues
End Sub

Private Sub cmdGetBelongTos_Click()
  Call FindBelongTo
End Sub

Private Sub cmdListZeroCRecs_Click()
  Call FixCustomerRecsandPins
End Sub

Private Sub FixZeroedOutCreditAtBilling()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim NextRec As Long
  Dim CustRec As Integer
  Dim LookRec As Long
  Dim CheckRec As Long
  Dim AHandle As Integer
  Dim cnt As Integer
  Dim BelongTo As Long
  Dim x As Integer
  Dim OK As Boolean
  Dim Amount As Double
  Dim CreditTrans As Long
  
  AHandle = FreeFile
  Open "test.txt" For Output As AHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  frmVATaxShowPctComp.Label1 = "Fixing Zeroed Out Credit at Billing"
  frmVATaxShowPctComp.Show , Me
  Unload frmVATaxShowPctComp
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo Skip
'    If x = 2148 Then Stop
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
        If TaxTrans.TranType = 9 Then
          CreditTrans = NextRec
          If TaxTrans.Revenue.PrePaidUsed = 0 Then
            BelongTo = TaxTrans.BelongTo
            Get TTHandle, BelongTo, TaxTrans
            Amount = TaxTrans.Amount
            If Amount <> 0 Then
              CheckRec = TaxTrans.LastTrans
              Do While CheckRec > 0
                Get TTHandle, CheckRec, TaxTrans
                If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
                  If TaxTrans.Revenue.PrePaidAmt = Amount Then
                    Get TTHandle, BelongTo, TaxTrans
                      If TaxTrans.Revenue.Principle1 = Amount Then
                        Get TTHandle, CreditTrans, TaxTrans
                          TaxTrans.Revenue.PrePaidUsed = Amount
                          TaxTrans.Revenue.Principle1Pd = Amount
                        Put TTHandle, CreditTrans, TaxTrans
                        Print #AHandle, CStr(x) + "~" + CStr(NextRec)
                        Exit Do
                      ElseIf TaxTrans.Revenue.Principle2 = Amount Then
                        Get TTHandle, CreditTrans, TaxTrans
                          TaxTrans.Revenue.PrePaidUsed = Amount
                          TaxTrans.Revenue.Principle2Pd = Amount
                        Put TTHandle, CreditTrans, TaxTrans
                        Print #AHandle, CStr(x) + "~" + CStr(NextRec)
                        Exit Do
                      ElseIf TaxTrans.Revenue.Principle3 = Amount Then
                        Get TTHandle, CreditTrans, TaxTrans
                          TaxTrans.Revenue.PrePaidUsed = Amount
                          TaxTrans.Revenue.Principle3Pd = Amount
                        Put TTHandle, CreditTrans, TaxTrans
                        Print #AHandle, CStr(x) + "~" + CStr(NextRec)
                        Exit Do
                      ElseIf TaxTrans.Revenue.Principle4 = Amount Then
                        Get TTHandle, CreditTrans, TaxTrans
                          TaxTrans.Revenue.PrePaidUsed = Amount
                          TaxTrans.Revenue.Principle4Pd = Amount
                        Put TTHandle, CreditTrans, TaxTrans
                        Print #AHandle, CStr(x) + "~" + CStr(NextRec)
                        Exit Do
                      ElseIf TaxTrans.Revenue.Principle5 = Amount Then
                        Get TTHandle, CreditTrans, TaxTrans
                          TaxTrans.Revenue.PrePaidUsed = Amount
                          TaxTrans.Revenue.Principle5Pd = Amount
                        Put TTHandle, CreditTrans, TaxTrans
                        Print #AHandle, CStr(x) + "~" + CStr(NextRec)
                        Exit Do
                      End If
                    
                  End If
                End If
                CheckRec = TaxTrans.LastTrans
              Loop
            End If
          End If
        End If
      Get TTHandle, NextRec, TaxTrans
      NextRec = TaxTrans.LastTrans
    Loop
    
Skip:
  Next x
  
'  For x = 1 To NumOfTCRecs
'    Get TCHandle, x, TaxCust
'    If TaxCust.Deleted <> 0 Then GoTo Skip1
'    NextRec = TaxCust.LastTrans
'    Do While NextRec > 0
'      Get TTHandle, NextRec, TaxTrans
'
'        If TaxTrans.TranType = 9 And TaxTrans.BillType = "R" Then
'          OK = False
'          LookRec = TaxCust.LastTrans
'          Do While LookRec > 0
'            Get TTHandle, LookRec, TaxTrans
'              If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
'                OK = True
'                Exit Do
'              End If
'            LookRec = TaxTrans.LastTrans
'          Loop
'          If OK = False Then
'            Print #AHandle, "R ~ " + CStr(x)
'          End If
'        End If
'      Get TTHandle, NextRec, TaxTrans
'      NextRec = TaxTrans.LastTrans
'    Loop
'
'Skip1:
'  Next x
  Unload frmVATaxShowPctComp
 
  
  Close
  MsgBox ("Done.")
End Sub

Private Sub FindTransInCustQueue()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim NextRec As Long
  Dim CustRec As Integer
  Dim LookRec As Long
  Dim AHandle As Integer
  Dim cnt As Integer
  Dim BelongTo As Long
  Dim x As Integer
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  Get TCHandle, 1781, TaxCust
  NextRec = TaxCust.LastTrans
  Do While NextRec > 0
    Get TTHandle, NextRec, TaxTrans
    BelongTo = TaxTrans.BelongTo
    If BelongTo = 0 Then GoTo Skip
    Get TTHandle, BelongTo, TaxTrans
    CustRec = TaxTrans.CustomerRec
    Debug.Print CStr(CustRec)
    If CustRec = 0 Then CustRec = TaxTrans.CustPin
    If CustRec <> 1781 Then
      Get TCHandle, CustRec, TaxCust
      If TaxCust.Deleted <> 0 Then
        Print #AHandle, CStr(CustRec)
        Call ClearTrans(NextRec)
        cnt = cnt + 1
      End If
    End If
Skip:
    Get TTHandle, NextRec, TaxTrans
    NextRec = TaxTrans.LastTrans
  Loop
  
  Close
  MsgBox ("Count = " + CStr(cnt))
  
End Sub

Private Sub cmdMisc_Click()
'  Call ClearDeletedCustTrans
  Call FindOrphanTrans
  Exit Sub
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  Dim Tot As Double
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If TaxTrans.BelongTo = 32948 And TaxTrans.TranType = 4 Then
      Tot = Tot + TaxTrans.Revenue.Interest
    End If
  Next x
  
  Close
  MsgBox (CStr(Tot))
End Sub
Private Sub FixBillsWithAdjustsMoreThanTotals()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Integer
  Dim NextRec As Long
  Dim LastRec As Long
  Dim BelongTo As Long
  Dim BillTot As Double
  Dim IntTot As Double
  Dim PenTot As Double
  Dim AdvTot As Double
  Dim AdjTot As Double
  Dim AHandle As Integer
  Dim cnt As Integer
  'this sub looks for adjust bills down transactions that adjust more than the bill amounts
  'but only includes interest, penalty or advertising transactions else it skips over the bill
  Call BuildMBvsCustHistArr
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  frmVATaxShowPctComp.Label1 = "Fixing Zeroed Out Credit at Billing"
  frmVATaxShowPctComp.Show , Me

  AHandle = FreeFile
  Open "AdjBillDownUpdates.txt" For Output As AHandle
  For x = 1 To CArrCnt
    Get TCHandle, CArr(x), TaxCust
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
        If TaxTrans.TranType = 13 Then
          AdjTot = TaxTrans.Amount
          BelongTo = TaxTrans.BelongTo
          Get TTHandle, BelongTo, TaxTrans
          BillTot = TaxTrans.Amount
          LastRec = NextRec
          Get TTHandle, LastRec, TaxTrans
          IntTot = 0
          PenTot = 0
          AdvTot = 0
          Do While LastRec >= BelongTo
             Get TTHandle, LastRec, TaxTrans
               If TaxTrans.TranType = 4 Or TaxTrans.TranType = 5 Or TaxTrans.TranType = 6 Then
                 Select Case TaxTrans.TranType
                   Case 4
                     IntTot = IntTot + TaxTrans.Amount
                   Case 5
                     PenTot = PenTot + TaxTrans.Amount
                   Case 6
                     AdvTot = AdvTot + TaxTrans.Amount
                 End Select
               ElseIf TaxTrans.TranType <> 13 And TaxTrans.TranType <> 1 Then
                 GoTo SkipIt
               End If
             LastRec = TaxTrans.LastTrans
          Loop
          If BillTot + IntTot + PenTot + AdvTot < AdjTot Then
            Get TTHandle, NextRec, TaxTrans
              If TaxTrans.Revenue.Principle1 > 0 Then
                TaxTrans.Revenue.Principle1 = BillTot
              ElseIf TaxTrans.Revenue.Principle2 > 0 Then
                TaxTrans.Revenue.Principle2 = BillTot
              ElseIf TaxTrans.Revenue.Principle3 > 0 Then
                TaxTrans.Revenue.Principle3 = BillTot
              ElseIf TaxTrans.Revenue.Principle4 > 0 Then
                TaxTrans.Revenue.Principle4 = BillTot
              ElseIf TaxTrans.Revenue.Principle5 > 0 Then
                TaxTrans.Revenue.Principle5 = BillTot
              End If
              TaxTrans.Amount = BillTot + IntTot + PenTot + AdvTot
              TaxTrans.Revenue.Interest = IntTot
              TaxTrans.Revenue.Collection = AdvTot
              TaxTrans.Revenue.Penalty = PenTot
              Put TTHandle, NextRec, TaxTrans
              cnt = cnt + 1
              Print #AHandle, CStr(CArr(x)) + "~" + CStr(NextRec)
          End If
        End If
SkipIt:
      Get TTHandle, NextRec, TaxTrans
      NextRec = TaxTrans.LastTrans
    Loop
  Next x
 
  Unload frmVATaxShowPctComp
 
  Close
  MsgBox ("There were " + CStr(cnt) + " adjust bill down transactions modified. Look for AdjBillDownUpdates.txt in the Citipak folder for results.")
End Sub

Private Sub FindAndFixZeroInterestCausedIssues()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Integer
  Dim NextRec As Long
  Dim ICnt As Integer
  Dim Found As Boolean
  Dim CBal As Double
  Dim IBal As Double
  Dim IRec As Long
  Dim PnRec As Long
  Dim PrBal1 As Double
  Dim PrBal2 As Double
  Dim PrBal3 As Double
  Dim PrBal4 As Double
  Dim PrBal5 As Double
  Dim OBal1 As Double
  Dim OBal2 As Double
  Dim OBal3 As Double
  Dim PnBal As Double
  Dim ABal As Double
  Dim LLBal As Double
  Dim MBBal As Double
  Dim ThisRec As Long
  Dim Dif As Double
  Dim FCnt As Integer
  Dim BelongTo As Long
  'this sub does NOT review all types of transactions
  'it looks for zeroed out penalty and interest trans and checks to see if the
  'bills affected have accurate totals
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  frmVATaxShowPctComp.Label1 = "Fixing Zeroed Out Interest Errors"
 
  Unload frmVATaxShowPctComp
 Call BuildMBvsCustHistArr
  FCnt = 0
'  CArrCnt = 1
'  CArr(1) = 5560
  For x = 1 To CArrCnt
    Get TCHandle, CArr(x), TaxCust
'    Found = False
'    ICnt = 0
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
        If TaxTrans.Revenue.Interest = 0 Then
          ICnt = ICnt + 1
        End If
      NextRec = TaxTrans.LastTrans
    Loop
    If ICnt > 0 Then
      IBal = 0
      PnBal = 0
      PrBal1 = 0
      PrBal2 = 0
      PrBal3 = 0
      PrBal4 = 0
      PrBal5 = 0
      OBal1 = 0
      OBal2 = 0
      OBal3 = 0
      ABal = 0
      LLBal = 0
      IRec = 0
      PnRec = 0
      NextRec = TaxCust.LastTrans
      Do While NextRec > 0
        Get TTHandle, NextRec, TaxTrans
          If TaxTrans.TranType = 1 Then
            IBal = 0
            PnBal = 0
            PrBal1 = 0
            PrBal2 = 0
            PrBal3 = 0
            PrBal4 = 0
            PrBal5 = 0
            OBal1 = 0
            OBal2 = 0
            OBal3 = 0
            ABal = 0
            LLBal = 0
            IRec = 0
              ThisRec = TaxCust.LastTrans
              BelongTo = NextRec
              Do While ThisRec > 0
                Get TTHandle, ThisRec, TaxTrans
                If TaxTrans.BelongTo = BelongTo Then
                  If TaxTrans.TranType = 13 Then 'adjust bill down
                    IBal = IBal - TaxTrans.Revenue.Interest
                    PnBal = PnBal - TaxTrans.Revenue.Penalty
                    PrBal1 = PrBal1 - TaxTrans.Revenue.Principle1
                    PrBal2 = PrBal2 - TaxTrans.Revenue.Principle2
                    PrBal3 = PrBal3 - TaxTrans.Revenue.Principle3
                    PrBal4 = PrBal4 - TaxTrans.Revenue.Principle4
                    PrBal5 = PrBal5 - TaxTrans.Revenue.Principle5
                    OBal1 = OBal1 - TaxTrans.Revenue.RevOpt1
                    OBal2 = OBal2 - TaxTrans.Revenue.RevOpt2
                    OBal3 = OBal3 - TaxTrans.Revenue.RevOpt3
                    ABal = ABal - TaxTrans.Revenue.Collection
                    LLBal = LLBal - TaxTrans.Revenue.LateList
                  ElseIf TaxTrans.TranType = 2 Then
                  ElseIf TaxTrans.TranType = 3 Then 'release
                    IBal = IBal - TaxTrans.Revenue.Interest
                    PnBal = PnBal - TaxTrans.Revenue.Penalty
                    PrBal1 = PrBal1 - TaxTrans.Revenue.Principle1
                    PrBal2 = PrBal2 - TaxTrans.Revenue.Principle2
                    PrBal3 = PrBal3 - TaxTrans.Revenue.Principle3
                    PrBal4 = PrBal4 - TaxTrans.Revenue.Principle4
                    PrBal5 = PrBal5 - TaxTrans.Revenue.Principle5
                    OBal1 = OBal1 - TaxTrans.Revenue.RevOpt1
                    OBal2 = OBal2 - TaxTrans.Revenue.RevOpt2
                    OBal3 = OBal3 - TaxTrans.Revenue.RevOpt3
                    ABal = ABal - TaxTrans.Revenue.Collection
                    LLBal = LLBal - TaxTrans.Revenue.LateList
                  ElseIf TaxTrans.TranType = 14 Then 'adjust bill up
                    IBal = IBal + TaxTrans.Revenue.Interest
                    PnBal = PnBal + TaxTrans.Revenue.Penalty
                    PrBal1 = PrBal1 + TaxTrans.Revenue.Principle1
                    PrBal2 = PrBal2 + TaxTrans.Revenue.Principle2
                    PrBal3 = PrBal3 + TaxTrans.Revenue.Principle3
                    PrBal4 = PrBal4 + TaxTrans.Revenue.Principle4
                    PrBal5 = PrBal5 + TaxTrans.Revenue.Principle5
                    OBal1 = OBal1 + TaxTrans.Revenue.RevOpt1
                    OBal2 = OBal2 + TaxTrans.Revenue.RevOpt2
                    OBal3 = OBal3 + TaxTrans.Revenue.RevOpt3
                    ABal = ABal + TaxTrans.Revenue.Collection
                    LLBal = LLBal + TaxTrans.Revenue.LateList
                  ElseIf TaxTrans.TranType = 4 Or TaxTrans.TranType = 5 Or TaxTrans.TranType = 6 Then
                    If TaxTrans.TranType = 4 And TaxTrans.Revenue.Interest = 0 Then
                      IRec = ThisRec
                    End If
                    If TaxTrans.TranType = 5 And TaxTrans.Revenue.Penalty = 0 Then
                      PnRec = ThisRec
                    End If
                    IBal = IBal + TaxTrans.Revenue.Interest
                    PnBal = PnBal + TaxTrans.Revenue.Penalty
                    PrBal1 = PrBal1 + TaxTrans.Revenue.Principle1
                    PrBal2 = PrBal2 + TaxTrans.Revenue.Principle2
                    PrBal3 = PrBal3 + TaxTrans.Revenue.Principle3
                    PrBal4 = PrBal4 + TaxTrans.Revenue.Principle4
                    PrBal5 = PrBal5 + TaxTrans.Revenue.Principle5
                    OBal1 = OBal1 + TaxTrans.Revenue.RevOpt1
                    OBal2 = OBal2 + TaxTrans.Revenue.RevOpt2
                    OBal3 = OBal3 + TaxTrans.Revenue.RevOpt3
                    ABal = ABal + TaxTrans.Revenue.Collection
                    LLBal = LLBal + TaxTrans.Revenue.LateList
                  Else
                    GoTo Skip
                  End If
                End If
GetNext:
                ThisRec = TaxTrans.LastTrans
              Loop
            Get TTHandle, BelongTo, TaxTrans
            Dif = TaxTrans.Revenue.Interest - IBal
            If Dif <> 0 And IRec <> 0 Then
              Get TTHandle, IRec, TaxTrans
                TaxTrans.Amount = Abs(Dif)
                TaxTrans.Revenue.Interest = Abs(Dif)
              Put TTHandle, IRec, TaxTrans
              FCnt = FCnt + 1
            End If
            Dif = TaxTrans.Revenue.Penalty - PnBal
            If Dif <> 0 And PnRec <> 0 Then
              Get TTHandle, PnRec, TaxTrans
                TaxTrans.Amount = Abs(Dif)
                TaxTrans.Revenue.Penalty = Abs(Dif)
              Put TTHandle, PnRec, TaxTrans
              FCnt = FCnt + 1
            End If
          End If
Skip:
        Get TTHandle, NextRec, TaxTrans
        NextRec = TaxTrans.LastTrans
      Loop
      
    End If
  Next x
  Unload frmVATaxShowPctComp
 
  Close
  MsgBox ("A total of " + CStr(FCnt) + " transactions were fixed.")
  
  Exit Sub
  
GetMBBal:
  ThisRec = TaxCust.LastTrans
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If TaxTrans.TranType = 10 Then 'adjust bill down affecting credit
       MBBal = MBBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'prepay adjust down
      MBBal# = OldRound#(MBBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'adjust bill up affecting credit
      MBBal# = OldRound#(MBBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
      MBBal# = OldRound#(MBBal# - TaxTrans.Revenue.PrePaidAmt)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay
      MBBal# = OldRound#(MBBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 9 Then 'added
      MBBal# = OldRound#(MBBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 1 Then
      MBBal# = OldRound#(MBBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      MBBal# = OldRound#(MBBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      MBBal# = OldRound#(MBBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      MBBal# = OldRound#(MBBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      MBBal# = OldRound#(MBBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
    End If
DoAgain:
    ThisRec = TaxTrans.LastTrans
  Loop

  Return
  

End Sub
Private Sub cmdOptRevInTaxBilling_Click()
  Dim AHandle As Integer
  Dim PTBillRec As VAPPTaxBillType
  Dim PTBHandle As Integer
  Dim NumOfPTBRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim RTBillRec As VARETaxBillType
  Dim RTBHandle As Integer
  Dim NumOfRTBRecs As Long
  Dim x As Long
  
  OpenPersTaxBillFile PTBHandle, NumOfPTBRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  If NumOfPTBRecs = 0 Then
    MsgBox ("A personal prebilling file does not exist.")
    GoTo TryReal
  End If
  AHandle = FreeFile
  Open "billoptrevs.txt" For Output As AHandle
  Print #AHandle, "Property Type" & "~" & "County #" & "~" & "Customer Name" & "~" & "Opt Tax #1" & "~" & "Opt Tax #2" & "~" & "Opt Tax #3"
  For x = 1 To NumOfPTBRecs
    Get PTBHandle, x, PTBillRec
    Get TCHandle, PTBillRec.CustRec, TaxCust
    If QPTrim$(TaxCust.CountyAcctString) <> "" Then
      Print #AHandle, "Personal" & "~" & QPTrim$(TaxCust.CountyAcctString) & "~" & QPTrim$(TaxCust.CustName) & "~" & Using("$##,###.##", PTBillRec.OptRevTax1) & "~" & Using("$##,###.##", PTBillRec.OptRevTax2) & "~" & Using("$##,###.##", PTBillRec.OptRevTax3)
    Else
      Print #AHandle, "Personal" & "~" & CStr(TaxCust.CountyAcct) & "~" & QPTrim$(TaxCust.CustName) & "~" & Using("$##,###.##", PTBillRec.OptRevTax1) & "~" & Using("$##,###.##", PTBillRec.OptRevTax2) & "~" & Using("$##,###.##", PTBillRec.OptRevTax3)
    End If
  Next x

TryReal:
  OpenRealTaxBillFile RTBHandle, NumOfRTBRecs
  If NumOfRTBRecs = 0 Then
    MsgBox ("A real prebilling file does not exist.")
    GoTo Done
  End If

  If NumOfPTBRecs = 0 Then
    AHandle = FreeFile
    Open "billoptrevs.txt" For Output As AHandle
    Print #AHandle, "Property Type" & "~" & "County #" & "~" & "Customer Name" & "~" & "Opt Tax #1" & "~" & "Opt Tax #2" & "~" & "Opt Tax #3"
  End If
  
  For x = 1 To NumOfRTBRecs
    Get RTBHandle, x, RTBillRec
    Get TCHandle, RTBillRec.CustRec, TaxCust
    If QPTrim$(TaxCust.CountyAcctString) <> "" Then
      Print #AHandle, "Real" & "~" & QPTrim$(TaxCust.CountyAcctString) & "~" & QPTrim$(TaxCust.CustName) & "~" & Using("$##,###.##", RTBillRec.OptRevTax1) & "~" & Using("$##,###.##", RTBillRec.OptRevTax2) & "~" & Using("$##,###.##", RTBillRec.OptRevTax3)
    Else
      Print #AHandle, "Real" & "~" & CStr(TaxCust.CountyAcct) & "~" & QPTrim$(TaxCust.CustName) & "~" & Using("$##,###.##", RTBillRec.OptRevTax1) & "~" & Using("$##,###.##", RTBillRec.OptRevTax2) & "~" & Using("$##,###.##", RTBillRec.OptRevTax3)
    End If
  Next x
  
Done:
  Close
  If NumOfPTBRecs > 0 And NumOfRTBRecs > 0 Then
    MsgBox ("A spreadsheet for optional revenue for real and personal billing records has been created in file 'billoptrevs.txt'.")
  ElseIf NumOfPTBRecs = 0 And NumOfRTBRecs > 0 Then
    MsgBox ("A spreadsheet for optional revenue for real billing records has been created in file 'billoptrevs.txt'.")
  ElseIf NumOfPTBRecs > 0 And NumOfRTBRecs = 0 Then
    MsgBox ("A spreadsheet for optional revenue for personal billing records has been created in file 'billoptrevs.txt'.")
  Else
    MsgBox ("No prebilling files exist. Spreadsheet was not created.")
  End If
End Sub

Private Sub cmdPenGapMCtoPers_Click()
  Dim PersRec As PersonalRecType
  Dim NumOfPRecs As Long
  Dim PHandle As Integer
  Dim x As Long, cnt As Long
  
  OpenPersPropFile PHandle, NumOfPRecs
  For x = 1 To NumOfPRecs
    Get PHandle, x, PersRec
    If PersRec.MCValue > 0 Then
      PersRec.PersVal = OldRound(PersRec.PersVal + PersRec.MCValue)
      PersRec.MCValue = 0
      Put PHandle, x, PersRec
      cnt = cnt + 1
    End If
  Next x
  
  Close
  Call TaxMsg(900, "A total of " + CStr(cnt) + " personal records were updated.")
    
End Sub

Private Sub cmdProcess1_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim Balance#
  Dim Principle1 As Double
  Dim Principle2 As Double
  Dim Principle3 As Double
  Dim Principle4 As Double
  Dim Principle5 As Double
  Dim Interest As Double
  Dim Future1 As Double
  Dim Future2 As Double
  Dim Collection As Double
  Dim LateList As Double
  Dim RevOpt1 As Double
  Dim RevOpt2 As Double
  Dim RevOpt3 As Double
  Dim Principle1Pd As Double
  Dim Principle2Pd As Double
  Dim Principle3Pd As Double
  Dim Principle4Pd As Double
  Dim Principle5Pd As Double
  Dim LateListPd As Double
  Dim Future1Pd As Double
  Dim Future2Pd As Double
  Dim CollectionPd As Double
  Dim InterestPd As Double
  Dim RevOpt1Pd As Double
  Dim RevOpt2Pd As Double
  Dim RevOpt3Pd As Double
  Dim Principle1Bill As Double
  Dim Principle2Bill As Double
  Dim Principle3Bill As Double
  Dim Principle4Bill As Double
  Dim Principle5Bill As Double
  Dim InterestBill As Double
  Dim Future1Bill As Double
  Dim Future2Bill As Double
  Dim CollectionBill As Double
  Dim LateListBill As Double
  Dim RevOpt1Bill As Double
  Dim RevOpt2Bill As Double
  Dim RevOpt3Bill As Double
  Dim NegCnt As Long
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim NextRec As Long
  Dim SaveHere As Long
  Dim SaveCnt As Long
  Dim StartDate As Integer
  Dim EndDate As Integer
  Dim TotRev As Double
'  Dim TotRevP As Double
  
'  StartDate = Date2Num(fptxtBegDate.Text)
'  EndDate = Date2Num(fptxtEndDate.Text)
  ReDim RevPay(1 To 9) As Double
  ReDim RevChg(1 To 9) As Double
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
     Select Case x
       Case 76303
         TaxTrans.Amount = TaxTrans.Amount
         TaxTrans.Revenue.Interest = 0 'TaxTrans.Revenue.Interest
         TaxTrans.Revenue.Principle1 = 0
         TaxTrans.Amount = 0
         Put TTHandle, x, TaxTrans
       Case 8436
         TaxTrans.Amount = TaxTrans.Amount
         TaxTrans.Revenue.Principle1 = 21.01 'TaxTrans.Revenue.Principle1
         TaxTrans.Revenue.Principle1Pd = 21.01 'TaxTrans.Revenue.Principle1Pd
         Put TTHandle, x, TaxTrans
       Case 11494
         TaxTrans.Amount = TaxTrans.Amount
         TaxTrans.Amount = 0
         TaxTrans.Revenue.Interest = 0 'TaxTrans.Revenue.Interest
         Put TTHandle, x, TaxTrans
       Case 11346
         TaxTrans.Amount = TaxTrans.Amount
         TaxTrans.Amount = 0
         TaxTrans.Revenue.Interest = 0 'TaxTrans.Revenue.Interest
         Put TTHandle, x, TaxTrans
       Case 11195
         TaxTrans.Amount = TaxTrans.Amount
         TaxTrans.Amount = 0
         TaxTrans.Revenue.Interest = 0 'TaxTrans.Revenue.Interest
         Put TTHandle, x, TaxTrans
       Case 11026
         TaxTrans.Amount = TaxTrans.Amount
         TaxTrans.Amount = 0
         TaxTrans.Revenue.Interest = 0 'TaxTrans.Revenue.Interest
         Put TTHandle, x, TaxTrans
       Case 10505
         TaxTrans.Amount = TaxTrans.Amount
         TaxTrans.Amount = 0
         TaxTrans.Revenue.Interest = 0 'TaxTrans.Revenue.Interest
         Put TTHandle, x, TaxTrans
       Case 8993
         TaxTrans.Amount = TaxTrans.Amount
          TaxTrans.Amount = 0
         TaxTrans.Revenue.Interest = 0 'TaxTrans.Revenue.Interest
         Put TTHandle, x, TaxTrans
       Case Else
     End Select
''    If TaxTrans.CustomerRec = 7052 Then '
'      If TaxTrans.TranType = 1 Then
''        Select Case x
'        Select Case TaxTrans.CustomerRec
'          Case 1800
'            TaxTrans.Description = "M Tax Bill 100"
''            RevChg(1) = TaxTrans.Revenue.Principle1
''            TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd
'          Case 1801
'            TaxTrans.Description = "M Tax Bill 200"
''            RevChg(2) = TaxTrans.Revenue.Principle1
'          Case 1802
'            TaxTrans.Description = "M Tax Bill 300"
''            RevChg(3) = TaxTrans.Revenue.Principle1
'          Case 1803
'            TaxTrans.Description = "M Tax Bill 400"
''            RevChg(4) = TaxTrans.Revenue.Principle1
'          Case 1804
'            TaxTrans.Description = "M Tax Bill 500"
''            RevChg(5) = TaxTrans.Revenue.Principle1
'          Case 1805
'            TaxTrans.Description = "M Tax Bill 600"
''            RevChg(6) = TaxTrans.Revenue.Principle1
'          Case 1806
'            TaxTrans.Description = "M Tax Bill 700"
''            RevChg(7) = TaxTrans.Revenue.Principle1
'          Case Else
''            RevChg(8) = TaxTrans.Revenue.Principle1
'        End Select
'        Put TTHandle, x, TaxTrans
''      Else
''        Select Case TaxTrans.BelongTo
''          Case 16892
''            RevPay(1) = OldRound(RevPay(1) + TaxTrans.Revenue.Principle1Pd)
'            TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd
'            TaxTrans.Revenue.Interest = TaxTrans.Revenue.Interest
'            TaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.CollectionPd
'            TaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateListPd
'            TaxTrans.Revenue.Future1Pd = TaxTrans.Revenue.Future1Pd
'            TaxTrans.Revenue.Future2Pd = TaxTrans.Revenue.Future2Pd
'            TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd
'            TaxTrans.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2Pd
'            TaxTrans.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3Pd
'            TaxTrans.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4Pd
'            TaxTrans.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5Pd
'            TaxTrans.Revenue.Collection = TaxTrans.Revenue.Collection
'            TaxTrans.Revenue.LateList = TaxTrans.Revenue.LateList
'            TaxTrans.Revenue.Future1 = TaxTrans.Revenue.Future1
'            TaxTrans.Revenue.Future2 = TaxTrans.Revenue.Future2
'            TaxTrans.Revenue.Principle1 = TaxTrans.Revenue.Principle1
'            TaxTrans.Revenue.Principle2 = TaxTrans.Revenue.Principle2
'            TaxTrans.Revenue.Principle3 = TaxTrans.Revenue.Principle3
'            TaxTrans.Revenue.Principle4 = TaxTrans.Revenue.Principle4
'            TaxTrans.Revenue.Principle5 = TaxTrans.Revenue.Principle5
'            TaxTrans.Amount = TaxTrans.Amount
''          Case 62842
''            RevPay(2) = OldRound(RevPay(2) + TaxTrans.Revenue.Principle1Pd)
''          Case 122662
''            RevPay(3) = OldRound(RevPay(3) + TaxTrans.Revenue.Principle1Pd)
''          Case 181625
''            RevPay(4) = OldRound(RevPay(4) + TaxTrans.Revenue.Principle1Pd)
''          Case 245680
''            RevPay(5) = OldRound(RevPay(5) + TaxTrans.Revenue.Principle1Pd)
''          Case 299621
''            RevPay(6) = OldRound(RevPay(6) + TaxTrans.Revenue.Principle1Pd)
''          Case 348941
''            RevPay(7) = OldRound(RevPay(7) + TaxTrans.Revenue.Principle1Pd)
''          Case Is > 0
''            RevPay(8) = OldRound(RevPay(8) + TaxTrans.Revenue.Principle1Pd)
''          Case Else
''            RevPay(9) = OldRound(RevPay(9) + TaxTrans.Revenue.Principle1Pd)
''       End Select
'     End If
'''      TotRevP = TotRev + TaxTrans.Revenue.Principle1Pd
'''      TotRev = TotRev + TaxTrans.Revenue.Principle1
'''      If TaxTrans.BelongTo > 0 Then
'''        Debug.Print CStr(TaxTrans.BelongTo)
'''      End If
''
'''    End If
  Next x
  Exit Sub
'
'  For x = 1 To 9
'    Debug.Print "Charges/Payments for " + CStr(x) + ":" + Using$("$##,##0.00", RevChg(x)) + " / " + Using$("$##,##0.00", RevPay(x))
'  Next x
  frmVATaxShowPctComp.Label1 = "Fixing Negative Balances"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  ReDim NegCust(1 To 1) As Long
  ReDim NegTrans(1 To 1) As Long
  
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    TaxTrans.CustomerRec = TaxTrans.CustomerRec
    If TaxTrans.TransDate < StartDate Or TaxTrans.TransDate > EndDate Then GoTo OD
    If TaxTrans.TranType = 1 Then
      TotRev = OldRound#(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TotRev = OldRound#(TotRev + TaxTrans.Revenue.LateList)
      Balance# = OldRound#(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.LateList)
      Balance# = OldRound#(Balance# + TaxTrans.Revenue.Future1 + TaxTrans.Revenue.Future2 + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Future1Pd + TaxTrans.Revenue.Future2Pd + TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd))
      If Balance# < 0 Then
        NegCnt = NegCnt + 1
        ReDim Preserve NegTrans(1 To NegCnt) As Long
        NegTrans(NegCnt) = x
        ReDim Preserve NegCust(1 To NegCnt) As Long
        NegCust(NegCnt) = TaxTrans.CustomerRec
      End If
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTTRecs
OD:
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  frmVATaxShowPctComp.Label1 = "Fixing Negative Balances"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
    
  For x = 1 To NegCnt
    If NegCust(x) <= 0 Then GoTo SkipIt
    Get TCHandle, NegCust(x), TaxCust
      SaveCnt = 0
      Principle1Pd = 0
      Future1Pd = 0
      Future2Pd = 0
      InterestPd = 0
      CollectionPd = 0
      LateListPd = 0
      Interest = 0
      For y = 1 To NumOfTTRecs
        Get TTHandle, y, TaxTrans
        If TaxTrans.TranType = 1 Then 'check all billing trans and strip out negatives
          If TaxTrans.Revenue.Principle1 < 0 Then
            TaxTrans.Revenue.Principle1 = 0
            TaxTrans.Amount = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList)
            Put TTHandle, y, TaxTrans
          End If
          If TaxTrans.Revenue.Future1 < 0 Then
            TaxTrans.Revenue.Future1 = 0
            Put TTHandle, y, TaxTrans
          End If
          If TaxTrans.Revenue.Future2 < 0 Then
            TaxTrans.Revenue.Future2 = 0
            Put TTHandle, y, TaxTrans
          End If
          If TaxTrans.Revenue.Interest < 0 Then
            TaxTrans.Revenue.Interest = 0
            Put TTHandle, y, TaxTrans
          End If
          If TaxTrans.Revenue.Collection < 0 Then
            TaxTrans.Revenue.Collection = 0
            Put TTHandle, y, TaxTrans
          End If
          If TaxTrans.Revenue.LateList < 0 Then
            TaxTrans.Revenue.LateList = 0
            TaxTrans.Amount = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList)
            Put TTHandle, y, TaxTrans
          End If
          GoTo NextTrans
        End If
        If TaxTrans.BelongTo = NegTrans(x) Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
          Principle1Bill = TaxTrans.Revenue.Principle1
          InterestBill = TaxTrans.Revenue.Interest
          Future1Bill = TaxTrans.Revenue.Future1
          Future2Bill = TaxTrans.Revenue.Future2
          CollectionBill = TaxTrans.Revenue.Collection
          LateListBill = TaxTrans.Revenue.LateList
          Get TTHandle, y, TaxTrans
          Select Case TaxTrans.TranType
            Case 2, 7 'payment, adjustment
                Principle1Pd = OldRound#(Principle1Pd + TaxTrans.Revenue.Principle1Pd) 'collect all revenues and
                'wait for a problem
                If Principle1Pd > Principle1Bill Then 'OK...we have a problem
                  If TaxTrans.Revenue.Principle1 < 0 Then TaxTrans.Revenue.Principle1 = 0
                  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - (Principle1Pd - Principle1Bill))
                  Principle1Pd = OldRound(Principle1Pd - (Principle1Pd - Principle1Bill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - Principle1Pd) 'reduce the existing total amount for
                  'this transaction by the amount of the overpay
                  Put TTHandle, y, TaxTrans 'save it to this transaction
                  SaveHere = TaxTrans.BelongTo 'establish the bill record
                  Get TTHandle, TaxTrans.BelongTo, TaxTrans 'pull the bill record
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.Principle1Pd = Principle1Pd 'correct the bill record for this resource
                  'TaxTrans.Amount is not affected because this value only reflects the original charges
                  'and not any updates to the bill
                  Put TTHandle, SaveHere, TaxTrans 'save it to the bill record
                  Get TTHandle, y, TaxTrans 'go back to original record
                End If
                Future1Pd = OldRound(Future1Pd + TaxTrans.Revenue.Future1Pd)
                If Future1Pd > Future1Bill Then
                  If TaxTrans.Revenue.Future1 < 0 Then TaxTrans.Revenue.Future1 = 0
                  TaxTrans.Revenue.Future1Pd = OldRound(TaxTrans.Revenue.Future1Pd - (Future1Pd - Future1Bill))
                  Future1Pd = OldRound(Future1Pd - (Future1Pd - Future1Bill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - Future1Pd)
                  Put TTHandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get TTHandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.Future1Pd = OldRound(TaxTrans.Revenue.Future1Pd - (Future1Pd - Future1Bill))
                  Put TTHandle, SaveHere, TaxTrans
                  Get TTHandle, y, TaxTrans
                End If
                Future2Pd = OldRound(Future2Pd + TaxTrans.Revenue.Future2Pd)
                If Future2Pd > Future2Bill Then
                  If TaxTrans.Revenue.Future2 < 0 Then TaxTrans.Revenue.Future2 = 0
                  TaxTrans.Revenue.Future2Pd = OldRound(TaxTrans.Revenue.Future2Pd - (Future2Pd - Future2Bill))
                  Future2Pd = OldRound(Future2Pd - (Future2Pd - Future2Bill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - Future2Pd)
                  Put TTHandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get TTHandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.Future2Pd = OldRound(TaxTrans.Revenue.Future2Pd - (Future2Pd - Future2Bill))
                  Put TTHandle, SaveHere, TaxTrans
                  Get TTHandle, y, TaxTrans
                End If
                InterestPd = OldRound(InterestPd + TaxTrans.Revenue.InterestPd)
                If InterestPd > InterestBill Then
                  If TaxTrans.Revenue.Interest < 0 Then TaxTrans.Revenue.Interest = 0
                  TaxTrans.Revenue.InterestPd = OldRound(TaxTrans.Revenue.InterestPd - (InterestPd - InterestBill))
                  InterestPd = OldRound(InterestPd - (InterestPd - InterestBill))
                  TaxTrans.Amount = TaxTrans.Amount - InterestPd
                  Put TTHandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get TTHandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.InterestPd = InterestPd
                  Put TTHandle, SaveHere, TaxTrans
                  Get TTHandle, y, TaxTrans
                End If
                CollectionPd = OldRound(CollectionPd + TaxTrans.Revenue.CollectionPd)
                If CollectionPd > CollectionBill Then
                  If TaxTrans.Revenue.Collection < 0 Then TaxTrans.Revenue.Collection = 0
                  TaxTrans.Revenue.CollectionPd = OldRound(TaxTrans.Revenue.CollectionPd - (CollectionPd - CollectionBill))
                  CollectionPd = OldRound(CollectionPd - (CollectionPd - CollectionBill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - CollectionPd)
                  Put TTHandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get TTHandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.CollectionPd = OldRound(TaxTrans.Revenue.CollectionPd - (CollectionPd - CollectionBill))
                  Put TTHandle, SaveHere, TaxTrans
                  Get TTHandle, y, TaxTrans
                End If
                LateListPd = OldRound(LateListPd + TaxTrans.Revenue.LateListPd)
                If LateListPd > LateListBill Then
                  If TaxTrans.Revenue.LateList < 0 Then TaxTrans.Revenue.LateList = 0
                  TaxTrans.Revenue.LateListPd = OldRound(TaxTrans.Revenue.LateListPd - (LateListPd - LateListBill))
                  LateListPd = OldRound(LateListPd - (LateListPd - LateListBill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - LateListPd)
                  Put TTHandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get TTHandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.LateListPd = OldRound(TaxTrans.Revenue.LateListPd - (LateListPd - LateListBill))
                  Put TTHandle, SaveHere, TaxTrans
                  Get TTHandle, y, TaxTrans
                End If
              Case 4:
                Interest = OldRound(Interest + TaxTrans.Revenue.Interest)
                If Interest > InterestBill Then
                  If TaxTrans.Revenue.Interest < 0 Then TaxTrans.Revenue.Interest = 0
                  TaxTrans.Revenue.Interest = OldRound(TaxTrans.Revenue.Interest - (Interest - InterestBill))
                  Interest = OldRound(Interest - (Interest - InterestBill))
                  TaxTrans.Amount = TaxTrans.Revenue.Interest ' OldRound(TaxTrans.Amount - Interest)
                  Put TTHandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get TTHandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.InterestPd = InterestBill ' OldRound(TaxTrans.Revenue.InterestPd - (Interest - InterestBill))
                  Put TTHandle, SaveHere, TaxTrans
                  Get TTHandle, y, TaxTrans
                End If
              Case 6, 8:
                Collection = OldRound(Collection + TaxTrans.Revenue.Collection)
                If Collection > CollectionBill Then
                  If TaxTrans.Revenue.Collection < 0 Then TaxTrans.Revenue.Collection = 0
                  TaxTrans.Revenue.Collection = OldRound(TaxTrans.Revenue.Collection - (Collection - CollectionBill))
                  Collection = OldRound(Collection - (Collection - CollectionBill))
                  TaxTrans.Amount = TaxTrans.Revenue.Collection 'OldRound(TaxTrans.Amount - Collection)
                  Put TTHandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get TTHandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.CollectionPd = CollectionBill 'OldRound(TaxTrans.Revenue.CollectionPd - (Collection - CollectionBill))
                  Put TTHandle, SaveHere, TaxTrans
                  Get TTHandle, y, TaxTrans
                End If
            End Select
         End If
NextTrans:
      Next y
      If SaveCnt = 0 Then 'been through all transactions for this negative balance and
      'couldn't find a transaction that belonged to it. This means something is negative that
      'shouldn't be
        Get TTHandle, NegTrans(x), TaxTrans 'always trans type #1
        If TaxTrans.Revenue.Principle1 < 0 Then TaxTrans.Revenue.Principle1 = 0
        TaxTrans.Amount = TaxTrans.Revenue.Principle1
        If TaxTrans.Revenue.Interest < 0 Then TaxTrans.Revenue.Interest = 0
        TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.Revenue.Interest)
        If TaxTrans.Revenue.Collection < 0 Then TaxTrans.Revenue.Collection = 0
        TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.Revenue.Collection)
        If TaxTrans.Revenue.LateList < 0 Then TaxTrans.Revenue.LateList = 0
        TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.Revenue.LateList)
        If TaxTrans.Revenue.Future1 < 0 Then TaxTrans.Revenue.Future1 = 0
        TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.Revenue.Future1)
        If TaxTrans.Revenue.Future2 < 0 Then TaxTrans.Revenue.Future2 = 0
        TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.Revenue.Future2)
        Put TTHandle, NegTrans(x), TaxTrans
      End If
    frmVATaxShowPctComp.ShowPctComp x, NegCnt
SkipIt:
  Next x
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
           
  frmVATaxShowPctComp.Label1 = "Deleting All Future Field Values"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  For y = 1 To NumOfTTRecs 'strip out all future1 and future2 fields
    Get TTHandle, y, TaxTrans
    If TaxTrans.Revenue.Future1Pd <> 0 Then
      TaxTrans.Revenue.Future1Pd = 0
    End If
    If TaxTrans.Revenue.Future1 <> 0 Then
      TaxTrans.Revenue.Future1 = 0
    End If
    If TaxTrans.Revenue.Future2Pd <> 0 Then
      TaxTrans.Revenue.Future2Pd = 0
    End If
    If TaxTrans.Revenue.Future2 <> 0 Then
      TaxTrans.Revenue.Future2 = 0
    End If
    Put TTHandle, y, TaxTrans
    frmVATaxShowPctComp.ShowPctComp y, NumOfTTRecs
  Next y
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
          
  Close
  Call Savemsg(900, "The negative and future vales repairs have completed successfully.")
End Sub


Private Sub cmdProcess11_Click()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long, NextRec As Long
  Dim ThisAmt As Double
  Dim Princ As Double
  Dim Adv As Double
  Dim LateList As Double
  Dim Interest As Double
  Dim Opt1 As Double
  Dim Opt2 As Double
  Dim Opt3 As Double
  Dim GetBill As Long
  Dim ThisDate$
  Dim SaveAmt As Double
  Dim ChngCnt As Long
  
  OpenTaxCustFile CHandle, NumOfCRecs
  OpenTaxTransFile THandle, NumOfTRecs
  For x = 1 To NumOfTRecs
    Get THandle, x, TaxTrans
    TaxTrans.Amount = TaxTrans.Amount
    TaxTrans.TranType = TaxTrans.TranType
    TaxTrans.CustomerRec = TaxTrans.CustomerRec
    If TaxTrans.BillType = "C" Then
      Get CHandle, TaxTrans.CustomerRec, TaxCust
      If TaxCust.FirstPersRec > 0 And TaxCust.FirstPropRec > 0 Then
        TaxTrans.BillType = ""
        Put THandle, x, TaxTrans
        ChngCnt = ChngCnt + 1
      ElseIf TaxCust.FirstPersRec > 0 Then
        TaxTrans.BillType = "P"
        Put THandle, x, TaxTrans
        ChngCnt = ChngCnt + 1
      ElseIf TaxCust.FirstPropRec > 0 Then
        TaxTrans.BillType = "R"
        Put THandle, x, TaxTrans
        ChngCnt = ChngCnt + 1
      Else
        TaxTrans.BillType = ""
        Put THandle, x, TaxTrans
        ChngCnt = ChngCnt + 1
      End If
    End If
  Next x

  Close CHandle
  Close THandle
  
  If ChngCnt > 0 Then
    Call Savemsg(900, "A total of " + CStr(ChngCnt) + " transaction records were corrected.")
  Else
    Call Savemsg(900, "No transaction records needed correcting.")
  End If
End Sub

Private Sub cmdProcess9_Click()
'  Call MakeAmtEqualRevs
  Dim TaxTrans As TaxTransactionType
  Dim NumOfTrans As Long
  Dim THandle As Integer
  Dim x As Long
  Dim TCnt As Long
  Dim CountIt As Boolean
  Dim ThisCollection As Double
  Dim ThisPenalty As Double
  Dim ThisInterest As Double
  Dim ThisLateList As Double
  Dim ThisPrinciple1 As Double
  Dim ThisPrinciple2 As Double
  Dim ThisPrinciple3 As Double
  Dim ThisPrinciple4 As Double
  Dim ThisPrinciple5 As Double
  Dim ThisRevOpt1 As Double
  Dim ThisRevOpt2 As Double
  Dim ThisRevOPt3 As Double
  Dim ThisBelongTo As Long
  
  frmVATaxShowPctComp.Label1 = "Updating Release Revenues"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  
  OpenTaxTransFile THandle, NumOfTrans
  For x = 1 To NumOfTrans
    CountIt = False
    Get THandle, x, TaxTrans
    If TaxTrans.TranType = 3 Then
      ThisCollection = 0
      ThisPenalty = 0
      ThisInterest = 0
      ThisLateList = 0
      ThisPrinciple1 = 0
      ThisPrinciple2 = 0
      ThisPrinciple3 = 0
      ThisPrinciple4 = 0
      ThisPrinciple5 = 0
      ThisRevOpt1 = 0
      ThisRevOpt2 = 0
      ThisRevOPt3 = 0
      If TaxTrans.Revenue.Penalty > 0 And TaxTrans.Revenue.PenaltyPd = 0 Then
        TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.Penalty
        TaxTrans.Revenue.Penalty = 0
        ThisPenalty = TaxTrans.Revenue.PenaltyPd
        CountIt = True
      End If
      If TaxTrans.Revenue.Collection > 0 And TaxTrans.Revenue.CollectionPd = 0 Then
        TaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.Collection
        TaxTrans.Revenue.Collection = 0
        ThisCollection = TaxTrans.Revenue.CollectionPd
        CountIt = True
      End If
      If TaxTrans.Revenue.Interest > 0 And TaxTrans.Revenue.InterestPd = 0 Then
        TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.Interest
        TaxTrans.Revenue.Interest = 0
        ThisInterest = TaxTrans.Revenue.InterestPd
        CountIt = True
      End If
      If TaxTrans.Revenue.LateList > 0 And TaxTrans.Revenue.LateListPd = 0 Then
        TaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateList
        TaxTrans.Revenue.LateList = 0
        ThisLateList = TaxTrans.Revenue.LateListPd
        CountIt = True
      End If
      If TaxTrans.Revenue.Principle1 > 0 And TaxTrans.Revenue.Principle1Pd = 0 Then
        TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1
        TaxTrans.Revenue.Principle1 = 0
        ThisPrinciple1 = TaxTrans.Revenue.Principle1Pd
        CountIt = True
      End If
      If TaxTrans.Revenue.Principle2 > 0 And TaxTrans.Revenue.Principle2Pd = 0 Then
        TaxTrans.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2
        TaxTrans.Revenue.Principle2 = 0
        ThisPrinciple2 = TaxTrans.Revenue.Principle2Pd
        CountIt = True
      End If
      If TaxTrans.Revenue.Principle3 > 0 And TaxTrans.Revenue.Principle3Pd = 0 Then
        TaxTrans.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3
        TaxTrans.Revenue.Principle3 = 0
        ThisPrinciple3 = TaxTrans.Revenue.Principle3Pd
        CountIt = True
      End If
      If TaxTrans.Revenue.Principle4 > 0 And TaxTrans.Revenue.Principle4Pd = 0 Then
        TaxTrans.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4
        TaxTrans.Revenue.Principle4 = 0
        ThisPrinciple4 = TaxTrans.Revenue.Principle4Pd
        CountIt = True
      End If
      If TaxTrans.Revenue.Principle5 > 0 And TaxTrans.Revenue.Principle5Pd = 0 Then
        TaxTrans.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5
        TaxTrans.Revenue.Principle5 = 0
        ThisPrinciple5 = TaxTrans.Revenue.Principle5Pd
        CountIt = True
      End If
      If TaxTrans.Revenue.RevOpt1 > 0 And TaxTrans.Revenue.RevOpt1Pd = 0 Then
        TaxTrans.Revenue.RevOpt1Pd = TaxTrans.Revenue.RevOpt1
        TaxTrans.Revenue.RevOpt1 = 0
        ThisRevOpt1 = TaxTrans.Revenue.RevOpt1Pd
        CountIt = True
      End If
      If TaxTrans.Revenue.RevOpt2 > 0 And TaxTrans.Revenue.RevOpt2Pd = 0 Then
        TaxTrans.Revenue.RevOpt2Pd = TaxTrans.Revenue.RevOpt2
        TaxTrans.Revenue.RevOpt2 = 0
        ThisRevOpt2 = TaxTrans.Revenue.RevOpt2Pd
        CountIt = True
      End If
      If TaxTrans.Revenue.RevOpt3 > 0 And TaxTrans.Revenue.RevOpt3Pd = 0 Then
        TaxTrans.Revenue.RevOpt3Pd = TaxTrans.Revenue.RevOpt3
        TaxTrans.Revenue.RevOpt3 = 0
        ThisRevOPt3 = TaxTrans.Revenue.RevOpt3Pd
        CountIt = True
      End If
      Put THandle, x, TaxTrans
   End If
   If CountIt = True Then
     ThisBelongTo = TaxTrans.BelongTo
     Get THandle, ThisBelongTo, TaxTrans
       If ThisCollection > 0 Then
         TaxTrans.Revenue.Collection = OldRound(TaxTrans.Revenue.Collection + ThisCollection)
         TaxTrans.Revenue.CollectionPd = OldRound(TaxTrans.Revenue.CollectionPd + ThisCollection)
       End If
       If ThisPenalty > 0 Then
         TaxTrans.Revenue.Penalty = OldRound(TaxTrans.Revenue.Penalty + ThisPenalty)
         TaxTrans.Revenue.PenaltyPd = OldRound(TaxTrans.Revenue.PenaltyPd + ThisPenalty)
       End If
       If ThisInterest > 0 Then
         TaxTrans.Revenue.Interest = OldRound(TaxTrans.Revenue.Interest + ThisInterest)
         TaxTrans.Revenue.InterestPd = OldRound(TaxTrans.Revenue.InterestPd + ThisInterest)
       End If
       If ThisLateList > 0 Then
         TaxTrans.Revenue.LateList = OldRound(TaxTrans.Revenue.LateList + ThisLateList)
         TaxTrans.Revenue.LateListPd = OldRound(TaxTrans.Revenue.LateListPd + ThisLateList)
       End If
       If ThisPrinciple1 > 0 Then
         TaxTrans.Revenue.Principle1 = OldRound(TaxTrans.Revenue.Principle1 + ThisPrinciple1)
         TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + ThisPrinciple1)
       End If
       If ThisPrinciple2 > 0 Then
         TaxTrans.Revenue.Principle2 = OldRound(TaxTrans.Revenue.Principle2 + ThisPrinciple2)
         TaxTrans.Revenue.Principle2Pd = OldRound(TaxTrans.Revenue.Principle2Pd + ThisPrinciple2)
       End If
       If ThisPrinciple3 > 0 Then
         TaxTrans.Revenue.Principle3 = OldRound(TaxTrans.Revenue.Principle3 + ThisPrinciple3)
         TaxTrans.Revenue.Principle3Pd = OldRound(TaxTrans.Revenue.Principle3Pd + ThisPrinciple3)
       End If
       If ThisPrinciple4 > 0 Then
         TaxTrans.Revenue.Principle4 = OldRound(TaxTrans.Revenue.Principle4 + ThisPrinciple4)
         TaxTrans.Revenue.Principle4Pd = OldRound(TaxTrans.Revenue.Principle4Pd + ThisPrinciple4)
       End If
       If ThisPrinciple5 > 0 Then
         TaxTrans.Revenue.Principle5 = OldRound(TaxTrans.Revenue.Principle5 + ThisPrinciple5)
         TaxTrans.Revenue.Principle5Pd = OldRound(TaxTrans.Revenue.Principle5Pd + ThisPrinciple5)
       End If
       If ThisRevOpt1 > 0 Then
         TaxTrans.Revenue.RevOpt1 = OldRound(TaxTrans.Revenue.RevOpt1 + ThisRevOpt1)
         TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + ThisRevOpt1)
       End If
       If ThisRevOpt2 > 0 Then
         TaxTrans.Revenue.RevOpt2 = OldRound(TaxTrans.Revenue.RevOpt2 + ThisRevOpt2)
         TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + ThisRevOpt2)
       End If
       If ThisRevOPt3 > 0 Then
         TaxTrans.Revenue.RevOpt3 = OldRound(TaxTrans.Revenue.RevOpt3 + ThisRevOPt3)
         TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + ThisRevOPt3)
       End If
       Put THandle, ThisBelongTo, TaxTrans
     TCnt = TCnt + 1
   End If
   frmVATaxShowPctComp.ShowPctComp x, NumOfTrans
  Next x
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
    
  Close THandle
  
  Call Savemsg(900, "A total of " + CStr(TCnt) + " release transactions were updated successfully.")
  
End Sub

Private Sub cmdRelinkPayToBill_Click()
Call RelinkBelongTosWithBills
End Sub

Private Sub cmdRelinkBelongTosToBill_Click()
  Call RelinkBelongTosWithBills
End Sub

Private Sub cmdRemovePenalties_Click()
  Dim TaxTrans As TaxTransactionType
  Dim NumOfTrans As Long
  Dim THandle As Integer
  Dim x As Long
  Dim cnt As Integer
  Dim BelongTo As Long
  Dim CmpDate As Integer
  
  CmpDate = Date2Num("12/11/2008")
  OpenTaxTransFile THandle, NumOfTrans
  For x = 1 To NumOfTrans
  Get THandle, x, TaxTrans
    If TaxTrans.TranType = 5 And TaxTrans.TransDate = CmpDate Then
      BelongTo = TaxTrans.BelongTo
      TaxTrans.Amount = 0
      TaxTrans.Revenue.Penalty = 0
      Put THandle, x, TaxTrans
      Get THandle, BelongTo, TaxTrans
        TaxTrans.Revenue.Penalty = 0
      Put THandle, BelongTo, TaxTrans
    End If
  Next x
  Close
  MsgBox ("Done.")

End Sub

Private Sub cmdRemoveRealPinSpaces_Click()
  Call UpdateChilhowieRealPins
  Exit Sub
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim x As Long
  
  OpenRealPropFile RHandle, NumOfRRecs
  For x = 1 To NumOfRRecs
    Get RHandle, x, RealRec
    RealRec.RealPin = ReplaceString(RealRec.RealPin, " ", "")
    Put RHandle, x, RealRec
  Next x
  
  Close
  MsgBox ("Done.")
End Sub
Private Sub UpdateChilhowieRealPins()
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim TaxCust As TaxCustType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim ch As String
  Dim x As Integer
  Dim TextLine As String
  Dim PHandle As Integer
  Dim RowLen As Integer
  Dim Collect As String
  Dim CustNum As Long
  Dim cnt As Long
  Dim Doc As String
  Dim ReadFile As String
  Dim RealNum As Long
  Dim TransRec As TaxTransactionType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Long
  Dim y As Long
  Dim RealPin As String
  Dim GoThru As Integer
  
  frmVATaxShowPctComp.Label1 = "Repairing Tax Years"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  
  ReadFile = "FixRealPns.csv"
  Doc = App.Path
  Dim DocPath As String
  DocPath = Doc + "\" + ReadFile
  PHandle = FreeFile
  If Exist(DocPath) Then
    Open DocPath For Input As #PHandle  ' Open file.
  End If
  OpenRealPropFile RHandle, NumOfRRecs
  OpenTaxCustFile THandle, NumOfTRecs
  OpenTaxTransFile TRHandle, NumOfTRRecs
  
  Line Input #PHandle, TextLine   ' Read first line into TextLine.
   Do While Not eof(PHandle)   ' Loop until end of file.
     Line Input #PHandle, TextLine   ' Read next line into Textline.
   GoThru = GoThru + 1
     RowLen = Len(TextLine)
     CustNum = 0
     Collect = ""
     For x = 1 To RowLen
'       If x = RowLen Then Stop
       ch = Mid(TextLine, x, 1)
       If ch <> "~" Then
         Collect = Collect + ch
       ElseIf ch = "~" Then
         If CustNum = 0 Then
           CustNum = CInt(Collect)
           Collect = ""
         End If
       End If
       If x = RowLen Then
         Get THandle, CustNum, TaxCust
         RealNum = TaxCust.FirstPropRec
         Get RHandle, RealNum, RealRec
         RealPin = QPTrim$(RealRec.RealPin)
         RealRec.RealPin = Collect
         Put RHandle, RealNum, RealRec
         cnt = cnt + 1
         
         For y = 1 To NumOfTRRecs
           Get TRHandle, y, TransRec
           If QPTrim$(TransRec.RealPin) = RealPin And RealPin <> "" Then
              TransRec.RealPin = Collect
              Put TRHandle, y, TransRec
           End If
         Next y
       End If
     Next x
     frmVATaxShowPctComp.ShowPctComp GoThru, 900
   Loop
   Unload frmVATaxShowPctComp
   Close
   MsgBox ("Done. A total of " + CStr(cnt) + " pins were updated.")
  
  
End Sub
Private Sub cmdRunningBal_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim Balance As Double
  Dim AHandle As Integer
  
  AHandle = FreeFile
  Open "runningbal.dat" For Output As AHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    Balance = GetCustBalance(TaxCust.Acct, -1)
    Print #AHandle, CStr(TaxCust.Acct) & "~" & QPTrim$(TaxCust.CustName) & "~" & Using("$##,###.##", Balance)
  Next x
  
  Close
  MsgBox ("Completed successfully.")

End Sub
Private Sub ExportFiles()
  Dim AHandle As Integer
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim TaxCust As TaxCustType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim PersType As String
  Dim PersAmt As Double
  Dim x As Long
  On Error Resume Next
  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxCustFile THandle, NumOfTRecs
  AHandle = FreeFile
  Open "taxexportfiles.txt" For Output As AHandle
  Print #AHandle, "Citipak Acct Num~String County Acct Num~Numeric County Acct Num~Personal Pin Num~Personal Value~Prop Type~PPTRA YN"
  For x = 1 To NumOfPRecs
    Get PHandle, x, PersRec
    If PersRec.CustPin > 0 Then
      Get THandle, PersRec.CustPin, TaxCust
      If PersRec.CVALUE > 0 Then
        PersType = "Farm Eq"
        PersAmt = PersRec.CVALUE
      ElseIf PersRec.MCValue > 0 Then
        PersType = "Merch Cap"
        PersAmt = PersRec.MCValue
      ElseIf PersRec.MHValue > 0 Then
        PersType = "Mobile Home"
        PersAmt = PersRec.MHValue
      ElseIf PersRec.MTValue > 0 Then
        PersType = "Machine Tools"
        PersAmt = PersRec.MTValue
      ElseIf PersRec.PersVal > 0 Then
        PersType = "Personal"
        PersAmt = PersRec.PersVal
      End If
    End If
    Print #AHandle, CStr(PersRec.CustPin) & "~" & QPTrim$(TaxCust.CountyAcctString) & "~" & CStr(TaxCust.CountyAcct) & "~" & QPTrim$(PersRec.PropPin) & "~" & Using("#,###,###.##", PersAmt) & "~" & PersType & "~"; PersRec.PPTRAYN
  Next x
  Close
  
  MsgBox ("Look for taxexportfiles.txt in the Citipak directory.")
End Sub

Private Sub cmdRunRealvsAcctBal_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim AHandle As Integer
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim NextRec As Long
  Dim RealBal As Double
  Dim AcctBal As Double
  AHandle = FreeFile
  Open "RealvsAcctBal.txt" For Output As AHandle
  Print #AHandle, "Cust Rec~Cust Balance~Real Pin~Real Balance"
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  frmVATaxShowPctComp.Label1 = "Running Comparison"
  frmVATaxShowPctComp.Show , Me
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo Skip
    If TaxCust.FirstPropRec > 0 Then
      AcctBal = GetCustBalance(x, -1)
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        If NextRec <= 0 Or NextRec > NumOfRealRecs Then GoTo Skip
        'MsgBox (CStr(NextRec))
        Get RHandle, NextRec, RealRec
        RealBal = GetRealBalance(RealRec.RealPin)
        NextRec = RealRec.NextRec
        Print #AHandle, CStr(x) + "~" + CStr(AcctBal) + "~" + QPTrim$(RealRec.RealPin) + "~" + CStr(RealBal)
       Loop
    End If
Skip:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
    End If
  Next x
  Close
  Unload frmVATaxShowPctComp
  MsgBox ("All done. Results are in RealvsAcctBal.txt.")
End Sub

Private Sub cmdUpdateAdd1andAdd2Long_Click()
  Call UpdateAddressFieldsLong
End Sub

Private Sub cmdUpdateAdd1andAdd2Short_Click()
  Call UpdateAddressFieldsShort
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
    Case vbKeyF2:
      Call cmdProcess9_Click
      KeyCode = 0
    Case vbKeyF9:
      SendKeys "%r"
      Call cmdProcess2_Click
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%0"
      Call cmdProcess3_Click
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%c"
      Call cmdProcess4_Click
      KeyCode = 0
    Case vbKeyF6:
      SendKeys "%e"
      Call cmdProcess5_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%s"
      Call cmdProcess6_Click
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
      ClearInUse PWcnt
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxDataRepair.")
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  MainLog ("User opened frmVATaxDataRepair.")
  Call LoadMe
End Sub

Private Sub LoadMe()
  Dim BHandle As Integer
  Dim CDateStr$
  Dim TaxTrans As TaxTransactionType
  Dim NumOfTRecs As Long
  Dim x As Long
  Dim THandle As Integer
  
'  Call FindBelongTo
'  Call AppendPtoPersCoNums
'  Call FindCoNum
'  Label14.Visible = False
'  cmdProcess9.Visible = False
'  Shape11.Visible = False
'  Call FixWarsawVA
'  Call FixAlberta
'  Call FixPocohantas
'  Call ForPatrick
'  Label3.Visible = False
'  Shape3.Visible = False
'  Label1.Visible = False
'  Label4.Visible = False
'  fptxtBegDate.Visible = False
'  fptxtEndDate.Visible = False
'  cmdProcess1.Visible = False
'  Call FixSpecificData
'  OpenTaxTransFile THandle, NumOfTRecs
'  For x = 1 To NumOfTRecs
'    Get THandle, x, NumOfTRecs
'    If TaxTrans.CustPin = 17484 And TaxTrans.TranType = 9 Then Stop
'    If TaxTrans.BelongTo = 115868 Then Stop
'    TaxTrans.TranType = TaxTrans.TranType
'  Next x
'  Close
  
'  If NumOfTRecs = 0 Then
'    Call TaxMsg(900, "No transactions stored.")
'    Close
'    Exit Sub
'  End If
  
'  For x = 1 To NumOfTRecs
'    Get THandle, x, TaxTrans
'    If TaxTrans.TransDate > 0 Then
'      fptxtBegDate.Text = MakeRegDate(x)
'      Close
'      Exit For
'    End If
'  Next x
    
'  lblBalloon.Visible = False
'  If Exist("cnvtdate.dat") Then
'    BHandle = FreeFile
'    Open "cnvtdate.dat" For Input As BHandle
'    Input #BHandle, CDateStr$
'    Close BHandle
'    If QPTrim(CDateStr$) = "" Then
'      fptxtEndDate.Text = Date
'      Exit Sub
'    End If
'    fptxtEndDate.Text = MakeRegDate(CInt(CDateStr$))
''    lblMessage.Visible = True
''    lblMessage.Caption = "The date in the 'Ending Date' field is the conversion date for DOS to Windows,  " + fptxtEndDate.Text + "."
'  Else
''    lblMessage.Visible = False
''    fptxtEndDate.Text = Date
'  End If
  
'  fptxtFiscalBeg.Text = "07/01"
'  fptxtFiscalEnd.Text = "06/30"
  
End Sub

Private Sub cmdProcess2_Click()
  Dim x As Long
  Dim TransRec As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxYear As Integer
  Dim YrCnt As Long
  
  frmVATaxShowPctComp.Label1 = "Repairing Tax Years"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  OpenTaxTransFile THandle, NumOfTRecs
  
  If NumOfTRecs = 0 Then
    Call TaxMsg(900, "There are no transactions saved.")
    Exit Sub
  End If
  
  For x = 1 To NumOfTRecs
    Get THandle, x, TransRec
    If TransRec.TranType <> 1 Then
      If TransRec.TaxYear = 0 Then
        If TransRec.BelongTo > 0 Then
          Get THandle, TransRec.BelongTo, TransRec
          TaxYear = TransRec.TaxYear
          Get THandle, x, TransRec
          TransRec.TaxYear = TaxYear
          Put THandle, x, TransRec
          YrCnt = YrCnt + 1
        End If
      End If
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTRecs
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  Call Savemsg(900, "A total of " + CStr(YrCnt) + " errant tax years were corrected successfully.")
  
End Sub

Private Sub ResequencePins()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  frmVATaxShowPctComp.Label1 = "Resequencing Customer Pin Numbers"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  OpenTaxCustFile TCHandle, NumOfTCRecs
  If NumOfTCRecs = 0 Then
    Close
    Call TaxMsg(900, "No tax customers saved. Re-sequencing aborted.")
    Exit Sub
  End If
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    TaxCust.Acct = x
    TaxCust.PIN = x
    Put TCHandle, x, TaxCust
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  Call Savemsg(900, "Re-sequencing of pin numbers completed successfully.")
  
End Sub

Private Sub cmdProcess3_Click()
  Dim x As Long
  Dim TransRec As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxYear As Integer
  Dim YrCnt As Long
  Dim UseThisYear As Integer
  Dim FiscBeg As Integer
  Dim TestYear$
  Dim EndOfYear As Integer
  Dim BegOfYear As Integer
  Dim JulBeg As Integer
  Dim TestJulBeg As Integer
  Dim TestJulEnd As Integer
  
  frmVATaxShowPctComp.Label1 = "Repairing Tax Years"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  OpenTaxTransFile THandle, NumOfTRecs
  
  UseThisYear = Year(fptxtFiscalBeg.Text)
  TestYear = "12/31/" + CStr(UseThisYear)
  EndOfYear = Date2Num(TestYear)
  FiscBeg = Date2Num(fptxtFiscalBeg.Text)
  JulBeg = EndOfYear - FiscBeg 'julian start date
  If NumOfTRecs = 0 Then
    Call TaxMsg(900, "There are no transactions saved.")
    Close
    Exit Sub
  End If
  
  For x = 1 To NumOfTRecs
    Get THandle, x, TransRec
    If TransRec.TaxYear = 0 Then
      If TransRec.TranType <> 1 Then GoTo SkipIt
      If TransRec.TransDate < 0 Then
        TransRec.TransDate = 0
      End If
      TestJulBeg = TransRec.TransDate
      TestYear = Mid(MakeRegDate(TransRec.TransDate), 7, 4)
      EndOfYear = Date2Num("12/31/" + TestYear)
      BegOfYear = Date2Num("01/01/" + TestYear)
      If (EndOfYear - TestJulBeg) >= (EndOfYear - (JulBeg + BegOfYear)) Then
        TransRec.TaxYear = CInt(TestYear)
      Else
        TransRec.TaxYear = CInt(TestYear) + 1
      End If
      YrCnt = YrCnt + 1
      Put THandle, x, TransRec
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTRecs
SkipIt:
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Call cmdProcess2_Click
  
  Close
  Call Savemsg(900, "A total of " + CStr(YrCnt) + " errant tax years were corrected successfully.")
  
End Sub

Private Sub cmdProcess4_Click()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim cnt As Long
  
  frmVATaxShowPctComp.Label1 = "Resequencing Customer Pin Numbers"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  OpenTaxCustFile TCHandle, NumOfTCRecs
  If NumOfTCRecs = 0 Then
    Close
    Call TaxMsg(900, "No tax customers saved. Re-sequencing aborted.")
    Exit Sub
  End If
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Acct <> x Then cnt = cnt + 1
    TaxCust.Acct = x
    TaxCust.PIN = x
    Put TCHandle, x, TaxCust
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  Call Savemsg(900, "Re-sequencing of " + CStr(cnt) + " pin numbers completed successfully.")
  
End Sub


'Private Sub fpcmdHelp_Click()
'  If InStr(fpcmdHelp.Text, "On") Then
'    fpcmdHelp.Text = "F1 &Turn Help Off"
'    btnHelp.AutoScan = fpAutoScanPopupOnly
'    lblBalloon.Visible = True
'  ElseIf InStr(fpcmdHelp.Text, "Off") Then
'    fpcmdHelp.Text = "F1 &Turn Help On"
'    btnHelp.AutoScan = fpAutoScanOff
'    lblBalloon.Visible = False
'  End If
'End Sub

'Private Sub cmdProcess5_Click()
'  Dim TaxTrans As TaxTransactionType
'  Dim TTHandle As Integer
'  Dim NumOfTTRecs As Long
'  Dim x As Long
'  Dim ThisAmtP As Double
'  Dim ThisAmtB As Double
'  Dim RepairCnt As Long
'
'  OpenTaxTransFile TTHandle, NumOfTTRecs
'  frmVATaxShowPctComp.Label1 = "Making Transaction Totals Equal Revenues"
'  frmVATaxShowPctComp.CmdCancel.Visible = False
'  frmVATaxShowPctComp.Show , Me
'  DoEvents
'  EnableCloseButton Me.hwnd, False
'  DoEvents
'  For x = 1 To NumOfTTRecs
'  Get TTHandle, x, TaxTrans
''    If x = 26519 Then Stop
'    ThisAmtP = OldRound(TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.Future1Pd + TaxTrans.Revenue.Future2Pd)
'    ThisAmtP = OldRound(ThisAmtP + TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.PenaltyPd)
'    ThisAmtP = OldRound(ThisAmtP + TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd)
'    ThisAmtP = OldRound(ThisAmtP + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd)
'    ThisAmtP = OldRound(ThisAmtP + TaxTrans.Revenue.Principle5Pd + TaxTrans.Revenue.RevOpt1Pd)
'    ThisAmtP = OldRound(ThisAmtP + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd)
''    If TaxTrans.TranType <> 1 Then
'      ThisAmtB = OldRound(TaxTrans.Revenue.Collection + TaxTrans.Revenue.Future1 + TaxTrans.Revenue.Future2)
'      ThisAmtB = OldRound(ThisAmtB + TaxTrans.Revenue.Interest + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Penalty)
''    Else
''      ThisAmtB = OldRound(TaxTrans.Revenue.Future1 + TaxTrans.Revenue.Future2)
''      ThisAmtB = OldRound(ThisAmtB + TaxTrans.Revenue.LateList) ' + TaxTrans.Revenue.Penalty)
''    End If
'    ThisAmtB = OldRound(ThisAmtB + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2)
'    ThisAmtB = OldRound(ThisAmtB + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4)
'    ThisAmtB = OldRound(ThisAmtB + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.RevOpt1)
'    ThisAmtB = OldRound(ThisAmtB + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
'
'    Select Case TaxTrans.TranType
'      Case 1:
'        If TaxTrans.Amount <> ThisAmtB Then
'          TaxTrans.Amount = ThisAmtB
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 2:
'        If TaxTrans.Amount <> ThisAmtP Then
'          TaxTrans.Amount = ThisAmtP
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 3:
'        If TaxTrans.Amount <> ThisAmtB Then
'          TaxTrans.Amount = ThisAmtB
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 4:
'        If TaxTrans.Amount <> ThisAmtB Then
'          TaxTrans.Amount = ThisAmtB
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 6, 8:
'        If TaxTrans.Amount <> ThisAmtB Then
'          TaxTrans.Amount = ThisAmtB
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 7:
'        If TaxTrans.Amount <> ThisAmtP Then
'          TaxTrans.Amount = ThisAmtP
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 8:
'        If TaxTrans.Amount <> ThisAmtB Then
'          TaxTrans.Amount = ThisAmtB
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 9:
'        If TaxTrans.Amount <> ThisAmtP Then
'          TaxTrans.Amount = ThisAmtP
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 13:
'        If TaxTrans.Amount <> ThisAmtB Then
'          TaxTrans.Amount = ThisAmtB
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 14:
'         If TaxTrans.Amount <> ThisAmtB Then
'          TaxTrans.Amount = ThisAmtB
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 21:
'        If TaxTrans.Amount <> ThisAmtP Then
'          TaxTrans.Amount = ThisAmtP
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 22:
'        If TaxTrans.Amount <> ThisAmtP Then
'          TaxTrans.Amount = ThisAmtP
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 10:
'        If TaxTrans.Amount <> ThisAmtP Then
'          TaxTrans.Amount = ThisAmtP
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 11:
'        If TaxTrans.Amount <> ThisAmtP Then
'          TaxTrans.Amount = ThisAmtP
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 12:
'        If TaxTrans.Amount <> ThisAmtP Then
'          TaxTrans.Amount = ThisAmtP
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'      Case 24:
'        If TaxTrans.Amount <> ThisAmtB Then
'          TaxTrans.Amount = ThisAmtB
'          Put TTHandle, x, TaxTrans
'          RepairCnt = RepairCnt + 1
'        End If
'    End Select
'
'    frmVATaxShowPctComp.ShowPctComp x, NumOfTTRecs
'  Next x
'  Unload frmVATaxShowPctComp
'  EnableCloseButton Me.hwnd, True
'
'  Close
'
'  Call Savemsg(900, CStr(RepairCnt) + " transaction amounts have been updated successfully")
'
'End Sub

Private Sub cmdProcess6_Click()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long, y As Long
  Dim TotRev As Double
  Dim BelongToCnt  As Integer
  Dim TotPrincPd As Double
  Dim TotLateListPd As Double
  Dim TotCollectionPd As Double
  Dim TotInterestPd As Double
  Dim TotPenaltyPd As Double
  Dim PrincBilledPaid As Double
  Dim LateListBilledPaid As Double
  Dim CollectionBilledPaid As Double
  Dim InterestBilledPaid As Double
  Dim PenaltyBilledPaid As Double
  Dim ErrCnt As Long
  
  If TaxMsgWOpts(900, "Run this only after you have already repaired negative values. THIS PROCESS IS VERY LENGTHY.", "F10 Continue", "ESC Quit Now") = "abort" Then
    Close
    Exit Sub
  End If
  OpenTaxTransFile THandle, NumOfTRecs
  ReDim ErrTrans(1 To 1) As Long
  For x = 4998 To NumOfTRecs
    Get THandle, x, TaxTrans
    If TaxTrans.TranType = 1 Then
      If OldRound(TaxTrans.Revenue.Principle1Pd) > OldRound(TaxTrans.Revenue.Principle1) Or OldRound(TaxTrans.Revenue.CollectionPd) > OldRound(TaxTrans.Revenue.Collection) _
      Or OldRound(TaxTrans.Revenue.LateListPd) > OldRound(TaxTrans.Revenue.LateList) Or OldRound(TaxTrans.Revenue.InterestPd) > OldRound(TaxTrans.Revenue.Interest) Or _
      TaxTrans.Revenue.PenaltyPd > TaxTrans.Revenue.Penalty Then
        ErrCnt = ErrCnt + 1
        ReDim Preserve ErrTrans(1 To ErrCnt) As Long
        ErrTrans(ErrCnt) = x
      End If
    End If
  Next x
  
  If ErrCnt = 0 Then
    Close
    Call TaxMsg(900, "There are no billing transactions affected by this error.")
    Exit Sub
  End If
  
  frmVATaxShowPctComp.Label1 = "Making Billing Paid Values Equal Belong To Values"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  For x = 1 To ErrCnt
    Get THandle, ErrTrans(x), TaxTrans
      PrincBilledPaid = TaxTrans.Revenue.Principle1Pd
      LateListBilledPaid = TaxTrans.Revenue.LateListPd
      CollectionBilledPaid = TaxTrans.Revenue.CollectionPd
      InterestBilledPaid = TaxTrans.Revenue.InterestPd
      PenaltyBilledPaid = TaxTrans.Revenue.PenaltyPd
      TotPrincPd = 0
      TotLateListPd = 0
      TotCollectionPd = 0
      TotInterestPd = 0
      TotPenaltyPd = 0
      For y = ErrTrans(x) + 1 To NumOfTRecs 'start at ErrTrans(x) because all trans related to bill
        'will come after the bill trans
        Get THandle, y, TaxTrans
        If TaxTrans.BelongTo = ErrTrans(x) Then
          TotPrincPd = OldRound#(TotPrincPd + TaxTrans.Revenue.Principle1Pd)
          TotLateListPd = OldRound#(TotLateListPd + TaxTrans.Revenue.LateListPd)
          TotCollectionPd = OldRound#(TotCollectionPd + TaxTrans.Revenue.CollectionPd)
          TotInterestPd = OldRound#(TotInterestPd + TaxTrans.Revenue.InterestPd)
          TotPenaltyPd = OldRound#(TotPenaltyPd + TaxTrans.Revenue.PenaltyPd)
        End If
'        frmVATaxShowPctComp.ShowPctComp2 y, NumOfTRecs
      Next y
'      frmVATaxShowPctComp.Label1 = "Making Billing Paid Values Equal Belong To Values"
'      frmVATaxShowPctComp.CmdCancel.Visible = False
'      frmVATaxShowPctComp.Show , Me
'      DoEvents
'      EnableCloseButton Me.hwnd, False
      If TotPrincPd < PrincBilledPaid Then
        Get THandle, ErrTrans(x), TaxTrans
        BelongToCnt = BelongToCnt + 1
        TaxTrans.Revenue.Principle1Pd = TotPrincPd
        Put THandle, ErrTrans(x), TaxTrans
        End If
      If TotLateListPd < LateListBilledPaid Then
        Get THandle, ErrTrans(x), TaxTrans
        BelongToCnt = BelongToCnt + 1
        TaxTrans.Revenue.LateListPd = TotLateListPd
        Put THandle, ErrTrans(x), TaxTrans
      End If
      If TotCollectionPd < CollectionBilledPaid Then
        Get THandle, ErrTrans(x), TaxTrans
        BelongToCnt = BelongToCnt + 1
        TaxTrans.Revenue.CollectionPd = TotCollectionPd
        Put THandle, ErrTrans(x), TaxTrans
      End If
      If TotInterestPd < InterestBilledPaid Then
        Get THandle, ErrTrans(x), TaxTrans
        BelongToCnt = BelongToCnt + 1
        TaxTrans.Revenue.InterestPd = TotInterestPd
        Put THandle, ErrTrans(x), TaxTrans
      End If
      If TotPenaltyPd < PenaltyBilledPaid Then
        Get THandle, ErrTrans(x), TaxTrans
        BelongToCnt = BelongToCnt + 1
        TaxTrans.Revenue.PenaltyPd = TotPenaltyPd
        Put THandle, ErrTrans(x), TaxTrans
      End If
    frmVATaxShowPctComp.ShowPctComp x, ErrCnt
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  
  Call Savemsg(900, CStr(BelongToCnt) + " billing transaction amounts have been updated successfully")

End Sub

Private Sub cmdProcess7_Click()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long, y As Long, z As Integer
  Dim AmtBalance As Double
  Dim RevBalance As Double
  Dim NextRec As Long
  Dim ErrCnt As Long
  Dim BTCnt As Integer
  Dim TotPrincPd As Double
  Dim TotLateListPd As Double
  Dim TotCollectionPd As Double
  Dim TotInterestPd As Double
  Dim TotPenaltyPd As Double
  Dim PrincBilledPaid As Double
  Dim LateListBilledPaid As Double
  Dim CollectionBilledPaid As Double
  Dim InterestBilledPaid As Double
  Dim PenaltyBilledPaid As Double
  Dim RevOpt1 As Double
  Dim RevOpt2 As Double
  Dim RevOpt3 As Double
  Dim Principle1Pd As Double
  Dim Principle2Pd As Double
  Dim Principle3Pd As Double
  Dim Principle4Pd As Double
  Dim Principle5Pd As Double
  Dim LateListPd As Double
  Dim Future1Pd As Double
  Dim Future2Pd As Double
  Dim CollectionPd As Double
  Dim Collection As Double
  Dim InterestPd As Double
  Dim Interest As Double
  Dim RevOpt1Pd As Double
  Dim RevOpt2Pd As Double
  Dim RevOpt3Pd As Double
  Dim Principle1BillPd As Double
  Dim Principle1Bill As Double
  Dim Principle2Bill As Double
  Dim Principle3Bill As Double
  Dim Principle4Bill As Double
  Dim Principle5Bill As Double
  Dim InterestBill As Double
  Dim InterestBillPd As Double
  Dim Future1Bill As Double
  Dim Future2Bill As Double
  Dim Future1BillPd As Double
  Dim Future2BillPd As Double
  Dim CollectionBill As Double
  Dim CollectionBillPd As Double
  Dim LateListBill As Double
  Dim LateListBillPd As Double
  Dim RevOpt1Bill As Double
  Dim RevOpt2Bill As Double
  Dim RevOpt3Bill As Double
  Dim SaveHere As Long
  Dim SaveCnt As Long
  Dim ThisPinNum As Long
  Dim FixNextRec As Long
  Dim SaveFixRec As Long
  Dim WrongCustRec As Long
  Dim NewRec As Long
  
  ReDim ErrCust(1 To 1) As Long
  OpenTaxCustFile CHandle, NumOfCRecs
  OpenTaxTransFile THandle, NumOfTRecs
'  For x = 1 To NumOfTRecs
'    Get THandle, x, TaxTrans
'    If x = 4998 Then Stop
'    If TaxTrans.BelongTo = 4998 Then Stop
'    If TaxTrans.Amount = 13.65 And TaxTrans.BelongTo = 198 Then
'      Stop
'      TaxTrans.CustomerRec = TaxTrans.CustomerRec
'      SaveCnt = SaveCnt + 1
'    End If
'  Next x
  SaveCnt = 0
  frmVATaxShowPctComp.Label1 = "Finding Errant Balances"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  For x = 1 To NumOfCRecs
    x = 1356
    Get CHandle, x, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo SkipMe
    RevBalance = 0
    AmtBalance = GetCustBalance(x, -1)
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get THandle, NextRec, TaxTrans
      If TaxTrans.TranType <> 1 Then GoTo NotThisOne
      RevBalance# = OldRound#(RevBalance# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      RevBalance# = OldRound#(RevBalance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      RevBalance# = OldRound#(RevBalance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      RevBalance# = OldRound#(RevBalance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      RevBalance# = OldRound#(RevBalance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
NotThisOne:
      If TaxTrans.Revenue.CollectionPd < 0 Then TaxTrans.Revenue.CollectionPd = 0
      Put THandle, NextRec, TaxTrans
      NextRec = TaxTrans.LastTrans
    Loop
    If OldRound(RevBalance) <> OldRound(AmtBalance) Then
      ErrCnt = ErrCnt + 1
      ReDim Preserve ErrCust(1 To ErrCnt) As Long
      ErrCust(ErrCnt) = x
'      AmtBalance = GetCustBalance(x, -1)
    End If
SkipMe:
    frmVATaxShowPctComp.ShowPctComp x, NumOfCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
    
  frmVATaxShowPctComp.Label1 = "Examining Transactions and Fixing Problems"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  
  For x = 1 To ErrCnt
    ReDim BelongTo(1 To 1) As Long
    BTCnt = 0
    Get CHandle, ErrCust(x), TaxCust
    
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get THandle, NextRec, TaxTrans
      If TaxTrans.TranType = 1 Then
        BTCnt = BTCnt + 1
        ReDim Preserve BelongTo(1 To BTCnt) As Long
        BelongTo(BTCnt) = NextRec
      End If
      NextRec = TaxTrans.LastTrans
    Loop
    For z = 1 To BTCnt
      Get THandle, BelongTo(z), TaxTrans
      Principle1Bill = OldRound(TaxTrans.Revenue.Principle1)
      InterestBill = OldRound(TaxTrans.Revenue.Interest)
      Future1Bill = OldRound(TaxTrans.Revenue.Future1)
      Future2Bill = OldRound(TaxTrans.Revenue.Future2)
      CollectionBill = OldRound(TaxTrans.Revenue.Collection)
      LateListBill = OldRound(TaxTrans.Revenue.LateList)
      Principle1BillPd = OldRound(TaxTrans.Revenue.Principle1Pd)
      InterestBillPd = OldRound(TaxTrans.Revenue.InterestPd)
      Future1BillPd = OldRound(TaxTrans.Revenue.Future1Pd)
      Future2BillPd = OldRound(TaxTrans.Revenue.Future2Pd)
      CollectionBillPd = OldRound(TaxTrans.Revenue.CollectionPd)
      LateListBillPd = OldRound(TaxTrans.Revenue.LateListPd)
      Principle1Pd = 0
      Future1Pd = 0
      Future2Pd = 0
      InterestPd = 0
      CollectionPd = 0
      LateListPd = 0
      Interest = 0
      Collection = 0
      For y = BelongTo(z) To NumOfTRecs
        Get THandle, y, TaxTrans
        If TaxTrans.BelongTo = BelongTo(z) Then
          Select Case TaxTrans.TranType
            Case 2, 7 'payment, adjustment
                Principle1Pd = OldRound#(Principle1Pd + TaxTrans.Revenue.Principle1Pd) 'collect all revenues and
                'wait for a problem
                If Principle1Pd > Principle1Bill Then 'OK...we have a problem
                  If TaxTrans.Revenue.Principle1 < 0 Then TaxTrans.Revenue.Principle1 = 0
                  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - (Principle1Pd - Principle1Bill))
                  Principle1Pd = OldRound(Principle1Pd - (Principle1Pd - Principle1Bill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - Principle1Pd) 'reduce the existing total amount for
                  'this transaction by the amount of the overpay
                  Put THandle, y, TaxTrans 'save it to this transaction
                  SaveHere = TaxTrans.BelongTo 'establish the bill record
                  Get THandle, TaxTrans.BelongTo, TaxTrans 'pull the bill record
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.Principle1Pd = Principle1Pd 'correct the bill record for this resource
                  'TaxTrans.Amount is not affected because this value only reflects the original charges
                  'and not any updates to the bill
                  Put THandle, SaveHere, TaxTrans 'save it to the bill record
                  Get THandle, y, TaxTrans 'go back to original record
                End If
                Future1Pd = OldRound(Future1Pd + TaxTrans.Revenue.Future1Pd)
                If Future1Pd > Future1Bill Then
                  If TaxTrans.Revenue.Future1 < 0 Then TaxTrans.Revenue.Future1 = 0
                  TaxTrans.Revenue.Future1Pd = OldRound(TaxTrans.Revenue.Future1Pd - (Future1Pd - Future1Bill))
                  Future1Pd = OldRound(Future1Pd - (Future1Pd - Future1Bill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - Future1Pd)
                  Put THandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get THandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.Future1Pd = OldRound(TaxTrans.Revenue.Future1Pd - (Future1Pd - Future1Bill))
                  Put THandle, SaveHere, TaxTrans
                  Get THandle, y, TaxTrans
                End If
                Future2Pd = OldRound(Future2Pd + TaxTrans.Revenue.Future2Pd)
                If Future2Pd > Future2Bill Then
                  If TaxTrans.Revenue.Future2 < 0 Then TaxTrans.Revenue.Future2 = 0
                  TaxTrans.Revenue.Future2Pd = OldRound(TaxTrans.Revenue.Future2Pd - (Future2Pd - Future2Bill))
                  Future2Pd = OldRound(Future2Pd - (Future2Pd - Future2Bill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - Future2Pd)
                  Put THandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get THandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.Future2Pd = OldRound(TaxTrans.Revenue.Future2Pd - (Future2Pd - Future2Bill))
                  Put THandle, SaveHere, TaxTrans
                  Get THandle, y, TaxTrans
                End If
                InterestPd = OldRound(InterestPd + TaxTrans.Revenue.InterestPd)
                If InterestPd > InterestBill Then
                  If TaxTrans.Revenue.Interest < 0 Then TaxTrans.Revenue.Interest = 0
                  TaxTrans.Revenue.InterestPd = OldRound(TaxTrans.Revenue.InterestPd - (InterestPd - InterestBill))
                  InterestPd = OldRound(InterestPd - (InterestPd - InterestBill))
                  TaxTrans.Amount = TaxTrans.Amount - InterestPd
                  Put THandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get THandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.InterestPd = InterestPd
                  Put THandle, SaveHere, TaxTrans
                  Get THandle, y, TaxTrans
                End If
                CollectionPd = OldRound(CollectionPd + TaxTrans.Revenue.CollectionPd)
                If CollectionPd > CollectionBill Then
                  If TaxTrans.Revenue.Collection < 0 Then TaxTrans.Revenue.Collection = 0
                  TaxTrans.Revenue.CollectionPd = OldRound(TaxTrans.Revenue.CollectionPd - (CollectionPd - CollectionBill))
                  CollectionPd = OldRound(CollectionPd - (CollectionPd - CollectionBill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - CollectionPd)
                  Put THandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get THandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.CollectionPd = OldRound(TaxTrans.Revenue.CollectionPd - (CollectionPd - CollectionBill))
                  Put THandle, SaveHere, TaxTrans
                  Get THandle, y, TaxTrans
                End If
                LateListPd = OldRound(LateListPd + TaxTrans.Revenue.LateListPd)
                If LateListPd > LateListBill Then
                  If TaxTrans.Revenue.LateList < 0 Then TaxTrans.Revenue.LateList = 0
                  TaxTrans.Revenue.LateListPd = OldRound(TaxTrans.Revenue.LateListPd - (LateListPd - LateListBill))
                  LateListPd = OldRound(LateListPd - (LateListPd - LateListBill))
                  TaxTrans.Amount = OldRound(TaxTrans.Amount - LateListPd)
                  Put THandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get THandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.LateListPd = OldRound(TaxTrans.Revenue.LateListPd - (LateListPd - LateListBill))
                  Put THandle, SaveHere, TaxTrans
                  Get THandle, y, TaxTrans
                End If
              Case 4:
                Interest = OldRound(Interest + TaxTrans.Revenue.Interest)
                If Interest > InterestBill Then
                  If TaxTrans.Revenue.Interest < 0 Then TaxTrans.Revenue.Interest = 0
                  TaxTrans.Revenue.Interest = OldRound(TaxTrans.Revenue.Interest - (Interest - InterestBill))
                  Interest = OldRound(Interest - (Interest - InterestBill))
                  TaxTrans.Amount = TaxTrans.Revenue.Interest ' OldRound(TaxTrans.Amount - Interest)
                  Put THandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get THandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.InterestPd = InterestBill ' OldRound(TaxTrans.Revenue.InterestPd - (Interest - InterestBill))
                  Put THandle, SaveHere, TaxTrans
                  Get THandle, y, TaxTrans
                End If
              Case 6, 8:
                Collection = OldRound(Collection + TaxTrans.Revenue.Collection)
                If Collection > CollectionBill Then
                  If TaxTrans.Revenue.Collection < 0 Then TaxTrans.Revenue.Collection = 0
                  TaxTrans.Revenue.Collection = OldRound(TaxTrans.Revenue.Collection - (Collection - CollectionBill))
                  Collection = OldRound(Collection - (Collection - CollectionBill))
                  TaxTrans.Amount = TaxTrans.Revenue.Collection 'OldRound(TaxTrans.Amount - Collection)
                  Put THandle, y, TaxTrans
                  SaveHere = TaxTrans.BelongTo
                  Get THandle, TaxTrans.BelongTo, TaxTrans
                  SaveCnt = SaveCnt + 1
                  TaxTrans.Revenue.CollectionPd = CollectionBill 'OldRound(TaxTrans.Revenue.CollectionPd - (Collection - CollectionBill))
                  Put THandle, SaveHere, TaxTrans
                  Get THandle, y, TaxTrans
                End If
            End Select
         End If
NextTrans:
      Next y
      If Principle1BillPd <> Principle1Pd Then
        Get THandle, BelongTo(z), TaxTrans
        TaxTrans.Revenue.Principle1 = Principle1Pd
        Put THandle, BelongTo(z), TaxTrans
        SaveCnt = SaveCnt + 1
      End If
      If LateListBillPd <> LateListPd Then
        Get THandle, BelongTo(z), TaxTrans
        TaxTrans.Revenue.LateList = LateListPd
        Put THandle, BelongTo(z), TaxTrans
        SaveCnt = SaveCnt + 1
      End If
      If CollectionBillPd <> Collection Then
        Get THandle, BelongTo(z), TaxTrans
        TaxTrans.Revenue.Collection = Collection
        Put THandle, BelongTo(z), TaxTrans
        SaveCnt = SaveCnt + 1
      End If
      If InterestBillPd <> InterestBill Then
        Get THandle, BelongTo(z), TaxTrans
        TaxTrans.Revenue.InterestPd = InterestPd
        Put THandle, BelongTo(z), TaxTrans
        SaveCnt = SaveCnt + 1
      End If
    Next z
    frmVATaxShowPctComp.ShowPctComp x, ErrCnt
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  Call Savemsg(900, "Procedure completed successfully.")
End Sub

Private Sub cmdProcess8_Click()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long, y As Long, z As Long, q As Integer
  Dim ErrCnt As Long
  Dim CustCnt As Long
  Dim NextRec As Long
  Dim BillCnt As Long
  Dim SaveHere As Long
  Dim SaveCnt As Long
  Dim ThisPinNum As Long
  Dim FixNextRec As Long
  Dim SaveFixRec As Long
  Dim WrongCustRec As Long
  Dim NewRec As Long
  Dim RightRec As Long
  Dim RightCnt As Integer
  Dim MaxBillCnt As Integer
  Dim UseThisRec As Long
  Dim Fixcnt As Long
    
  OpenTaxCustFile CHandle, NumOfCRecs
  OpenTaxTransFile THandle, NumOfTRecs
  
  ReDim CustList(1 To 1) As Long
  For x = 1 To NumOfCRecs
    Get CHandle, x, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo Deleted
    CustCnt = CustCnt + 1
    ReDim Preserve CustList(1 To CustCnt) As Long
    CustList(CustCnt) = x
Deleted:
  Next x
  
'  ReDim CustList(1 To 2) As Long
'  CustList(1) = 66
'  CustList(2) = 5899
'  CustCnt = 2

  frmVATaxShowPctComp.Label1 = "Building Arrays Of Customer Transactions"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  
  ReDim CustBillCnt(1 To CustCnt) As Integer
  MaxBillCnt = 0
  For x = 1 To CustCnt
    Get CHandle, CustList(x), TaxCust
    BillCnt = 0
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get THandle, NextRec, TaxTrans
      If TaxTrans.TranType = 1 Then
        BillCnt = BillCnt + 1
      End If
      NextRec = TaxTrans.LastTrans
    Loop
    If BillCnt > MaxBillCnt Then
      MaxBillCnt = BillCnt
    End If
    frmVATaxShowPctComp.ShowPctComp x, CustCnt
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  ReDim CustBillQ(1 To CustCnt, 1 To MaxBillCnt) As Long
 
  frmVATaxShowPctComp.Label1 = "Building Arrays Of Customer Transactions"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  For x = 1 To CustCnt
    Get CHandle, CustList(x), TaxCust
    BillCnt = 0
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get THandle, NextRec, TaxTrans
      If TaxTrans.TranType = 1 Then
        BillCnt = BillCnt + 1
        CustBillQ(x, BillCnt) = NextRec
      End If
      NextRec = TaxTrans.LastTrans
    Loop
    If BillCnt = 0 Then
      CustBillQ(x, 1) = -1
      BillCnt = 1
    End If
    CustBillCnt(x) = BillCnt
    frmVATaxShowPctComp.ShowPctComp x, CustCnt
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  frmVATaxShowPctComp.Label1 = "Repairing Orphan Transactions (Final Procedure)"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  DoEvents
  For x = 1 To CustCnt 'looking for transactions that do not belong in this Queue
    Get CHandle, CustList(x), TaxCust
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get THandle, NextRec, TaxTrans
      UseThisRec = TaxTrans.LastTrans
      If TaxTrans.TranType <> 1 Then 'type 1 trans are in the CustBillQ array
        For y = 1 To CustBillCnt(x)
          If CustBillQ(x, y) < 0 Then GoTo NewLoop
          If TaxTrans.BelongTo = CustBillQ(x, y) Then Exit For 'match em up by
          'making sure that each non 1 trans matches up with a 1 trans for this customer
NewLoop:
        Next y
        If y > CustBillCnt(x) Then 'NextRec is an orphan
          GoSub GetRightRec 'go find who it belongs to
          If RightCnt <> 1 Then
            GoTo NoRightRec
          End If
          Fixcnt = Fixcnt + 1
          Get CHandle, CustList(x), TaxCust 'get the original cust
          FixNextRec = NextRec 'remove this trans from wrong cust Q
          Get THandle, FixNextRec, TaxTrans
          If FixNextRec = TaxCust.LastTrans Then
            TaxCust.LastTrans = TaxTrans.LastTrans 'links past the bad trans
            Put CHandle, CustList(x), TaxCust
          Else
            Get THandle, NextRec, TaxTrans
            SaveHere = TaxTrans.LastTrans
            Get THandle, SaveFixRec, TaxTrans
            TaxTrans.LastTrans = SaveHere 'otherwise it links up with next trans
            Put THandle, SaveFixRec, TaxTrans
          End If
          Get CHandle, RightRec, TaxCust 'now go to correct cust and add the
          'trans posted in error to this customer's Q
          NewRec = TaxCust.LastTrans 'save this to add to the changing trans
          TaxCust.LastTrans = NextRec 'add the new link to the top of the Q
          Put CHandle, RightRec, TaxCust 'save it
          Get THandle, NextRec, TaxTrans 'go back to the trans that is changing Q's
          TaxTrans.LastTrans = NewRec 'insert new link
          Put THandle, NextRec, TaxTrans 'save
          Get CHandle, CustList(x), TaxCust 'return to original customer
        End If
      End If
NoRightRec:
      SaveFixRec = NextRec 'save this rec for linking later on if necessary
      NextRec = UseThisRec
    Loop
    frmVATaxShowPctComp.ShowPctComp x, CustCnt
  Next x
  
  Close
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Call Savemsg(800, "A total of " + CStr(Fixcnt) + " orphan transactions were fixed successfully.")
  
  Exit Sub
  
GetRightRec:
  RightCnt = 0
  For z = 1 To CustCnt
    Get CHandle, CustList(z), TaxCust
    For q = 1 To CustBillCnt(z)
      If CustBillQ(z, q) < 0 Then GoTo AnotherLoop
      If CustBillQ(z, q) = TaxTrans.BelongTo Then
        RightRec = CustList(z)
        RightCnt = RightCnt + 1
      End If
AnotherLoop:
    Next q
  Next z
  
  Return
  
  End Sub
  
Private Sub FixSpecificData()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long, NextRec As Long
  Dim ThisAmt As Double
  Dim Princ As Double
  Dim Adv As Double
  Dim LateList As Double
  Dim Interest As Double
  Dim Opt1 As Double
  Dim Opt2 As Double
  Dim Opt3 As Double
  Dim GetBill As Long
  Dim ThisDate$
  Dim SaveAmt As Double
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim SCnt As Long
  Dim OCnt As Long
  
'  OpenRealPropFile RHandle, NumOfRRecs
'  For x = 1 To NumOfRRecs
'    Get RHandle, x, RealRec
'    If RealRec.EXMPOTHR > 0 Then OCnt = OCnt + 1
'    If RealRec.EXMPSENI > 0 Then SCnt = SCnt + 1
'    If RealRec.EXMPOTHR > 0 And RealRec.EXMPSENI > 0 Then Stop
'    RealRec.EXMPSENI = 0
'    RealRec.EXMPOTHR = 0
'    Put RHandle, x, RealRec
'  Next x
'  Close RHandle
    
  OpenPersPropFile PHandle, NumOfPRecs
  For x = 1 To NumOfPRecs
    Get PHandle, x, PersRec
    
'      If PersRec.EXMPOTHR > 0 And PersRec.EXMPSENI > 0 Then Stop
'      If PersRec.EXMPOTHR > 0 Or PersRec.EXMPSENI > 0 Then Cnt = Cnt + 1
''    If PersRec.Deleted = False Then
''      PersRec.TaxBillYear = 2006
''      Put PHandle, x, PersRec
''    End If
  Next x
'  Close PHandle
'  OpenTaxCustFile CHandle, NumOfCRecs
'  For x = 1 To NumOfCRecs
'    Get CHandle, x, TaxCust
'    If TaxCust.CountyAcct = 55459 Or TaxCust.CountyAcctString = "55459" Then
'      TaxCust.CustName = TaxCust.CustName
'      Stop
'    End If
'  Next x
'  OpenTaxTransFile THandle, NumOfTRecs
'  For x = 1 To NumOfTRecs
''    If x = 2016 Then Stop
'    Get THandle, x, TaxTrans
'    If TaxTrans.BelongTo = 16892 And TaxTrans.TranType = 2 Then Stop
'    TaxTrans.Amount = TaxTrans.Amount
'    TaxTrans.TranType = TaxTrans.TranType
'    TaxTrans.CustomerRec = TaxTrans.CustomerRec
'    If TaxTrans.BillType = "C" Then
'      Get CHandle, TaxTrans.CustomerRec, TaxCust
'      If TaxCust.FirstPersRec > 0 Then
'        TaxTrans.BillType = "P"
'        Put THandle, x, TaxTrans
'      ElseIf TaxCust.FirstPropRec > 0 Then
'        TaxTrans.BillType = "R"
'        Put THandle, x, TaxTrans
'      Else
''        Stop
'        TaxTrans.BillType = ""
'        Put THandle, x, TaxTrans
'      End If
'    End If
''    ThisDate$ = MakeRegDate(TaxTrans.TransDate)
''   -----------Fix for Harrisburg's 1/5/06 errors---------
''    If x = 8760 Then
''      SaveAmt = TaxTrans.Amount
''      TaxTrans.BelongTo = 2264
''      TaxTrans.Description = "2264"
''      Put THandle, x, TaxTrans
''      Get THandle, 3190, TaxTrans
''      TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - SaveAmt)
''      Put THandle, 3190, TaxTrans
''      Get THandle, 2264, TaxTrans
''      TaxTrans.Revenue.Principle1Pd = SaveAmt
''      Put THandle, 2264, TaxTrans
''    End If
''
''    If x = 8752 Then
''      SaveAmt = TaxTrans.Amount
''      TaxTrans.BelongTo = 2303
''      TaxTrans.Description = "2303"
''      Put THandle, x, TaxTrans
''      Get THandle, 3744, TaxTrans
''      TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - SaveAmt)
''      Put THandle, 3744, TaxTrans
''      Get THandle, 2303, TaxTrans
''      TaxTrans.Revenue.Principle1Pd = SaveAmt
''      Put THandle, 2303, TaxTrans
''    End If
''
''    If x = 8744 Then
''      SaveAmt = TaxTrans.Amount
''      TaxTrans.BelongTo = 3925
''      TaxTrans.Description = "3925"
''      Put THandle, x, TaxTrans
''      Get THandle, 3531, TaxTrans
''      TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - SaveAmt)
''      Put THandle, 3531, TaxTrans
''      Get THandle, 3925, TaxTrans
''      TaxTrans.Revenue.Principle1Pd = SaveAmt
''      Put THandle, 3925, TaxTrans
''    End If
''
''    If x = 8712 Then
''      SaveAmt = TaxTrans.Amount
''      TaxTrans.BelongTo = 457
''      TaxTrans.Description = "457"
''      Put THandle, x, TaxTrans
''      Get THandle, 587, TaxTrans
''      TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - SaveAmt)
''      Put THandle, 587, TaxTrans
''      Get THandle, 457, TaxTrans
''      TaxTrans.Revenue.Principle1Pd = SaveAmt
''      Put THandle, 457, TaxTrans
''    End If
''
''    If x = 8716 Then
''      SaveAmt = TaxTrans.Amount
''      TaxTrans.BelongTo = 3895
''      TaxTrans.Description = "3895"
''      Put THandle, x, TaxTrans
''      Get THandle, 3355, TaxTrans
''      TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd - SaveAmt)
''      Put THandle, 3355, TaxTrans
''      Get THandle, 3895, TaxTrans
''      TaxTrans.Revenue.Principle1Pd = SaveAmt
''      Put THandle, 3895, TaxTrans
''    End If
''
''    Get CHandle, 3578, TaxCust
''    NextRec = TaxCust.LastTrans
''    Do While NextRec > 0
''      Get THandle, NextRec, TaxTrans
''      If NextRec = 8883 Then
''        GetBill = TaxTrans.BelongTo
''        TaxTrans.BelongTo = 1841 'change from 453 to 1881
''        ThisAmt = TaxTrans.Amount
''        Princ = TaxTrans.Revenue.Principle1Pd
''        Adv = TaxTrans.Revenue.CollectionPd
''        LateList = TaxTrans.Revenue.LateListPd
''        Interest = TaxTrans.Revenue.InterestPd
''        Opt1 = TaxTrans.Revenue.RevOpt1Pd
''        Opt2 = TaxTrans.Revenue.RevOpt2Pd
''        Opt3 = TaxTrans.Revenue.RevOpt3Pd
''        TaxTrans.Description = "1049 1841"
''        Put THandle, 8883, TaxTrans
''        Get THandle, GetBill, TaxTrans
''        TaxTrans.Revenue.Principle1Pd = OldRound#(TaxTrans.Revenue.Principle1Pd - Princ)
''        TaxTrans.Revenue.CollectionPd = OldRound#(TaxTrans.Revenue.CollectionPd - Adv)
''        TaxTrans.Revenue.LateListPd = OldRound#(TaxTrans.Revenue.LateListPd - LateList)
''        TaxTrans.Revenue.InterestPd = OldRound#(TaxTrans.Revenue.InterestPd - Interest)
''        TaxTrans.Revenue.RevOpt1Pd = OldRound#(TaxTrans.Revenue.RevOpt1Pd - Opt1)
''        TaxTrans.Revenue.RevOpt2Pd = OldRound#(TaxTrans.Revenue.RevOpt2Pd - Opt2)
''        TaxTrans.Revenue.RevOpt3Pd = OldRound#(TaxTrans.Revenue.RevOpt3Pd - Opt3)
''        Put THandle, GetBill, TaxTrans
''        Get THandle, 1841, TaxTrans
''        TaxTrans.Revenue.Principle1Pd = Princ
''        TaxTrans.Revenue.CollectionPd = Adv
''        TaxTrans.Revenue.LateListPd = LateList
''        TaxTrans.Revenue.InterestPd = Interest
''        TaxTrans.Revenue.RevOpt1Pd = Opt1
''        TaxTrans.Revenue.RevOpt2Pd = Opt2
''        TaxTrans.Revenue.RevOpt3Pd = Opt3
''        Put THandle, 1841, TaxTrans
''        Exit Do
''      End If
''      NextRec = TaxTrans.LastTrans
''    Loop
''   ^^^^^^^^^^^Fix for Harrisburg's 1/5/06 errors^^^^^^^^^^^^
'
''    If TaxTrans.BelongTo = 3190 Then Stop
''    TaxTrans.TranType = TaxTrans.TranType
''    TaxTrans.Amount = TaxTrans.Amount
''    TaxTrans.CustomerRec = TaxTrans.CustomerRec
''    fix for sunset beach
''    If x = 11485 Or x = 11337 Or x = 11186 Or x = 11017 Or x = 10496 Or x = 8944 Then
''      SaveAmt = TaxTrans.Amount
''      Get THandle, 8272, TaxTrans
''      TaxTrans.Revenue.Interest = OldRound(TaxTrans.Revenue.Interest + SaveAmt)
''      Put THandle, 8272, TaxTrans
''    End If
''    If x = 11484 Or x = 11336 Or x = 11185 Or x = 11016 Or x = 10495 Or x = 8938 Then
''      SaveAmt = TaxTrans.Amount
''      Get THandle, 8269, TaxTrans
''      TaxTrans.Revenue.Interest = OldRound(TaxTrans.Revenue.Interest + SaveAmt)
''      Put THandle, 8269, TaxTrans
''    End If
''    If x = 11491 Or x = 11343 Or x = 11192 Or x = 11023 Or x = 10502 Or x = 8966 Then
''      SaveAmt = TaxTrans.Amount
''      Get THandle, 8285, TaxTrans
''      TaxTrans.Revenue.Interest = OldRound(TaxTrans.Revenue.Interest + SaveAmt)
''      Put THandle, 8285, TaxTrans
''    End If
'  Next x
'
'  Close CHandle
'  Close THandle

End Sub

Private Sub cmdProcess5_Click()
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NewTaxTrans As TaxTransactionType
  Dim NewTHandle As Integer
  Dim NumOfTRecs As Long
  Dim NextRec As Long
  Dim Balance As Double
  Dim TransCnt As Long
  Dim FirstTrans As Boolean
  Dim TotCustBal As Double
  Dim BHandle As Integer
  Dim CnvtDate As Integer
  Dim CDateStr$
  Dim LastTrans As Long
  Dim TotAmt As Double
  
  If Exist("cnvtdate.dat") Then
    BHandle = FreeFile
    Open "cnvtdate.dat" For Input As BHandle
    Input #BHandle, CDateStr$
    Close BHandle
    CDateStr$ = MakeRegDate(CInt(CDateStr$))
    CnvtDate = Date2Num(CDateStr$)
  Else
    CnvtDate = 0
  End If

  OpenTaxCustFile CHandle, NumOfCRecs
  OpenTaxTransFile THandle, NumOfTRecs
  OpenNewTaxTransFile NewTHandle
  frmVATaxShowPctComp.Label1 = "Clearing Negative Customer Balances"
  frmVATaxShowPctComp.Show
  frmVATaxMainMenu.cmdExit.Enabled = False
  For x = 1 To NumOfCRecs
    Get CHandle, x, TaxCust
    FirstTrans = True
    TotCustBal = 0
    TotAmt = 0
    If TaxCust.Deleted <> 0 Then
      TaxCust.LastTrans = 0
      Put CHandle, x, TaxCust
      GoTo Skip
    End If
    TaxCust.CustName = TaxCust.CustName
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get THandle, NextRec, TaxTrans
      If TaxTrans.TranType = 22 Then 'prepaid only is treated as a stand alone transaction
        Balance# = TaxTrans.Revenue.PrePaidAmt
        GoTo KeepInQ
      End If
      If TaxTrans.TranType = 21 Then 'prepaid with billpay is changed and saved to 22 and then treated
        'as a stand alone transaction
        TaxTrans.TranType = 22
        Balance# = TaxTrans.Revenue.PrePaidAmt
        GoTo KeepInQ
      End If
      If TaxTrans.TranType = 1 Then
        Balance# = OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
        Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
        Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
        Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
        If Balance <= 0 Then GoTo LoopAgain
KeepInQ:
        TotAmt = OldRound(TotAmt + TaxTrans.Amount)
        NewTaxTrans.PPTRADisc = TaxTrans.PPTRADisc
        NewTaxTrans.TransDate = TaxTrans.TransDate 'Date2Num(Date)
        NewTaxTrans.TaxYear = TaxTrans.TaxYear
        NewTaxTrans.BillType = TaxTrans.BillType 'TaxTrans.BillType           'R=Real P=Personal Property C=Combined (NC/GA)
        NewTaxTrans.TranType = TaxTrans.TranType                '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing
'        If TaxTrans.TransDate <= CnvtDate Then
          NewTaxTrans.Amount = Balance# ' OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1)
'        Else
'          If TaxTrans.TranType = 22 Then
'            NewTaxTrans.Amount = Balance
'          Else
'            NewTaxTrans.Amount = TaxTrans.Amount
'          End If
'        End If
        NewTaxTrans.Revenue.Principle1 = TaxTrans.Revenue.Principle1
        NewTaxTrans.Revenue.Principle2 = TaxTrans.Revenue.Principle2
        NewTaxTrans.Revenue.Principle3 = TaxTrans.Revenue.Principle3
        NewTaxTrans.Revenue.Principle4 = TaxTrans.Revenue.Principle4
        NewTaxTrans.Revenue.Principle5 = TaxTrans.Revenue.Principle5
        NewTaxTrans.Revenue.Interest = TaxTrans.Revenue.Interest
        NewTaxTrans.Revenue.LateList = TaxTrans.Revenue.LateList
        NewTaxTrans.Revenue.RevOpt1 = TaxTrans.Revenue.RevOpt1
        NewTaxTrans.Revenue.RevOpt2 = TaxTrans.Revenue.RevOpt2
        NewTaxTrans.Revenue.RevOpt3 = TaxTrans.Revenue.RevOpt3
        NewTaxTrans.Revenue.Penalty = TaxTrans.Revenue.Penalty
        NewTaxTrans.Revenue.Collection = TaxTrans.Revenue.Collection
        NewTaxTrans.Revenue.Future1 = 0
        NewTaxTrans.Revenue.Future2 = 0
        NewTaxTrans.Revenue.PrePaidAmt = TaxTrans.Revenue.PrePaidAmt
        NewTaxTrans.Revenue.PrePaidUsed = TaxTrans.Revenue.PrePaidUsed
        NewTaxTrans.Revenue.PrePaidBal = TaxTrans.Revenue.PrePaidBal
        If TaxTrans.TranType = 22 Then
          NewTaxTrans.Revenue.Principle1Pd = 0
        Else
          NewTaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd
        End If
        NewTaxTrans.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2Pd
        NewTaxTrans.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3Pd
        NewTaxTrans.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4Pd
        NewTaxTrans.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5Pd
        NewTaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd
        NewTaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd
        NewTaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.CollectionPd
        NewTaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateListPd
        NewTaxTrans.Revenue.Future1Pd = 0
        NewTaxTrans.Revenue.Future2Pd = 0
        NewTaxTrans.Revenue.RevOpt1Pd = TaxTrans.Revenue.RevOpt1Pd
        NewTaxTrans.Revenue.RevOpt2Pd = TaxTrans.Revenue.RevOpt2Pd
        NewTaxTrans.Revenue.RevOpt3Pd = TaxTrans.Revenue.RevOpt3Pd
        NewTaxTrans.Revenue.pad = ""
        NewTaxTrans.DiscXDate = TaxTrans.DiscXDate
        NewTaxTrans.DiscAmt = TaxTrans.DiscAmt
        NewTaxTrans.OperNum = TaxTrans.OperNum
        NewTaxTrans.InternalPin = TaxTrans.InternalPin
        NewTaxTrans.RealPin = TaxTrans.RealPin
        NewTaxTrans.InternalPin = TaxTrans.InternalPin
        NewTaxTrans.PersPin = TaxTrans.PersPin
        NewTaxTrans.FromPrePay = QPTrim$(TaxTrans.FromPrePay)
        If TaxTrans.TransDate <= CnvtDate Then
          NewTaxTrans.Description = "Initialize " + CStr(TransCnt + 1)
        Else
          NewTaxTrans.Description = TaxTrans.Description
        End If
        NewTaxTrans.Altered = TaxTrans.Altered
        NewTaxTrans.Posted2GL = TaxTrans.Posted2GL
        NewTaxTrans.CustomerRec = x
        NewTaxTrans.BelongTo = TaxTrans.BelongTo
        NewTaxTrans.Padding = ""
        NewTaxTrans.CntyPara = TaxTrans.CntyPara
        NewTaxTrans.CustPin = x
        NewTaxTrans.DMVSubmitted = TaxTrans.DMVSubmitted
        NewTaxTrans.DMVBatch = TaxTrans.DMVBatch
        NewTaxTrans.TShpPara = TaxTrans.TShpPara
        If FirstTrans = True Then
          FirstTrans = False
          NewTaxTrans.LastTrans = 0
        Else
          NewTaxTrans.LastTrans = TransCnt
        End If
        TransCnt = TransCnt + 1
        Put NewTHandle, TransCnt, NewTaxTrans
      End If
LoopAgain:
      NextRec = TaxTrans.LastTrans
    Loop
    If TotAmt = 0 Then
      TaxCust.LastTrans = 0
      Put CHandle, x, TaxCust
    Else
      TaxCust.LastTrans = TransCnt
      Put CHandle, x, TaxCust
    End If
Skip:
    frmVATaxShowPctComp.ShowPctComp x, NumOfCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      frmVATaxMainMenu.cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  frmVATaxMainMenu.cmdExit.Enabled = True
  Close
  
  KillFile App.Path + "\OLDWINTAXTRANS.DAT"
  Name App.Path + "\TAXTRANS.DAT" As App.Path + "\OLDWINTAXTRANS.DAT"
  Name App.Path + "\NEWTAXTRANS.DAT" As App.Path + "\TAXTRANS.DAT"
'  Call MakeAmtEqualRevsAfterClearingNegs
  Call TaxMsg(800, "Balance update has completed successfully. The old transaction file is now saved as 'OLDWINTAXTRANS.DAT'.")
  
  Exit Sub
  
CheckBillBal:

  Balance# = OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
  Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
  Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
  Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
  Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))

  Return

End Sub

Private Sub MakeAmtEqualRevs()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  Dim TotRev As Double
  
  If Exist("negscleared.dat") Then
    Call TaxMsg(900, "This procedure has already been successfully executed.")
    Exit Sub
  End If
  
  ReDim TypeCnt(1 To 17) As Long
  
  frmVATaxShowPctComp.Label1 = "Making Amounts Equal Revenues"
  frmVATaxShowPctComp.Show
  frmVATaxMainMenu.cmdExit.Enabled = False
  OpenTaxTransFile THandle, NumOfTRecs
  For x = 1 To NumOfTRecs
    Get THandle, x, TaxTrans
'    If x = 87 Then Stop
    Select Case TaxTrans.TranType
      Case 1 'billing
        TotRev = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        TotRev = OldRound#(TotRev + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4)
        TotRev = OldRound#(TotRev + TaxTrans.Revenue.Principle5)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(1) = TypeCnt(1) + 1
        End If
        TotRev = 0
      Case 2 'payment
        TotRev = OldRound(TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.CollectionPd)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd + TaxTrans.Revenue.InterestPd)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(2) = TypeCnt(2) + 1
        End If
        TotRev = 0
      Case 3 'release
        '7/12/06 revenues changed from charges to paid
        TotRev = OldRound(TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.CollectionPd)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd + TaxTrans.Revenue.InterestPd)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(3) = TypeCnt(3) + 1
        End If
        TotRev = 0
      Case 4 'Interest
        TotRev = TaxTrans.Revenue.Interest
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(4) = TypeCnt(4) + 1
        End If
        TotRev = 0
      Case 5 'Penalty
        TotRev = TaxTrans.Revenue.Penalty
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(5) = TypeCnt(5) + 1
        End If
        TotRev = 0
      Case 6 'Advertising
        TotRev = TaxTrans.Revenue.Collection
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(6) = TypeCnt(6) + 1
        End If
        TotRev = 0
      Case 7 'adjust pay down
        TotRev = OldRound(TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.CollectionPd)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd + TaxTrans.Revenue.InterestPd)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(7) = TypeCnt(7) + 1
        End If
        TotRev = 0
      Case 9 'credit applied at billing
        TotRev = TaxTrans.Revenue.PrePaidUsed
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(8) = TypeCnt(8) + 1
        End If
        TotRev = 0
      Case 13 'adjust bill down
        TotRev = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Collection)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.Interest)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(9) = TypeCnt(9) + 1
        End If
        TotRev = 0
      Case 14 'adjust bill up
        TotRev = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Collection)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.Interest)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(10) = TypeCnt(10) + 1
        End If
        TotRev = 0
      Case 21 'bill pay/overpay
        TotRev = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Collection)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.Interest)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(11) = TypeCnt(11) + 1
        End If
        TotRev = 0
      Case 22
        'not needed
      Case 24 'adjust bill up affecting credit
        TotRev = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Collection)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.Interest)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(12) = TypeCnt(12) + 1
        End If
        TotRev = 0
      Case 10 'adjust bill down affecting credit
        TotRev = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Collection)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.Interest)
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(13) = TypeCnt(13) + 1
        End If
        TotRev = 0
      Case 11 'adjust prepay down
        'not needed
      Case 12 'refund prepay
        'not needed
      Case Else
    End Select
    frmVATaxShowPctComp.ShowPctComp x, NumOfTRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      frmVATaxMainMenu.cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
      
  Unload frmVATaxShowPctComp
  frmVATaxMainMenu.cmdExit.Enabled = True
  Close
  For x = 1 To 17
    If TypeCnt(x) > 0 Then
      Exit For
    End If
  Next x
  If x <= 17 Then
    frmVATaxAmtToRevsList.Show vbModal
  Else
    Call TaxMsg(900, "There were no transactions that needed this repair.")
  End If
  
End Sub

Private Sub MakeAmtEqualRevsAfterClearingNegs()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  Dim TotRev As Double
  Dim One As Integer
  Dim ThisFile As Integer
  Dim FileName$
  
  ReDim TypeCnt(1 To 17) As Long
  
  frmVATaxShowPctComp.Label1 = "Making Amounts Equal Revenues"
  frmVATaxShowPctComp.Show
  frmVATaxMainMenu.cmdExit.Enabled = False
  OpenTaxTransFile THandle, NumOfTRecs
  For x = 1 To NumOfTRecs
    Get THandle, x, TaxTrans
    If TaxTrans.TranType <> 1 Then
      Call TaxMsg(900, "ERROR: There is a transaction type that is not a billing transaction. Please call Southern Software at 1-800-842-8190 for assistance.")
      Exit For
    End If
  Next x
  If x <= NumOfTRecs Then
    Close
    Exit Sub
  End If
  
  For x = 1 To NumOfTRecs
    Get THandle, x, TaxTrans
    Select Case TaxTrans.TranType
      Case 1 'billing
        TotRev = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.LateList)
        TotRev = OldRound(TotRev + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        TotRev = OldRound#(TotRev + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4)
        TotRev = OldRound#(TotRev + TaxTrans.Revenue.Principle5)
        TotRev = OldRound(TotRev - OldRound(TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.LateListPd))
        TotRev = OldRound(TotRev - OldRound(TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd))
        TotRev = OldRound#(TotRev - OldRound(TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd))
        TotRev = OldRound#(TotRev - OldRound(TaxTrans.Revenue.Principle5Pd))
        If TaxTrans.Amount <> TotRev Then
          TaxTrans.Amount = TotRev
          Put THandle, x, TaxTrans
          TypeCnt(1) = TypeCnt(1) + 1
        End If
        TotRev = 0
      Case Else
    End Select
    frmVATaxShowPctComp.ShowPctComp x, NumOfTRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      frmVATaxMainMenu.cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
      
  Unload frmVATaxShowPctComp
  frmVATaxMainMenu.cmdExit.Enabled = True
  Close
  For x = 1 To 17
    If TypeCnt(x) > 0 Then
      Exit For
    End If
  Next x
  If x <= 17 Then
    frmVATaxAmtToRevsList.Show vbModal
  Else
    Call TaxMsg(900, "There were no transactions that needed this repair.")
  End If
  
  FileName = "negscleared.dat"
  ThisFile = FreeFile
  Open FileName For Output As ThisFile
  One = 1
  Print #ThisFile, One
  Close ThisFile
  
End Sub

Private Sub FixWarsawVA()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  
  OpenTaxTransFile THandle, NumOfTRecs
  'fix for cust# 1290
  Get THandle, 237, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 80.66
  Put THandle, 237, TaxTrans
  
  'fix for 347
  Get THandle, 1871, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 9.72
  Put THandle, 1871, TaxTrans
  
  'fix for 1855
  Get THandle, 943, TaxTrans
  TaxTrans.Revenue.PenaltyPd = 4.97
  Put THandle, 943, TaxTrans
  
  Get THandle, 944, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Penalty = 0
  Put THandle, 944, TaxTrans
  
  Get THandle, 15, TaxTrans
  TaxTrans.Revenue.Principle1 = 54.64
  TaxTrans.Revenue.Principle1Pd = 54.64
  TaxTrans.Revenue.Penalty = 0
  TaxTrans.Revenue.PenaltyPd = 0
  Put THandle, 15, TaxTrans
  
  Get THandle, 2288, TaxTrans
  TaxTrans.Amount = 0
  TaxTrans.Revenue.Principle1Pd = 0
  Put THandle, 2288, TaxTrans
  
  'fix for cust# 1341
  Get THandle, 335, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 2.4
  Put THandle, 335, TaxTrans
  
  Close THandle
  
  Call TaxMsg(900, "Completed successfully.")
End Sub

Private Sub FixAlberta()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  
  OpenTaxTransFile THandle, NumOfTRecs
  'fix for cust# 758
  Get THandle, 10327, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 534
  Put THandle, 10327, TaxTrans
  
  Close THandle
End Sub

Private Sub FixPocohantas()
  Dim TaxTrans As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim x As Long
  Dim Sum#
  
  Exit Sub
  OpenTaxTransFile THandle, NumOfTRecs
  For x = 1 To NumOfTRecs
    Get THandle, x, TaxTrans
    If TaxTrans.TranType = 1 Then
      Sum# = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
      Sum# = OldRound(Sum# + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.Collection)
      Sum# = OldRound(Sum# + TaxTrans.Revenue.Future1 + TaxTrans.Revenue.Future2 + TaxTrans.Revenue.Interest)
      Sum# = OldRound(Sum# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.RevOpt1)
      Sum# = OldRound(Sum# + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      If TaxTrans.Amount < 0.05 And Sum# = 0 Then
        TaxTrans.Revenue.Principle1 = TaxTrans.Amount
        Put THandle, x, TaxTrans
      End If
    End If
  Next x
  'fix for #286
'  Get THandle, 8066, TaxTrans
'  TaxTrans.Revenue.Principle1 = 0.01
'  Put THandle, 8066, TaxTrans
'  'fix for #98
'  Get THandle, 2569, TaxTrans
'  TaxTrans.Revenue.Principle1 = 0.01
'  TaxTrans.Amount = TaxTrans.Amount
'  Put THandle, 2569, TaxTrans
'  Get THandle, 4370, TaxTrans
'  TaxTrans.Revenue.Principle1 = 0.01
'  TaxTrans.Amount = TaxTrans.Amount
'  Put THandle, 4370, TaxTrans
'  Get THandle, 6224, TaxTrans
'  TaxTrans.Revenue.Principle1 = 0.01
'  TaxTrans.Amount = TaxTrans.Amount
'  Put THandle, 6224, TaxTrans
'  Get THandle, 8094, TaxTrans
'  TaxTrans.Amount = TaxTrans.Amount
'  TaxTrans.Revenue.Principle1 = 0.01
'  Put THandle, 8094, TaxTrans
'
'  'for #104
'  Get THandle, 8103, TaxTrans
'  TaxTrans.Revenue.Principle1 = 0.01
'  TaxTrans.Amount = TaxTrans.Amount
'  Put THandle, 8103, TaxTrans
'  Get THandle, 6232, TaxTrans
'  TaxTrans.Revenue.Principle1 = 0.01
'  TaxTrans.Amount = TaxTrans.Amount
'  Put THandle, 6232, TaxTrans
'  Get THandle, 4379, TaxTrans
'  TaxTrans.Revenue.Principle1 = 0.01
'  TaxTrans.Amount = TaxTrans.Amount
'  Put THandle, 4379, TaxTrans
'  Get THandle, 2573, TaxTrans
'  TaxTrans.Amount = TaxTrans.Amount
'  TaxTrans.Revenue.Principle1 = 0.01
'  Put THandle, 2573, TaxTrans
'
'  'for #57
'  Get THandle, 6182, TaxTrans
'  TaxTrans.Revenue.Principle1 = 0.01
'  TaxTrans.Amount = TaxTrans.Amount
'  Put THandle, 6182, TaxTrans
'  Get THandle, 4328, TaxTrans
'  TaxTrans.Revenue.Principle1 = 0.01
'  TaxTrans.Amount = TaxTrans.Amount
'  Put THandle, 4328, TaxTrans
'  Get THandle, 2535, TaxTrans
'  TaxTrans.Revenue.Principle1 = 0.01
'  TaxTrans.Amount = TaxTrans.Amount
'  Put THandle, 2535, TaxTrans
'
  Close
End Sub

Private Sub ForPatrick()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, 1, TaxCust
  TaxCust.CountyAcctString = "1"
  Put TCHandle, 1, TaxCust
  Close
End Sub

Private Sub FixPenGap()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, 74, TaxCust
  TaxCust.Addr1 = "C/O WSWV RADIO STATION"
  TaxCust.Addr2 = "PO BOX 630"
  TaxCust.City = "PENNINGTON GAP"
  TaxCust.State = "VA"
  TaxCust.Zip = "24277"
  TaxCust.CustName = "IBS COMMUNICATIONS LLC"
  TaxCust.SName = "IBS COMMUN"
  TaxCust.CSSN = "384"
  TaxCust.Active = "Y"
  TaxCust.TaxExempt = "N"
  TaxCust.Interest = "Y"
  TaxCust.LateNotice = "Y"
  TaxCust.Penalty = "Y"
  TaxCust.Bankrupt = "N"
  TaxCust.LastTrans = 53040
  TaxCust.CountyAcctString = "7117130"
  Put TCHandle, 74, TaxCust
  
  Close
  Call TaxMsg(900, "Finished.")
End Sub

Private Sub cmdFixVictoria_Click()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim ThisBill$
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim NextRec As Long
  Dim BelongTo As Long
  Dim Fixcnt As Long
  Dim ThisAmt As Double
  Dim ThisCustRec As Long
  Dim IntCnt As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  frmVATaxShowPctComp.Label1 = "Fixing Victoria"
  frmVATaxShowPctComp.Show
  For x = 1 To NumOfTTRecs
  Get TTHandle, x, TaxTrans
    If TaxTrans.TranType = 1 Then
      ThisAmt = 0
      ThisBill = ParseBillNum(TaxTrans.Description)
      If ThisBill = "0" Then ThisBill = CStr(x)
      ThisBill = "Initialize Bill #" + ThisBill
      TaxTrans.Description = ThisBill
      Put TTHandle, x, TaxTrans
      If TaxTrans.BillType = "P" Then
        If TaxTrans.Revenue.Collection > 0 Then
          ThisAmt = OldRound(TaxTrans.Revenue.Collection)
          TaxTrans.Revenue.Collection = 0
          If TaxTrans.Amount = ThisAmt Then TaxTrans.Amount = 0
          Put TTHandle, x, TaxTrans
          Fixcnt = Fixcnt + 1
'          Get TCHandle, TaxTrans.CustomerRec, TaxCust
'          ThisCustRec = TaxTrans.CustomerRec
'          For y = 1 To NumOfTTRecs
'            Get TTHandle, y, TaxTrans
'            If TaxTrans.TranType = 1 Then GoTo SkipIt
'            If TaxTrans.CustomerRec = ThisCustRec Then
'              If TaxTrans.Revenue.Interest = ThisAmt Then
'                TaxTrans.Amount = TaxTrans.Amount - ThisAmt
'                TaxTrans.Revenue.Interest = 0
'                Put TTHandle, y, TaxTrans
'                IntCnt = IntCnt + 1
'              End If
'            End If
'SkipIt:
'          Next y
        End If
      End If
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTTRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  Close
  
  Call TaxMsg(800, "A total of " + CStr(Fixcnt) + " bill personal transactions have had their collection revenues removed.")
End Sub

Private Sub FindCoNum()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If x = 2131 Then Stop
    TaxCust.Deleted = TaxCust.Deleted
'    If QPTrim$(TaxCust.CountyAcctString) = "5913" Then Stop
  Next x
  
  Close
End Sub
Private Sub AppendPtoPersCoNums()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.FirstPersRec > 0 Then
      If QPTrim$(TaxCust.CountyAcctString) <> "" Then
        TaxCust.CountyAcctString = QPTrim$(TaxCust.CountyAcctString) + "P"
        Put TCHandle, x, TaxCust
      ElseIf TaxCust.CountyAcct <> 0 Then
        TaxCust.CountyAcctString = CStr(TaxCust.CountyAcct) + "P"
        Put TCHandle, x, TaxCust
      End If
    End If
  Next x
  
  Close
  
  Call TaxMsg(900, "Finished.")
      
End Sub

Private Sub FixLawrenceville()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
'    If TaxTrans.CustomerRec = 726 And TaxTrans.TranType <> 1 And TaxTrans.TranType <> 4 Then Stop
'    If TaxTrans.Amount = 103.5 Then Stop
    TaxTrans.Amount = TaxTrans.Amount
    TaxTrans.TranType = TaxTrans.TranType
    TaxTrans.BelongTo = TaxTrans.BelongTo
'    If TaxTrans.BelongTo = 298 Then Stop
  Next x
  
  Close TTHandle
'  Call TaxMsg(900, "Finished.")
End Sub

Private Sub ClearDeletedCustTrans()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim x As Long, y As Long
  Dim NextRec As Long
  ReDim DArr(1 To 1) As Integer
  Dim dcnt As Integer
  Dim AHandle As Integer
  
  AHandle = FreeFile
  Open "DeletedCustomers.txt" For Output As AHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
   For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted <> 0 Then
      dcnt = dcnt + 1
      ReDim Preserve DArr(1 To dcnt) As Integer
      DArr(dcnt) = x
    End If
  Next x
  frmVATaxShowPctComp.Label1 = "Clearing deleted customer transactions"
  frmVATaxShowPctComp.Show , Me
 
 For x = 1 To dcnt
   Get TCHandle, DArr(x), TaxCust
      Print #AHandle, "Cust #   " + CStr(DArr(x))
      NextRec = TaxCust.LastTrans
      Do While NextRec > 0
        Get TTHandle, NextRec, TaxTrans
          Call ClearTrans(NextRec)
          Print #AHandle, "    Trans # " + CStr(NextRec)
          For y = 1 To NumOfTTRecs
            Get TTHandle, y, TaxTrans
            If TaxTrans.BelongTo = NextRec Then
              Call ClearTrans(y)
             Print #AHandle, "       BelongTo Trans # " + CStr(y)
            End If
          Next y
        Get TTHandle, NextRec, TaxTrans
        NextRec = TaxTrans.LastTrans
      Loop
     frmVATaxShowPctComp.ShowPctComp x, dcnt
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
 Next x
  Unload frmVATaxShowPctComp
  Close
  MsgBox ("A total of " + CStr(dcnt) + " transactions have been cleared. Look for DeletedCustomers.txt in the Citipak folder for results.")
  
End Sub

Private Sub FindBelongTo()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim x As Long, y As Integer
  Dim NextRec As Long
  ReDim TArr(1 To 1) As Long
  Dim TCnt As Integer
  Dim AHandle As Integer
  AHandle = FreeFile
  Open "DeletedCustomers.txt" For Output As AHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
'    If x = 7559 Then Stop
'    If TaxTrans.BelongTo = 1847 Then Stop
    TaxTrans.CustPin = TaxTrans.CustPin
      'Stop
'      Get TCHandle, TaxTrans.CustomerRec, TaxTrans
'      NextRec = TaxCust.LastTrans
'      Do While NextRec > 0
'        Get TTHandle, NextRec, TaxTrans
'        If TaxTrans.BelongTo = 63 Then
'          TCnt = TCnt + 1
'          ReDim Preserve TArr(1 To TCnt) As Long
'          TArr(TCnt) = x 'NextRec
'        End If

'        NextRec = TaxTrans.LastTrans
'      Loop
'    End If
    TaxTrans.CustPin = TaxTrans.CustPin
    TaxTrans.TranType = TaxTrans.TranType
  Next x
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
'        For y = 1 To TCnt
           If NextRec = 7735 Then 'TArr(y) Then
             Debug.Print CStr(x)
           End If
'        Next y
      NextRec = TaxTrans.LastTrans
    Loop
  Next x
  
  Close TTHandle
  MsgBox ("Done.")
End Sub
Private Sub FindOrphanTrans()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim AHandle As Integer
  Dim NextRec As Long
  Dim BelongTo As Long
  Dim cnt As Long
  Dim CustRec As Long
  Dim Found As Boolean
  
  AHandle = FreeFile
  Open "OrphanTrans.txt" For Output As AHandle
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  frmVATaxShowPctComp.Label1 = "Locating Orphan Transactions"
  frmVATaxShowPctComp.Show , Me
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
      CustRec = TaxTrans.CustomerRec
      If CustRec = 0 Then
        CustRec = TaxTrans.CustPin
      End If
      If CustRec = 0 Then
        Print #AHandle, CStr(CustRec) + "~" + CStr(x)
        cnt = cnt + 1
        GoTo Skip
      End If
      Found = False
      Get TCHandle, CustRec, TaxCust
      NextRec = TaxCust.LastTrans
      Do While NextRec > 0
        Get TTHandle, NextRec, TaxTrans
          If NextRec = x Then
            GoTo Skip
          End If
        NextRec = TaxTrans.LastTrans
      Loop
      Print #AHandle, CStr(CustRec) + "~" + CStr(x)
      cnt = cnt + 1
Skip:
     frmVATaxShowPctComp.ShowPctComp x, NumOfTTRecs
     If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  Close
  MsgBox ("A total of " + CStr(cnt) + " orphaned transactions have been located. Look for OrphanTrans.txt in the Citipak folder for results.")
End Sub
Private Sub FindAndFixBillsWithWrongTotals()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim AHandle As Integer
  Dim NextRec As Long
  Dim LastRec As Long
  Dim ECnt As Integer
  Dim BelongTo As Long
  Dim P1Pd As Double
  Dim P2Pd As Double
  Dim P3Pd As Double
  Dim P4Pd As Double
  Dim P5Pd As Double
  Dim IntPd As Double
  Dim AdvPd As Double
  Dim LLPd As Double
  Dim PenPd As Double
  Dim Opt1Pd As Double
  Dim Opt2Pd As Double
  Dim Opt3Pd As Double
  Dim P1 As Double
  Dim P2 As Double
  Dim P3 As Double
  Dim P4 As Double
  Dim P5 As Double
  Dim Intr As Double
  Dim Adv As Double
  Dim LL As Double
  Dim Pen As Double
  Dim Opt1 As Double
  Dim Opt2 As Double
  Dim Opt3 As Double
  Dim PPTRADisc As Double
  Dim Found As Boolean
  Call BuildMBvsCustHistArr
  AHandle = FreeFile
  Open "CorrectedBills.txt" For Output As AHandle
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  frmVATaxShowPctComp.Label1 = "Fixing Zeroed Out Credit at Billing"
  frmVATaxShowPctComp.Show , Me
 
'  CArrCnt = 1
'  CArr(1) = 5270
  For x = 1 To CArrCnt
    Get TCHandle, CArr(x), TaxCust
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
'        If NextRec = 31986 Then Stop
        If TaxTrans.TranType = 1 Then
          PPTRADisc = TaxTrans.PPTRADisc
          P1Pd = 0
          P2Pd = 0
          P3Pd = 0
          P4Pd = 0
          P5Pd = 0
          IntPd = 0
          AdvPd = 0
          LLPd = 0
          PenPd = 0
          Opt1Pd = 0
          Opt2Pd = 0
          Opt3Pd = 0
          P1 = 0
          P2 = 0
          P3 = 0
          P4 = 0
          P5 = 0
          Intr = 0
          Adv = 0
          LL = 0
          Pen = 0
          Opt1 = 0
          Opt2 = 0
          Opt3 = 0
          If TaxTrans.Revenue.Principle1 > 0 Then
            P1 = TaxTrans.Amount ' - TaxTrans.PPTRADisc
          ElseIf TaxTrans.Revenue.Principle2 > 0 Then
            P2 = TaxTrans.Amount '- TaxTrans.PPTRADisc
          ElseIf TaxTrans.Revenue.Principle3 > 0 Then
            P3 = TaxTrans.Amount '- TaxTrans.PPTRADisc
          ElseIf TaxTrans.Revenue.Principle4 > 0 Then
            P4 = TaxTrans.Amount '- TaxTrans.PPTRADisc
           ElseIf TaxTrans.Revenue.Principle5 > 0 Then
            P5 = TaxTrans.Amount '- TaxTrans.PPTRADisc
         End If
          LastRec = TaxCust.LastTrans
          Do While LastRec > 0
             Get TTHandle, LastRec, TaxTrans
                  If TaxTrans.TranType = 13 And TaxTrans.BelongTo = NextRec Then 'adjust bill down
                    Intr = Intr - TaxTrans.Revenue.Interest
                    Pen = Pen - TaxTrans.Revenue.Penalty
                    P1 = P1 - TaxTrans.Revenue.Principle1
                    P2 = P2 - TaxTrans.Revenue.Principle2
                    P3 = P3 - TaxTrans.Revenue.Principle3
                    P4 = P4 - TaxTrans.Revenue.Principle4
                    P5 = P5 - TaxTrans.Revenue.Principle5
                    Opt1 = Opt1 - TaxTrans.Revenue.RevOpt1
                    Opt2 = Opt2 - TaxTrans.Revenue.RevOpt2
                    Opt3 = Opt3 - TaxTrans.Revenue.RevOpt3
                    Adv = Adv - TaxTrans.Revenue.Collection
                    LL = LL - TaxTrans.Revenue.LateList
                  ElseIf (TaxTrans.TranType = 2 Or TaxTrans.TranType = 9 Or TaxTrans.TranType = 21) And TaxTrans.BelongTo = NextRec Then
                    IntPd = IntPd + TaxTrans.Revenue.InterestPd
                    PenPd = PenPd + TaxTrans.Revenue.PenaltyPd
                    P1Pd = P1Pd + TaxTrans.Revenue.Principle1Pd
                    P2Pd = P2Pd + TaxTrans.Revenue.Principle2Pd
                    P3Pd = P3Pd + TaxTrans.Revenue.Principle3Pd
                    P4Pd = P4Pd + TaxTrans.Revenue.Principle4Pd
                    P5Pd = P5Pd + TaxTrans.Revenue.Principle5Pd
                    Opt1Pd = Opt1Pd + TaxTrans.Revenue.RevOpt1Pd
                    Opt2Pd = Opt2Pd + TaxTrans.Revenue.RevOpt2Pd
                    Opt3Pd = Opt3Pd + TaxTrans.Revenue.RevOpt3Pd
                    AdvPd = AdvPd + TaxTrans.Revenue.CollectionPd
                    LLPd = LLPd + TaxTrans.Revenue.LateListPd
                  ElseIf TaxTrans.TranType = 3 And TaxTrans.BelongTo = NextRec Then  'release
                    Intr = Intr - TaxTrans.Revenue.Interest
                    Pen = Pen - TaxTrans.Revenue.Penalty
                    P1 = P1 - TaxTrans.Revenue.Principle1
                    P2 = P2 - TaxTrans.Revenue.Principle2
                    P3 = P3 - TaxTrans.Revenue.Principle3
                    P4 = P4 - TaxTrans.Revenue.Principle4
                    P5 = P5 - TaxTrans.Revenue.Principle5
                    Opt1 = Opt1 - TaxTrans.Revenue.RevOpt1
                    Opt2 = Opt2 - TaxTrans.Revenue.RevOpt2
                    Opt3 = Opt3 - TaxTrans.Revenue.RevOpt3
                    Adv = Adv - TaxTrans.Revenue.Collection
                    LL = LL - TaxTrans.Revenue.LateList
                  ElseIf TaxTrans.TranType = 14 And TaxTrans.BelongTo = NextRec Then  'adjust bill up
                    Intr = Intr + TaxTrans.Revenue.Interest
                    Pen = Pen + TaxTrans.Revenue.Penalty
                    P1 = P1 + TaxTrans.Revenue.Principle1
                    P2 = P2 + TaxTrans.Revenue.Principle2
                    P3 = P3 + TaxTrans.Revenue.Principle3
                    P4 = P4 + TaxTrans.Revenue.Principle4
                    P5 = P5 + TaxTrans.Revenue.Principle5
                    Opt1 = Opt1 + TaxTrans.Revenue.RevOpt1
                    Opt2 = Opt2 + TaxTrans.Revenue.RevOpt2
                    Opt3 = Opt3 + TaxTrans.Revenue.RevOpt3
                    Adv = Adv + TaxTrans.Revenue.Collection
                    LL = LL + TaxTrans.Revenue.LateList
                  ElseIf (TaxTrans.TranType = 4 Or TaxTrans.TranType = 5 Or TaxTrans.TranType = 6) And TaxTrans.BelongTo = NextRec Then
                    Intr = Intr + TaxTrans.Revenue.Interest
                    Pen = Pen + TaxTrans.Revenue.Penalty
                    P1 = P1 + TaxTrans.Revenue.Principle1
                    P2 = P2 + TaxTrans.Revenue.Principle2
                    P3 = P3 + TaxTrans.Revenue.Principle3
                    P4 = P4 + TaxTrans.Revenue.Principle4
                    P5 = P5 + TaxTrans.Revenue.Principle5
                    Opt1 = Opt1 + TaxTrans.Revenue.RevOpt1
                    Opt2 = Opt2 + TaxTrans.Revenue.RevOpt2
                    Opt3 = Opt3 + TaxTrans.Revenue.RevOpt3
                    Adv = Adv + TaxTrans.Revenue.Collection
                    LL = LL + TaxTrans.Revenue.LateList
                  End If
             
               LastRec = TaxTrans.LastTrans
              Loop
              Get TTHandle, NextRec, TaxTrans
              Found = False
             If P1 + P2 + P3 + P4 + P5 < 0 Then GoTo SkipIt
              If P1 + PPTRADisc <> TaxTrans.Revenue.Principle1 And TaxTrans.Revenue.Principle1 > 0 Then
               TaxTrans.Revenue.Principle1 = P1 + PPTRADisc
               Found = True
              End If
              If P2 + PPTRADisc <> TaxTrans.Revenue.Principle2 And TaxTrans.Revenue.Principle2 > 0 Then
               TaxTrans.Revenue.Principle2 = P2 + PPTRADisc
               Found = True
              End If
              If P3 + PPTRADisc <> TaxTrans.Revenue.Principle3 And TaxTrans.Revenue.Principle3 > 0 Then
               TaxTrans.Revenue.Principle3 = P3 + PPTRADisc
               Found = True
              End If
              If P4 + PPTRADisc <> TaxTrans.Revenue.Principle4 And TaxTrans.Revenue.Principle4 > 0 Then
               TaxTrans.Revenue.Principle4 = P4 + PPTRADisc
               Found = True
              End If
              If P5 + PPTRADisc <> TaxTrans.Revenue.Principle5 And TaxTrans.Revenue.Principle5 > 0 Then
               TaxTrans.Revenue.Principle5 = P5 + PPTRADisc
               Found = True
              End If
              If P1Pd + P2Pd + P3Pd + P4Pd + P5Pd < 0 Then GoTo SkipIt
              If TaxTrans.Revenue.Principle1Pd <> P1Pd Then
                TaxTrans.Revenue.Principle1Pd = P1Pd
                Found = True
              End If
              If TaxTrans.Revenue.Principle2Pd <> P2Pd Then
                TaxTrans.Revenue.Principle2Pd = P2Pd
                Found = True
              End If
              If TaxTrans.Revenue.Principle3Pd <> P3Pd Then
                TaxTrans.Revenue.Principle3Pd = P3Pd
                Found = True
              End If
              If TaxTrans.Revenue.Principle4Pd <> P4Pd Then
                TaxTrans.Revenue.Principle4Pd = P4Pd
                Found = True
              End If
              If TaxTrans.Revenue.Principle5Pd <> P5Pd Then
                TaxTrans.Revenue.Principle5Pd = P5Pd
                Found = True
              End If
              If Adv < 0 Or Intr < 0 Or Pen < 0 Or LL < 0 Then GoTo SkipIt
              If TaxTrans.Revenue.Collection <> Adv Then
                TaxTrans.Revenue.Collection = Adv
                 Found = True
             End If
              If TaxTrans.Revenue.Interest <> Intr Then
                TaxTrans.Revenue.Interest = Intr
                 Found = True
             End If
              If TaxTrans.Revenue.Penalty <> Pen Then
                TaxTrans.Revenue.Penalty = Pen
                 Found = True
             End If
              If TaxTrans.Revenue.LateList <> LL Then
               TaxTrans.Revenue.LateList = LL
                Found = True
             End If
              If AdvPd < 0 Or IntPd < 0 Or PenPd < 0 Or LLPd < 0 Then GoTo SkipIt
              If TaxTrans.Revenue.CollectionPd <> AdvPd Then
                TaxTrans.Revenue.CollectionPd = AdvPd
                Found = True
              End If
              If TaxTrans.Revenue.InterestPd <> IntPd Then
               TaxTrans.Revenue.InterestPd = IntPd
               Found = True
              End If
             If TaxTrans.Revenue.PenaltyPd <> PenPd Then
               TaxTrans.Revenue.PenaltyPd = PenPd
               Found = True
             End If
             If TaxTrans.Revenue.LateListPd <> LLPd Then
              TaxTrans.Revenue.LateListPd = LLPd
              Found = True
            End If
              If Opt1 < 0 Or Opt2 < 0 Or Opt3 < 0 Then GoTo SkipIt
              If TaxTrans.Revenue.RevOpt1 <> Opt1 Then
                TaxTrans.Revenue.RevOpt1 = Opt1
                Found = True
              End If
              If TaxTrans.Revenue.RevOpt2 <> Opt2 Then
                TaxTrans.Revenue.RevOpt2 = Opt2
                Found = True
              End If
              If TaxTrans.Revenue.RevOpt3 <> Opt3 Then
                TaxTrans.Revenue.RevOpt3 = Opt3
                Found = True
              End If
              If Opt1Pd < 0 Or Opt2Pd < 0 Or Opt3Pd < 0 Then GoTo SkipIt
              If TaxTrans.Revenue.RevOpt1Pd <> Opt1Pd Then
                TaxTrans.Revenue.RevOpt1Pd = Opt1Pd
                Found = True
              End If
              If TaxTrans.Revenue.RevOpt2Pd <> Opt2Pd Then
                TaxTrans.Revenue.RevOpt2Pd = Opt2Pd
                Found = True
              End If
              If TaxTrans.Revenue.RevOpt3Pd <> Opt3Pd Then
                TaxTrans.Revenue.RevOpt3Pd = Opt3Pd
                Found = True
              End If
              If Found = True Then
                Put TTHandle, NextRec, TaxTrans
                ECnt = ECnt + 1
                Print #AHandle, CStr(CArr(x)) + "~" + CStr(NextRec)
              End If
        End If
SkipIt:
      Get TTHandle, NextRec, TaxTrans
      NextRec = TaxTrans.LastTrans
    Loop

  Next x
  
  Unload frmVATaxShowPctComp
 
  Close
  MsgBox ("A total of " + CStr(ECnt) + " bills were corrected. Look for CorrectedBills.txt in Citipak.")
End Sub
Private Sub FindCorrectQueue()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim AHandle As Integer
  Dim NextRec As Long
  Dim cnt As Integer
  ReDim ECust(1 To 1) As Long
  Dim ECnt As Integer
  Dim PHandle As Integer
  Dim LineLen As Integer
  Dim NumHold As String
  Dim ch As String
  Dim TextLine As String
  Dim BillPaid As Double
  Dim Paid As Double
  Dim Billed As Double
  Dim TestBal As Double
  Dim NewTestBal As Long
  Dim NewRec As Long
  Dim LookRec As Long
  Dim CustFound As Integer
  Dim BottomRec As Long
  Dim TopRec As Long
  Dim ErrRec As Long
  Dim FindRec As Long
  Dim Found As Integer
  
  AHandle = FreeFile
  Open "CorrectQueue.txt" For Output As AHandle
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  GoSub BuildArray
  frmVATaxShowPctComp.Label1 = "Correcting orphan or crossed transaction"
  frmVATaxShowPctComp.Show , Me
'  ECnt = 1
'  ECust(1) = 880
  For x = 1 To ECnt
   Get TCHandle, ECust(x), TaxCust
   NextRec = TaxCust.LastTrans
'   If ECust(x) = 189 Then Stop
   Do While NextRec > 0
     Get TTHandle, NextRec, TaxTrans
'     If ECust(x) = 24 And NextRec = 22292 Then Stop
     If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
      If TestBal = 0 Then
        BillPaid# = OldRound#(TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd)
        BillPaid# = OldRound#(BillPaid# + TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd)
        BillPaid# = OldRound#(BillPaid# + TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd) ' + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl)
        NewRec = TaxCust.LastTrans
        Paid = 0
        Do While NewRec > 0
          Get TTHandle, NewRec, TaxTrans
            If TaxTrans.BelongTo = NextRec Then
              Paid# = OldRound#(Paid# + TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd)
              Paid# = OldRound#(Paid# + TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd)
              Paid# = OldRound#(Paid# + TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt) ' + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl)
            End If
          NewRec = TaxTrans.LastTrans
        Loop
        If Paid <> BillPaid Then 'selected bill is paid but payment missing in queue
          'now look thru customers to find where any payments might be
          CustFound = 0
          For y = 1 To NumOfTCRecs
            Get TCHandle, y, TaxCust
'            If y = 988 Then Stop
            LookRec = TaxCust.LastTrans
            Do While LookRec > 0
              Get TTHandle, LookRec, TaxTrans
'                Print #AHandle, "LookRec = " + CStr(LookRec)
                If TaxTrans.BelongTo = NextRec Then
                  CustFound = y
                  Exit For
                End If
              If TaxTrans.LastTrans = LookRec Then Stop
              LookRec = TaxTrans.LastTrans
            Loop
          Next y
          If CustFound > 0 Then
            GoSub SwapTrans
          Else
            GoSub InsertTrans
          End If
        End If
      End If
     End If
     Get TTHandle, NextRec, TaxTrans
     NextRec = TaxTrans.LastTrans
   Loop
    frmVATaxShowPctComp.ShowPctComp x, ECnt
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  
  Next x
  Unload frmVATaxShowPctComp
  Close
  MsgBox ("Count = " + CStr(cnt))
  Exit Sub

BuildArray:
  Call CompareMBWithCustHistory
  PHandle = FreeFile
  Open App.Path + "\MBvsCustHist.txt" For Input As #PHandle  ' Open file.
   Do While Not eof(PHandle)  ' Loop until end of file.
     Line Input #PHandle, TextLine   ' Read next line into Textline.
     LineLen = Len(TextLine)
     For x = 1 To LineLen
       ch = Mid(TextLine, x, 1)
       If ch <> "~" Then
         NumHold = NumHold + ch
       Else
        ECnt = ECnt + 1
        ReDim Preserve ECust(1 To ECnt) As Long
        ECust(ECnt) = CLng(NumHold)
        NumHold = ""
        Exit For
       End If
     Next x
   Loop
   Close PHandle
  Return
  
SwapTrans:
  Get TCHandle, CustFound, TaxCust
  ErrRec = TaxCust.LastTrans
  Found = 0
  Do While ErrRec > 0
    Get TTHandle, ErrRec, TaxTrans
      If TaxTrans.BelongTo = NextRec Then
        If ErrRec = TaxCust.LastTrans Then
          TaxCust.LastTrans = TaxTrans.LastTrans 'new cust last trans is trans below lost trans
           Put TCHandle, CustFound, TaxCust
          Print #AHandle, "Moved trans # " + CStr(ErrRec) + " from cust # " + CStr(CustFound) + " to cust # " + CStr(ECust(x))
          cnt = cnt + 1
          Found = 1
          Exit Do
        Else
          BottomRec = TaxTrans.LastTrans
          Get TTHandle, TopRec, TaxTrans
            TaxTrans.LastTrans = BottomRec 'remove from incorrect queue
          Put TTHandle, TopRec, TaxTrans
          Print #AHandle, "Moved trans # " + CStr(ErrRec) + " from cust # " + CStr(CustFound) + " to cust # " + CStr(ECust(x))
          cnt = cnt + 1
           Found = 1
          Exit Do
        End If
      End If
    TopRec = ErrRec
    ErrRec = TaxTrans.LastTrans
  Loop
  If Found = 0 Then
    GoSub InsertTrans
    Return
  End If
  'now add to correct queue
  Get TCHandle, ECust(x), TaxCust
    If TaxCust.LastTrans = NextRec Then 'top trans is selected bill trans
      Get TTHandle, ErrRec, TaxTrans 'assign the new link (cust top trans) to the lost trans
        TaxTrans.LastTrans = TaxCust.LastTrans
        TaxTrans.CustomerRec = ECust(x)
        TaxTrans.CustPin = ECust(x)
        Put TTHandle, ErrRec, TaxTrans
        TaxCust.LastTrans = ErrRec 'now assign the lost trans to cust as new last link
        Put TCHandle, ECust(x), TaxCust
        Return
    Else
      FindRec = TaxCust.LastTrans
      Do While FindRec > 0
        Get TTHandle, FindRec, TaxTrans
        If FindRec = NextRec Then
          Get TTHandle, ErrRec, TaxTrans
            TaxTrans.LastTrans = FindRec
            TaxTrans.CustomerRec = ECust(x)
            TaxTrans.CustPin = ECust(x)
            Put TTHandle, ErrRec, TaxTrans
            Get TTHandle, TopRec, TaxTrans
              TaxTrans.LastTrans = ErrRec
              Put TTHandle, TopRec, TaxTrans
              Exit Do
        End If
        TopRec = FindRec
        FindRec = TaxTrans.LastTrans
      Loop
      
    End If
    
  Return
  
InsertTrans:
  
  ErrRec = 0
  For y = 1 To NumOfTTRecs
    Get TTHandle, y, TaxTrans
    If TaxTrans.BelongTo = NextRec Then
      ErrRec = y
      Get TCHandle, ECust(x), TaxCust
        If TaxCust.LastTrans = NextRec Then 'top cust trans is selected bill trans
          Get TTHandle, ErrRec, TaxTrans 'assign the new link (cust top trans) to the lost trans
            TaxTrans.LastTrans = TaxCust.LastTrans
            TaxTrans.CustomerRec = ECust(x)
            TaxTrans.CustPin = ECust(x)
            Put TTHandle, ErrRec, TaxTrans
            TaxCust.LastTrans = ErrRec 'now assign the lost trans to cust as new last link
            Print #AHandle, "Moved trans # " + CStr(ErrRec) + " from cust # " + CStr(0) + " to cust # " + CStr(ECust(x))
            Put TCHandle, ECust(x), TaxCust
            cnt = cnt + 1
            Return
        Else
          FindRec = TaxCust.LastTrans
          Do While FindRec > 0
            Get TTHandle, FindRec, TaxTrans
            If FindRec = NextRec Then 'nextrec is trans that needs inserting
              Get TTHandle, ErrRec, TaxTrans
                TaxTrans.LastTrans = FindRec
                TaxTrans.CustomerRec = ECust(x)
                TaxTrans.CustPin = ECust(x)
                Put TTHandle, ErrRec, TaxTrans
                Get TTHandle, TopRec, TaxTrans
                  TaxTrans.LastTrans = ErrRec
                  Put TTHandle, TopRec, TaxTrans
                    Exit Do
            End If
            TopRec = FindRec
            FindRec = TaxTrans.LastTrans
          Loop
       End If
     End If
  Next y
     
  Return
End Sub

Private Sub FixCustomerRecsandPins()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim NumOfTCRecs As Long
  Dim TCHandle As Integer
  Dim x As Long, y As Long, z As Long
  Dim AHandle As Integer
  Dim BHandle As Integer
  Dim cnt As Integer
  Dim CustNum As Integer
  Dim NextRec As Long
  Dim Fixed As Boolean
  Dim OldCust As Integer
  Dim NewCust As Integer
  Dim TransRecError As Long
  Dim BelongTo As Long
  Dim TopRec As Long
  Dim BottomRec As Long
  
  cnt = 0
  AHandle = FreeFile
  Open "CustRecIsZero.txt" For Output As AHandle
  BHandle = FreeFile
  Open "CustRecIsZeroFixed.txt" For Output As BHandle
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  frmVATaxShowPctComp.Label1 = "Fixing Transactions/Owner Errors"
  frmVATaxShowPctComp.Show , Me
  
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
'    If x = 7720 Then Stop
    BelongTo = TaxTrans.BelongTo
    If TaxTrans.CustomerRec = 0 Or TaxTrans.CustPin = 0 Then 'locate error
      TransRecError = x
      CustNum = -1
      OldCust = 0
      NewCust = 0
      For y = 1 To NumOfTCRecs
        Get TCHandle, y, TaxCust
        NextRec = TaxCust.LastTrans
        Do While NextRec > 0
          Get TTHandle, NextRec, TaxTrans
            If NextRec = x Then
             CustNum = y 'who's got it now is this CustNum
             Exit For
            End If
          NextRec = TaxTrans.LastTrans
        Loop
      Next y
      If CustNum = -1 And BelongTo > 0 Then
        GoSub FixItForNewOnly 'can't find it in anybody's queue
        GoTo SkipIt
      End If
      If BelongTo = 0 Then
        Print #BHandle, "Could not fix # " + CStr(x) + " because the belongto is zero."
        GoTo SkipIt
      End If
      Fixed = False
      Get TTHandle, x, TaxTrans
      If TaxTrans.CustPin > 0 Then 'its in the correct customer's queue so just assign cust rec
        If CustNum = TaxTrans.CustPin Then
          TaxTrans.CustomerRec = CustNum
          Put TTHandle, x, TaxTrans
          Print #BHandle, CStr(x) + "~" + CStr(TaxTrans.CustPin) + "~" + "Rec = 0" + "~" + CStr(CustNum) + "~" + "Fixed"
        Else
          OldCust = TaxTrans.CustPin
          TransRecError = x
'          BelongTo = TaxTrans.BelongTo
          GoSub FixIt
'          Print #AHandle, CStr(x) + "~" + CStr(TaxTrans.CustPin) + "~" + "Rec = 0" + "~" + CStr(CustNum)
        End If
      Else
        GoSub FixItBothZeros
'        Print #AHandle, CStr(x) + "~" + "Rec and Pin = 0" + "~" + CStr(CustNum)
      End If
    
      If TaxTrans.CustomerRec > 0 Then
        If CustNum = TaxTrans.CustomerRec Then 'its in the correct customer's queue so just assign cust rec
          TaxTrans.CustPin = CustNum
          Put TTHandle, x, TaxTrans
          Print #BHandle, CStr(x) + "~" + CStr(TaxTrans.CustomerRec) + "~" + "Pin = 0" + "~" + CStr(CustNum) + "~" + "Fixed"
        Else
          OldCust = TaxTrans.CustomerRec
          TransRecError = x
'          BelongTo = TaxTrans.BelongTo
          GoSub FixIt
'          Print #AHandle, CStr(x) + "~" + CStr(TaxTrans.CustomerRec) + "~" + "Pin = 0" + "~" + CStr(CustNum)
        End If
      Else
        GoSub FixItBothZeros
'        Print #AHandle, CStr(x) + "~" + "Rec and Pin = 0" + "~" + CStr(CustNum)
      End If
    End If
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTTRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  Close
  MsgBox ("A total of " + CStr(cnt) + " transactions were found.")
  Exit Sub
  
FixIt:
  For y = 1 To NumOfTCRecs 'need to locate the customer that should have this trans
    Get TCHandle, y, TaxCust
    NextRec = TaxCust.LastTrans
    If NextRec = BelongTo Then 'found it...interating just to make sure that we
      'have the correct cust (could have gotten the cust # from the belong to trans)
      Get TCHandle, CustNum, TaxCust 'get old cust
      NextRec = TaxCust.LastTrans
      If NextRec = TransRecError Then 'if the bad trans is at the top
        TaxCust.LastTrans = TaxTrans.LastTrans 'pull it out of the queue
        Put TCHandle, CustNum, TaxCust
        GoTo UpDateNew
      Else
        Do While NextRec > 0 'if its not at the top
          Get TTHandle, NextRec, TaxTrans 'move down the queue till you find it
          If NextRec = TransRecError Then 'found it
            BottomRec = TaxTrans.LastTrans 'link this to the one above the bad trans
            Get TTHandle, TopRec, TaxTrans
              TaxTrans.LastTrans = BottomRec 'save new link
            Put TTHandle, TopRec, TaxTrans
            Get TTHandle, TransRecError, TaxTrans
              TaxTrans.LastTrans = NextRec
              TaxTrans.CustomerRec = y
              TaxTrans.CustPin = y
            Put TTHandle, TransRecError, TaxTrans
            GoTo UpDateNew
          End If
          TopRec = NextRec 'keep up with the last trans
          NextRec = TaxTrans.LastTrans
        Loop
      End If
UpDateNew: 'now handle the new cust
      Get TCHandle, y, TaxCust 'bring the correct cust back up
      NextRec = TaxCust.LastTrans
      If NextRec = BelongTo Then 'if the top trans is the correct bill trans
        Get TTHandle, TransRecError, TaxTrans 'pull the new trans up
          TaxTrans.LastTrans = BelongTo 'give it the link of the former top trans
        Put TTHandle, TransRecError, TaxTrans
          TaxCust.LastTrans = TransRecError 'now give the new cust the new trans as its last trans
        Put TCHandle, y, TaxCust
        Exit For
        Print #BHandle, "Moved trans # ~" + CStr(TransRecError) + " ~ to cust # ~ " + CStr(y) + " ~ from cust # ~ " + CStr(CustNum)
      Else 'top trans is not the belongto
        Do While NextRec > 0
          Get TTHandle, NextRec, TaxTrans
          If NextRec = BelongTo Then
            Get TTHandle, TopRec, TaxTrans 'reassign links
              TaxTrans.LastTrans = TransRecError
            Put TTHandle, TopRec, TaxTrans
            Get TTHandle, TransRecError, TaxTrans
              TaxTrans.LastTrans = NextRec
              TaxTrans.CustomerRec = y
              TaxTrans.CustPin = y
            Put TTHandle, TransRecError, TaxTrans
            Print #BHandle, "Moved trans # ~" + CStr(TransRecError) + " ~ to cust # ~ " + CStr(y) + " ~ from cust # ~ " + CStr(CustNum)
            Exit For
          End If
          TopRec = NextRec
          NextRec = TaxTrans.LastTrans
        Loop
      End If
    End If
  
  Next y
  Return
  
FixItBothZeros:
  Get TCHandle, CustNum, TaxCust
  NextRec = TaxCust.LastTrans
  If NextRec = TransRecError Then 'bad trans is at the top
    Get TTHandle, NextRec, TaxTrans
      TaxCust.LastTrans = TaxTrans.LastTrans 'pull the bad trans out and reassign last trans
      'to cust
    Put TCHandle, CustNum, TaxCust
    cnt = cnt + 1

    GoTo FixNewCust
  Else
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
      If NextRec = TransRecError Then
        BottomRec = TaxTrans.LastTrans
        Get TTHandle, TopRec, TaxTrans
          TaxTrans.LastTrans = BottomRec
        Put TTHandle, TopRec, TaxTrans
        GoTo FixNewCust
      End If
      TopRec = NextRec
      NextRec = TaxTrans.LastTrans
    Loop
  End If
  
FixNewCust:
  For z = 1 To NumOfTCRecs
    Get TCHandle, z, TaxCust
    NextRec = TaxCust.LastTrans
    If NextRec = BelongTo Then
      TaxTrans.LastTrans = NextRec
      TaxCust.LastTrans = TransRecError
      Get TTHandle, TransRecError, TaxTrans
      Put TTHandle, TransRecError, TaxTrans
      Put TCHandle, z, TaxCust
      Print #BHandle, "Moved trans # ~" + CStr(TransRecError) + " ~ to cust # ~ " + CStr(z) + " ~ from cust # ~ " + CStr(CustNum)
      Return
    Else
     Do While NextRec > 0
       Get TTHandle, NextRec, TaxTrans
         If NextRec = BelongTo Then
           Get TTHandle, TransRecError, TaxTrans
             TaxTrans.LastTrans = NextRec
             TaxTrans.CustomerRec = z
             TaxTrans.CustPin = z
           Put TTHandle, TransRecError, TaxTrans
           Get TTHandle, TopRec, TaxTrans
           TaxTrans.LastTrans = TransRecError
           Put TTHandle, TopRec, TaxTrans
           Print #BHandle, "Moved trans # ~" + CStr(TransRecError) + " ~ to cust # ~ " + CStr(z) + " ~ from cust # ~ " + CStr(CustNum)
           Return
         End If
       TopRec = NextRec
       NextRec = TaxTrans.LastTrans
     Loop
    
    End If
  
  Next z

  Return
  
FixItForNewOnly:
  cnt = cnt + 1
  For z = 1 To NumOfTCRecs
    Get TCHandle, z, TaxCust
    NextRec = TaxCust.LastTrans
    If NextRec = BelongTo Then
      TaxTrans.LastTrans = NextRec
      TaxCust.LastTrans = TransRecError
      Get TTHandle, TransRecError, TaxTrans
      Put TTHandle, TransRecError, TaxTrans
      Put TCHandle, x, TaxCust
      Print #BHandle, "Moved trans # ~" + CStr(TransRecError) + " ~ to cust # ~ " + CStr(z) + " ~ from cust # ~ " + CStr(CustNum)
      Return
    Else
     Do While NextRec > 0
       Get TTHandle, NextRec, TaxTrans
         If NextRec = BelongTo Then
           Get TTHandle, TransRecError, TaxTrans
             TaxTrans.LastTrans = NextRec
             TaxTrans.CustomerRec = z
             TaxTrans.CustPin = z
           Put TTHandle, TransRecError, TaxTrans
           Get TTHandle, TopRec, TaxTrans
           TaxTrans.LastTrans = TransRecError
           Put TTHandle, TopRec, TaxTrans
           Print #BHandle, "Moved trans # ~" + CStr(TransRecError) + " ~ to cust # ~ " + CStr(z) + " ~ from cust # ~ " + CStr(CustNum)
           Return
         End If
       TopRec = NextRec
       NextRec = TaxTrans.LastTrans
     Loop
    End If
  
  Next z

  Return
End Sub
Private Sub cmdCnvtPstedBills_Click()
  Dim PRRec As VARETaxBillType
  Dim PRHandle As Integer
  Dim NumOfPRRecs As Long
  Dim PRRecOld As VARETaxBillTypeOld
  Dim PROldHandle As Integer
  Dim NumOfOldPRRecs As Long
  Dim MyPath$, ThisFile$
  Dim x As Integer, y As Integer
  Dim PostRec As TaxBillPostDateType
  Dim PostHandle As Integer
  Dim NumOfPostRecs As Long
  Dim RCnt As Integer
  Dim ThisCnt As Integer
  Dim BigBCnt As Long
  
  
  ReDim RFileNames(1 To 1) As String
  OpenBillPostDateFile PostHandle, NumOfPostRecs
  For x = 1 To NumOfPostRecs
    Get PostHandle, x, PostRec
    ThisFile = PostRec.BackUpName
    If PostRec.BillType = "R" Then
      OpenRealPostedReprintFileOld PROldHandle, NumOfOldPRRecs, ThisFile
      If NumOfOldPRRecs > BigBCnt Then BigBCnt = NumOfOldPRRecs
      RCnt = RCnt + 1
      ReDim Preserve RFileNames(1 To RCnt) As String
      RFileNames(RCnt) = PostRec.BackUpName
    End If
  Next x
  
  ReDim CustRec(1 To RCnt, 1 To BigBCnt) As Long
  ReDim CustName(1 To RCnt, 1 To BigBCnt) As String
  ReDim CustAdd1(1 To RCnt, 1 To BigBCnt) As String
  ReDim CustAdd2(1 To RCnt, 1 To BigBCnt) As String
  ReDim CustAdd3(1 To RCnt, 1 To BigBCnt) As String
  ReDim CustZip(1 To RCnt, 1 To BigBCnt) As String
  ReDim RDesc1(1 To RCnt, 1 To BigBCnt) As String
  ReDim RDesc2(1 To RCnt, 1 To BigBCnt) As String
  ReDim RealPin(1 To RCnt, 1 To BigBCnt) As String
  ReDim RealValue(1 To RCnt, 1 To BigBCnt) As Double
  ReDim TotalValue(1 To RCnt, 1 To BigBCnt) As Double
  ReDim ExptValue(1 To RCnt, 1 To BigBCnt) As Double
  ReDim RealTaxDue(1 To RCnt, 1 To BigBCnt) As Double
  ReDim BldgValue(1 To RCnt, 1 To BigBCnt) As Double
  ReDim LateTaxDue(1 To RCnt, 1 To BigBCnt) As Double
  ReDim TotalBillDue(1 To RCnt, 1 To BigBCnt) As Double
  ReDim BillNumber(1 To RCnt, 1 To BigBCnt) As Long
  ReDim TaxYear(1 To RCnt, 1 To BigBCnt) As Integer
  ReDim BillPrinted(1 To RCnt, 1 To BigBCnt) As Integer
  ReDim RealPropRecord(1 To RCnt, 1 To BigBCnt) As Long
  ReDim PriorYrBalance(1 To RCnt, 1 To BigBCnt) As Double
  ReDim RealTaxRate(1 To RCnt, 1 To BigBCnt) As Double
  ReDim CustPin(1 To RCnt, 1 To BigBCnt) As Long
  ReDim TownShip(1 To RCnt, 1 To BigBCnt) As String
  ReDim MORTCODE(1 To RCnt, 1 To BigBCnt) As String
  ReDim LotOrAcre(1 To RCnt, 1 To BigBCnt) As String
  ReDim LASize(1 To RCnt, 1 To BigBCnt) As String
  ReDim MortRec(1 To RCnt, 1 To BigBCnt) As Integer
  ReDim RDesc3(1 To RCnt, 1 To BigBCnt) As String
  ReDim InternalPin(1 To RCnt, 1 To BigBCnt) As Long
  ReDim OptRevTax1(1 To RCnt, 1 To BigBCnt) As Double
  ReDim OptRevTax2(1 To RCnt, 1 To BigBCnt) As Double
  ReDim OptRevTax3(1 To RCnt, 1 To BigBCnt) As Double
  ReDim OverPayAmt(1 To RCnt, 1 To BigBCnt) As Double
  ReDim DueDate(1 To RCnt, 1 To BigBCnt) As Integer
  ReDim PostDate(1 To RCnt, 1 To BigBCnt) As Integer
  ReDim TransRec(1 To RCnt, 1 To BigBCnt) As Long
  ReDim Opt1Desc(1 To RCnt, 1 To BigBCnt) As String
  ReDim Opt2Desc(1 To RCnt, 1 To BigBCnt) As String
  ReDim Opt3Desc(1 To RCnt, 1 To BigBCnt) As String
  ReDim Padding(1 To RCnt, 1 To BigBCnt) As String
  ReDim Comment(1 To RCnt, 1 To BigBCnt) As String
  For y = 1 To NumOfPostRecs
    Get PostHandle, y, PostRec
    ThisFile = QPTrim$(PostRec.BackUpName)
    If PostRec.BillType = "R" Then
      OpenRealPostedReprintFileOld PROldHandle, NumOfOldPRRecs, ThisFile
      ThisCnt = ThisCnt + 1
      For x = 1 To NumOfOldPRRecs
        Get PROldHandle, x, PRRecOld
        If QPTrim$(PRRecOld.Padding) = "Converted" Then
          Close
          MsgBox ("This data has already been converted.")
          Exit Sub
        End If
        CustRec(ThisCnt, x) = PRRecOld.CustRec
        CustName(ThisCnt, x) = PRRecOld.CustName
        CustAdd1(ThisCnt, x) = PRRecOld.CustAdd1
        CustAdd2(ThisCnt, x) = PRRecOld.CustAdd2
        CustAdd3(ThisCnt, x) = PRRecOld.CustAdd3
        CustZip(ThisCnt, x) = PRRecOld.CustZip
        RDesc1(ThisCnt, x) = PRRecOld.RDesc1
        RDesc2(ThisCnt, x) = PRRecOld.RDesc2
        RealPin(ThisCnt, x) = PRRecOld.RealPin
        RealValue(ThisCnt, x) = PRRecOld.RealValue
        TotalValue(ThisCnt, x) = PRRecOld.TotalValue
        ExptValue(ThisCnt, x) = PRRecOld.ExptValue
        RealTaxDue(ThisCnt, x) = PRRecOld.RealTaxDue
        BldgValue(ThisCnt, x) = PRRecOld.BldgValue
        LateTaxDue(ThisCnt, x) = PRRecOld.LateTaxDue
        TotalBillDue(ThisCnt, x) = PRRecOld.TotalBillDue
        BillNumber(ThisCnt, x) = PRRecOld.BillNumber
        TaxYear(ThisCnt, x) = PRRecOld.TaxYear
        BillPrinted(ThisCnt, x) = PRRecOld.BillPrinted
        RealPropRecord(ThisCnt, x) = PRRecOld.RealPropRecord
        PriorYrBalance(ThisCnt, x) = PRRecOld.PriorYrBalance
        RealTaxRate(ThisCnt, x) = PRRecOld.RealTaxRate
        CustPin(ThisCnt, x) = PRRecOld.CustPin
        TownShip(ThisCnt, x) = PRRecOld.TownShip
        MORTCODE(ThisCnt, x) = PRRecOld.MORTCODE
        LotOrAcre(ThisCnt, x) = PRRecOld.LotOrAcre
        LASize(ThisCnt, x) = PRRecOld.LASize
        MortRec(ThisCnt, x) = PRRecOld.MortRec
        RDesc3(ThisCnt, x) = PRRecOld.RDesc3
        InternalPin(ThisCnt, x) = PRRecOld.InternalPin
        OptRevTax1(ThisCnt, x) = PRRecOld.OptRevTax1
        OptRevTax2(ThisCnt, x) = PRRecOld.OptRevTax2
        OptRevTax3(ThisCnt, x) = PRRecOld.OptRevTax3
        OverPayAmt(ThisCnt, x) = PRRecOld.OverPayAmt
        DueDate(ThisCnt, x) = PRRecOld.DueDate
        PostDate(ThisCnt, x) = PRRecOld.PostDate
        TransRec(ThisCnt, x) = PRRecOld.TransRec
        Opt1Desc(ThisCnt, x) = PRRecOld.Opt1Desc
        Opt2Desc(ThisCnt, x) = PRRecOld.Opt2Desc
        Opt3Desc(ThisCnt, x) = PRRecOld.Opt3Desc
        Padding(ThisCnt, x) = QPTrim$(PRRecOld.Padding)
        Comment(ThisCnt, x) = QPTrim$(PRRecOld.Comment)
      Next x
      Close PROldHandle
    End If
  Next y
  Close
  
  For x = 1 To RCnt
    OpenRealPostedReprintFile PRHandle, NumOfPRRecs, RFileNames(x)
    For y = 1 To NumOfPRRecs
      Get PRHandle, x, PRRec
      PRRec.CustRec = CustRec(x, y)
      PRRec.CustName = CustName(x, y)
      PRRec.CustAdd1 = CustAdd1(x, y)
      PRRec.CustAdd2 = CustAdd2(x, y)
      PRRec.CustAdd3 = CustAdd3(x, y)
      PRRec.CustZip = CustZip(x, y)
      PRRec.RDesc1 = RDesc1(x, y)
      PRRec.RDesc2 = RDesc2(x, y)
      PRRec.RealPin = RealPin(x, y)
      PRRec.RealValue = RealValue(x, y)
      PRRec.TotalValue = TotalValue(x, y)
      PRRec.ExptValue = ExptValue(x, y)
      PRRec.RealTaxDue = RealTaxDue(x, y)
      PRRec.BldgValue = BldgValue(x, y)
      PRRec.LateTaxDue = LateTaxDue(x, y)
      PRRec.TotalBillDue = TotalBillDue(x, y)
      PRRec.BillNumber = BillNumber(x, y)
      PRRec.TaxYear = TaxYear(x, y)
      PRRec.BillPrinted = BillPrinted(x, y)
      PRRec.RealPropRecord = RealPropRecord(x, y)
      PRRec.PriorYrBalance = PriorYrBalance(x, y)
      PRRec.RealTaxRate = RealTaxRate(x, y)
      PRRec.CustPin = CustPin(x, y)
      PRRec.TownShip = TownShip(x, y)
      PRRec.MORTCODE = MORTCODE(x, y)
      PRRec.LotOrAcre = LotOrAcre(x, y)
      PRRec.LASize = LASize(x, y)
      PRRec.MortRec = MortRec(x, y)
      PRRec.RDesc3 = RDesc3(x, y)
      PRRec.InternalPin = InternalPin(x, y)
      PRRec.OptRevTax1 = OptRevTax1(x, y)
      PRRec.OptRevTax2 = OptRevTax2(x, y)
      PRRec.OptRevTax3 = OptRevTax3(x, y)
      PRRec.OverPayAmt = OverPayAmt(x, y)
      PRRec.DueDate = DueDate(x, y)
      PRRec.PostDate = PostDate(x, y)
      PRRec.TransRec = TransRec(x, y)
      PRRec.Opt1Desc = Opt1Desc(x, y)
      PRRec.Opt2Desc = Opt2Desc(x, y)
      PRRec.Opt3Desc = Opt3Desc(x, y)
      PRRec.Padding = "Converted"
      PRRec.Comment = Comment(x, y)
      PRRec.Comment2 = ""
      PRRec.CommentPlace = ""
      PRRec.SetDscvry2No = "N"
      Put PRHandle, y, PRRec
    Next y
  Next x
  Close
  
  Call ConvertPersPostedBills
  
  MsgBox ("Finished.")


End Sub

Private Sub ConvertPersPostedBills()

  Dim PPRec As VAPPTaxBillType
  Dim PRHandle As Integer
  Dim NumOfPPRecs As Long
  Dim PPRecOld As VAPPTaxBillTypeOld
  Dim PROldHandle As Integer
  Dim NumOfOldPPRecs As Long
  Dim MyPath$, ThisFile$
  Dim x As Integer, y As Integer
  Dim PostRec As TaxBillPostDateType
  Dim PostHandle As Integer
  Dim NumOfPostRecs As Long
  Dim PCnt As Integer
  Dim ThisCnt As Integer
  Dim BigBCnt As Long
  
  ReDim PFileNames(1 To 1) As String
  OpenBillPostDateFile PostHandle, NumOfPostRecs
  For x = 1 To NumOfPostRecs
    Get PostHandle, x, PostRec
    ThisFile = PostRec.BackUpName
    If PostRec.BillType = "P" Then
      OpenPersPostedReprintFileOld PROldHandle, NumOfOldPPRecs, ThisFile
      If NumOfOldPPRecs > BigBCnt Then BigBCnt = NumOfOldPPRecs
      PCnt = PCnt + 1
      ReDim Preserve PFileNames(1 To PCnt) As String
      PFileNames(PCnt) = PostRec.BackUpName
    End If
  Next x
  
  ReDim CustRec(1 To PCnt, 1 To BigBCnt) As Long
  ReDim CustName(1 To PCnt, 1 To BigBCnt) As String
  ReDim CustAdd1(1 To PCnt, 1 To BigBCnt) As String
  ReDim CustAdd2(1 To PCnt, 1 To BigBCnt) As String
  ReDim CustAdd3(1 To PCnt, 1 To BigBCnt) As String
  ReDim CustZip(1 To PCnt, 1 To BigBCnt) As String
  ReDim RDesc1(1 To PCnt, 1 To BigBCnt) As String
  ReDim RDesc2(1 To PCnt, 1 To BigBCnt) As String
  ReDim RealPin(1 To PCnt, 1 To BigBCnt) As String
  ReDim PersValue(1 To PCnt, 1 To BigBCnt) As Double
  ReDim MHValue(1 To PCnt, 1 To BigBCnt) As Double
  ReDim MCValue(1 To PCnt, 1 To BigBCnt) As Double
  ReDim FEValue(1 To PCnt, 1 To BigBCnt) As Double
  ReDim MTValue(1 To PCnt, 1 To BigBCnt) As Double
  ReDim ExptValue(1 To PCnt, 1 To BigBCnt) As Double
  ReDim PersTaxDue(1 To PCnt, 1 To BigBCnt) As Double
  ReDim MHTaxDue(1 To PCnt, 1 To BigBCnt) As Double
  ReDim MCTaxDue(1 To PCnt, 1 To BigBCnt) As Double
  ReDim FETaxDue(1 To PCnt, 1 To BigBCnt) As Double
  ReDim MTTaxDue(1 To PCnt, 1 To BigBCnt) As Double
  ReDim LateTaxDue(1 To PCnt, 1 To BigBCnt) As Double
  ReDim TotalBillDue(1 To PCnt, 1 To BigBCnt) As Double
  ReDim BillNumber(1 To PCnt, 1 To BigBCnt) As Long
  ReDim TaxYear(1 To PCnt, 1 To BigBCnt) As Integer
  ReDim BillPrinted(1 To PCnt, 1 To BigBCnt) As Integer
  ReDim PersPropRecord(1 To PCnt, 1 To BigBCnt) As Long
  ReDim PriorYrBalance(1 To PCnt, 1 To BigBCnt) As Double
  ReDim PersTaxRate(1 To PCnt, 1 To BigBCnt) As Double
  ReDim MTTaxRate(1 To PCnt, 1 To BigBCnt) As Double
  ReDim MCTaxRate(1 To PCnt, 1 To BigBCnt) As Double
  ReDim FETaxRate(1 To PCnt, 1 To BigBCnt) As Double
  ReDim MHTaxRate(1 To PCnt, 1 To BigBCnt) As Double
  ReDim CustPin(1 To PCnt, 1 To BigBCnt) As Long
  ReDim ChillHowieFudge(1 To PCnt, 1 To BigBCnt) As Single
  ReDim PPTRAValue(1 To PCnt, 1 To BigBCnt) As Double
  ReDim PPTRADiscnt(1 To PCnt, 1 To BigBCnt) As Double
  ReDim InternalPin(1 To PCnt, 1 To BigBCnt) As Long
  ReDim OptRevTax1(1 To PCnt, 1 To BigBCnt) As Double
  ReDim OptRevTax2(1 To PCnt, 1 To BigBCnt) As Double
  ReDim OptRevTax3(1 To PCnt, 1 To BigBCnt) As Double
  ReDim OverPayAmt(1 To PCnt, 1 To BigBCnt) As Double
  ReDim RDesc3(1 To PCnt, 1 To BigBCnt) As String
  ReDim PersPin(1 To PCnt, 1 To BigBCnt) As String
  ReDim Prorate(1 To PCnt, 1 To BigBCnt) As String
  ReDim PersTaxNet(1 To PCnt, 1 To BigBCnt) As Double
  ReDim MultiYrVal(1 To PCnt, 1 To BigBCnt) As Integer
  ReDim DueDate(1 To PCnt, 1 To BigBCnt) As Integer
  ReDim OptRevDesc1(1 To PCnt, 1 To BigBCnt) As String
  ReDim OptRevDesc2(1 To PCnt, 1 To BigBCnt) As String
  ReDim OptRevDesc3(1 To PCnt, 1 To BigBCnt) As String
  ReDim PostDate(1 To PCnt, 1 To BigBCnt) As Integer
  ReDim TransRec(1 To PCnt, 1 To BigBCnt) As Long
  ReDim Comment(1 To PCnt, 1 To BigBCnt) As String
  ReDim Padding(1 To PCnt, 1 To BigBCnt) As String
  For y = 1 To NumOfPostRecs
    Get PostHandle, y, PostRec
    ThisFile = QPTrim$(PostRec.BackUpName)
    If PostRec.BillType = "P" Then
      OpenPersPostedReprintFileOld PROldHandle, NumOfOldPPRecs, ThisFile
      ThisCnt = ThisCnt + 1
      For x = 1 To NumOfOldPPRecs
        Get PROldHandle, x, PPRecOld
        CustRec(ThisCnt, x) = PPRecOld.CustRec
        CustName(ThisCnt, x) = PPRecOld.CustName
        CustAdd1(ThisCnt, x) = PPRecOld.CustAdd1
        CustAdd2(ThisCnt, x) = PPRecOld.CustAdd2
        CustAdd3(ThisCnt, x) = PPRecOld.CustAdd3
        CustZip(ThisCnt, x) = PPRecOld.CustZip
        RDesc1(ThisCnt, x) = PPRecOld.RDesc1
        RDesc2(ThisCnt, x) = PPRecOld.RDesc2
        RealPin(ThisCnt, x) = PPRecOld.RealPin
        PersValue(ThisCnt, x) = PPRecOld.PersValue
        MHValue(ThisCnt, x) = PPRecOld.MHValue
        MCValue(ThisCnt, x) = PPRecOld.MCValue
        FEValue(ThisCnt, x) = PPRecOld.FEValue
        MTValue(ThisCnt, x) = PPRecOld.MTValue
        ExptValue(ThisCnt, x) = PPRecOld.ExptValue
        PersTaxDue(ThisCnt, x) = PPRecOld.PersTaxDue
        MHTaxDue(ThisCnt, x) = PPRecOld.MHTaxDue
        MCTaxDue(ThisCnt, x) = PPRecOld.MCTaxDue
        FETaxDue(ThisCnt, x) = PPRecOld.FETaxDue
        MTTaxDue(ThisCnt, x) = PPRecOld.MTTaxDue
        LateTaxDue(ThisCnt, x) = PPRecOld.LateTaxDue
        TotalBillDue(ThisCnt, x) = PPRecOld.TotalBillDue
        BillNumber(ThisCnt, x) = PPRecOld.BillNumber
        TaxYear(ThisCnt, x) = PPRecOld.TaxYear
        BillPrinted(ThisCnt, x) = PPRecOld.BillPrinted
        PersPropRecord(ThisCnt, x) = PPRecOld.PersPropRecord
        PriorYrBalance(ThisCnt, x) = PPRecOld.PriorYrBalance
        PersTaxRate(ThisCnt, x) = PPRecOld.PersTaxRate
        MTTaxRate(ThisCnt, x) = PPRecOld.MTTaxRate
        MCTaxRate(ThisCnt, x) = PPRecOld.MCTaxRate
        FETaxRate(ThisCnt, x) = PPRecOld.FETaxRate
        MHTaxRate(ThisCnt, x) = PPRecOld.MHTaxRate
        CustPin(ThisCnt, x) = PPRecOld.CustPin
        ChillHowieFudge(ThisCnt, x) = PPRecOld.ChillHowieFudge
        PPTRAValue(ThisCnt, x) = PPRecOld.PPTRAValue
        PPTRADiscnt(ThisCnt, x) = PPRecOld.PPTRADiscnt
        InternalPin(ThisCnt, x) = PPRecOld.InternalPin
        OptRevTax1(ThisCnt, x) = PPRecOld.OptRevTax1
        OptRevTax2(ThisCnt, x) = PPRecOld.OptRevTax2
        OptRevTax3(ThisCnt, x) = PPRecOld.OptRevTax3
        OverPayAmt(ThisCnt, x) = PPRecOld.OverPayAmt
        RDesc3(ThisCnt, x) = PPRecOld.RDesc3
        PersPin(ThisCnt, x) = PPRecOld.PersPin
        Prorate(ThisCnt, x) = PPRecOld.Prorate
        PersTaxNet(ThisCnt, x) = PPRecOld.PersTaxNet
        MultiYrVal(ThisCnt, x) = PPRecOld.MultiYrVal
        DueDate(ThisCnt, x) = PPRecOld.DueDate
        OptRevDesc1(ThisCnt, x) = PPRecOld.OptRevDesc1
        OptRevDesc2(ThisCnt, x) = PPRecOld.OptRevDesc2
        OptRevDesc3(ThisCnt, x) = PPRecOld.OptRevDesc3
        PostDate(ThisCnt, x) = PPRecOld.PostDate
        TransRec(ThisCnt, x) = PPRecOld.TransRec
        Comment(ThisCnt, x) = QPTrim$(PPRecOld.Comment)
        Padding(ThisCnt, x) = QPTrim$(PPRecOld.Padding)
      Next x
      Close PROldHandle
    End If
  Next y
  Close
  
  For x = 1 To PCnt
    OpenPersPostedReprintFile PRHandle, NumOfPPRecs, PFileNames(x)
    For y = 1 To NumOfPPRecs
      Get PRHandle, x, PPRec
      PPRec.CustRec = CustRec(x, y)
      PPRec.CustName = CustName(x, y)
      PPRec.CustAdd1 = CustAdd1(x, y)
      PPRec.CustAdd2 = CustAdd2(x, y)
      PPRec.CustAdd3 = CustAdd3(x, y)
      PPRec.CustZip = CustZip(x, y)
      PPRec.RDesc1 = RDesc1(x, y)
      PPRec.RDesc2 = RDesc2(x, y)
      PPRec.RealPin = RealPin(x, y)
      PPRec.PersValue = PersValue(x, y)
      PPRec.MHValue = MHValue(x, y)
      PPRec.MCValue = MCValue(x, y)
      PPRec.FEValue = FEValue(x, y)
      PPRec.MTValue = MTValue(x, y)
      PPRec.ExptValue = ExptValue(x, y)
      PPRec.PersTaxDue = PersTaxDue(x, y)
      PPRec.MHTaxDue = MHTaxDue(x, y)
      PPRec.MCTaxDue = MCTaxDue(x, y)
      PPRec.FETaxDue = FETaxDue(x, y)
      PPRec.MTTaxDue = MTTaxDue(x, y)
      PPRec.LateTaxDue = LateTaxDue(x, y)
      PPRec.TotalBillDue = TotalBillDue(x, y)
      PPRec.BillNumber = BillNumber(x, y)
      PPRec.TaxYear = TaxYear(x, y)
      PPRec.BillPrinted = BillPrinted(x, y)
      PPRec.PersPropRecord = PersPropRecord(x, y)
      PPRec.PriorYrBalance = PriorYrBalance(x, y)
      PPRec.PersTaxRate = PersTaxRate(x, y)
      PPRec.MTTaxRate = MTTaxRate(x, y)
      PPRec.MCTaxRate = MCTaxRate(x, y)
      PPRec.FETaxRate = FETaxRate(x, y)
      PPRec.MHTaxRate = MHTaxRate(x, y)
      PPRec.CustPin = CustPin(x, y)
      PPRec.ChillHowieFudge = ChillHowieFudge(x, y)
      PPRec.PPTRAValue = PPTRAValue(x, y)
      PPRec.PPTRADiscnt = PPTRADiscnt(x, y)
      PPRec.InternalPin = InternalPin(x, y)
      PPRec.OptRevTax1 = OptRevTax1(x, y)
      PPRec.OptRevTax2 = OptRevTax2(x, y)
      PPRec.OptRevTax3 = OptRevTax3(x, y)
      PPRec.OverPayAmt = OverPayAmt(x, y)
      PPRec.RDesc3 = RDesc3(x, y)
      PPRec.PersPin = PersPin(x, y)
      PPRec.Prorate = Prorate(x, y)
      PPRec.PersTaxNet = PersTaxNet(x, y)
      PPRec.MultiYrVal = MultiYrVal(x, y)
      PPRec.DueDate = DueDate(x, y)
      PPRec.OptRevDesc1 = OptRevDesc1(x, y)
      PPRec.OptRevDesc2 = OptRevDesc2(x, y)
      PPRec.OptRevDesc3 = OptRevDesc3(x, y)
      PPRec.PostDate = PostDate(x, y)
      PPRec.TransRec = TransRec(x, y)
      PPRec.Comment = QPTrim$(Comment(x, y))
      PPRec.Comment2 = ""
      PPRec.CommentPlace = ""
      PPRec.SetDscvry2No = "N"
      PPRec.Padding = QPTrim$(Padding(x, y))
      Put PRHandle, y, PPRec
    Next y
  Next x
  Close

End Sub

Private Sub StripOutTrans()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim FileName As String
  Dim ThisFile As Integer
  Dim NextRec As Long
  Dim EmptyTaxTrans As TaxTransactionType
  Dim cnt As Long
  Dim PriorRec As Long
  
  FileName = "LunenErrors.txt"
  ThisFile = FreeFile
  Open FileName For Output As ThisFile
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For y = 1 To NumOfTCRecs
    Get TCHandle, y, TaxCust
    If TaxCust.Acct > 19930 Then
      If y = 19933 Or y = 20017 Or y = 20019 Or y = 20183 Or y = 20205 Or y = 20222 Or y = 20225 Then
        GoSub FixThese
      End If
      If y = 19957 Or y = 20087 Or y = 20088 Or y = 20156 Or y = 20214 Or y = 20259 Then
        GoSub FixThese
      End If
      NextRec = TaxCust.LastTrans
      If NextRec > 0 Then
        Get TTHandle, NextRec, TaxTrans
        If TaxTrans.TransDate < Date2Num("10/30/2007") Then
          TaxCust.LastTrans = 0
          Put TCHandle, y, TaxCust
        End If
      End If
      Do While NextRec > 0
        Get TTHandle, NextRec, TaxTrans
        If TaxTrans.TransDate < Date2Num("10/30/2007") And TaxTrans.CustomerRec <> y Then
          cnt = cnt + 1
          Print #ThisFile, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName)
          If PriorRec > 0 Then
            Get TTHandle, PriorRec, TaxTrans
            TaxTrans.LastTrans = 0
            Put TTHandle, PriorRec, TaxTrans
            GoTo Loop2
          End If
        End If
        PriorRec = NextRec
        NextRec = TaxTrans.LastTrans
      Loop
    End If
Loop2:
  Next y
    
  Close ThisFile
  
  MsgBox ("A total of " + CStr(cnt) + " transactions were stripped. Look for 'LunenErrors.txt' to see all transactions stripped.")
  Exit Sub
  
FixThese:
  Select Case y
    Case 19933
      Get TTHandle, 62440, TaxTrans
      TaxTrans.Revenue.Principle1Pd = 186.47
      Put TTHandle, 62440, TaxTrans
      
      Get TTHandle, 62441, TaxTrans
      TaxTrans.Amount = 0
      TaxTrans.Revenue.PrePaidAmt = 0
      TaxTrans.Revenue.PrePaidBal = 0
      TaxTrans.Revenue.PrePaidUsed = 0
      TaxTrans.Revenue.Principle1Pd = 0
      Put TTHandle, 62441, TaxTrans
    Case 20017
      Get TTHandle, 69763, TaxTrans
      TaxTrans.Revenue.Principle1Pd = 3.54
      Put TTHandle, 69763, TaxTrans
      
      Get TTHandle, 69764, TaxTrans
      TaxTrans.Amount = 0
      TaxTrans.Revenue.PrePaidAmt = 0
      TaxTrans.Revenue.PrePaidBal = 0
      TaxTrans.Revenue.PrePaidUsed = 0
      TaxTrans.Revenue.Principle1Pd = 0
      Put TTHandle, 69764, TaxTrans
    Case 20019
      Get TTHandle, 69766, TaxTrans
      TaxTrans.Revenue.Principle1Pd = 2.62
      Put TTHandle, 69766, TaxTrans
      
      Get TTHandle, 69767, TaxTrans
      TaxTrans.Amount = 0
      TaxTrans.Revenue.PrePaidAmt = 0
      TaxTrans.Revenue.PrePaidBal = 0
      TaxTrans.Revenue.PrePaidUsed = 0
      TaxTrans.Revenue.Principle1Pd = 0
      Put TTHandle, 69767, TaxTrans
    Case 20183
      Get TTHandle, 64999, TaxTrans
      TaxTrans.Revenue.Principle1Pd = 177.66
      Put TTHandle, 64999, TaxTrans
      
      Get TTHandle, 65000, TaxTrans
      TaxTrans.Amount = 0
      TaxTrans.Revenue.PrePaidAmt = 0
      TaxTrans.Revenue.PrePaidBal = 0
      TaxTrans.Revenue.PrePaidUsed = 0
      TaxTrans.Revenue.Principle1Pd = 0
      Put TTHandle, 65000, TaxTrans
    Case 20205
      Get TTHandle, 68322, TaxTrans
      TaxTrans.Revenue.Principle1Pd = 493.32
      Put TTHandle, 68322, TaxTrans
      
      Get TTHandle, 68323, TaxTrans
      TaxTrans.Amount = 0
      TaxTrans.Revenue.PrePaidAmt = 0
      TaxTrans.Revenue.PrePaidBal = 0
      TaxTrans.Revenue.PrePaidUsed = 0
      TaxTrans.Revenue.Principle1Pd = 0
      Put TTHandle, 68323, TaxTrans
    Case 20222
      Get TTHandle, 69768, TaxTrans
      TaxTrans.Revenue.Principle1Pd = 5.84
      Put TTHandle, 69768, TaxTrans
      
      Get TTHandle, 69769, TaxTrans
      TaxTrans.Amount = 0
      TaxTrans.Revenue.PrePaidAmt = 0
      TaxTrans.Revenue.PrePaidBal = 0
      TaxTrans.Revenue.PrePaidUsed = 0
      TaxTrans.Revenue.Principle1Pd = 0
      Put TTHandle, 69769, TaxTrans
    Case 20225
      Get TTHandle, 69770, TaxTrans
      TaxTrans.Revenue.Principle1Pd = 3.54
      Put TTHandle, 69770, TaxTrans
      
      Get TTHandle, 69771, TaxTrans
      TaxTrans.Amount = 0
      TaxTrans.Revenue.PrePaidAmt = 0
      TaxTrans.Revenue.PrePaidBal = 0
      TaxTrans.Revenue.PrePaidUsed = 0
      TaxTrans.Revenue.Principle1Pd = 0
      Put TTHandle, 69771, TaxTrans
    Case 20214
      Get TTHandle, 93935, TaxTrans
      TaxTrans.LastTrans = 89496
      Put TTHandle, 93935, TaxTrans
      
      'for cust #9375
      Get TTHandle, 48416, TaxTrans
      TaxTrans.LastTrans = 93934
      Put TTHandle, 48416, TaxTrans
      
      Get TTHandle, 93934, TaxTrans
      TaxTrans.LastTrans = 37796
      TaxTrans.CustomerRec = 9375
      TaxTrans.CustPin = 9375
      TaxTrans.BelongTo = 17227
      Put TTHandle, 93934, TaxTrans
      
    Case 19957
      Get TTHandle, 82130, TaxTrans
      TaxTrans.LastTrans = 62418
      Put TTHandle, 82130, TaxTrans
      
      'for cust #3049
      Get TCHandle, 3049, TaxCust
      TaxCust.LastTrans = 82128
      Put TCHandle, 3049, TaxCust
      Get TCHandle, y, TaxCust
      
      Get TTHandle, 82128, TaxTrans
      TaxTrans.LastTrans = 38423
      TaxTrans.CustomerRec = 3049
      TaxTrans.CustPin = 3049
      TaxTrans.BelongTo = 20790
      Put TTHandle, 82128, TaxTrans
      
    Case 20087
      Get TTHandle, 77146, TaxTrans
      TaxTrans.LastTrans = 62037
      Put TTHandle, 77146, TaxTrans
      
      'for cust #5073
      Get TCHandle, 5073, TaxCust
      TaxCust.LastTrans = 77145
      Put TCHandle, 5073, TaxCust
      Get TCHandle, y, TaxCust
      
      Get TTHandle, 36616, TaxTrans
      TaxTrans.LastTrans = 10707
      Put TTHandle, 36616, TaxTrans
      
      Get TTHandle, 77145, TaxTrans
      TaxTrans.LastTrans = 36616
      TaxTrans.CustomerRec = 5073
      TaxTrans.CustPin = 5073
      TaxTrans.BelongTo = 10707
      Put TTHandle, 77145, TaxTrans
      
    Case 20088
      Get TTHandle, 77150, TaxTrans
      TaxTrans.LastTrans = 62027
      Put TTHandle, 77150, TaxTrans
      
      'for cust #5074
      Get TCHandle, 5074, TaxCust
      TaxCust.LastTrans = 77148
      Put TCHandle, 5074, TaxCust
      Get TCHandle, y, TaxCust
      
      Get TTHandle, 77148, TaxTrans
      TaxTrans.LastTrans = 36614
      TaxTrans.CustomerRec = 5074
      TaxTrans.CustPin = 5074
      TaxTrans.BelongTo = 10697
      Put TTHandle, 77148, TaxTrans
      
      Get TTHandle, 36614, TaxTrans
      TaxTrans.LastTrans = 10697
      Put TTHandle, 36614, TaxTrans
      
    Case 20156
      Get TTHandle, 95282, TaxTrans
      TaxTrans.LastTrans = 89479
      Put TTHandle, 95282, TaxTrans
      
      'for cust #5334
      Get TCHandle, 5334, TaxCust
      TaxCust.LastTrans = 95281
      Put TCHandle, 5334, TaxCust
      Get TCHandle, y, TaxCust
      
      Get TTHandle, 95281, TaxTrans
      TaxTrans.LastTrans = 36833
      TaxTrans.CustomerRec = 5334
      TaxTrans.CustPin = 5334
      TaxTrans.BelongTo = 11976
      Put TTHandle, 95281, TaxTrans
      
    Case 20259
      TaxCust.LastTrans = 85708
      Put TCHandle, y, TaxCust
    
      'for cust #11206
      Get TCHandle, 11206, TaxCust
      TaxCust.LastTrans = 93770
      Put TCHandle, 11206, TaxCust
      Get TCHandle, y, TaxCust
      
      Get TTHandle, 93770, TaxTrans
      TaxTrans.LastTrans = 36741
      TaxTrans.CustomerRec = 11206
      TaxTrans.CustPin = 11206
      TaxTrans.BelongTo = 78
      Put TTHandle, 93770, TaxTrans
      
      Get TTHandle, 36741, TaxTrans
      TaxTrans.LastTrans = 85715
      Put TTHandle, 36741, TaxTrans
      
      Get TTHandle, 85715, TaxTrans
      TaxTrans.LastTrans = 11446
      TaxTrans.CustomerRec = 11206
      TaxTrans.CustPin = 11206
      TaxTrans.BelongTo = 11446
      Put TTHandle, 85715, TaxTrans
      
  End Select
  Return


End Sub

Private Sub UpdateAddressFieldsShort()
  Dim x As Long, y As Long
  Dim TextLine$
  Dim ThisRFile$
  Dim ThisPFile$
  Dim RHandle As Integer
  Dim PHandle As Integer
  Dim WordCnt As Integer
  Dim TextLen As Integer
  Dim ThisCh As String
  Dim ThisWord$
  Dim CntyNum As String
  Dim Add1 As String
  Dim Add2 As String
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim dlm As String
  Dim cnt As Integer
  Dim track As Integer
  
  If MsgBox("Did you make the last line(s) = 'End~~'?", vbYesNo) = vbNo Then
    Exit Sub
  End If
  dlm = "~"
  track = 0
  WordCnt = 0
  ReDim Words(1 To 1) As String
  frmVATaxShowPctComp.Label1 = "Real Addresses Update"
  frmVATaxShowPctComp.Show , Me

  If Exist("raddresses.csv") Then
    RHandle = FreeFile
    ThisRFile = "raddresses.csv"
    Open ThisRFile For Input As #RHandle
    Do While ThisWord <> "End"
      Line Input #RHandle, TextLine
      If InStr(TextLine, "End~~") Then Exit Do
      track = track + 1
    Loop
    Close
    OpenTaxCustFile TCHandle, NumOfTCRecs
    RHandle = FreeFile
    Open ThisRFile For Input As #RHandle
    Do While ThisWord <> "End"
      Line Input #RHandle, TextLine
      TextLen = Len(TextLine)
      TextLine = TextLine + dlm
      For x = 1 To TextLen + 1
        ThisCh = Mid(TextLine, x, 1)
        If ThisCh = dlm Then
          WordCnt = WordCnt + 1
          ReDim Preserve Words(1 To WordCnt) As String
          If WordCnt = 1 Then
            CntyNum = ThisWord
            ThisWord = ""
            GoTo NewWordReal
          ElseIf WordCnt = 2 Then
            Add1 = ThisWord
            ThisWord = ""
            GoTo NewWordReal
          ElseIf WordCnt = 3 Then
            Add2 = ThisWord
            GoSub SaveAdd
            Add1 = ""
            Add2 = ""
            CntyNum = ""
            ThisWord = ""
            WordCnt = 0
            GoTo NewLoopReal
          End If
        End If
        ThisWord = ThisWord + ThisCh
        If ThisWord = "End" Then Exit Do

NewWordReal:
      Next x
NewLoopReal:
    frmVATaxShowPctComp.ShowPctComp cnt, track
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
    End If
    Loop
  Else
    MsgBox ("The file 'raddresses.csv' cannot be found.")
    Exit Sub
  End If
  Unload frmVATaxShowPctComp
  MsgBox (CStr(cnt) & " real addresses were updated successfully.")
  Close
  DoEvents
  
  If Exist("paddresses.csv") Then
    track = 0
    WordCnt = 0
    cnt = 0
    ThisWord = ""
    ThisCh = ""
    ReDim Words(1 To 1) As String
    frmVATaxShowPctComp.Label1 = "Personal Addresses Update"
    frmVATaxShowPctComp.Show , Me
    PHandle = FreeFile
    ThisPFile = "paddresses.csv"
    Open ThisPFile For Input As #PHandle
    Do While ThisWord <> "End"
      Line Input #PHandle, TextLine
      If InStr(TextLine, "End~~") Then Exit Do
      track = track + 1
    Loop
    Close
    OpenTaxCustFile TCHandle, NumOfTCRecs
    PHandle = FreeFile
    Open ThisPFile For Input As #PHandle
    Do While ThisWord <> "End"
      Line Input #PHandle, TextLine
      TextLen = Len(TextLine)
      TextLine = TextLine + dlm
      For x = 1 To TextLen + 1
        ThisCh = Mid(TextLine, x, 1)
        If ThisCh = dlm Then
          WordCnt = WordCnt + 1
          ReDim Preserve Words(1 To WordCnt) As String
          If WordCnt = 1 Then
            CntyNum = ThisWord
            ThisWord = ""
            GoTo NewWordPers
          ElseIf WordCnt = 2 Then
            Add1 = ThisWord
            ThisWord = ""
            GoTo NewWordPers
          ElseIf WordCnt = 3 Then
            Add2 = ThisWord
            GoSub SaveAdd
            Add1 = ""
            Add2 = ""
            CntyNum = ""
            ThisWord = ""
            WordCnt = 0
            GoTo NewLoopPers
          End If
        End If
        ThisWord = ThisWord + ThisCh
        If ThisWord = "End" Then Exit Do
NewWordPers:
      Next x
NewLoopPers:
    frmVATaxShowPctComp.ShowPctComp cnt, track
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
    Loop
  Else
    MsgBox ("The file 'paddresses.csv' cannot be found.")
    Exit Sub
  End If
  Unload frmVATaxShowPctComp
  
  MsgBox (CStr(cnt) & " personal addresses were updated successfully.")
  Close
  Exit Sub
  
SaveAdd:
   cnt = cnt + 1
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If QPTrim$(TaxCust.CountyAcctString) = "" Then
      TaxCust.CountyAcctString = CStr(TaxCust.CountyAcct)
    End If
    If QPTrim$(TaxCust.CountyAcctString) = CntyNum Then
      TaxCust.Addr1 = Add1
      TaxCust.Addr2 = Add2
      Put TCHandle, x, TaxCust
      Exit For
    End If
Skip:
  Next x
  
  Return

End Sub

Private Sub UpdateAddressFieldsLong()
  Dim x As Long, y As Long
  Dim TextLine$
  Dim ThisRFile$
  Dim ThisPFile$
  Dim RHandle As Integer
  Dim PHandle As Integer
  Dim WordCnt As Integer
  Dim TextLen As Integer
  Dim ThisCh As String
  Dim ThisWord$
  Dim CntyNum As String
  Dim Add1 As String
  Dim Add2 As String
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim dlm As String
  Dim cnt As Integer
  Dim track As Integer
  
  If MsgBox("Did you make the last line(s) = 'End~~'?", vbYesNo) = vbNo Then
    Exit Sub
  End If
  dlm = "~"
  track = 0
  WordCnt = 0
  ReDim Words(1 To 1) As String
  frmVATaxShowPctComp.Label1 = "Real Addresses Update"
  frmVATaxShowPctComp.Show , Me

  If Exist("raddresses.csv") Then
    RHandle = FreeFile
    ThisRFile = "raddresses.csv"
    Open ThisRFile For Input As #RHandle
    Do While ThisWord <> "End"
      Line Input #RHandle, TextLine
      If InStr(TextLine, "End~~") Then Exit Do
      track = track + 1
    Loop
    Close
    OpenTaxCustFile TCHandle, NumOfTCRecs
    RHandle = FreeFile
    Open ThisRFile For Input As #RHandle
    Do While ThisWord <> "End"
      Line Input #RHandle, TextLine
      TextLen = Len(TextLine)
      TextLine = TextLine + dlm
      For x = 1 To TextLen + 1
        ThisCh = Mid(TextLine, x, 1)
        If ThisCh = dlm Then
          WordCnt = WordCnt + 1
          ReDim Preserve Words(1 To WordCnt) As String
          If WordCnt = 1 Then
            CntyNum = ThisWord
            ThisWord = ""
            GoTo NewWordReal
          ElseIf WordCnt = 2 Then
            Add1 = ThisWord
            ThisWord = ""
            GoTo NewWordReal
          ElseIf WordCnt = 3 Then
            Add2 = ThisWord
            GoSub SaveAdd
            Add1 = ""
            Add2 = ""
            CntyNum = ""
            ThisWord = ""
            WordCnt = 0
            GoTo NewLoopReal
          End If
        End If
        ThisWord = ThisWord + ThisCh
        If ThisWord = "End" Then Exit Do

NewWordReal:
      Next x
NewLoopReal:
    frmVATaxShowPctComp.ShowPctComp cnt, track
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
    End If
    Loop
  Else
    MsgBox ("The file 'raddresses.csv' cannot be found.")
    Exit Sub
  End If
  Unload frmVATaxShowPctComp
  MsgBox (CStr(cnt) & " real addresses were updated successfully.")
  Close
  DoEvents
  
  If Exist("paddresses.csv") Then
    track = 0
    WordCnt = 0
    cnt = 0
    ThisWord = ""
    ThisCh = ""
    ReDim Words(1 To 1) As String
    frmVATaxShowPctComp.Label1 = "Personal Addresses Update"
    frmVATaxShowPctComp.Show , Me
    PHandle = FreeFile
    ThisPFile = "paddresses.csv"
    Open ThisPFile For Input As #PHandle
    Do While ThisWord <> "End"
      Line Input #PHandle, TextLine
      If InStr(TextLine, "End~~") Then Exit Do
      track = track + 1
    Loop
    Close
    OpenTaxCustFile TCHandle, NumOfTCRecs
    PHandle = FreeFile
    Open ThisPFile For Input As #PHandle
    Do While ThisWord <> "End"
      Line Input #PHandle, TextLine
      TextLen = Len(TextLine)
      TextLine = TextLine + dlm
      For x = 1 To TextLen + 1
        ThisCh = Mid(TextLine, x, 1)
        If ThisCh = dlm Then
          WordCnt = WordCnt + 1
          ReDim Preserve Words(1 To WordCnt) As String
          If WordCnt = 1 Then
            CntyNum = ThisWord
            ThisWord = ""
            GoTo NewWordPers
          ElseIf WordCnt = 2 Then
            Add1 = ThisWord
            ThisWord = ""
            GoTo NewWordPers
          ElseIf WordCnt = 3 Then
            Add2 = ThisWord
            GoSub SaveAdd
            Add1 = ""
            Add2 = ""
            CntyNum = ""
            ThisWord = ""
            WordCnt = 0
            GoTo NewLoopPers
          End If
        End If
        ThisWord = ThisWord + ThisCh
        If ThisWord = "End" Then Exit Do
NewWordPers:
      Next x
NewLoopPers:
    frmVATaxShowPctComp.ShowPctComp cnt, track
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
    Loop
  Else
    MsgBox ("The file 'paddresses.csv' cannot be found.")
    Exit Sub
  End If
  Unload frmVATaxShowPctComp
  
  MsgBox (CStr(cnt) & " personal addresses were updated successfully.")
  Close
  Exit Sub
  
SaveAdd:
   cnt = cnt + 1
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If QPTrim$(TaxCust.CountyAcctString) = "" Then
      TaxCust.CountyAcctString = CStr(TaxCust.CountyAcct)
    End If
    If QPTrim$(TaxCust.CountyAcctString) = CntyNum Then
      TaxCust.Addr1 = Add1
      TaxCust.Addr2 = Add2
      Put TCHandle, x, TaxCust
    End If
Skip:
  Next x
  
  Return

End Sub

Private Sub FixLunenburgZeroYears()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim TheDate As String
  Dim AHandle As Integer
  Dim BillType$
  Dim TaxYear$, BillNum$
  Dim Amount As Double
  Dim cnt As Integer
  AHandle = FreeFile
  Open "yearzero.txt" For Output As AHandle
  Print #AHandle, "Updated Tax Year" & "~" & "Transaction #" & "~" & "Customer Pin" & "~" & "Trans Type" & "~" & "Amount"
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    TheDate = MakeRegDate(TaxTrans.TransDate)
    If TaxTrans.TaxYear = 0 Then
      TaxTrans.TaxYear = Mid(MakeRegDate(TaxTrans.TransDate), 7, 4)
      If Not IsNumeric(TaxTrans.TaxYear) Then GoTo Skip
      Put TTHandle, x, TaxTrans
      Amount = 0
      If TaxTrans.Amount > 0 Then
        Amount = TaxTrans.Amount
      ElseIf TaxTrans.Revenue.PrePaidAmt > 0 Then
        Amount = TaxTrans.Revenue.PrePaidAmt
      ElseIf TaxTrans.Revenue.PrePaidUsed > 0 Then
        Amount = TaxTrans.Revenue.PrePaidUsed
      End If
      cnt = cnt + 1
      GoSub GetTransType
      Print #AHandle, CStr(TaxTrans.TaxYear) & "~" & CStr(x) & "~" & CStr(TaxTrans.CustomerRec) & "~" & BillType & "~" & Using("$##,###.##", Amount)
    End If
Skip:
  Next x
  Close
  MsgBox ("A total of " & CStr(cnt) & " transactions were updated. Look for 'yearzero.txt' for spreadsheet.")
  Exit Sub

GetTransType:
  Select Case TaxTrans.TranType
  Case 1
    Select Case TaxTrans.BillType
    Case "R"
      BillType$ = "Real-Estate Bill"
    Case "P"
      BillType$ = "Personal Property Bill"
    Case "C"
      BillType$ = "Combined Bill"
    Case "M"
      BillType$ = "Manual Bill"
    End Select
    TaxYear$ = QPTrim$(Str$(TaxTrans.TaxYear))
  Case 2
    BillNum$ = ParseBillNum$(TaxTrans.Description)
    If Len(BillNum$) = 0 Then
      If QPTrim$(TaxTrans.Description) = "Prepay" Then
        BillType = "Prepayment"
      Else
        BillType$ = "Payment ??? "
      End If
    Else
      If TaxTrans.Revenue.PrePaidAmt > 0 Then
        BillType = "Pre/Payment on: "
      Else
        BillType$ = "Payment on: "
      End If
    End If
    BillType$ = BillType$ + BillNum$
  Case 3
    BillType$ = "Release"
  Case 4
    BillType$ = "Interest"
  Case 5
    BillType$ = "Penalty"
  Case 6
    BillType$ = "Collection/Ad Cost"
  Case 7
    BillType$ = "Adjust Paid Down"
  Case 9
    BillType$ = "Credit Applied at Billing"
  Case 13
    BillType$ = "Adjust Bill Down"
  Case 14
    BillType$ = "Adjust Bill Up"
  Case 21
    BillNum$ = ParseBillNum$(TaxTrans.Description)
    BillType$ = "Paid Bill Plus Prepay"
  Case 22
    BillType$ = "Prepayment"
  Case 10
    BillType = "Adjust Pay Dwn Affecting Credit"
  Case 24
    BillType = "Adjust Bill Up Affecting Credit"
  Case 11
    BillType = "Adjust Prepay Down" 'added 1/29/08
  Case Else
    BillType$ = Str$(TaxTrans.TranType) + "??"
    
  End Select

Return

End Sub
Private Sub FixPayPlusOPThatShouldHaveBeenApplied(ByVal BillTrans As Long, ByVal LastCustTrans As Long, ByVal ThisDate As String, ByVal Amount As Double, ByVal CustPin As Integer, ByVal TaxYear As Integer, ByVal PrePay As Double, ByVal BelongTo As Long, ByVal PropType As String, ByVal TopTrans As Long, ByVal BottomTrans As Long, ByVal BillNum As String)
  'this sub fixes places where a payment/overpayment trans should have been applied but wasn't
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim NewRec As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  NewRec = NumOfTTRecs + 1
'  If NewRec = 45599 Then Stop
  Get TTHandle, NewRec, TaxTrans
  TaxTrans.Revenue.Interest = 0#
  TaxTrans.Amount = 0
  TaxTrans.TransDate = Date2Num%(ThisDate)
  TaxTrans.TranType = 9
  TaxTrans.Revenue.CollectionPd = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.LateListPd = 0
  TaxTrans.Revenue.PenaltyPd = 0
  TaxTrans.Revenue.Principle1Pd = 0
  TaxTrans.Revenue.Principle2Pd = 0
  TaxTrans.Revenue.Principle3Pd = 0
  TaxTrans.Revenue.Principle4Pd = 0
  TaxTrans.Revenue.Principle5Pd = 0
  TaxTrans.Revenue.RevOpt1Pd = 0
  TaxTrans.Revenue.RevOpt2Pd = 0
  TaxTrans.Revenue.RevOpt3Pd = 0
  If Amount > TaxTrans.Revenue.Interest Then
    TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.Interest
    Amount = Amount - TaxTrans.Revenue.Interest
  Else
    TaxTrans.Revenue.InterestPd = Amount
    Amount = 0
    GoTo Paid
  End If
  If Amount > TaxTrans.Revenue.Penalty Then
    TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.Penalty
    Amount = Amount - TaxTrans.Revenue.Penalty
  Else
    TaxTrans.Revenue.PenaltyPd = Amount
    Amount = 0
    GoTo Paid
  End If
  If Amount > TaxTrans.Revenue.LateList Then
    TaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateList
    Amount = Amount - TaxTrans.Revenue.LateList
  Else
    TaxTrans.Revenue.LateListPd = Amount
    Amount = 0
    GoTo Paid
  End If
  If Amount > TaxTrans.Revenue.Collection Then
    TaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.Collection
    Amount = Amount - TaxTrans.Revenue.Collection
  Else
    TaxTrans.Revenue.CollectionPd = Amount
    Amount = 0
    GoTo Paid
  End If
  If Amount > TaxTrans.Revenue.Principle1 Then
    TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1
    Amount = Amount - TaxTrans.Revenue.Principle1
  Else
    TaxTrans.Revenue.Principle1Pd = Amount
    Amount = 0
    GoTo Paid
  End If
  If Amount > TaxTrans.Revenue.Principle2 Then
    TaxTrans.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2
    Amount = Amount - TaxTrans.Revenue.Principle2
  Else
    TaxTrans.Revenue.Principle2Pd = Amount
    Amount = 0
    GoTo Paid
  End If
  If Amount > TaxTrans.Revenue.Principle3 Then
    TaxTrans.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3
    Amount = Amount - TaxTrans.Revenue.Principle3
  Else
    TaxTrans.Revenue.Principle3Pd = Amount
    Amount = 0
    GoTo Paid
  End If
  If Amount > TaxTrans.Revenue.Principle4 Then
    TaxTrans.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4
    Amount = Amount - TaxTrans.Revenue.Principle4
  Else
    TaxTrans.Revenue.Principle4Pd = Amount
    Amount = 0
    GoTo Paid
  End If
  If Amount > TaxTrans.Revenue.Principle5 Then
    TaxTrans.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5
    Amount = Amount - TaxTrans.Revenue.Principle5
  Else
    TaxTrans.Revenue.Principle5Pd = Amount
    Amount = 0
    GoTo Paid
  End If
  If Amount > TaxTrans.Revenue.RevOpt1 Then
    TaxTrans.Revenue.RevOpt1Pd = TaxTrans.Revenue.RevOpt1
    Amount = Amount - TaxTrans.Revenue.RevOpt1
  Else
    TaxTrans.Revenue.RevOpt1Pd = Amount
    Amount = 0
    GoTo Paid
  End If
  If Amount > TaxTrans.Revenue.RevOpt2 Then
    TaxTrans.Revenue.RevOpt2Pd = TaxTrans.Revenue.RevOpt2
    Amount = Amount - TaxTrans.Revenue.RevOpt2
  Else
    TaxTrans.Revenue.RevOpt2Pd = Amount
    Amount = 0
    GoTo Paid
  End If
  If Amount > TaxTrans.Revenue.RevOpt3 Then
    TaxTrans.Revenue.RevOpt3Pd = TaxTrans.Revenue.RevOpt3
    Amount = Amount - TaxTrans.Revenue.RevOpt3
  Else
    TaxTrans.Revenue.RevOpt3Pd = Amount
    Amount = 0
    GoTo Paid
  End If
Paid:
'  If CustPin = 44 Then Stop
  TaxTrans.CustPin = CustPin
  TaxTrans.DiscXDate = 0
  TaxTrans.RealPin = ""
  TaxTrans.PersPin = ""
  TaxTrans.Posted2GL = "N"
  TaxTrans.TaxYear = TaxYear
  TaxTrans.DiscAmt = 0
  TaxTrans.OperNum = 0
  TaxTrans.Amount = 0
  TaxTrans.FromPrePay = PrePay
  TaxTrans.Description = "Credit Applied to Bill# " + BillNum
  TaxTrans.CustomerRec = CustPin
  TaxTrans.BelongTo = BelongTo
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidUsed = PrePay
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.InternalPin = CustPin
  TaxTrans.CntyPara = ""
  TaxTrans.CyclPara = ""
  TaxTrans.TShpPara = ""
  TaxTrans.BillType = PropType
  Put TTHandle, NewRec, TaxTrans
  
  If BillTrans <> LastCustTrans Then 'Last Trans was not the bill trans needing paying
  'this trans needs embedding in list
    Get TTHandle, TopTrans, TaxTrans 'TopTrans is the one above the bill
    TaxTrans.LastTrans = NewRec
    Put TTHandle, TopTrans, TaxTrans
    
    Get TTHandle, NewRec, TaxTrans
    TaxTrans.LastTrans = BottomTrans
    Put TTHandle, NewRec, TaxTrans
  Else 'place at very top
    Get TCHandle, CustPin, TaxCust
    TaxCust.LastTrans = NewRec
    Put TCHandle, CustPin, TaxCust
    
    Get TTHandle, NewRec, TaxTrans
    TaxTrans.LastTrans = BelongTo
    Put TTHandle, NewRec, TaxTrans
  End If
  
  Close TTHandle
  Close TCHandle
End Sub

Private Sub MakeMasterBalEqualBillBal()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim SaveRec As Long
  Dim BelongTo As Long
  Dim P1 As Double
  Dim P2 As Double
  Dim P3 As Double
  Dim P4 As Double
  Dim P5 As Double
  Dim Pen As Double
  Dim Adv As Double
  Dim LL As Double
  Dim Intr As Double
  Dim P1P As Double
  Dim P2P As Double
  Dim P3P As Double
  Dim P4P As Double
  Dim P5P As Double
  Dim PenP As Double
  Dim AdvP As Double
  Dim LLP As Double
  Dim IntrP As Double
  
  Call FixErrorInOPAtBilling
  Exit Sub
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If TaxTrans.TranType = 1 And TaxTrans.CustomerRec = 3600 Then
    P1P = 0
    P2P = 0
    P3P = 0
    P4P = 0
    P5P = 0
    PenP = 0
    AdvP = 0
    LLP = 0
    IntrP = 0
    P1 = 0
    P2 = 0
    P3 = 0
    P4 = 0
    P5 = 0
    Pen = 0
    Adv = 0
    LL = 0
    Intr = 0
      For y = 1 To NumOfTTRecs
      Get TTHandle, y, TaxTrans
      If TaxTrans.BelongTo = x Then
        Select Case TaxTrans.TranType
        Case 2
          P1P = P1P + TaxTrans.Revenue.Principle1Pd
          P2P = P2P + TaxTrans.Revenue.Principle2Pd
          P3P = P3P + TaxTrans.Revenue.Principle3Pd
          P4P = P4P + TaxTrans.Revenue.Principle4Pd
          P5P = P5P + TaxTrans.Revenue.Principle5Pd
          PenP = PenP + TaxTrans.Revenue.PenaltyPd
          AdvP = AdvP + TaxTrans.Revenue.CollectionPd
          LLP = LLP + TaxTrans.Revenue.LateListPd
          IntrP = IntrP + TaxTrans.Revenue.InterestPd
        Case 4
          Intr = Intr + TaxTrans.Revenue.Interest
        Case 5
          Pen = Pen + TaxTrans.Revenue.Penalty
        Case 6
          Adv = Adv + TaxTrans.Revenue.Collection
        Case Else
          Exit For
        End Select
      End If
      Next y
      If y <= NumOfTTRecs Then
'        GoTo Skip
      Else
        Get TTHandle, x, TaxTrans
        If TaxTrans.Revenue.Principle1Pd <> P1P Then
          TaxTrans.Revenue.Principle1Pd = P1P
        End If
        If TaxTrans.Revenue.Principle2Pd <> P2P Then
          TaxTrans.Revenue.Principle2Pd = P2P
        End If
         If TaxTrans.Revenue.Principle3Pd <> P3P Then
          TaxTrans.Revenue.Principle3Pd = P3P
        End If
        If TaxTrans.Revenue.Principle4Pd <> P4P Then
          TaxTrans.Revenue.Principle1Pd = P4P
        End If
        If TaxTrans.Revenue.Principle5Pd <> P5P Then
          TaxTrans.Revenue.Principle5Pd = P5P
        End If
        If TaxTrans.Revenue.InterestPd <> IntrP Then
          TaxTrans.Revenue.InterestPd = IntrP
        End If
        If TaxTrans.Revenue.PenaltyPd <> PenP Then
          TaxTrans.Revenue.PenaltyPd = PenP
        End If
        If TaxTrans.Revenue.CollectionPd <> AdvP Then
          TaxTrans.Revenue.CollectionPd = AdvP
        End If
        If TaxTrans.Revenue.LateListPd <> LLP Then
          TaxTrans.Revenue.LateListPd = LLP
        End If
        If TaxTrans.Revenue.Interest <> Intr Then
          TaxTrans.Revenue.Interest = Intr
        End If
        If TaxTrans.Revenue.Penalty <> Pen Then
          TaxTrans.Revenue.Penalty = Pen
        End If
        If TaxTrans.Revenue.Collection <> Adv Then
          TaxTrans.Revenue.Collection = Adv
        End If
        If TaxTrans.Revenue.LateList <> LL Then
          TaxTrans.Revenue.LateList = LL
        End If
        Put TTHandle, x, TaxTrans
      End If
      
    End If
Skip:
  Next x
  
  Close
  MsgBox ("Finished.")

End Sub
Private Sub BuildMBvsCustHistArr()
  Dim PHandle As Integer
  Dim TextLine$
  Dim LineLen As Integer
  Dim NumHold As String
  Dim ch As String
  Dim x As Integer
  
  CArrCnt = 0
  ReDim CArr(1 To 1) As Long
  PHandle = FreeFile
'  If Not Exist(App.Path + "\txbalerrors.txt") Then
'    MsgBox ("Please put an empty text file named txbalerrors.txt in the Citipak folder.")
'    Exit Sub
'  End If
  
  Open "txbalerrors.txt" For Input As #PHandle  ' Open file.

'  Line Input #PHandle, TextLine   ' Read first line into TextLine.
   Do While Not eof(PHandle)  ' Loop until end of file.
     Line Input #PHandle, TextLine   ' Read next line into Textline.
     LineLen = Len(TextLine)
     For x = 1 To LineLen
       ch = Mid(TextLine, x, 1)
       If ch = "," Then
         CArrCnt = CArrCnt + 1
         ReDim Preserve CArr(1 To CArrCnt) As Long
         CArr(CArrCnt) = CInt(NumHold)
         NumHold = ""
       Else
         NumHold = NumHold + ch
       End If
     Next x
   Loop
   Close PHandle
   
   
End Sub
Private Sub BuildCrossTransArr()
  Dim PHandle As Integer
  Dim TextLine$
  Dim LineLen As Integer
  Dim NumHold As String
  Dim ch As String
  Dim x As Integer
  Dim Start As Boolean
  Dim Num1 As String
  Dim Num2 As String
  Dim Num3 As String
  
  CrossCnt = 0
  ReDim CrossArr(1 To 1) As Long
  ReDim CrossGoodArr(1 To 1) As Long
  ReDim CrossBadArr(1 To 1) As Long
  PHandle = FreeFile
  Open App.Path + "\TransWithCrossCusts.txt" For Input As #PHandle  ' Open file.
  Start = False
   Do While Not eof(PHandle)  ' Loop until end of file.
     Line Input #PHandle, TextLine   ' Read next line into Textline.
     LineLen = Len(TextLine)
     Num1 = ""
     Num2 = ""
     Num3 = ""
     For x = 1 To LineLen
       ch = Mid(TextLine, x, 1)
       If ch <> "~" Then
         NumHold = NumHold + ch
       Else
         If Len(Num1) = 0 Then
           Num1 = NumHold
           CrossCnt = CrossCnt + 1
           ReDim Preserve CrossArr(1 To CrossCnt) As Long
           CrossArr(CrossCnt) = CLng(NumHold)
           NumHold = ""
         ElseIf Len(Num2) = 0 Then
           Num2 = NumHold
           ReDim Preserve CrossGoodArr(1 To CrossCnt) As Long
           CrossGoodArr(CrossCnt) = CLng(NumHold)
           NumHold = ""
         ElseIf Len(Num3) = 0 Then
           Num3 = NumHold
           ReDim Preserve CrossBadArr(1 To CrossCnt) As Long
           CrossBadArr(CrossCnt) = CLng(NumHold)
           NumHold = ""
           Exit For
         End If
       End If
     Next x
   Loop
   Close PHandle
   
End Sub
Private Sub FixErrorInOPAtBilling()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim SaveRec As Long
  Dim PPAmount As Double
  Dim BelongTo As Long
  Dim BDate As Integer
  Dim EDate As Integer
  Dim CustBal As Double
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim LastRec As Long
  Dim BTrans As Long
  Dim ETrans As Long
  Dim BillNum As String
  Dim PropType As String
  Dim AHandle As Integer
  Dim cnt As Long
  Dim BillAmt As Double
  Dim PHandle As Integer
  Dim TextLine$
  Dim LineLen As Integer
  Dim NumHold As String
  Dim ch As String
  Dim RealBal As Double
  Dim PersBal As Double
  Dim LastTrans2 As Long
  Dim LastTrans As Long
  Dim CustLastTrans As Long
  Dim NewLastTrans As Long
  
'  Call ChangeCreditAtBillingtoRegPayment(16080)
  
  Call CompareMBWithCustHistory
  BDate = Date2Num("06/01/2006")
  EDate = Date2Num("06/30/2007")
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  AHandle = FreeFile
  Open "OPErrorFix.txt" For Output As AHandle
  
  frmVATaxShowPctComp.Label1 = "Processing"
  frmVATaxShowPctComp.Show , Me
  For x = 1 To CArrCnt
    Get TCHandle, CArr(x), TaxCust
    CustBal = GetCustBalance(CArr(x), 0)
    RealBal = GetCustRealBalance(CArr(x), 0)
    PersBal = GetCustPersBalance(CArr(x), 0)
    LastTrans = TaxCust.LastTrans
    CustLastTrans = TaxCust.LastTrans
    Do While LastTrans > 0
      Get TTHandle, LastTrans, TaxTrans
      If TaxTrans.TranType = 21 And TaxTrans.TransDate >= BDate And TaxTrans.TransDate <= EDate Then
        PPAmount = 0
        Call ChangePayOverPayToJustPay(LastTrans, CustLastTrans, PPAmount)
        Print #AHandle, CStr(CArr(x)) + "~" + CStr(PPAmount) '+ "~" + BillNum
        cnt = cnt + 1
        GoTo NextTrans
      End If
      Get TTHandle, LastTrans, TaxTrans
      LastTrans = TaxTrans.LastTrans
    Loop
NextTrans:
    frmVATaxShowPctComp.ShowPctComp x, CArrCnt
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
    
  
  Close
  MsgBox ("A total of " + CStr(cnt) + " customers have been updated. Look for OPErrorFix.txt in the Citipak folder for details.")

End Sub
Private Sub RelinkBelongTosWithBills(Optional ByVal PrintIt As Boolean = False)
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim NextRec As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim AHandle As Integer
  Dim BHandle As Integer
  Dim cnt As Integer
  Dim BelongTo As Long
  Dim P1Pd As Double
  Dim P2Pd As Double
  Dim P3Pd As Double
  Dim P4Pd As Double
  Dim P5Pd As Double
  Dim IntPd As Double
  Dim AdvPd As Double
  Dim LLPd As Double
  Dim PenPd As Double
  Dim Opt1Pd As Double
  Dim Opt2Pd As Double
  Dim Opt3Pd As Double
  Dim P1 As Double
  Dim P2 As Double
  Dim P3 As Double
  Dim P4 As Double
  Dim P5 As Double
  Dim Intr As Double
  Dim Adv As Double
  Dim LL As Double
  Dim Pen As Double
  Dim Opt1 As Double
  Dim Opt2 As Double
  Dim Opt3 As Double
  Dim Paid1 As Double
  Dim Billed1 As Double
  Dim Balance1 As Double
  Dim Paid2 As Double
  Dim Billed2 As Double
  Dim Balance2 As Double
  Dim BillAmt As Double
  Dim SaveRec As Long
  Dim ProcessCrosses As Boolean
  Dim ThisCust As Integer
  Dim ABDTot As Double
  Dim BillTot As Double
  Dim AdjustedAll As Boolean
  Dim AdjustedABU As Integer
  
  ProcessCrosses = False
  ReDim CArr(1 To 1) As Long
  CArrCnt = 0
  Call CompareMBWithCustHistory
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  AHandle = FreeFile
  Open "ReindexedBills.txt" For Output As AHandle
  BHandle = FreeFile
  Open "TransWithCrossCusts.txt" For Output As BHandle
  frmVATaxShowPctComp.Label1 = "Relinking Bill Transactions"
  frmVATaxShowPctComp.Show , Me
'  CArr(1) = 880
'  CArrCnt = 1
  For x = 1 To CArrCnt
   Get TCHandle, CArr(x), TaxCust
   If TaxCust.Deleted <> 0 Then GoTo GSkip
    
   NextRec = TaxCust.LastTrans
   Do While NextRec > 0
     Get TTHandle, NextRec, TaxTrans
'      If NextRec = 22292 Then Stop
      If TaxTrans.TranType = 1 Then
        If InStr(UCase(TaxTrans.Description), "INITIALIZE") > 0 Then
          GoTo GSkip
        End If
      BillAmt = TaxTrans.Amount
      P1Pd = 0
      P2Pd = 0
      P3Pd = 0
      P4Pd = 0
      P5Pd = 0
      IntPd = 0
      AdvPd = 0
      LLPd = 0
      PenPd = 0
      Opt1Pd = 0
      Opt2Pd = 0
      Opt3Pd = 0
      Intr = 0
      Adv = 0
      LL = 0
      Pen = 0
      P1 = TaxTrans.Revenue.Principle1
      P2 = TaxTrans.Revenue.Principle2
      P3 = TaxTrans.Revenue.Principle3
      P4 = TaxTrans.Revenue.Principle4
      P5 = TaxTrans.Revenue.Principle5
      Opt1 = TaxTrans.Revenue.RevOpt1
      Opt2 = TaxTrans.Revenue.RevOpt2
      Opt3 = TaxTrans.Revenue.RevOpt3
      For y = 1 To NumOfTTRecs 'go thru all trans to locate all that apply even if they
      'are not in this customer's queue
        Get TTHandle, y, TaxTrans
'        If y = 8304 Then Stop
'        If TaxTrans.TranType = 13 Then TaxTrans.TranType = 2
        TaxTrans.CustomerRec = TaxTrans.CustomerRec
        Select Case TaxTrans.TranType
          Case 8, 22, 10, 11, 12, 24, 30
            GoTo Skip
          Case Else
        End Select
        ThisCust = TaxTrans.CustomerRec
        If ThisCust = 0 Then
          ThisCust = TaxTrans.CustPin
        End If
        If TaxTrans.BelongTo = NextRec Then
          If CArr(x) <> ThisCust Then 'TaxTrans.CustomerRec Then
            Print #BHandle, CStr(CArr(x)) + "~" + CStr(NextRec) + "~" + CStr(y) + "~"
            ProcessCrosses = True
            GoTo GSkip
          End If
          If TaxTrans.TranType = 13 Then 'adjust bill down
            GoSub AdjustBillDown
            If AdjustedAll = True Then
              AdjustedAll = False
              Print #AHandle, CStr(CArr(x)) + "~" + CStr(NextRec)
              cnt = cnt + 1
            End If
           GoTo GSkip
          End If
          If TaxTrans.TranType = 14 Then 'adjust bill up
            AdjustedABU = 0
            GoSub AdjustBillUp
            If AdjustedABU = 1 Then
              Print #AHandle, CStr(CArr(x)) + "~" + CStr(NextRec)
              cnt = cnt + 1
            ElseIf AdjustedABU = 2 Then
              Print #AHandle, CStr(CArr(x)) + "~" + CStr(NextRec) + "~ is an adjust bill up that could not be fixed."
            End If
            GoTo GSkip
          End If
          If TaxTrans.TranType = 7 Then
            GoSub DifferentAdjPayDown
            GoTo Skip
          End If
          TaxTrans.BelongTo = TaxTrans.BelongTo 'collect all payments in list
          'for this bill
          P1Pd = P1Pd + TaxTrans.Revenue.Principle1Pd
          P2Pd = P2Pd + TaxTrans.Revenue.Principle2Pd
          P3Pd = P3Pd + TaxTrans.Revenue.Principle3Pd
          P4Pd = P4Pd + TaxTrans.Revenue.Principle4Pd
          P5Pd = P5Pd + TaxTrans.Revenue.Principle5Pd
          IntPd = IntPd + TaxTrans.Revenue.InterestPd
          AdvPd = AdvPd + TaxTrans.Revenue.CollectionPd
          LLPd = LLPd + TaxTrans.Revenue.LateListPd
          PenPd = PenPd + TaxTrans.Revenue.PenaltyPd
          Opt1Pd = Opt1Pd + TaxTrans.Revenue.RevOpt1Pd
          Opt2Pd = Opt2Pd + TaxTrans.Revenue.RevOpt2Pd
          Opt3Pd = Opt3Pd + TaxTrans.Revenue.RevOpt3Pd
          Intr = Intr + TaxTrans.Revenue.Interest
          Adv = Adv + TaxTrans.Revenue.Collection
          LL = LL + TaxTrans.Revenue.LateList
          Pen = Pen + TaxTrans.Revenue.Penalty
          TaxTrans.Revenue.Principle1 = TaxTrans.Revenue.Principle1
          TaxTrans.TranType = TaxTrans.TranType
        End If
Skip:
      Next y
      Get TTHandle, NextRec, TaxTrans 'now compare collected data with that in the bill
      Billed1 = P1 + P2 + P3 + P4 + P5 + Opt1 + Opt2 + Opt3 + Intr + Adv + LL + Pen
      Paid1 = P1Pd + P2Pd + P3Pd + P4Pd + P5Pd + Opt1Pd + Opt2Pd + Opt3Pd + IntPd + AdvPd + LLPd + PenPd
      Paid1 = Paid1 + TaxTrans.PPTRADisc
      Balance1 = OldRound(Billed1 - Paid1)
      Billed2 = TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2
      Billed2 = Billed2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4
      Billed2 = Billed2 + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.Interest
      Billed2 = Billed2 + TaxTrans.Revenue.Collection + TaxTrans.Revenue.LateList
      Billed2 = Billed2 + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.RevOpt1
      Billed2 = Billed2 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3
      Paid2 = TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd
      Paid2 = Paid2 + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd
      Paid2 = Paid2 + TaxTrans.Revenue.Principle5Pd + TaxTrans.Revenue.InterestPd
      Paid2 = Paid2 + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd
      Paid2 = Paid2 + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.RevOpt1Pd
      Paid2 = Paid2 + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.PPTRADisc
      Balance2 = OldRound(Billed2 - Paid2)
      If Balance1 = Balance2 Then
        GoTo GSkip
      Else
        Dim TBal As Double
        Dim TCnt As Integer
        Dim HoldRec As Long
        TCnt = 0
        'go back thru and see if anything had been zeroed out that needs to go back
        If Pen <> TaxTrans.Revenue.Penalty Then
          TBal = Abs(Pen - TaxTrans.Revenue.Penalty)
          SaveRec = TaxCust.LastTrans
          Do While SaveRec > 0
            Get TTHandle, SaveRec, TaxTrans
            If TaxTrans.TranType = 5 And SaveRec > NextRec Then
              If TaxTrans.Amount = 0 And TaxTrans.Revenue.Penalty = 0 Then
                HoldRec = SaveRec
                TCnt = TCnt + 1
              End If
            End If
            SaveRec = TaxTrans.LastTrans
          Loop
          If TCnt = 1 Then 'don't want to do this if there are more than one cleared
            'just for safety reasons
            Get TTHandle, HoldRec, TaxTrans
            TaxTrans.Amount = TBal
            TaxTrans.Revenue.Penalty = TBal
            Put TTHandle, HoldRec, TaxTrans
            
            Get TTHandle, NextRec, TaxTrans
            GoTo ClearBill
          End If
        End If
        If Intr <> TaxTrans.Revenue.Interest Then
          TBal = Abs(Intr - TaxTrans.Revenue.Interest)
          SaveRec = TaxCust.LastTrans
          Do While SaveRec > 0
            Get TTHandle, SaveRec, TaxTrans
            If TaxTrans.TranType = 4 And SaveRec > NextRec Then
              If TaxTrans.Amount = 0 And TaxTrans.Revenue.Interest = 0 Then
                HoldRec = SaveRec
                TCnt = TCnt + 1
              End If
            End If
            SaveRec = TaxTrans.LastTrans
          Loop
          If TCnt = 1 Then
            Get TTHandle, HoldRec, TaxTrans
            TaxTrans.Amount = TBal
            TaxTrans.Revenue.Interest = TBal
            Put TTHandle, HoldRec, TaxTrans
            
            Get TTHandle, NextRec, TaxTrans
            GoTo ClearBill
          End If
        End If
        If Adv <> TaxTrans.Revenue.Collection Then
          TBal = Abs(Adv - TaxTrans.Revenue.Penalty)
          SaveRec = TaxCust.LastTrans
          Do While SaveRec > 0
            Get TTHandle, SaveRec, TaxTrans
            If TaxTrans.TranType = 6 And SaveRec > NextRec Then
              If TaxTrans.Amount = 0 And TaxTrans.Revenue.Collection = 0 Then
                HoldRec = SaveRec
                TCnt = TCnt + 1
              End If
            End If
            SaveRec = TaxTrans.LastTrans
          Loop
          If TCnt = 1 Then
            Get TTHandle, HoldRec, TaxTrans
            TaxTrans.Amount = TBal
            TaxTrans.Revenue.Collection = TBal
            Put TTHandle, HoldRec, TaxTrans
            
            Get TTHandle, NextRec, TaxTrans
            GoTo ClearBill
          End If
        End If
      End If
Checked:
      TaxTrans.Revenue.Principle1 = P1
      TaxTrans.Revenue.Principle2 = P2
      TaxTrans.Revenue.Principle3 = P3
      TaxTrans.Revenue.Principle4 = P4
      TaxTrans.Revenue.Principle5 = P5
      TaxTrans.Revenue.RevOpt1 = Opt1
      TaxTrans.Revenue.RevOpt2 = Opt2
      TaxTrans.Revenue.RevOpt3 = Opt3
      TaxTrans.Revenue.Interest = Intr
      TaxTrans.Revenue.Collection = Adv
      TaxTrans.Revenue.LateList = LL
      TaxTrans.Revenue.Penalty = Pen
      TaxTrans.Revenue.Principle1Pd = P1Pd
      TaxTrans.Revenue.Principle2Pd = P2Pd
      TaxTrans.Revenue.Principle3Pd = P3Pd
      TaxTrans.Revenue.Principle4Pd = P4Pd
      TaxTrans.Revenue.Principle5Pd = P5Pd
      TaxTrans.Revenue.RevOpt1Pd = Opt1Pd
      TaxTrans.Revenue.RevOpt2Pd = Opt2Pd
      TaxTrans.Revenue.RevOpt3Pd = Opt3Pd
      TaxTrans.Revenue.InterestPd = IntPd
      TaxTrans.Revenue.CollectionPd = AdvPd
      TaxTrans.Revenue.LateListPd = LLPd
      TaxTrans.Revenue.PenaltyPd = PenPd
      
ClearBill:
      Print #AHandle, CStr(CArr(x)) + "~" + CStr(NextRec)
      Put TTHandle, NextRec, TaxTrans
      cnt = cnt + 1
GSkip:
     End If
     Get TTHandle, NextRec, TaxTrans
     NextRec = TaxTrans.LastTrans
   Loop
    frmVATaxShowPctComp.ShowPctComp x, CArrCnt
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  
  Close
  MsgBox ("A total of " + CStr(cnt) + " bills have been processed. Look for ReindexedBills.txt in the Citipak folder for results.")
'  If ProcessCrosses = True Then
'    frmVATaxMsgWOpts.Label1.Caption = "Transactions were discovered saved under incorrect customers. OK to correct these?"
'    frmVATaxMsgWOpts.Label1.Top = 900
'    frmVATaxMsgWOpts.cmdCont.Text = "F10 Correct"
'    frmVATaxMsgWOpts.cmdExit.Text = "ESC No"
'    frmVATaxMsgWOpts.Show vbModal
'    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
'      Unload frmVATaxMsgWOpts
'      Call FindAndFixTransWithCrossCusts
'    Else
'      Unload frmVATaxMsgWOpts
'    End If
'  End If
  Unload frmVATaxShowPctComp

  Exit Sub
  
AdjustBillUp:
  Dim ABUTot As Double
  Dim ABUBillTot As Double
  Dim ABUCnt As Integer
  BelongTo = TaxTrans.BelongTo
  AdjustedAll = False
  ABUTot = TaxTrans.Revenue.Principle1
  ABUTot = ABUTot + TaxTrans.Revenue.Principle2
  ABUTot = ABUTot + TaxTrans.Revenue.Principle3
  ABUTot = ABUTot + TaxTrans.Revenue.Principle4
  ABUTot = ABUTot + TaxTrans.Revenue.Principle5
  ABUTot = ABUTot + TaxTrans.Revenue.RevOpt1
  ABUTot = ABUTot + TaxTrans.Revenue.RevOpt2
  ABUTot = ABUTot + TaxTrans.Revenue.RevOpt3
  ABUTot = ABUTot + TaxTrans.Revenue.Interest
  ABUTot = ABUTot + TaxTrans.Revenue.Collection
  ABUTot = ABUTot + TaxTrans.Revenue.LateList
  ABUTot = ABUTot + TaxTrans.Revenue.Penalty
  Get TTHandle, BelongTo, TaxTrans
'  BillTot = TaxTrans.Amount
  ABUCnt = 0
  BillTot = TaxTrans.Revenue.Principle1
  If TaxTrans.Revenue.Principle1 > 0 Then ABUCnt = ABUCnt + 1
  BillTot = BillTot + TaxTrans.Revenue.Principle2
  If TaxTrans.Revenue.Principle2 > 0 Then ABUCnt = ABUCnt + 1
  BillTot = BillTot + TaxTrans.Revenue.Principle3
  If TaxTrans.Revenue.Principle3 > 0 Then ABUCnt = ABUCnt + 1
  BillTot = BillTot + TaxTrans.Revenue.Principle4
  If TaxTrans.Revenue.Principle4 > 0 Then ABUCnt = ABUCnt + 1
  BillTot = BillTot + TaxTrans.Revenue.Principle5
  If TaxTrans.Revenue.Principle5 > 0 Then ABUCnt = ABUCnt + 1
  BillTot = BillTot + TaxTrans.Revenue.RevOpt1
  If TaxTrans.Revenue.RevOpt1 > 0 Then ABUCnt = ABUCnt + 1
  BillTot = BillTot + TaxTrans.Revenue.RevOpt2
  If TaxTrans.Revenue.RevOpt2 > 0 Then ABUCnt = ABUCnt + 1
  BillTot = BillTot + TaxTrans.Revenue.RevOpt3
  If TaxTrans.Revenue.RevOpt3 > 0 Then ABUCnt = ABUCnt + 1
  BillTot = BillTot + TaxTrans.Revenue.Interest
  If TaxTrans.Revenue.Interest > 0 Then ABUCnt = ABUCnt + 1
  BillTot = BillTot + TaxTrans.Revenue.Collection
  If TaxTrans.Revenue.Collection > 0 Then ABUCnt = ABUCnt + 1
  BillTot = BillTot + TaxTrans.Revenue.LateList
  If TaxTrans.Revenue.LateList > 0 Then ABUCnt = ABUCnt + 1
  BillTot = BillTot + TaxTrans.Revenue.Penalty
  If TaxTrans.Revenue.Penalty > 0 Then ABUCnt = ABUCnt + 1
  ABUBillTot = (BillTot - TaxTrans.PPTRADisc) - ABUTot
  If TaxTrans.Amount <> ABUBillTot Then 'if there is only one revenue affected then
  'the total should be equal to the total of the adjustedbillup and the original bill amount
    If ABUCnt = 1 Then
      If TaxTrans.Revenue.Principle1 > 0 Then
        TaxTrans.Revenue.Principle1 = ABUTot + TaxTrans.Amount + TaxTrans.PPTRADisc
      ElseIf TaxTrans.Revenue.Principle2 > 0 Then
        TaxTrans.Revenue.Principle2 = ABUTot + TaxTrans.Amount + TaxTrans.PPTRADisc
      ElseIf TaxTrans.Revenue.Principle3 > 0 Then
        TaxTrans.Revenue.Principle3 = ABUTot + TaxTrans.Amount + TaxTrans.PPTRADisc
      ElseIf TaxTrans.Revenue.Principle4 > 0 Then
        TaxTrans.Revenue.Principle4 = ABUTot + TaxTrans.Amount + TaxTrans.PPTRADisc
       ElseIf TaxTrans.Revenue.Principle5 > 0 Then
        TaxTrans.Revenue.Principle5 = ABUTot + TaxTrans.Amount + TaxTrans.PPTRADisc
      ElseIf TaxTrans.Revenue.RevOpt1 > 0 Then
        TaxTrans.Revenue.RevOpt1 = ABUTot + TaxTrans.Amount + TaxTrans.PPTRADisc
      ElseIf TaxTrans.Revenue.RevOpt2 > 0 Then
        TaxTrans.Revenue.RevOpt2 = ABUTot + TaxTrans.Amount + TaxTrans.PPTRADisc
      ElseIf TaxTrans.Revenue.RevOpt3 > 0 Then
        TaxTrans.Revenue.RevOpt3 = ABUTot + TaxTrans.Amount + TaxTrans.PPTRADisc
      End If
      Put TTHandle, BelongTo, TaxTrans
      AdjustedABU = 1
    Else
      AdjustedABU = 2
    End If
  End If
  
  Return
  
AdjustBillDown:
  BelongTo = TaxTrans.BelongTo
  AdjustedAll = False
  ABDTot = TaxTrans.Revenue.Principle1
  ABDTot = ABDTot + TaxTrans.Revenue.Principle2
  ABDTot = ABDTot + TaxTrans.Revenue.Principle3
  ABDTot = ABDTot + TaxTrans.Revenue.Principle4
  ABDTot = ABDTot + TaxTrans.Revenue.Principle5
  ABDTot = ABDTot + TaxTrans.Revenue.RevOpt1
  ABDTot = ABDTot + TaxTrans.Revenue.RevOpt2
  ABDTot = ABDTot + TaxTrans.Revenue.RevOpt3
  ABDTot = ABDTot + TaxTrans.Revenue.Interest
  ABDTot = ABDTot + TaxTrans.Revenue.Collection
  ABDTot = ABDTot + TaxTrans.Revenue.LateList
  ABDTot = ABDTot + TaxTrans.Revenue.Penalty
  Get TTHandle, BelongTo, TaxTrans
  BillTot = TaxTrans.Amount
  BillTot = BillTot + TaxTrans.Revenue.Interest
  BillTot = BillTot + TaxTrans.Revenue.Collection
  BillTot = BillTot + TaxTrans.Revenue.LateList
  BillTot = BillTot + TaxTrans.Revenue.Penalty
  If ABDTot = BillTot Then
    TaxTrans.Revenue.Principle1 = 0
    TaxTrans.Revenue.Principle1Pd = 0
    TaxTrans.Revenue.Principle2 = 0
    TaxTrans.Revenue.Principle2Pd = 0
    TaxTrans.Revenue.Principle3 = 0
    TaxTrans.Revenue.Principle3Pd = 0
    TaxTrans.Revenue.Principle4 = 0
    TaxTrans.Revenue.Principle4Pd = 0
    TaxTrans.Revenue.Principle5 = 0
    TaxTrans.Revenue.Principle5Pd = 0
    TaxTrans.Revenue.RevOpt1 = 0
    TaxTrans.Revenue.RevOpt1Pd = 0
    TaxTrans.Revenue.RevOpt2 = 0
    TaxTrans.Revenue.RevOpt2Pd = 0
    TaxTrans.Revenue.RevOpt3 = 0
    TaxTrans.Revenue.RevOpt3Pd = 0
    TaxTrans.Revenue.Interest = 0
    TaxTrans.Revenue.InterestPd = 0
    TaxTrans.Revenue.Collection = 0
    TaxTrans.Revenue.CollectionPd = 0
    TaxTrans.Revenue.LateList = 0
    TaxTrans.Revenue.LateListPd = 0
    TaxTrans.Revenue.Penalty = 0
    TaxTrans.Revenue.Penalty = 0
    TaxTrans.PPTRADisc = 0
    Put TTHandle, BelongTo, TaxTrans
    AdjustedAll = True
  End If

  Return
DifferentPlus:
  P1 = P1 + TaxTrans.Revenue.Principle1
  P2 = P2 + TaxTrans.Revenue.Principle2
  P3 = P3 + TaxTrans.Revenue.Principle3
  P4 = P4 + TaxTrans.Revenue.Principle4
  P5 = P5 + TaxTrans.Revenue.Principle5
  Intr = Intr + TaxTrans.Revenue.Interest
  Adv = Adv + TaxTrans.Revenue.Collection
  LL = LL + TaxTrans.Revenue.LateList
  Pen = Pen + TaxTrans.Revenue.Penalty
  Opt1 = Opt1 + TaxTrans.Revenue.RevOpt1
  Opt2 = Opt2 + TaxTrans.Revenue.RevOpt2
  Opt3 = Opt3 + TaxTrans.Revenue.RevOpt3
  TaxTrans.Revenue.Principle1 = TaxTrans.Revenue.Principle1
  TaxTrans.TranType = TaxTrans.TranType

  Return
  
DifferentAdjPayDown:
  P1Pd = P1Pd - TaxTrans.Revenue.Principle1Pd
  P2Pd = P2Pd - TaxTrans.Revenue.Principle2Pd
  P3Pd = P3Pd - TaxTrans.Revenue.Principle3Pd
  P4Pd = P4Pd - TaxTrans.Revenue.Principle4Pd
  P5Pd = P5Pd - TaxTrans.Revenue.Principle5Pd
  IntPd = IntPd - TaxTrans.Revenue.InterestPd
  AdvPd = AdvPd - TaxTrans.Revenue.CollectionPd
  LLPd = LLPd - TaxTrans.Revenue.LateListPd
  PenPd = PenPd - TaxTrans.Revenue.PenaltyPd
  Opt1Pd = Opt1Pd - TaxTrans.Revenue.RevOpt1Pd
  Opt2Pd = Opt2Pd - TaxTrans.Revenue.RevOpt2Pd
  Opt3Pd = Opt3Pd - TaxTrans.Revenue.RevOpt3Pd
  Return
  
End Sub
Private Sub CompareMBWithCustHistory(Optional ByVal PrintIt As Boolean = True)
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim NextRec As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim SaveRec As Long
  Dim Amount As Double
  Dim MBBal As Double
  Dim CHBal As Double
  Dim AHandle As Integer
  Dim ThisRec As Long
  Dim cnt As Long
  Dim BHandle As Integer
  Dim Reject As Boolean
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  AHandle = FreeFile
  Open "MBvsCustHist.txt" For Output As AHandle
  BHandle = FreeFile
  Open "txbalerrors.txt" For Output As BHandle
  frmVATaxShowPctComp.Label1 = "Finding MB vs Cust History Errors"
  frmVATaxShowPctComp.Show , Me
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    Reject = False
    If TaxCust.Deleted <> 0 Then GoTo Skip
    CHBal = GetCustBalance(x, -1)
    GoSub RunMB
    If Reject = True Then GoTo Skip
    If OldRound(CHBal) <> OldRound(MBBal) Then
'      If x = 852 Then Stop
      Print #AHandle, CStr(x) + "~" + CStr(CHBal) + "~" + CStr(MBBal)
      Print #BHandle, CStr(x) + ","
      cnt = cnt + 1
    End If
Skip:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  
  Close TCHandle
  Close TTHandle
  Close AHandle
  Close BHandle
  Call BuildMBvsCustHistArr
  If PrintIt = False Then Exit Sub
  MsgBox ("A total of " + CStr(cnt) + " have been processed. Look for MBvsCustHist.txt in the Citipak folder for results.")
  Exit Sub
  
RunMB:
'  If x = 3677 Then Stop
  ThisRec = TaxCust.LastTrans
  If ThisRec = 0 Then
    Reject = True
    Return
  End If
  MBBal = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If TaxTrans.TranType = 10 Then 'adjust bill down affecting credit
       MBBal = MBBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'prepay adjust down
      MBBal# = OldRound#(MBBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'adjust bill up affecting credit
      MBBal# = OldRound#(MBBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
      MBBal# = OldRound#(MBBal# - TaxTrans.Revenue.PrePaidAmt)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay
      MBBal# = OldRound#(MBBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 9 Then 'added
      MBBal# = OldRound#(MBBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 1 Then
      MBBal# = OldRound#(MBBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      MBBal# = OldRound#(MBBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      MBBal# = OldRound#(MBBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      MBBal# = OldRound#(MBBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      MBBal# = OldRound#(MBBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
    End If
DoAgain:
    ThisRec = TaxTrans.LastTrans
  Loop
  Return

End Sub
Private Sub ExamineAReal()
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim Trans As Long
 OpenRealPropFile RHandle, NumOfRRecs
 Trans = 3
 Get RHandle, Trans, RealRec
 RealRec.Blank = ""
 RealRec.BldgVal = 0
 RealRec.BLOCK = ""
 RealRec.CustPin = 0
 RealRec.Deleted = "N"
 RealRec.EXMPOTHR = 0
 RealRec.EXMPSENI = 0
 RealRec.Fill1 = ""
 RealRec.GISPOS = ""
 RealRec.ICPDesc = ""
 RealRec.Image = 0
 RealRec.InternalPin = 0
 RealRec.LastYrPrinted = 0
 RealRec.LateList = "N"
 RealRec.LienDesc = ""
 RealRec.LOTACRE = ""
 RealRec.LOTNUMB = ""
 RealRec.Map = ""
 RealRec.Mock = "N"
 RealRec.MORTCODE = ""
 RealRec.NextRec = 0
 RealRec.OptRev1Chrg = 0
 RealRec.OptRev2Chrg = 0
 RealRec.OptRev3Chrg = 0
 RealRec.OptSearch = ""
 RealRec.PropAddr = ""
 RealRec.PROPDATE = 0
 RealRec.PROPDISC = "N"
 RealRec.PROPNOT1 = ""
 RealRec.PROPNOT2 = ""
 RealRec.PROPNOT3 = ""
 RealRec.PropSize = 0
 RealRec.PROPVALU = 0
 RealRec.RealPin = ""
 RealRec.TownShip = ""
 Close
 
End Sub


Private Sub ExamineATrans()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim Trans As Long
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  Trans = 4199

  Get TTHandle, Trans, TaxTrans
  TaxTrans.Revenue.Interest = TaxTrans.Revenue.Interest
  TaxTrans.CustPin = TaxTrans.CustPin
  TaxTrans.BelongTo = TaxTrans.BelongTo
  TaxTrans.Amount = TaxTrans.Amount
  TaxTrans.TransDate = TaxTrans.TransDate
  TaxTrans.TranType = TaxTrans.TranType
  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd
  TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd
  TaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.CollectionPd
  TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd
  TaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateListPd
  TaxTrans.Revenue.RevOpt1Pd = TaxTrans.Revenue.RevOpt1Pd
  TaxTrans.Revenue.RevOpt2Pd = TaxTrans.Revenue.RevOpt2Pd
  TaxTrans.Revenue.RevOpt3Pd = TaxTrans.Revenue.RevOpt3Pd
  TaxTrans.PPTRADisc = TaxTrans.PPTRADisc
  TaxTrans.DiscXDate = TaxTrans.DiscXDate
  TaxTrans.RealPin = TaxTrans.RealPin
  TaxTrans.PersPin = TaxTrans.PersPin

  TaxTrans.Posted2GL = TaxTrans.Posted2GL
  TaxTrans.TaxYear = TaxTrans.TaxYear
  TaxTrans.DiscAmt = TaxTrans.DiscAmt
  TaxTrans.OperNum = TaxTrans.OperNum
  TaxTrans.FromPrePay = TaxTrans.FromPrePay
  TaxTrans.Description = TaxTrans.Description
  TaxTrans.CustomerRec = TaxTrans.CustomerRec
  TaxTrans.Revenue.PrePaidAmt = TaxTrans.Revenue.PrePaidAmt
  TaxTrans.Revenue.PrePaidUsed = TaxTrans.Revenue.PrePaidUsed
  TaxTrans.Revenue.PrePaidBal = TaxTrans.Revenue.PrePaidBal
  TaxTrans.InternalPin = TaxTrans.InternalPin
  TaxTrans.CntyPara = TaxTrans.CntyPara
  TaxTrans.CyclPara = TaxTrans.CyclPara
  TaxTrans.TShpPara = TaxTrans.TShpPara
  TaxTrans.BillType = TaxTrans.BillType
  TaxTrans.Revenue.Principle1 = TaxTrans.Revenue.Principle1
  TaxTrans.Revenue.Interest = TaxTrans.Revenue.Interest
  TaxTrans.Revenue.Collection = TaxTrans.Revenue.Collection
  TaxTrans.Revenue.Penalty = TaxTrans.Revenue.Penalty
  TaxTrans.Revenue.LateList = TaxTrans.Revenue.LateList
  TaxTrans.Revenue.RevOpt1 = TaxTrans.Revenue.RevOpt1
  TaxTrans.Revenue.RevOpt2 = TaxTrans.Revenue.RevOpt2
  TaxTrans.Revenue.RevOpt3 = TaxTrans.Revenue.RevOpt3
  TaxTrans.LastTrans = TaxTrans.LastTrans
'  Put TTHandle, 117, TaxTrans
  Close TTHandle
End Sub

Private Sub Check4CreditAtBillingWithNoPrepayment()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim NextRec As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim SaveRec As Long
  Dim Amount As Double
  Dim MBBal As Double
  Dim CHBal As Double
  Dim AHandle As Integer
  Dim ThisRec As Long
  Dim cnt As Long
  Dim BHandle As Integer
  Dim Is9 As Boolean
  Dim Trans9 As Long
'22 overpayment only
'21 payment plus overpayment
'9 credit applied at billing
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  AHandle = FreeFile
  Open "CredAtBillNoPrePay.txt" For Output As AHandle
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
'    If x = 125 Then Stop
    TaxCust.CustName = TaxCust.CustName
    If TaxCust.Deleted = True Then GoTo Skip
    Is9 = False
    Trans9 = 0
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
'      If NextRec = 4673 Then Stop
      If TaxTrans.TranType = 9 Then
        Is9 = True
        Trans9 = NextRec
      End If
      If TaxTrans.TranType = 22 Or TaxTrans.TranType = 21 And Is9 = True Then
        Is9 = False
        Trans9 = 0
      End If
      NextRec = TaxTrans.LastTrans
    Loop
    If Is9 = True Then
      Print #AHandle, CStr(x) + "~" + CStr(Trans9)
      cnt = cnt + 1
    End If
Skip:
  Next x
  
  Close
  MsgBox ("A total of " + CStr(cnt) + " bad credit at billings have been found. Look for CreditAtBillNoPrepay in the Citipak folder.")
End Sub

Private Sub FixUnusedPrepayIn2006()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim NextRec As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Long
  Dim SaveRec As Long
  Dim Amount As Double
  Dim AHandle As Integer
  Dim ThisRec As Long
  Dim cnt As Long
  Dim BHandle As Integer
  ReDim TransArr(1 To 1) As Long
  Dim TCnt As Integer
  Dim UseIt As Boolean
  Dim BDate As Integer
  Dim EDate As Integer
  Dim OPAmt As Double
  Call ApplyOPToPaidBillsWithNoPayTrans
  Exit Sub
  BDate = Date2Num("06/01/2006")
  EDate = Date2Num("08/30/2006")
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  AHandle = FreeFile
  Open "UnusedPrepay.txt" For Output As AHandle
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    NextRec = TaxCust.LastTrans
    TCnt = 0
    UseIt = False
    ReDim TransArr(1 To 1) As Long
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
      TCnt = TCnt + 1
      ReDim Preserve TransArr(1 To TCnt) As Long
      TransArr(TCnt) = NextRec
      If TaxTrans.TranType = 21 And TaxTrans.TransDate >= BDate And TaxTrans.TransDate <= EDate Then UseIt = True
      NextRec = TaxTrans.LastTrans
    Loop
    If UseIt = True Then
      For y = TCnt To 1 Step -1
        Get TTHandle, CLng(TransArr(y)), TaxTrans
         If TaxTrans.TranType = 21 And TaxTrans.TransDate >= BDate And TaxTrans.TransDate <= EDate Then
           OPAmt = TaxTrans.Revenue.PrePaidAmt
         End If
         If TaxTrans.TranType = 1 And TaxTrans.Amount = OPAmt Or TaxTrans.Amount = OPAmt - 0.01 Then
           Call ClearTrans(CLng(TransArr(y)))
           cnt = cnt + 1
           Print #AHandle, CStr(x) + "~" + CStr(TransArr(y))
           Exit For
         End If
      Next y
    End If
  Next x
  
  Close
  MsgBox ("A total of " + CStr(cnt) + " unused prepays have been fixed. Look for UnusedPrepay.txt in the Citipak folder for results.")
End Sub

Private Sub ApplyOPToPaidBillsWithNoPayTrans()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim NextRec As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Amount As Double
  Dim TopTrans As Long
  Dim BottomTrans As Long
  Dim SaveRec As Long
  Dim x As Long, y As Long
  Dim PayFound As Boolean
  Dim AHandle As Integer
  Dim cnt As Long
  Dim BillTrans As Long
  Dim LastCustTrans As Long
  Dim CustPin As Integer
  Dim TaxYear As Integer
  Dim PrePay As Double
  Dim BelongTo As Long
  Dim PropType As String
  Dim BillNum As String
  Dim ThisDate As String
  Dim LookRec As Long
  ReDim CArr(1 To 1) As Long
  CArrCnt = 0
  Call CompareMBWithCustHistory
  AHandle = FreeFile
  Open "PaidBillsNoPayTrans.txt" For Output As AHandle
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  frmVATaxShowPctComp.Label1 = "Processing"
  frmVATaxShowPctComp.Show , Me
  For x = 1 To CArrCnt
    Get TCHandle, CArr(x), TaxCust
'    If CArr(x) = 3 Then Stop
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
        If TaxTrans.TranType = 21 Then GoTo CarryOn
      NextRec = TaxTrans.LastTrans
    Loop
    GoTo Skip
CarryOn:
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
      LookRec = TaxCust.LastTrans
      Do While LookRec > 0
        Get TTHandle, LookRec, TaxTrans
        If TaxTrans.LastTrans = NextRec Then
          TopTrans = LookRec
          Exit Do
        End If
        LookRec = TaxTrans.LastTrans
      Loop
      Get TTHandle, NextRec, TaxTrans
      BottomTrans = NextRec
      SaveRec = TaxCust.LastTrans
      If TaxTrans.TranType = 1 And TaxTrans.Amount > 0 Then
        Amount = TaxTrans.Amount
        PayFound = False
        Do While SaveRec > 0
          Get TTHandle, SaveRec, TaxTrans
          If (TaxTrans.TranType = 21 Or TaxTrans.TranType = 9) And TaxTrans.BelongTo = NextRec Then
            Exit Do 'Prepay is used properly for this bill
          End If
          If TaxTrans.TranType = 2 And TaxTrans.BelongTo = NextRec Then
            Exit Do 'a payment transaction is entered for this bill
          End If
          If TaxTrans.TranType = 21 Then
            If TaxTrans.Revenue.PrePaidAmt = Amount Or TaxTrans.Revenue.PrePaidAmt = (Amount - 0.01) Or TaxTrans.Revenue.PrePaidAmt = (Amount + 0.01) Then
              If TaxTrans.Revenue.PrePaidAmt = (Amount - 0.01) Then
                TaxTrans.Amount = TaxTrans.Amount + 0.01
                TaxTrans.Revenue.PrePaidAmt = TaxTrans.Revenue.PrePaidAmt + 0.01
                Put TTHandle, SaveRec, TaxTrans
              ElseIf TaxTrans.Revenue.PrePaidAmt = (Amount + 0.01) Then
                TaxTrans.Amount = TaxTrans.Amount - 0.01
                TaxTrans.Revenue.PrePaidAmt = TaxTrans.Revenue.PrePaidAmt - 0.01
                Put TTHandle, SaveRec, TaxTrans
              End If
              PrePay = TaxTrans.Revenue.PrePaidAmt
              PayFound = True
              Exit Do
            End If
          End If
          SaveRec = TaxTrans.LastTrans
        Loop
        If PayFound = True Then
          GoSub FixIt
          Print #AHandle, CStr(CArr(x)) + "~" + CStr(NextRec)
          cnt = cnt + 1
        End If
      End If
      Get TTHandle, NextRec, TaxTrans
      NextRec = TaxTrans.LastTrans
    Loop
Skip:
    frmVATaxShowPctComp.ShowPctComp x, CArrCnt
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  Close
  MsgBox ("A total of " + CStr(cnt) + " pay errors were found and fixed. Look for PaidBillsNoPayTrans.txt in the Citipak folder for results.")
  Exit Sub

FixIt:
  BelongTo = NextRec
  Get TTHandle, NextRec, TaxTrans
  BillTrans = NextRec
  LastCustTrans = TaxCust.LastTrans
  TaxYear = TaxTrans.TaxYear
  BillNum = TaxTrans.Description
  PropType = TaxTrans.BillType
  ThisDate = Now
  CustPin = CInt(CArr(x))
  Call FixPayPlusOPThatShouldHaveBeenApplied(BillTrans, LastCustTrans, ThisDate, Amount, CustPin, TaxYear, PrePay, BelongTo, PropType, TopTrans, BottomTrans, BillNum)

Return

ClearPennyTransPlus:
  
Return

ClearPennyTransMinus:

Return

End Sub
Private Sub FindAndFixTransWithCrossCusts()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim NextRec As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Amount As Double
  Dim TopTrans As Long
  Dim BottomTrans As Long
  Dim SaveRec As Long
  Dim x As Long, y As Long
  Dim PayFound As Boolean
  Dim AHandle As Integer
  Dim cnt As Long
  Dim BillDate As Long
  Dim LastCustDate As Long
  Dim CustPin As Integer
  Dim TaxYear As Integer
  Dim PrePay As Double
  Dim BelongTo As Long
  Dim PropType As String
  Dim BillNum As String
  Dim ThisDate As String
  Dim LookRec As Long
  Dim BadTop As Long
  Dim BadBottom As Long
  Dim GoodTop As Long
  Dim GoodBottom As Long
  Dim BadCust As Integer
  Dim Fixed As Boolean
  Dim LastCustRec As Long
  Dim TTRecCnt As Long
  Dim SaveLast As Long
  Dim NextInQ As Long
  Dim BHandle As Integer
  
  BHandle = FreeFile
  Open "Oddballerrors.txt" For Output As BHandle
  If Not Exist(App.Path + "\TransWithCrossCusts.txt") Then
    MsgBox ("Please run 'Relink BelongTos To Bills' first.")
    Exit Sub
  End If
  AHandle = FreeFile
  Open "TransWithCrossCustsFixed.txt" For Output As AHandle
  Call BuildCrossTransArr
  On Error Resume Next
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  frmVATaxShowPctComp.Label1 = "Fixing Transactions Assigned to Wrong Customers"
  frmVATaxShowPctComp.Show , Me
  
  For x = 1 To CrossCnt
    If CrossArr(x) > NumOfTCRecs Then GoTo Skip
    Get TCHandle, CrossArr(x), TaxCust
    GoSub FixIt
Skip:
    frmVATaxShowPctComp.ShowPctComp x, CrossCnt 'NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  Close
  MsgBox ("A total of " + CStr(cnt) + " transactions have been found and fixed. Look for TransWithCrossCustsFixed.txt in the Citipak folder for results.")
  Exit Sub
  
FixIt:
  Get TTHandle, CrossBadArr(x), TaxTrans
  BadCust = TaxTrans.CustomerRec
  If BadCust = 0 Then Return
  Get TCHandle, BadCust, TaxCust
  SaveRec = TaxCust.LastTrans
  If CrossBadArr(x) = SaveRec Then 'if the bad tran is the next one in the queue for the bad cust
  'then save that trans next trans as the customer's last trans
    Get TTHandle, SaveRec, TaxTrans
    TaxCust.LastTrans = TaxTrans.LastTrans
    Put TTHandle, SaveRec, TaxTrans 'now the bad trans is no longer in the bad cust's queue
  Else
    Get TTHandle, CrossBadArr(x), TaxTrans
    SaveLast = TaxTrans.LastTrans
    Do While SaveRec > 0 'bad trans is inside the queue so find it
      Get TTHandle, SaveRec, TaxTrans
      If TaxTrans.LastTrans = CrossBadArr(x) Then 'find trans right above the bad one
        TaxTrans.LastTrans = SaveLast 'and make its last trans the one below the bad one
        'effectively removing the bad one from the queue
        Put TTHandle, SaveRec, TaxTrans
        Exit Do
      End If
      SaveRec = TaxTrans.LastTrans
    Loop
  End If
  
    'now go back to the good cust
  Get TCHandle, CrossArr(x), TaxCust
  LastCustRec = TaxCust.LastTrans
  If TaxCust.LastTrans = CrossGoodArr(x) Then 'the bill was the last trans for this cust
    Get TTHandle, CrossBadArr(x), TaxTrans 'get the bad trans and assign the next trans
    'as the new custs old last trans
    TaxTrans.LastTrans = TaxCust.LastTrans
    Put TTHandle, CrossBadArr(x), TaxTrans 'save to bad and your done with that trans
    
    Get TTHandle, LastCustRec, TaxTrans 'need to drop the former good cust last trans down a notch
    SaveLast = TaxTrans.LastTrans
    Get TTHandle, CrossGoodArr(x), TaxTrans 'assign new last trans to the old last trans
    TaxTrans.LastTrans = SaveLast
    Put TTHandle, CrossGoodArr(x), TaxTrans 'save it to this trans and your done with it
    
    TTRecCnt = LOF(TTHandle) / Len(TaxTrans) 'get the latest number of tax trans (not
    'NumOfTTRecs)
    TTRecCnt = TTRecCnt + 1 'assign the last trans to this cust
    TaxCust.LastTrans = TTRecCnt 'now the newly inserted trans is at the top of the queue
    Put TCHandle, CrossArr(x), TaxCust 'save cust
    cnt = cnt + 1
    Print #AHandle, "Trans " + CStr(CrossBadArr(x)) + " was switched from customer # " + CStr(BadCust) + " to customer # " + CStr(CrossArr(x)) + "."
    Return
  Else
    NextInQ = LastCustRec
    Do While NextInQ > 0
      Get TTHandle, NextInQ, TaxTrans
        If TaxTrans.LastTrans = CrossGoodArr(x) Then
          TaxTrans.LastTrans = CrossBadArr(x)
          Put TTHandle, NextInQ, TaxTrans
          Get TTHandle, CrossBadArr(x), TaxTrans
          TaxTrans.LastTrans = CrossGoodArr(x)
          Put TTHandle, CrossBadArr(x), TaxTrans
          cnt = cnt + 1
          Print #AHandle, "Trans " + CStr(CrossBadArr(x)) + " was switched from customer # " + CStr(BadCust) + " to customer # " + CStr(CrossArr(x)) + "."
          Return
        End If
      NextInQ = TaxTrans.LastTrans
    Loop
  End If
  
  Return
End Sub

Private Sub FindPennyPlusInErrorAndFix()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim NextRec As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Amount As Double
  Dim TopTrans As Long
  Dim BottomTrans As Long
  Dim SaveRec As Long
  Dim x As Long, y As Long
  Dim PayFound As Boolean
  Dim AHandle As Integer
  Dim cnt As Long
  Dim BillDate As Long
  Dim LastCustDate As Long
  Dim CustPin As Integer
  Dim TaxYear As Integer
  Dim BillBal As Double
  Dim Billed As Double
  Dim Paid As Double
  
  AHandle = FreeFile
  Open "PennyErrorFoundandFixed.txt" For Output As AHandle
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  frmVATaxShowPctComp.Label1 = "Processing"
  frmVATaxShowPctComp.Show , Me
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
      
      If TaxTrans.TranType = 1 Then
        Billed = TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2
        Billed = Billed + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4
        Billed = Billed + TaxTrans.Revenue.Principle5 + TaxTrans.Revenue.Interest
        Billed = Billed + TaxTrans.Revenue.Collection + TaxTrans.Revenue.LateList
        Billed = Billed + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.RevOpt1
        Billed = Billed + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3
        Paid = TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd
        Paid = Paid + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd
        Paid = Paid + TaxTrans.Revenue.Principle5Pd + TaxTrans.Revenue.InterestPd
        Paid = Paid + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd
        Paid = Paid + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.RevOpt1Pd
        Paid = Paid + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.PPTRADisc
        BillBal = OldRound(Billed - Paid)
        If BillBal = -0.01 Then
          SaveRec = TaxCust.LastTrans
          Do While SaveRec > 0
            Get TTHandle, SaveRec, TaxTrans
            If TaxTrans.TranType = 2 And TaxTrans.BelongTo = NextRec And TaxTrans.Amount = 0.01 Then
              Call ClearTrans(SaveRec)
              Get TTHandle, NextRec, TaxTrans
              If OldRound(TaxTrans.Revenue.Principle1 - (TaxTrans.Revenue.Principle1Pd + TaxTrans.PPTRADisc)) = -0.01 Then
                TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd - 0.01
              ElseIf OldRound(TaxTrans.Revenue.Principle2 - TaxTrans.Revenue.Principle2Pd + TaxTrans.PPTRADisc) = -0.01 Then
                TaxTrans.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2Pd - 0.01
              ElseIf OldRound(TaxTrans.Revenue.Principle3 - TaxTrans.Revenue.Principle3Pd + TaxTrans.PPTRADisc) = -0.01 Then
                TaxTrans.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3Pd - 0.01
              ElseIf OldRound(TaxTrans.Revenue.Principle4 - TaxTrans.Revenue.Principle4Pd + TaxTrans.PPTRADisc) = -0.01 Then
                TaxTrans.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4Pd - 0.01
              ElseIf OldRound(TaxTrans.Revenue.Principle5 - TaxTrans.Revenue.Principle5Pd + TaxTrans.PPTRADisc) = -0.01 Then
                TaxTrans.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5Pd - 0.01
              ElseIf OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd) = -0.01 Then
                TaxTrans.Revenue.RevOpt1Pd = TaxTrans.Revenue.RevOpt1Pd - 0.01
              ElseIf OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd) = -0.01 Then
                TaxTrans.Revenue.RevOpt2Pd = TaxTrans.Revenue.RevOpt2Pd - 0.01
              ElseIf OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd) = -0.01 Then
                TaxTrans.Revenue.RevOpt3Pd = TaxTrans.Revenue.RevOpt3Pd - 0.01
              ElseIf OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd) = -0.01 Then
                TaxTrans.Revenue.InterestPd = TaxTrans.Revenue.InterestPd - 0.01
              ElseIf OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd) = -0.01 Then
                TaxTrans.Revenue.CollectionPd = TaxTrans.Revenue.CollectionPd - 0.01
              ElseIf OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd) = -0.01 Then
                TaxTrans.Revenue.LateListPd = TaxTrans.Revenue.LateListPd - 0.01
              ElseIf OldRound(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd) = -0.01 Then
                TaxTrans.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd - 0.01
              End If
              Put TTHandle, NextRec, TaxTrans
              Print #AHandle, CStr(x) + "~" + CStr(NextRec) + "~" + CStr(SaveRec)
              cnt = cnt + 1
              Exit Do
            End If
            SaveRec = TaxTrans.LastTrans
          Loop
        End If
      End If
      NextRec = TaxTrans.LastTrans
      frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
      If frmVATaxShowPctComp.Out = True Then
        Close
        frmVATaxShowPctComp.Out = False
        Unload frmVATaxShowPctComp
        Exit Sub
      End If
    Loop
  Next x
  Unload frmVATaxShowPctComp
  Close TCHandle
  Close TTHandle
  Close AHandle
  
  MsgBox ("A total of " + CStr(cnt) + " transactions were updated. Look for PennyErrorFoundandFixed.txt in the Citipak folder for results.")
End Sub

Private Sub FixInitializedTrans()
  Dim TaxTrans As TaxTransactionType
  Dim NewTaxTrans As TaxTransactionType
  Dim PayTranRec As TaxTransactionType
  Dim ClearTaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Integer
  Dim IntAmt As Double
  Dim PenAmt As Double
  Dim AdvAmt As Double
  Dim TopRec As Long
  Dim BottomRec As Long
  Dim cnt As Long
  Dim AHandle As Integer
  Dim NextTTTrans As Long
  Dim NextRec As Long
  Dim NextPayRec As Long
  Dim ThisCust As Long
  Dim RealPin As String
  Dim PersPin As String
  Dim SaveRec As Long
  Dim Savex As Long
  ReDim DoneArr(1 To 1) As Long
  Dim DoneCnt As Long
  Dim Fixed As Boolean
  Dim Paid As Double
  Dim DoPaid As Boolean
  AHandle = FreeFile
  Open "InitializedFixes.txt" For Output As AHandle
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  cnt = 0
  frmVATaxShowPctComp.Label1 = "Fixing Initialization Transactions"
  frmVATaxShowPctComp.Show , Me
'  GoTo Start
  For x = 1 To NumOfTTRecs
    Fixed = False
    Get TTHandle, x, TaxTrans
    Savex = x
    If InStr(UCase(TaxTrans.Description), "INITIALIZE") > 0 Then
      NextTTTrans = LOF(TTHandle) \ Len(TaxTrans) + 1
      IntAmt = TaxTrans.Revenue.Interest
      PenAmt = TaxTrans.Revenue.Penalty
      AdvAmt = TaxTrans.Revenue.Collection
      RealPin = QPTrim$(TaxTrans.RealPin)
      PersPin = QPTrim$(TaxTrans.PersPin)
      ThisCust = TaxTrans.CustomerRec
      If ThisCust = 0 Then
        ThisCust = TaxTrans.CustPin
      End If
      If ThisCust = 0 Then
        Print #AHandle, "Could not fix trans # " + CStr(x) + " because owning customer could not be determined."
        GoTo SkipIt
      End If
      Get TCHandle, ThisCust, TaxCust
      NextRec = TaxCust.LastTrans
      Paid = 0
      Paid = TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd
      Paid = Paid + TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.LateListPd
'      DoPaid = False
'      If Paid > 0 Then
'        DoPaid = True
'      End If
      
      If IntAmt > 0 Then
        NextTTTrans = LOF(TTHandle) \ Len(TaxTrans) + 1
        cnt = cnt + 1
        NewTaxTrans = ClearTaxTrans
        If Savex = NextRec Then 'initialized trans is at the top of the queue
          TaxCust.LastTrans = NextTTTrans
          Put TCHandle, ThisCust, TaxCust
'          NewTaxTrans.LastTrans = NextRec
          NewTaxTrans.LastTrans = Savex
          NextRec = NextTTTrans
        Else
          SaveRec = NextRec
          TopRec = SaveRec
          Do While SaveRec > 0
            Get TTHandle, SaveRec, TaxTrans
            For y = 1 To DoneCnt
              If DoneArr(y) = SaveRec Then GoTo SkipInt
            Next y
            If TaxTrans.LastTrans = Savex Then
'              Get TTHandle, TopRec, TaxTrans
              TaxTrans.LastTrans = NextTTTrans
'              Put TTHandle, TopRec, TaxTrans
              Put TTHandle, SaveRec, TaxTrans
'              Get TTHandle, NextTTTrans, TaxTrans
              NewTaxTrans.LastTrans = Savex
'              Put TTHandle, NextTTTrans, TaxTrans
              Exit Do
            End If
SkipInt:
'            TopRec = SaveRec
            SaveRec = TaxTrans.LastTrans
          Loop
        End If
        Get TTHandle, Savex, TaxTrans
        NewTaxTrans.TransDate = TaxTrans.TransDate
        NewTaxTrans.TaxYear = TaxTrans.TaxYear
        NewTaxTrans.TranType = 4       '4=Interest
        NewTaxTrans.BillType = TaxTrans.BillType
        NewTaxTrans.Amount = IntAmt  'Total Transaction Amount
        NewTaxTrans.Revenue.Interest = IntAmt
        NewTaxTrans.Description = "Tax Int on Bill# " + ParseBillNum$(TaxTrans.Description)
        NewTaxTrans.Posted2GL = "N"
        NewTaxTrans.CustomerRec = ThisCust
        NewTaxTrans.CustPin = ThisCust
        NewTaxTrans.RealPin = RealPin
        NewTaxTrans.PersPin = PersPin
        NewTaxTrans.BelongTo = Savex
        NewTaxTrans.Revenue.PrePaidAmt = 0
        If TaxTrans.BillType = "R" Or TaxTrans.BillType = "P" Then
          NewTaxTrans.Revenue.PrePaidBal = OldRound(GetOverPayBalance(ThisCust, TaxTrans.BillType))
        Else
          NewTaxTrans.Revenue.PrePaidBal = 0
        End If
        NewTaxTrans.Revenue.PrePaidUsed = 0
        NewTaxTrans.OperNum = 0
        LSet NewTaxTrans.Padding = ""
        Put TTHandle, NextTTTrans, NewTaxTrans
'        Savex = NextTTTrans 'must move the target up one slot
        Fixed = True
        Print #AHandle, "Inserted interest trans # " + CStr(NextTTTrans) + " into Cust # " + CStr(ThisCust) + "."
      End If
      If PenAmt > 0 Then
        NextTTTrans = LOF(TTHandle) \ Len(TaxTrans) + 1
        Get TTHandle, Savex, TaxTrans
        NextRec = TaxCust.LastTrans
        cnt = cnt + 1
        NewTaxTrans = ClearTaxTrans
        If x = NextRec Then 'initialized trans is at the top of the queue
          TaxCust.LastTrans = NextTTTrans
          Put TCHandle, ThisCust, TaxCust
          NewTaxTrans.LastTrans = NextRec
          NextRec = NextTTTrans
        Else
          SaveRec = TaxCust.LastTrans 'NextRec
          TopRec = SaveRec
          Do While SaveRec > 0
            Get TTHandle, SaveRec, TaxTrans
            For y = 1 To DoneCnt
              If DoneArr(y) = SaveRec Then GoTo SkipPen
            Next y
            If TaxTrans.LastTrans = Savex Then
'              Get TTHandle, TopRec, TaxTrans
              TaxTrans.LastTrans = NextTTTrans
'              Put TTHandle, TopRec, TaxTrans
              Put TTHandle, SaveRec, TaxTrans
              NewTaxTrans.LastTrans = Savex
              Exit Do
            End If
SkipPen:
'            TopRec = SaveRec
            SaveRec = TaxTrans.LastTrans
          Loop
        End If
'        Get TTHandle, 4908, TaxTrans
'        TaxTrans.LastTrans = TaxTrans.LastTrans
        Get TTHandle, Savex, TaxTrans
        NewTaxTrans.TransDate = TaxTrans.TransDate
        NewTaxTrans.TaxYear = TaxTrans.TaxYear
        NewTaxTrans.TranType = 5
        NewTaxTrans.BillType = TaxTrans.BillType
        NewTaxTrans.Amount = PenAmt
        NewTaxTrans.Revenue.Penalty = PenAmt
        NewTaxTrans.Description = "Tax Pen on Bill# " + ParseBillNum$(TaxTrans.Description)
        NewTaxTrans.Posted2GL = "N"
        NewTaxTrans.CustomerRec = ThisCust
        NewTaxTrans.CustPin = ThisCust
        NewTaxTrans.RealPin = RealPin
        NewTaxTrans.PersPin = PersPin
        NewTaxTrans.BelongTo = Savex
        NewTaxTrans.Revenue.PrePaidAmt = 0
        If TaxTrans.BillType = "R" Or TaxTrans.BillType = "P" Then
          NewTaxTrans.Revenue.PrePaidBal = OldRound(GetOverPayBalance(ThisCust, TaxTrans.BillType))
        Else
          NewTaxTrans.Revenue.PrePaidBal = 0
        End If
        NewTaxTrans.Revenue.PrePaidUsed = 0
        NewTaxTrans.OperNum = 0
        LSet NewTaxTrans.Padding = ""
        Put TTHandle, NextTTTrans, NewTaxTrans
        Fixed = True
'        Savex = NextTTTrans 'must move the target up one slot
        Print #AHandle, "Inserted penalty trans # " + CStr(NextTTTrans) + " into Cust # " + CStr(ThisCust) + "."
        End If
      If AdvAmt > 0 Then
        NextTTTrans = LOF(TTHandle) \ Len(TaxTrans) + 1
        Get TTHandle, Savex, TaxTrans
        NextRec = TaxCust.LastTrans
        cnt = cnt + 1
        NewTaxTrans = ClearTaxTrans
        If x = NextRec Then 'initialized trans is at the top of the queue
          TaxCust.LastTrans = NextTTTrans
          Put TCHandle, ThisCust, TaxCust
          NewTaxTrans.LastTrans = NextRec
          NextRec = NextTTTrans
        Else
          SaveRec = NextRec
          TopRec = SaveRec
          Do While SaveRec > 0
            Get TTHandle, SaveRec, TaxTrans
            For y = 1 To DoneCnt
              If DoneArr(y) = SaveRec Then GoTo SkipAdv
            Next y
            If TaxTrans.LastTrans = Savex Then
'              Get TTHandle, TopRec, TaxTrans
              TaxTrans.LastTrans = NextTTTrans
'              Put TTHandle, TopRec, TaxTrans
              Put TTHandle, TopRec, TaxTrans
             NewTaxTrans.LastTrans = Savex
              Exit Do
            End If
SkipAdv:
'            TopRec = SaveRec
            SaveRec = TaxTrans.LastTrans
          Loop
        End If
        Get TTHandle, Savex, TaxTrans
        NewTaxTrans.TransDate = TaxTrans.TransDate
        NewTaxTrans.TaxYear = TaxTrans.TaxYear
        NewTaxTrans.TranType = 6
        NewTaxTrans.BillType = TaxTrans.BillType
        NewTaxTrans.Amount = AdvAmt
        NewTaxTrans.Revenue.Collection = AdvAmt
        NewTaxTrans.Description = "Tax Adv on Bill# " + ParseBillNum$(TaxTrans.Description)
        NewTaxTrans.Posted2GL = "N"
        NewTaxTrans.CustomerRec = ThisCust
        NewTaxTrans.CustPin = ThisCust
        NewTaxTrans.RealPin = RealPin
        NewTaxTrans.PersPin = PersPin
        NewTaxTrans.BelongTo = Savex
        NewTaxTrans.Revenue.PrePaidAmt = 0
        If TaxTrans.BillType = "R" Or TaxTrans.BillType = "P" Then
          NewTaxTrans.Revenue.PrePaidBal = OldRound(GetOverPayBalance(ThisCust, TaxTrans.BillType))
        Else
          NewTaxTrans.Revenue.PrePaidBal = 0
        End If
        NewTaxTrans.Revenue.PrePaidUsed = 0
        NewTaxTrans.OperNum = 0
        LSet NewTaxTrans.Padding = ""
        Put TTHandle, NextTTTrans, NewTaxTrans
        Fixed = True
'        Savex = NextTTTrans 'must move the target up one slot
        Print #AHandle, "Inserted advertising trans # " + CStr(NextTTTrans) + " into Cust # " + CStr(ThisCust) + "."
        End If
    End If
SkipIt:
    If DoPaid = True Then GoSub HandlePayments

    If Fixed = True Then
      DoneCnt = DoneCnt + 1
      ReDim Preserve DoneArr(1 To DoneCnt) As Long
      DoneArr(DoneCnt) = x
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTTRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next x
Unload frmVATaxShowPctComp
Close

MsgBox ("A total of " + CStr(cnt) + " initialized transactions have been corrected.")
'Call RelinkBelongTosWithBills
'
'Call FixCustomerRecsandPins
'Call FixUnusedPrepayIn2006
'Call FindCorrectQueue
'Start:
'Call FixZeroedOutCreditAtBilling
Exit Sub

HandlePayments:
      NextTTTrans = (LOF(TTHandle) \ Len(TaxTrans)) + 1
      NextPayRec = TaxCust.LastTrans
      If x = NextPayRec Then
        BottomRec = NextRec
        TaxCust.LastTrans = NextTTTrans
        Put TCHandle, ThisCust, TaxCust
      Else
        Do While NextPayRec > 0
           Get TTHandle, NextPayRec, TaxTrans
             If TaxTrans.LastTrans = Savex Then
'                Get TTHandle, TopRec, TaxTrans
                TaxTrans.LastTrans = NextTTTrans
'                Put TTHandle, TopRec, TaxTrans
                Put TTHandle, NextPayRec, TaxTrans
                BottomRec = Savex
                Get TTHandle, x, TaxTrans 'get the original bill rec back
                Exit Do
             End If
'           TopRec = NextPayRec
           NextPayRec = TaxTrans.LastTrans
        Loop
      End If
      cnt = cnt + 1
      PayTranRec = ClearTaxTrans
      PayTranRec.TransDate = TaxTrans.TransDate
      PayTranRec.TranType = 2
 
      PayTranRec.Revenue.Principle1Pd = TaxTrans.Revenue.Principle1Pd
      PayTranRec.Revenue.InterestPd = TaxTrans.Revenue.InterestPd
      PayTranRec.Revenue.CollectionPd = TaxTrans.Revenue.CollectionPd
      PayTranRec.Revenue.LateListPd = TaxTrans.Revenue.LateListPd
      PayTranRec.Revenue.PenaltyPd = TaxTrans.Revenue.PenaltyPd
      PayTranRec.Revenue.RevOpt1Pd = 0
      PayTranRec.Revenue.RevOpt2Pd = 0
      PayTranRec.Revenue.RevOpt3Pd = 0
      PayTranRec.Revenue.Principle2Pd = TaxTrans.Revenue.Principle2Pd
      PayTranRec.Revenue.Principle3Pd = TaxTrans.Revenue.Principle3Pd
      PayTranRec.Revenue.Principle4Pd = TaxTrans.Revenue.Principle4Pd
      PayTranRec.Revenue.Principle5Pd = TaxTrans.Revenue.Principle5Pd
      PayTranRec.CustPin = TaxTrans.CustPin
      PayTranRec.DiscXDate = TaxTrans.DiscXDate
      PayTranRec.RealPin = QPTrim$(TaxTrans.RealPin)
      PayTranRec.PersPin = QPTrim$(TaxTrans.PersPin)
      PayTranRec.Posted2GL = "N"
      PayTranRec.TaxYear = TaxTrans.TaxYear
      PayTranRec.DiscAmt = 0
      PayTranRec.OperNum = OperNum
      PayTranRec.Amount = Paid
      PayTranRec.Description = ParseBillNum$(TaxTrans.Description)
      PayTranRec.CustomerRec = TaxTrans.CustPin

      PayTranRec.BelongTo = x
      PayTranRec.Revenue.PrePaidAmt = 0
      PayTranRec.Revenue.PrePaidUsed = 0
      PayTranRec.Revenue.PrePaidBal = 0
      PayTranRec.InternalPin = TaxTrans.InternalPin
      PayTranRec.BillType = TaxTrans.BillType

      PayTranRec.LastTrans = BottomRec
      Put TTHandle, NextTTTrans, PayTranRec
      Fixed = True
'      Savex = NextTTTrans
      Print #AHandle, "Inserted payment trans # " + CStr(NextTTTrans) + " into Cust # " + CStr(ThisCust) + "."
     
  Return
End Sub

Private Sub FixPaidBillsWithNoPayments()
  Dim TaxTrans As TaxTransactionType
  Dim NewTaxTrans As TaxTransactionType
  Dim ClearTaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Long, z As Long
  Dim AHandle As Integer
  Dim cnt As Integer
  Dim NextRec As Long
  Dim NextRecToo As Long
  Dim NextRecErr As Long
  Dim BelongTo As Long
  Dim BillPaid As Double
  Dim CollectPaid As Double
  Dim LostPaid As Double
  Dim FoundIt As Boolean
  ReDim PayArr(1 To 1) As Long
  Dim PayCnt As Integer
  ReDim CustArr(1 To 1) As Integer
  Dim ThisCust As Integer
  Dim TopRec As Long
  Dim BottomRec As Long
  Dim NewTrans As Long
  AHandle = FreeFile
  Open "BillsPaidWithNoPayments.txt" For Output As AHandle
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  frmVATaxShowPctComp.Label1 = "Fixing Paid Bills With No Payments"
  frmVATaxShowPctComp.Show , Me
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
'    If x = 97 Then Stop
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
'      If NextRec = 7731 Then Stop
      If TaxTrans.TranType = 1 Then
        BillPaid = OldRound(TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.LateListPd)
        BillPaid = BillPaid + OldRound(TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.Principle1Pd)
        BillPaid = BillPaid + OldRound(TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd)
        BillPaid = BillPaid + OldRound(TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd)
        BillPaid = BillPaid + OldRound(TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd)
        BillPaid = BillPaid + OldRound(TaxTrans.Revenue.RevOpt3Pd)
        BelongTo = NextRec
        NextRecToo = TaxCust.LastTrans
        CollectPaid = 0
        Do While NextRecToo > 0
          Get TTHandle, NextRecToo, TaxTrans
          If TaxTrans.BelongTo = BelongTo Then
            CollectPaid = CollectPaid + OldRound(TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.LateListPd)
            CollectPaid = CollectPaid + OldRound(TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.Principle1Pd)
            CollectPaid = CollectPaid + OldRound(TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd)
            CollectPaid = CollectPaid + OldRound(TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd)
            CollectPaid = CollectPaid + OldRound(TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd)
            CollectPaid = CollectPaid + OldRound(TaxTrans.Revenue.RevOpt3Pd)
          End If
          NextRecToo = TaxTrans.LastTrans
        Loop
        FoundIt = False
        PayCnt = 0
        ReDim PayArr(1 To 1) As Long
        ReDim CustArr(1 To 1) As Integer
        If CollectPaid <> BillPaid Then 'trans in cust queue does not equal bill payments
          For y = 1 To NumOfTTRecs 'go see if they are somewhere in the file
            Get TTHandle, y, TaxTrans
            If TaxTrans.BelongTo = BelongTo Then 'here's one
              LostPaid = OldRound(TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.LateListPd)
              LostPaid = LostPaid + OldRound(TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.Principle1Pd)
              LostPaid = LostPaid + OldRound(TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd)
              LostPaid = LostPaid + OldRound(TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd)
              LostPaid = LostPaid + OldRound(TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd)
              LostPaid = LostPaid + OldRound(TaxTrans.Revenue.RevOpt3Pd + TaxTrans.PPTRADisc)
              If LostPaid > 0 Then 'selected bill trans has qualifying records
                ThisCust = TaxTrans.CustomerRec
                If ThisCust = 0 Then
                  ThisCust = TaxTrans.CustPin
                End If
                PayCnt = PayCnt + 1
                ReDim Preserve PayArr(1 To PayCnt) As Long
                PayArr(PayCnt) = y
                ReDim Preserve CustArr(1 To PayCnt) As Integer
                CustArr(PayCnt) = ThisCust
              End If
            End If
          Next y
          If PayCnt = 0 Then 'no trans found
            Get TTHandle, BelongTo, TaxTrans
              TaxTrans.Revenue.CollectionPd = 0
              TaxTrans.Revenue.InterestPd = 0
              TaxTrans.Revenue.LateListPd = 0
              TaxTrans.Revenue.PenaltyPd = 0
              TaxTrans.Revenue.Principle1Pd = 0
              TaxTrans.Revenue.Principle2Pd = 0
              TaxTrans.Revenue.Principle3Pd = 0
              TaxTrans.Revenue.Principle4Pd = 0
              TaxTrans.Revenue.Principle5Pd = 0
              TaxTrans.Revenue.RevOpt1Pd = 0
              TaxTrans.Revenue.RevOpt2Pd = 0
              TaxTrans.Revenue.RevOpt3Pd = 0
              TaxTrans.PPTRADisc = 0
            Put TTHandle, BelongTo, TaxTrans 'no pay found so make all payments equal zero
            Print #AHandle, "Cleared payments from trans # " + CStr(BelongTo) + " under Cust # " + CStr(x) + "."

          ElseIf PayCnt > 0 Then 'some qualifying trans found
            For y = 1 To PayCnt
              If CustArr(y) <> x Then 'trans is in wrong cust queue
                'remove this trans from curr cust
                Get TCHandle, CustArr(y), TaxCust 'get in error cust
                NextRecErr = TaxCust.LastTrans
                Do While NextRecErr > 0 'find the in error trans in the queue
                  Get TTHandle, NextRecErr, TaxTrans
                    BottomRec = TaxTrans.LastTrans
                    If NextRecErr = PayArr(y) Then 'this finds it
                      Get TTHandle, TopRec, TaxTrans
                        TaxTrans.LastTrans = BottomRec 'relinks the trans above the selected trans
                        'to the one below it...selected trans now is an orphan
                      Put TTHandle, TopRec, TaxTrans
                      Get TCHandle, x, TaxCust 'now go back and get the correct cust
                        TopRec = TaxCust.LastTrans 'hold the old last trans
                        TaxCust.LastTrans = PayArr(y) 'assign the in error trans as cust last trans
                      Put TCHandle, x, TaxCust
                      Get TTHandle, PayArr(y), TaxTrans 'now get the in error trans and make
                        'its last trans the old cust's last trans
                        TaxTrans.LastTrans = TopRec
                        TaxTrans.CustomerRec = x
                        TaxTrans.CustPin = x
                      Put TTHandle, PayArr(y), TaxTrans
                      Print #AHandle, "Removed trans # " + CStr(PayArr(y)) + " from cust # " + CStr(CustArr(y)) + " to " + CStr(x) + "."
                      Exit Do
                    End If
                  TopRec = NextRecErr
                  NextRecErr = TaxTrans.LastTrans
                Loop
              End If
            Next y
          End If
        End If
      End If
      Get TTHandle, NextRec, TaxTrans
      NextRec = TaxTrans.LastTrans
    Loop
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next x
  
  
  Close
  MsgBox ("A total of " + CStr(PayCnt) + " transactions were found and fixed. Look for BillsPaidWithNoPayments.txt in the Citipak folder for results.")
  
End Sub
Private Sub FixFinalFew()
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Long
  Dim NextRec As Long
  Dim BottomRec As Long
  Dim TopRec As Long
  Dim Found As Boolean
  Dim AHandle As Integer
  Dim BelongTo As Long
  Dim Bal As Double
  Dim NewRec As Long
  Dim SaveRec As Long
  
  On Error GoTo ERRORSTUFF
  SaveRec = 0
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  AHandle = FreeFile
  Open "Trans9Error.txt" For Output As AHandle
  'fix for 5277
  Get TTHandle, 45438, TaxTrans
  TaxTrans.TranType = 2
  TaxTrans.Amount = 2.68
  TaxTrans.Revenue.Principle1Pd = 2.68
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 45438, TaxTrans
  
  'fix for 4421
  Call InsertCreditAtBillingTrans("5/30/2007", 17.91, 4421, 2006, 17.91, 8531, "Personal", 8635, 8531, "4421")
  
  Get TTHandle, 8531, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 17.91
  TaxTrans.Revenue.Interest = 0
  Put TTHandle, 8531, TaxTrans
  
  Call ClearTrans(53957)
  
  'fix for 4093
  Get TTHandle, 45071, TaxTrans
  TaxTrans.TranType = 2
  TaxTrans.Amount = 28
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 45071, TaxTrans
  
  'fix for 3810
  Call InsertCreditAtBillingTrans("5/30/2007", 45.37, 3810, 2007, 45.37, 8593, "Personal", 8976, 8593, "5302007")

  Get TTHandle, 8593, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 45.37
  Put TTHandle, 8593, TaxTrans
  
  Call ClearTrans(21385)
  Call ClearTrans(20753)
  Get TTHandle, 18359, TaxTrans
  TaxTrans.Revenue.Principle1 = 45.37
  Put TTHandle, 18359, TaxTrans
  
  'fix for 3621
'  Call ClearTrans(4079)
'  Get TTHandle, 7162, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0
'   Put TTHandle, 7162, TaxTrans
'
'  Get TTHandle, 6098, TaxTrans
'  TaxTrans.Amount = 69.85
'  TaxTrans.Revenue.Principle1 = 69.85
'  TaxTrans.Revenue.Principle1Pd = 69.85
'  Put TTHandle, 6098, TaxTrans
'
'  Get TTHandle, 17954, TaxTrans
'  TaxTrans.TranType = 2
'  TaxTrans.Amount = 69.85
'  TaxTrans.Revenue.Principle1Pd = 69.85
'  TaxTrans.Revenue.PrePaidAmt = 0
'  TaxTrans.Revenue.PrePaidBal = 0
'  TaxTrans.Revenue.PrePaidUsed = 0
'  Put TTHandle, 17954, TaxTrans
'
'  Get TTHandle, 8792, TaxTrans
'  TaxTrans.Revenue.Principle1 = 143.19
'  TaxTrans.Revenue.Principle1Pd = 143.19
'  TaxTrans.Revenue.Interest = 3.49
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Amount = 143.19
'  Put TTHandle, 8792, TaxTrans
  
  'fix for 1735
  Get TTHandle, 7972, TaxTrans
  TaxTrans.Revenue.Principle2 = 0
  Put TTHandle, 7972, TaxTrans
  
  Get TTHandle, 510, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 184.87
  Put TTHandle, 510, TaxTrans

  Get TTHandle, 3886, TaxTrans
  TaxTrans.TranType = 2
  TaxTrans.Amount = 184.87
  TaxTrans.Revenue.Principle1Pd = 184.87
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 3886, TaxTrans
  
  'fix for 933
  Call InsertCreditAtBillingTrans("11/03/2009", 9.82, 933, 2009, 13.3, 31644, "Personal", 32946, 31644, "187")
  Get TTHandle, 31644, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 34.04
  TaxTrans.Revenue.InterestPd = 2.32
  TaxTrans.Revenue.PenaltyPd = 1.16
  Put TTHandle, 31644, TaxTrans
  
  Get TTHandle, 5920, TaxTrans
  TaxTrans.BillType = "Personal"
  Put TTHandle, 5920, TaxTrans
  
  Get TTHandle, 3347, TaxTrans
  TaxTrans.Amount = 12.49
  TaxTrans.Revenue.Principle1Pd = 12.49
  Put TTHandle, 3347, TaxTrans
  
  Get TTHandle, 142, TaxTrans
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.PenaltyPd = 0
  Put TTHandle, 142, TaxTrans

  Get TTHandle, 18616, TaxTrans
  TaxTrans.TranType = 2
  TaxTrans.Amount = 24.81
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 18616, TaxTrans
  
  Get TTHandle, 12202, TaxTrans
  TaxTrans.TranType = 2
  TaxTrans.Amount = 44.8
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 12202, TaxTrans
  Get TTHandle, 9812, TaxTrans
  TaxTrans.Revenue.Principle1 = 0
  TaxTrans.PPTRADisc = 0
  Put TTHandle, 9812, TaxTrans
  
  Get TTHandle, 11377, TaxTrans
  TaxTrans.Amount = 44.8
  Put TTHandle, 11377, TaxTrans
  
'  'fix for 502
'  Get TTHandle, 12059, TaxTrans
'  TaxTrans.Revenue.Principle3Pd = 63.22
'  Put TTHandle, 12059, TaxTrans
'
'  Get TTHandle, 12060, TaxTrans
'  TaxTrans.Revenue.PrePaidBal = 0
'  Put TTHandle, 12060, TaxTrans
'
'  Call ClearTrans(12060)
'  Get TTHandle, 21537, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Principle3Pd = 44.74
'  TaxTrans.Revenue.Interest = 0
'  TaxTrans.Revenue.Penalty = 0
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.PenaltyPd = 0
'  Put TTHandle, 21537, TaxTrans
'
'  Get TTHandle, 31474, TaxTrans
'  TaxTrans.Revenue.Principle1Pd = 0
'  TaxTrans.Revenue.Interest = 16.1
'  TaxTrans.Revenue.Penalty = 8.05
'  TaxTrans.Revenue.Principle3Pd = 23.82
'  TaxTrans.Revenue.InterestPd = 2.38
'  TaxTrans.Revenue.PenaltyPd = 1.19
'  Put TTHandle, 31474, TaxTrans
'
'  Get TTHandle, 11863, TaxTrans
'  TaxTrans.Amount = 52.98
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.PenaltyPd = 0
'  Put TTHandle, 11863, TaxTrans
'
'  Get TTHandle, 74, TaxTrans
'  TaxTrans.Revenue.InterestPd = 0
'  TaxTrans.Revenue.PenaltyPd = 0
'  Put TTHandle, 74, TaxTrans
  
  'fix for 712
  Call InsertCreditAtBillingTrans("5/30/2007", 19.16, 712, 2007, 7, 8511, "Personal", 9772, 8511, "5302007")
  Get TTHandle, 8511, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 19.16
  TaxTrans.Revenue.InterestPd = 0
  Put TTHandle, 8511, TaxTrans

  'fix for #280
  Get TTHandle, 24, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 34.24
  Put TTHandle, 24, TaxTrans
  
  
  Get TTHandle, 3517, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 34.24
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.Revenue.PrePaidAmt = 9.55
  TaxTrans.Revenue.PrePaidBal = 9.55
  Put TTHandle, 3517, TaxTrans
  
  'fix for #15
  Get TTHandle, 7945, TaxTrans
  TaxTrans.Amount = 7.55
  TaxTrans.Revenue.Principle1Pd = 7.55
  TaxTrans.TranType = 2
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 7945, TaxTrans
  
  'fix for 3100
  Call ClearTrans(44711)
  Call ClearTrans(40748)
  Call ClearTrans(16746)
  Call ClearTrans(25575)
  Call ClearTrans(6294)
  Call InsertCreditAtBillingTrans("12/27/2007", 2.02, 3100, 2007, 2.02, 16405, "Real", 16746, 16405, "5")
  
  Get TTHandle, 6277, TaxTrans
  TaxTrans.Amount = 1.3
  TaxTrans.Revenue.Principle1Pd = 1.3
  TaxTrans.Description = "6"
  TaxTrans.BelongTo = 5826
  Put TTHandle, 6277, TaxTrans
  
  Get TTHandle, 16405, TaxTrans
  TaxTrans.Revenue.Principle1 = 2.02
  TaxTrans.Revenue.Principle1Pd = 2.02
  TaxTrans.Revenue.Interest = 0
  Put TTHandle, 16405, TaxTrans
  
  'fix for 2712
  Get TTHandle, 3972, TaxTrans
  TaxTrans.Amount = 22.88
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.TranType = 2
  TaxTrans.Description = "816"
  TaxTrans.Revenue.Principle1Pd = 22.88
  TaxTrans.BillType = "Personal"
  Put TTHandle, 3972, TaxTrans

  Get TTHandle, 816, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 22.88
  Put TTHandle, 816, TaxTrans

    'fix for 2539
  Get TTHandle, 25322, TaxTrans
  TaxTrans.Amount = 56.32
  TaxTrans.Revenue.PrePaidAmt = 0
  TaxTrans.Revenue.PrePaidBal = 0
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.TranType = 2
  TaxTrans.Description = "808"
  TaxTrans.Revenue.Principle1Pd = 56.32
  Put TTHandle, 25322, TaxTrans
  
  Get TTHandle, 808, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 56.32
  Put TTHandle, 808, TaxTrans
  
   'fix for 2251
  ClearTrans (8595)
 Call InsertPayTrans("03/19/2007", 13.63, 2251, 2007, 13.63, 780, "Personal", 0, 8595, "780", 0, 0, SaveRec)
 Call InsertPayTrans("03/19/2007", 13.63, 2251, 2007, 13.63, 781, "Personal", 0, SaveRec, "781", 0, 0, SaveRec)
 Get TCHandle, 2251, TaxCust
 TaxCust.LastTrans = SaveRec
 Put TCHandle, 2251, TaxCust
  

  'fix for 1258
  Call InsertCreditAtBillingTrans("11/25/2008", 2158.21, 1258, 2007, 2158.21, 24238, "Real", 25170, 24238, "901")
  Get TTHandle, 24238, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 2158.21
  TaxTrans.Revenue.InterestPd = 0
  Put TTHandle, 24238, TaxTrans

  'fix for 97
  Get TTHandle, 7908, TaxTrans
  TaxTrans.TranType = 2
  TaxTrans.Revenue.Principle1Pd = 8.26
  TaxTrans.Amount = 8.26
'  TaxTrans.Amount = TaxTrans.Revenue.PrePaidUsed
'  TaxTrans.Revenue.Principle1Pd = TaxTrans.Revenue.PrePaidUsed
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 7908, TaxTrans
  
  
  'fix for 3805
  Call ClearTrans(8617)
  Get TTHandle, 7325, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 38.83
  Put TTHandle, 7325, TaxTrans
  
  Get TTHandle, 18321, TaxTrans
  TaxTrans.TranType = 2
  TaxTrans.Amount = 38.83
  TaxTrans.Revenue.Principle1Pd = 38.83
  TaxTrans.Revenue.PrePaidUsed = 0
  Put TTHandle, 18321, TaxTrans
  
  Get TTHandle, 6797, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  Put TTHandle, 6797, TaxTrans
  
  Get TTHandle, 7025, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0
  Put TTHandle, 7025, TaxTrans
  
  Get TTHandle, 45510, TaxTrans
  TaxTrans.Revenue.PrePaidBal = 0
  Put TTHandle, 45510, TaxTrans
 
  'fix for 1781
  Call ClearTrans(7949)
  'fix for 1781 -> 1015
  Get TTHandle, 8307, TaxTrans
  TaxTrans.LastTrans = 8185
  Put TTHandle, 8307, TaxTrans

  Get TTHandle, 8261, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 14.05
  Put TTHandle, 8261, TaxTrans

  Get TTHandle, 11611, TaxTrans
  TaxTrans.LastTrans = 8262
  Put TTHandle, 11611, TaxTrans

  Get TTHandle, 8262, TaxTrans
  TaxTrans.LastTrans = 8261
  TaxTrans.CustomerRec = 1015
  TaxTrans.CustPin = 1015
  Put TTHandle, 8262, TaxTrans

  'fix for 4140
  Get TTHandle, 7187, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 40.02
  Put TTHandle, 7187, TaxTrans
 
  Call BuildMBvsCustHistArr

  For x = 1 To CArrCnt
    Get TCHandle, CArr(x), TaxCust
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
'        If NextRec = 8262 Then Stop
      '  Print #AHandle, CStr(CArr(x)) + "~" + CStr(NextRec)
        If TaxTrans.TranType = 9 And TaxTrans.CustPin <> CArr(x) Then
           BelongTo = TaxTrans.BelongTo
'          If TaxTrans.CustPin = 1806 Then Stop
          Get TCHandle, TaxTrans.CustPin, TaxCust
          If TaxCust.Deleted <> 0 Then
            Call ClearTrans(NextRec)
          Else
            Get TTHandle, BelongTo, TaxTrans
            Bal = 0
            Bal# = OldRound#(Bal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
            Bal# = OldRound#(Bal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
            Bal# = OldRound#(Bal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
            Bal# = OldRound#(Bal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
            Bal# = OldRound#(Bal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
            If Bal = 0 Then
              Call ClearTrans(NextRec)
              Print #AHandle, CStr(TaxTrans.CustPin) + "~" + CStr(NextRec)
            Else
              Print #AHandle, CStr(TaxTrans.CustPin) + "~" + CStr(NextRec) + "~" + Using("##,###.00", Bal)
           End If
          End If
        End If
          Get TTHandle, NextRec, TaxTrans
         NextRec = TaxTrans.LastTrans
    Loop
   Next x
  ' GoTo Skip

  'fix for 4393
  Get TTHandle, 7951, TaxTrans
  TaxTrans.Amount = 13.49
  TaxTrans.Revenue.Principle2Pd = 13.49
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.TranType = 2
  Put TTHandle, 7951, TaxTrans
  
  Call ClearTrans(7863)
  Call ClearTrans(7853)

  'fix for 4253
  Call ClearTrans(8639)
  
  'fix for 4158
  Get TTHandle, 7356, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 0.76
  Put TTHandle, 7356, TaxTrans
  
  
  'fix for 3886
  Call ClearTrans(7834)
  Call ClearTrans(7822)
  Call ClearTrans(7820)

  'fix for 1832
  Call ClearTrans(5583)
  Call ClearTrans(6579)
  
  Get TTHandle, 5100, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 10.95
  Put TTHandle, 5100, TaxTrans
  
  Get TTHandle, 6578, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 10.94
  Put TTHandle, 6578, TaxTrans
  
  Get TTHandle, 7490, TaxTrans
  TaxTrans.Amount = 10.94
  TaxTrans.Revenue.Principle1Pd = 10.94
  Put TTHandle, 7490, TaxTrans
  
  'fix for 1735
  Get TTHandle, 7972, TaxTrans
  TaxTrans.Revenue.Principle2 = 0
  Put TTHandle, 7972, TaxTrans

  'fix for 1713
  Call ClearTrans(26894)
  Get TTHandle, 25475, TaxTrans
  TaxTrans.Amount = 7.35
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Revenue.Penalty = 0
  Put TTHandle, 25475, TaxTrans
  
  Get TTHandle, 26893, TaxTrans
  TaxTrans.Revenue.Principle1 = 19.58
  TaxTrans.Revenue.Principle1Pd = 0
  Put TTHandle, 26893, TaxTrans
  
  Get TTHandle, 32753, TaxTrans
  TaxTrans.Amount = 52.92
  TaxTrans.Revenue.Principle1 = 52.92
  Put TTHandle, 32753, TaxTrans
  
  Get TTHandle, 12097, TaxTrans
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Revenue.Penalty = 0
  Put TTHandle, 12097, TaxTrans

  Get TTHandle, 9975, TaxTrans
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Revenue.Penalty = 0
  Put TTHandle, 9975, TaxTrans
   
    'fix for 1568
  Call ClearTrans(7304)
  
  Get TTHandle, 37432, TaxTrans
  TaxTrans.Amount = 151.49
  TaxTrans.Revenue.PenaltyPd = 0
  TaxTrans.Revenue.Penalty = 151.49
  TaxTrans.Revenue.PrePaidUsed = 0
  TaxTrans.TranType = 14
  Put TTHandle, 37432, TaxTrans

  Get TTHandle, 31510, TaxTrans
  TaxTrans.Revenue.Principle1 = 64.06
  TaxTrans.Revenue.Principle1Pd = 52.21
  TaxTrans.Revenue.Penalty = 75.84
  TaxTrans.Revenue.PenaltyPd = 50.94
  Put TTHandle, 31510, TaxTrans
  
  Get TTHandle, 4966, TaxTrans
  TaxTrans.Revenue.Principle1Pd = 28.84
  Put TTHandle, 4966, TaxTrans
  
  Call ClearTrans(5537)
 
  'fix for 933
  Get TTHandle, 11377, TaxTrans
  TaxTrans.Amount = 44.8
  Put TTHandle, 11377, TaxTrans
  
  Get TTHandle, 9812, TaxTrans
  TaxTrans.Revenue.Principle1 = 0
  TaxTrans.PPTRADisc = 0
  Put TTHandle, 9812, TaxTrans
  'fix for 427
  Get TTHandle, 29614, TaxTrans
  TaxTrans.Amount = 10.65
  TaxTrans.Revenue.Principle1 = 6.96
  TaxTrans.Revenue.Interest = 3.34
  TaxTrans.Revenue.Penalty = 0.35
  Put TTHandle, 29614, TaxTrans

  Call ClearTrans(29615)
  
  Get TTHandle, 63, TaxTrans
  TaxTrans.Revenue.Interest = 0
  TaxTrans.Revenue.Penalty = 0
  TaxTrans.Revenue.InterestPd = 0
  TaxTrans.Revenue.PenaltyPd = 0
  Put TTHandle, 63, TaxTrans
  
  'fix for 880
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    NextRec = TaxCust.LastTrans
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
      If NextRec = 7873 Then
        Found = False
        BottomRec = TaxTrans.LastTrans
        If TaxCust.LastTrans = NextRec Then
          TaxCust.LastTrans = BottomRec
          Put TCHandle, x, TaxCust
          Exit Do
         Else
          Get TTHandle, TopRec, TaxTrans
          TaxTrans.LastTrans = BottomRec
          Put TTHandle, TopRec, TaxTrans
          Exit Do
        End If
      End If
      TopRec = NextRec
      NextRec = TaxTrans.LastTrans
    Loop
  Next x

  Get TCHandle, 880, TaxCust
  NextRec = TaxCust.LastTrans
  Do While NextRec > 0
     Get TTHandle, NextRec, TaxTrans
     If NextRec = 7874 Then
       BottomRec = TaxTrans.LastTrans
       TaxTrans.LastTrans = 7873
       Put TTHandle, 7874, TaxTrans
       Get TTHandle, 7873, TaxTrans
       TaxTrans.LastTrans = BottomRec
       TaxTrans.Revenue.Principle1Pd = 0
       Put TTHandle, 7873, TaxTrans
       Exit Do
     End If
     TopRec = NextRec
     NextRec = TaxTrans.LastTrans
  Loop
  Call ClearTrans(7874)
  
  'fix for 173
  Call ClearTrans(29591)
  Call cmdFixStephensCity_Click
Skip:
  Close
'  MsgBox ("Done.")
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxDataRepair", "FixFinalFew", Erl)
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

