VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmCustAddEdit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "1frmCustAddEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EditLib.fpLongInteger fpCustRecNo 
      Height          =   300
      Left            =   768
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   144
      Visible         =   0   'False
      Width           =   684
      _Version        =   196608
      _ExtentX        =   1206
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      Appearance      =   1
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Timer MsgAlertTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   96
      Top             =   144
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdSave 
      Height          =   384
      Left            =   8256
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7176
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   384
      Left            =   9624
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7152
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":0AA5
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdOwner 
      Height          =   384
      Left            =   6432
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7968
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":0C80
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdConHist 
      Height          =   384
      Left            =   6240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7128
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":0E5B
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdTranHist 
      Height          =   384
      Left            =   4872
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7104
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":1037
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdWorkHist 
      Height          =   384
      Left            =   3408
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7104
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":1213
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdPrintInfo 
      Height          =   384
      Left            =   1872
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7104
      Width           =   1248
      _Version        =   131072
      _ExtentX        =   2201
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
      ButtonDesigner  =   "1frmCustAddEdit.frx":13EF
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   336
      Left            =   0
      TabIndex        =   7
      Top             =   8532
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   593
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
            TextSave        =   "4:34 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "6/3/2005"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpText fpSearch 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   3720
      TabIndex        =   9
      Top             =   2208
      Width           =   1548
      _Version        =   196608
      _ExtentX        =   2730
      _ExtentY        =   572
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
      MarginTop       =   0
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "~ "
      MaxLength       =   10
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
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
   Begin EditLib.fpText fpCustName 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   3720
      TabIndex        =   10
      Top             =   2604
      Width           =   3924
      _Version        =   196608
      _ExtentX        =   6921
      _ExtentY        =   572
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
      MarginTop       =   0
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
   Begin EditLib.fpText fpAddr1 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   3720
      TabIndex        =   11
      Top             =   2988
      Width           =   3924
      _Version        =   196608
      _ExtentX        =   6921
      _ExtentY        =   572
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
      MarginTop       =   0
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
   Begin EditLib.fpText fpAddr2 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   3720
      TabIndex        =   12
      Top             =   3372
      Width           =   3924
      _Version        =   196608
      _ExtentX        =   6921
      _ExtentY        =   572
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
      MarginTop       =   0
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
   Begin EditLib.fpMask fpZip 
      Height          =   324
      Left            =   7536
      TabIndex        =   17
      Top             =   3768
      Width           =   1548
      _Version        =   196608
      _ExtentX        =   2730
      _ExtentY        =   572
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
   Begin EditLib.fpText fpCity 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   3720
      TabIndex        =   18
      Top             =   3768
      Width           =   2100
      _Version        =   196608
      _ExtentX        =   3704
      _ExtentY        =   572
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
      MarginTop       =   0
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
      MaxLength       =   18
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
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
   Begin EditLib.fpText fpState 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   6552
      TabIndex        =   19
      Top             =   3768
      Width           =   420
      _Version        =   196608
      _ExtentX        =   741
      _ExtentY        =   572
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
      MarginTop       =   0
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   "~-0123456789"
      MaxLength       =   2
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
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
   Begin EditLib.fpMask fpSoSec 
      Height          =   300
      Left            =   4032
      TabIndex        =   23
      Top             =   4824
      Width           =   1620
      _Version        =   196608
      _ExtentX        =   2857
      _ExtentY        =   529
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
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   "###-##-####"
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
   Begin EditLib.fpText fpDrvLic 
      CausesValidation=   0   'False
      Height          =   300
      Left            =   4032
      TabIndex        =   24
      Top             =   5208
      Width           =   2100
      _Version        =   196608
      _ExtentX        =   3704
      _ExtentY        =   529
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
      MarginTop       =   0
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
      MaxLength       =   16
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
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
   Begin EditLib.fpBoolean fpCashOnly 
      Height          =   300
      Left            =   8544
      TabIndex        =   28
      Top             =   4824
      Width           =   324
      _Version        =   196608
      _ExtentX        =   572
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "N"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   0
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   0
      ToggleTrue      =   "Yy"
      TextTrue        =   "Y"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "N"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpLateFee 
      Height          =   300
      Left            =   8544
      TabIndex        =   30
      Top             =   5952
      Width           =   324
      _Version        =   196608
      _ExtentX        =   572
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "N"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   0
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   0
      ToggleTrue      =   "Yy"
      TextTrue        =   "Y"
      Value           =   1
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "N"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpCutOffYN 
      Height          =   300
      Left            =   8544
      TabIndex        =   31
      Top             =   5580
      Width           =   324
      _Version        =   196608
      _ExtentX        =   572
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "N"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   0
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   0
      ToggleTrue      =   "Yy"
      TextTrue        =   "Y"
      Value           =   1
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "N"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpBoolean fpTaxExpt 
      Height          =   300
      Left            =   8544
      TabIndex        =   32
      Top             =   5196
      Width           =   324
      _Version        =   196608
      _ExtentX        =   572
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "N"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   0
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   0
      ToggleTrue      =   "Yy"
      TextTrue        =   "Y"
      Value           =   0
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "N"
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpDateTime txtopnDate 
      Height          =   300
      Left            =   4032
      TabIndex        =   36
      Top             =   5592
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
      _ExtentY        =   529
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
      Text            =   "11/06/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "19800101"
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
   Begin EditLib.fpLongInteger fpAcct 
      Height          =   324
      Left            =   3720
      TabIndex        =   37
      Top             =   1824
      Width           =   1872
      _Version        =   196608
      _ExtentX        =   3302
      _ExtentY        =   572
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
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit an Existing Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2700
      TabIndex        =   39
      Top             =   1008
      Width           =   5652
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000005&
      Height          =   612
      Left            =   2592
      Top             =   888
      Width           =   5772
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   2172
      Left            =   1344
      Top             =   4320
      Width           =   8052
   End
   Begin VB.Shape Shape3 
      Height          =   2220
      Left            =   1320
      Top             =   4296
      Width           =   8100
   End
   Begin VB.Label Label2b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   1
      Left            =   696
      TabIndex        =   38
      Top             =   1872
      Width           =   2856
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicles On File (Y/N):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   6192
      TabIndex        =   35
      Top             =   5976
      Width           =   2220
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Owner (Y/N):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   6192
      TabIndex        =   34
      Top             =   5604
      Width           =   2220
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Residential (Y/N):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   6192
      TabIndex        =   33
      Top             =   5220
      Width           =   2220
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Only (Y/N):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   6192
      TabIndex        =   29
      Top             =   4848
      Width           =   2220
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Account Opened:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1392
      TabIndex        =   27
      Top             =   5640
      Width           =   2484
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Social Security Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1416
      TabIndex        =   26
      Top             =   4872
      Width           =   2460
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Drivers Licenses No:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1656
      TabIndex        =   25
      Top             =   5256
      Width           =   2220
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1332
      TabIndex        =   22
      Top             =   3792
      Width           =   2220
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   5784
      TabIndex        =   21
      Top             =   3792
      Width           =   684
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Zip:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   6912
      TabIndex        =   20
      Top             =   3792
      Width           =   564
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Search Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1332
      TabIndex        =   16
      Top             =   2256
      Width           =   2220
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1332
      TabIndex        =   15
      Top             =   2652
      Width           =   2220
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 1:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1332
      TabIndex        =   14
      Top             =   3036
      Width           =   2220
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 2:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1332
      TabIndex        =   13
      Top             =   3420
      Width           =   2220
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   732
      Left            =   2592
      Top             =   768
      Width           =   5772
   End
   Begin VB.Shape Shape1 
      Height          =   2988
      Left            =   1320
      Top             =   1320
      Width           =   8100
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   2940
      Left            =   1344
      Top             =   1344
      Width           =   8052
   End
End
Attribute VB_Name = "frmCustAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim RecNo As Long, CntL As Long
Dim TransRec As Long, MsgRec As Long
Dim UBSetupLen As Integer, cnt As Integer
Dim OldBook As String, NBook As String
Dim FinalFlag As Boolean, UpDateOwner As Boolean
Dim BeenDone As Boolean
Dim BtnFnt As Double
Dim fromform As Form, toform As Form, codeopt As Integer
Dim dontdoit As Boolean
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
End Sub
Private Sub Form_Load()
  BlockInput True
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  DoEvents
  dontdoit = False
  BlockInput False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If fpCmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        DCLog "Closed via CustAddEdit by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    'DoEvents
    Temp_Class.ResizeControls Me
   ' DoEvents
   ' Me.Visible = True
   ' Me.AutoRedraw = False
   ' DoEvents
  End If
  DoEvents
End Sub

Private Sub Form_Activate()
  BlockInput True
  If Val(frmCustAddEdit.fpCustRecNo) > 0 And Not BeenDone Then
    BeenDone = True
   ' LoadCustInfo2Form
    DoEvents
  ElseIf Val(frmCustAddEdit.fpCustRecNo) <= 0 And Not BeenDone Then
    DoEvents
    fpCmdTranHist.Enabled = False
    fpCmdPrintInfo.Enabled = False
    fpCmdWorkHist.Enabled = False
    fpCmdConHist.Enabled = False
    NewCustDefaults
  End If
  BlockInput False
End Sub
'
'Mouse/Keyboard/Button events
'
Private Sub NewCustDefaults()
'  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
'  fpCity = QPTrim(UBSetUpRec(1).DEFCITY)
'  fpState = QPTrim(UBSetUpRec(1).DEFSTATE)
'  fpZip = QPTrim(UBSetUpRec(1).ZIPCODE)
'  fpstatus.ListIndex = 0
'  fpOpenDate = Format(Now, "mm/dd/yyyy")
'  fpBillTo.ListIndex = 0
'  fpBillCopy = 1
'  For cnt = 0 To 6
'    fpMtrMulti(cnt) = 1
'    fpMtrUser(cnt) = 1
'  Next
'  fpGroupCde.ListIndex = 0
' ' LblInfo.Caption = "New"
' UBLog PWUser + " New Cust Entry"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
  '    Call fpCmdExit_Click
'    Case vbKeyPageDown
'    '  If Not ListIsDown Then
'        KeyCode = 0
'        If vaTabPro1.ActiveTab < 3 Then
'          vaTabPro1.ActiveTab = vaTabPro1.ActiveTab + 1
'        Else
'          vaTabPro1.ActiveTab = 0
'        End If
'      'End If
'    Case vbKeyPageUp
''      If Not ListIsDown Then
'        KeyCode = 0
'        If vaTabPro1.ActiveTab > 0 Then
'          vaTabPro1.ActiveTab = vaTabPro1.ActiveTab - 1
'        Else
'          vaTabPro1.ActiveTab = 3
'        End If
''      Else
''        KeyCode = 0
''      End If
'    Case vbKeyF2
'      KeyCode = 0
'      Call fpcmdPrintinfo_Click
'    Case vbKeyF3
'      KeyCode = 0
'    '  Call fpCmdWorkHist_Click
'    Case vbKeyF4
'      KeyCode = 0
'      Call fpCmdTranHist_Click
'      'trans history
'    Case vbKeyF6
'      KeyCode = 0
'      Call fpCmdConHist_Click
'    Case vbKeyF7
'      KeyCode = 0
'      Call fpCmdMsg_Click
'    Case vbKeyF8
'      KeyCode = 0
'      Call fpCmdOwner_Click
'    Case vbKeyF9
'      KeyCode = 0
'      Call fpCmdWOE_Click
'    Case vbKeyF10
'      KeyCode = 0
'      DoEvents
'      If ChkCustInfoOK% Then
'        If dontdoit = False Then
'          Call SaveCustInfo2Disk
'        End If
'        DoEvents
'        Call ExitCustAddEdit
'      End If
'032003
'    Case vbKeyReturn
'      KeyCode = 0
'      SendKeys "{tab}", True   ' Set the focus to the next control.
'      DoEvents
'    Case Else:
  End Select
End Sub


'Private Sub Form_KeyPress(KeyAscii As Integer)
'  If KeyAscii = vbKeyReturn Then  ' The ENTER key.
'    KeyAscii = 0        ' Ignore this key.
'    SendKeys "{tab}", True   ' Set the focus to the next control.
'    DoEvents
'  End If
'End Sub

Private Sub fpcmdPrintinfo_Click()
  If RecNo& > 0 Then
    frmReportOpt.Show 1
    DeActivateControls Me
    If rptopt = 1 Then
    'do the graphics
  '    PrintCustInfo RecNo&, 1
    ElseIf rptopt = 2 Then
    'do the text
  '    PrintCustInfo RecNo&, 2
    End If
   ActivateControls Me
  Else
    ActivateControls Me
  End If

End Sub


Private Sub btnPgUp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 Then
    DoEvents
    SendKeys "{PgUp}", True
  End If
End Sub

Private Sub btnPgDn_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 Then
    DoEvents
    SendKeys "{PgDn}", True
  End If
End Sub

'---------------------------------
'&&&&&& Page 1 Keydowns
'---------------------------------


'Private Sub fpCmdConHist_Click()
'  If RecNo > 0 Then
'    If Exist(UBPath$ + "UBTRANS.DAT") Then
'      frmRptConsumpHist.ShowCustConsHist (RecNo&)
'    Else
'      MsgBox "No Transactions to Display.", vbOKOnly, "No Transactions"
'    End If
'  End If
'End Sub

'Private Sub fpCmdWOE_Click()
'  If RecNo& <= 0 Then
'  'need to give option to save
'    If MsgBox("You must save new customer info before entering workorders.", vbOKCancel, "Save info?") = vbCancel Then
'      Exit Sub
'    Else
'      If ChkCustInfoOK% Then
'        If dontdoit = False Then
'          Call SaveCustInfo2Disk
'          dontdoit = False
'        End If
'        WorkOrders
'      End If
'    End If
'  Else
'    Select Case CheckSaveCustFile%
'    Case True:  '-1 save chenges
'    If ChkCustInfoOK% Then
'      If dontdoit = False Then
'        Call SaveCustInfo2Disk
'        dontdoit = False
'      End If
'      WorkOrders
'    End If
'    Case False:  '0= exit
'      WorkOrders
'    Case Else     '1 is review
'      'stay right where you are
'    End Select
'  End If
'End Sub
'Private Sub fpCmdWorkHist_Click()
'  If RecNo > 0 Then
'    frmRptWrkOrdHist.ShowWrkOrdHistory (RecNo&)
'  End If
'End Sub

'Private Sub fpCmdOwner_Click()
'  frmCustOwnerEdit.RecNo = RecNo
'  frmCustOwnerEdit.Show vbModal
'  DoEvents
'  UpDateOwner = frmCustOwnerEdit.ActionFlag
'  If UpDateOwner And RecNo > 0 Then  'an existing cust account
'    Call UBSaveOwnerInfo(RecNo)      'update owner info now. (user may not update cust)
'    UpDateOwner = False
'  End If                        'hay, Just forget about it.
'  DoEvents
'  'Call UNLoadOwnerForm
'  'Unload frmCustOwnerEdit
'End Sub

'Private Sub UBSaveOwnerInfo(OwnerRecNo As Long)
'  Dim UBFile As Integer, OwnerRecLen As Integer
'  OwnerRecLen = Len(UBOwnerRec)
'  UBOwnerRec.OwnFName = frmCustOwnerEdit.fpFirstName  'new owner info until user
'  UBOwnerRec.OwnLName = frmCustOwnerEdit.fpLastName   'saves new cust account.
'  UBOwnerRec.ADDR1 = frmCustOwnerEdit.fpAddr1
'  UBOwnerRec.ADDR2 = frmCustOwnerEdit.fpAddr2
'  UBOwnerRec.city = frmCustOwnerEdit.fpCity
'  UBOwnerRec.STATE = frmCustOwnerEdit.fpState
'  UBOwnerRec.ZIPCODE = frmCustOwnerEdit.fpZip
'  UBOwnerRec.HPHONE = frmCustOwnerEdit.fpHPhone
'  UBOwnerRec.WPHONE = frmCustOwnerEdit.fpWPhone
'  UBOwnerRec.ChkByte = Chr$(1)
'  UBFile = FreeFile
'  Open UBOwnerFile For Random Shared As UBFile Len = OwnerRecLen
'  Put UBFile, OwnerRecNo, UBOwnerRec
'  Close UBFile
'End Sub
'
'Private Sub fpCmdSave_Click()  'f10
' If ChkCustInfoOK% Then
'    If dontdoit = False Then
'      Call SaveCustInfo2Disk
'      Call ExitCustAddEdit
'    End If
'  End If
'End Sub
'
'Private Sub fpCmdExit_Click()
'  Select Case CheckSaveCustFile%
'  Case True:  '-1 save changes
'  If ChkCustInfoOK% Then
'    If dontdoit = False Then
'      Call SaveCustInfo2Disk
'    End If
'    Call ExitCustAddEdit
'  End If
'  Case False:  '0= exit
''    ExitingForm = True
'    Call ExitCustAddEdit
'  Case Else     '1 is review
'    'continue editing
'  End Select
'End Sub
''F7
''Display customer transaction history
'Private Sub fpCmdTranHist_Click()
'  ReDim MsgText(0 To 5) As String
'  Dim FntSize As Integer
'  If TransRec > 0 Then
'    'DeActivateControls Me
'    DisplayCustTransList RecNo
'    'ActivateControls Me
'    Select Case vaTabPro1.ActiveTab
'    Case 0
'      If Me.fpBook.Enabled = True Then
'        Me.fpBook.SetFocus
'      Else
'        Me.fpstatus.SetFocus
'      End If
'    Case 1
'      Me.fpCashOnly.SetFocus
'    Case 2
'      Me.fpServCode(0).SetFocus
'    Case 3
'      Me.fpMonOwed(0).SetFocus
'    End Select
'  Else
'  MsgBox "No Transactions to Display.", vbOKOnly, "No Transactions"
''    frmMsgDialog.RetLabel = "-2"
''    FntSize = frmMsgDialog.Label(2).FontSize
''    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
''    MsgText(0) = "ERROR:"
''    MsgText(1) = ""
''    MsgText(2) = ""
''    MsgText(3) = "There are NO transactions to display."
''    MsgText(4) = ""
''    MsgText(5) = ""
''    GetOKorNot MsgText(), True
'  End If
'End Sub
'
'Private Sub fpSeqNumb_LostFocus()
'  Call ChkFormatBookSeqN
'End Sub
'
'Private Sub ExitCustAddEdit()
'On Local Error Resume Next
'  DoEvents
'  RecNo = 0
'  BeenDone = False
'  TransRec = 0
'  fpCustRecNo = 0
'  NBook$ = ""
'  MsgRec = 0
'  OldBook = ""
'  FinalFlag = False
'  UpDateOwner = False
''  Load frmUBCustMenu
''  DoEvents
''  frmUBCustMenu.Show
'  DoEvents
'  If codeopt = 1 Then
'    ActivateControls frmCustEditLookUP
'  ElseIf codeopt = 2 Then
'    ActivateControls frmDisplayList
'  End If
'  If codeopt = 0 Then
'    frmUBCustMenu.Show
'  End If
'  UBLog PWUser + " Exit CustAddEdit"
'  Unload frmCustAddEdit
'  Unload frmCustOwnerEdit
'  'Call UNLoadOwnerForm
''  DoEvents
'End Sub

'Private Sub SaveCustInfo2Disk()
' ' Dim ClearRFlag As Boolean
'  DeActivateControls frmCustAddEdit
'
'  ReDim tmpCustRec(1 To 2) As NewUBCustRecType
'  Dim UBHandle As Integer, CustRecLen As Integer
'  Dim ReindexFlag As Boolean
'  Dim NextCRec As Long
'  BlockInput True
'  dontdoit = True
'  CustRecLen = Len(tmpCustRec(1))
'  If RecNo& > 0 Then
'    UBHandle = FreeFile
'    Open UBCustFile For Random Shared As UBHandle Len = CustRecLen
'    Get #UBHandle, RecNo&, tmpCustRec(1)
'    Close UBHandle
'    LSet tmpCustRec(2) = tmpCustRec(1) 'copy for reindex comparison check below
'  End If
''  If tmpCustRec(2).Status <> "A" Then
''    ClearRFlag = True
''  Else
''    ClearRFlag = True
''  End If
'  tmpCustRec(1).Book = QPTrim$(fpBook.Text)
'  tmpCustRec(1).SEQNUMB = QPTrim$(fpSeqNumb.Text)
'  tmpCustRec(1).Status = QPTrim$(fpstatus.Text)
'  tmpCustRec(1).OPENDATE = Date2Num(fpOpenDate.Text)
'  tmpCustRec(1).SEARCH = QPTrim$(fpSearch.Text)
'  tmpCustRec(1).CustName = QPTrim$(fpCustName.Text)
'  tmpCustRec(1).ADDR1 = QPTrim$(fpAddr1.Text)
'  tmpCustRec(1).ADDR2 = QPTrim$(fpAddr2.Text)
'
'  tmpCustRec(1).ServAddr = QPTrim$(fpServAddr.Text)
'  tmpCustRec(1).city = QPTrim$(fpCity.Text)
'  tmpCustRec(1).STATE = QPTrim$(fpState.Text)
''check
'  tmpCustRec(1).ZIPCODE = QPTrim$(fpZip.Text)
'  tmpCustRec(1).DPCode = QPTrim$(fpDPCode.Text)
'  tmpCustRec(1).HPHONE = QPTrim$(fpHPhone.Text)
'  tmpCustRec(1).WPHONE = QPTrim$(fpWPhone.Text)
'  tmpCustRec(1).SOSEC = QPTrim$(fpSoSec.Text)
'  tmpCustRec(1).DRVLIC = QPTrim$(fpDrvLic.Text)
'  tmpCustRec(1).CUSTTYPE = QPTrim$(fpCustType.Text)
'  tmpCustRec(1).Addr911 = QPTrim$(fpAddr911.Text)
'  If fpBillTo.ListIndex = 1 Then
'    tmpCustRec(1).BillTo = "O"
'  Else
'    tmpCustRec(1).BillTo = "C"
'  End If
'  tmpCustRec(1).BILLCOPY = Val(fpBillCopy.Text)
'  tmpCustRec(1).POSTRTE = QPTrim$(fpPostRte.Text)
'  If Len(QPTrim$(fpBillCycl.Text)) = 0 Then
'    tmpCustRec(1).BILLCYCL = -32767
'  Else
'    tmpCustRec(1).BILLCYCL = Val(fpBillCycl.Text)
'  End If
'  tmpCustRec(1).ZONE = QPTrim$(fpZone.Text)
'  If Len(QPTrim$(fpSeq.Text)) = 0 Then
'    tmpCustRec(1).Seq = -32767
'  Else
'    tmpCustRec(1).Seq = Val(fpSeq.Text)
'  End If
'  fpGroupCde.col = 0
'  tmpCustRec(1).GroupCodeRec = Val(fpGroupCde.ColText)
'  tmpCustRec(1).CASHONLY = fpCashOnly.Text
'  tmpCustRec(1).LATEFEE = fpLateFee.Text
'  tmpCustRec(1).CUTOFFYN = fpCutOffYN.Text
'  tmpCustRec(1).TAXEXPT = fpTaxExpt.Text
'  tmpCustRec(1).SRCIT = fpSrCit.Text
'  tmpCustRec(1).USEDRAFT = fpUseDraft.Text
'  tmpCustRec(1).AcctType = QPTrim$(fpAcctType.Text)
'  tmpCustRec(1).BankName = QPTrim$(fpBankName.Text)
'  tmpCustRec(1).BANKLOC = QPTrim$(fpBankLoc.Text)
'  tmpCustRec(1).TRANSIT = QPTrim$(fpTransit.Text)
'  tmpCustRec(1).BankAcct = QPTrim$(fpBankAcct.Text)
'  tmpCustRec(1).BILLCMNT = QPTrim$(fpBillCmnt.Text)
'  tmpCustRec(1).PAYCMNT = QPTrim$(fpPayCmnt.Text)
'  tmpCustRec(1).PumpCode = QPTrim$(fpPumpCode.Text)
'  tmpCustRec(1).USERCODE1 = QPTrim$(fpUserCode1.Text)
'  tmpCustRec(1).USERCODE2 = QPTrim$(fpUserCode2.Text)
'  tmpCustRec(1).ProRatePCT = Val(QPTrim$(Str$(fpProRatePCT.Text)))
'  tmpCustRec(1).HHMSG1 = QPTrim$(fpHHMsg1.Text)
'  tmpCustRec(1).HHMSG2 = QPTrim$(fpHHMsg2.Text)
'  tmpCustRec(1).HHMSG3 = QPTrim$(fpHHMsg3.Text)
'
'  For cnt = 0 To 14
'    tmpCustRec(1).serv(cnt + 1).Ratecode = QPTrim$(fpServCode(cnt).Text)
'    tmpCustRec(1).serv(cnt + 1).RMtrType = QPTrim$(fpServMType(cnt).Text)
'  Next
'  For cnt = 0 To 3
'    tmpCustRec(1).FlatRates(cnt + 1).FRDESC = QPTrim$(fpFlatDesc(cnt).Text)
'    tmpCustRec(1).FlatRates(cnt + 1).FRAMT = Val(QPTrim$(Str$(fpFlatAmt(cnt).Text)))
'    If fpFlatFreq(cnt).ListIndex = 0 Then
'      tmpCustRec(1).FlatRates(cnt + 1).FRFREQ = "R"
'    ElseIf fpFlatFreq(cnt).ListIndex = 1 Then
'      tmpCustRec(1).FlatRates(cnt + 1).FRFREQ = "N"
'    Else
'      tmpCustRec(1).FlatRates(cnt + 1).FRFREQ = " "
'    End If
'    tmpCustRec(1).FlatRates(cnt + 1).REVSRC = Val(QPTrim$(Str$(fpFlatRevSrc(cnt).Text)))
'    tmpCustRec(1).FlatRates(cnt + 1).NumMin = Val(QPTrim$(Str$(fpFlatMin(cnt).Text)))
'  Next
'  For cnt = 0 To 1
'    tmpCustRec(1).Monthly(cnt + 1).AMTOWED = fpMonOwed(cnt)
'    tmpCustRec(1).Monthly(cnt + 1).TotAmtPD = fpMonPaid(cnt)
'    tmpCustRec(1).Monthly(cnt + 1).PayAmt = fpMonAmt(cnt)
'    tmpCustRec(1).Monthly(cnt + 1).RevSource = fpMonRev(cnt)
'  Next
'  tmpCustRec(1).MFEE1 = fpMemFee(0)
'  tmpCustRec(1).MFEE2 = fpMemFee(1)
'
'  For cnt = 0 To 6
'    tmpCustRec(1).LocMeters(cnt + 1).MtrNum = QPTrim$(fpMtrSerial(cnt))
'    If Len(QPTrim$(fpMtrMulti(cnt).Text)) > 0 Then
'      tmpCustRec(1).LocMeters(cnt + 1).MTRMulti = fpMtrMulti(cnt)
'    Else
'      tmpCustRec(1).LocMeters(cnt + 1).MTRMulti = -1
'    End If
'    tmpCustRec(1).LocMeters(cnt + 1).MtrType = QPTrim$(fpLocMType(cnt).Text)
'    tmpCustRec(1).LocMeters(cnt + 1).MtrUnit = QPTrim$(fpLocUnit(cnt).Text)
'    If Len(QPTrim$(fpMtrUser(cnt).Text)) > 0 Then
'      tmpCustRec(1).LocMeters(cnt + 1).NumUser = fpMtrUser(cnt)
'    Else
'      tmpCustRec(1).LocMeters(cnt + 1).NumUser = -1
'    End If
'    tmpCustRec(1).LocMeters(cnt + 1).InsDate = Date2Num(fpLocMtrIns(cnt).Text)
'    If Len(QPTrim$(fpLocMtrCur(cnt).Text)) > 0 Then
'    'If Not Len(QPTrim$(fpLocMtrCur(cnt).Text)) < 0 Then
'      tmpCustRec(1).LocMeters(cnt + 1).CurRead = fpLocMtrCur(cnt)
'    End If
'    If Len(QPTrim$(fpLocMtrPre(cnt).Text)) > 0 Then
'    'If Len(QPTrim$(fpLocMtrPre(cnt).Text)) < 0 Then
'      tmpCustRec(1).LocMeters(cnt + 1).PrevRead = fpLocMtrPre(cnt)
'    End If
'    tmpCustRec(1).LocMeters(cnt + 1).CurDate = Date2Num(fpLocMLRDate(cnt).Text)
'    tmpCustRec(1).LocMeters(cnt + 1).MtrIDNO = QPTrim$(fpMtrIDNO(cnt).Text)
''put new field here
'    If Not RecNo& > 0 Then
'      tmpCustRec(1).LocMeters(cnt + 1).MtrLat = 0
'      tmpCustRec(1).LocMeters(cnt + 1).MtrLng = 0
'    End If
''
''put new field here
'    'no no can't do the clear thing because of editing cust during meter read entry etc.
''    If ClearRFlag Then
''      tmpCustRec(1).LocMeters(cnt + 1).ReadFlag = "N"
''    Else
'
'      tmpCustRec(1).LocMeters(cnt + 1).ReadFlag = tmpCustRec(2).LocMeters(cnt + 1).ReadFlag
''    End If
'  Next
'  tmpCustRec(1).FillPad = ""
'  tmpCustRec(1).ChkByte = Chr$(5) 'changed this on 2/10/05 because of conversion
'
'  DoEvents
'
'  UBHandle = FreeFile
'  Open UBCustFile For Random Shared As UBHandle Len = CustRecLen
'
'  If RecNo& > 0 Then
'    Put #UBHandle, RecNo&, tmpCustRec(1)
'  Else
'    RecNo& = (LOF(UBHandle) / CustRecLen) + 1
'    Put #UBHandle, RecNo&, tmpCustRec(1)
'  End If
'  Close UBHandle
'  UBLog PWUser + " Saved Acct: " + Str(RecNo&) + "," + QPTrim$(fpstatus.Text) + "," + QPTrim$(fpCustName.Text) + "," + QPTrim$(fpBook.Text) + "-" + QPTrim$(fpSeqNumb.Text)
'
'  If UpDateOwner Then             'need to save new owner rec also
'    Call UBSaveOwnerInfo(RecNo&)
'  End If
'
'  If RecNo& > 0 Then
'    If tmpCustRec(1).SEARCH <> tmpCustRec(2).SEARCH Then
'      ReindexFlag = True
'    End If
'    If tmpCustRec(1).CustName <> tmpCustRec(2).CustName Then
'      ReindexFlag = True
'    End If
'    If (tmpCustRec(1).Book <> tmpCustRec(2).Book) Then
'      ReindexFlag = True
'    End If
'    If (tmpCustRec(1).SEQNUMB <> tmpCustRec(2).SEQNUMB) Then
'      ReindexFlag = True
'    End If
'    For cnt = 1 To 7
'      If tmpCustRec(1).LocMeters(cnt).CurRead <> tmpCustRec(2).LocMeters(cnt).CurRead Then
'        UBLog PWUser + " Saved Acct: " + Str(RecNo&) + ",changed Curr read - " + Str(tmpCustRec(2).LocMeters(cnt).CurRead) + " to " + Str(tmpCustRec(1).LocMeters(cnt).CurRead)
'      End If
'      If tmpCustRec(1).LocMeters(cnt).PrevRead <> tmpCustRec(2).LocMeters(cnt).PrevRead Then
'        UBLog PWUser + " Saved Acct: " + Str(RecNo&) + ",changed Prev read - " + Str(tmpCustRec(2).LocMeters(cnt).PrevRead) + " to " + Str(tmpCustRec(1).LocMeters(cnt).PrevRead)
'      End If
'    Next
'  Else  'adding new account set flag to reindex
'    ReindexFlag = True
'  End If
'  DoEvents
'  If ReindexFlag Then
'    ReIndexSystem False
'    DoEvents
'  End If
'  Close
'  Erase tmpCustRec
'  BlockInput False
'  Call UPDateOK
'  ActivateControls frmCustAddEdit
'
'End Sub
'
'Private Sub LoadCustInfo2Form()
'  Dim tmpCustRec As NewUBCustRecType
'  Dim UBHandle As Integer, CustRecLen As Integer
'  CustRecLen = Len(tmpCustRec)
'
'  RecNo& = Val(frmCustAddEdit.fpCustRecNo)
'  frmCustAddEdit.fpCustRecNo = 0
'  UBLog PWUser + " Edit Acct: " + Str(RecNo&)
'  UBHandle = FreeFile
'  Open UBCustFile For Random Shared As UBHandle Len = CustRecLen
'
'  Get #UBHandle, RecNo&, tmpCustRec
'  Close UBHandle
'
'  If CustHasMsg(RecNo) Then
'    MsgAlertTimer.Enabled = True
'    'MsgRec = tmpCustRec.MessageRec
'  End If
'
'  If tmpCustRec.LastTrans > 0 Then
'    TransRec = tmpCustRec.LastTrans
'  End If
'
'  If tmpCustRec.Status = "F" Then
'    FinalFlag = True
'    fpBook.Enabled = False
'    fpSeqNumb.Enabled = False
'  Else
'    FinalFlag = False
'    fpBook.Enabled = True
'    fpSeqNumb.Enabled = True
'  End If
'
'  OldBook$ = tmpCustRec.Book + "-" + tmpCustRec.SEQNUMB
'
'  LabelAcctNo.Caption = RecNo&
'  fpBook = tmpCustRec.Book
'  fpSeqNumb = tmpCustRec.SEQNUMB
'  fpstatus.Text = " " + tmpCustRec.Status
'  fpOpenDate = Num2Date(tmpCustRec.OPENDATE)
'  fpSearch = QPTrim$(tmpCustRec.SEARCH)
'  fpCustName = QPTrim$(tmpCustRec.CustName)
'  'LblInfo.Caption = QPTrim$(tmpCustRec.CustName)
'  fpAddr1 = QPTrim$(tmpCustRec.ADDR1)
'  fpAddr2 = QPTrim$(tmpCustRec.ADDR2)
'  fpServAddr = QPTrim$(tmpCustRec.ServAddr)
'  fpCity = QPTrim$(tmpCustRec.city)
'  fpState = QPTrim$(tmpCustRec.STATE)
'  fpZip = QPTrim$(tmpCustRec.ZIPCODE)
'  fpDPCode = QPTrim$(tmpCustRec.DPCode)
'  fpHPhone.Text = QPTrim$(tmpCustRec.HPHONE)
'  fpWPhone = QPTrim$(tmpCustRec.WPHONE)
''Stop
''here
'  fpGroupCde.col = 0
'  fpGroupCde.SearchText = Str$(tmpCustRec.GroupCodeRec)
'  fpGroupCde.Action = 0
'  If fpGroupCde.SearchIndex <> -1 Then
'    fpGroupCde.ListIndex = fpGroupCde.SearchIndex
'  Else
'    fpGroupCde.ListIndex = 0
'  End If
'
'  fpSoSec = QPTrim$(tmpCustRec.SOSEC)
'  fpDrvLic = QPTrim$(tmpCustRec.DRVLIC)
'  fpCustType = QPTrim$(tmpCustRec.CUSTTYPE)
'  fpAddr911 = QPTrim$(tmpCustRec.Addr911)
'  If QPTrim$(tmpCustRec.BillTo) = "O" Then
'    fpBillTo.ListIndex = 1
'  Else
'    fpBillTo.ListIndex = 0
'  End If
'  fpBillCopy = QPTrim$(Str$(tmpCustRec.BILLCOPY))
'  fpPostRte = QPTrim$(tmpCustRec.POSTRTE)
'  If tmpCustRec.BILLCYCL >= 0 Then
'    fpBillCycl = QPTrim$(Str$(tmpCustRec.BILLCYCL))
'  Else
'    fpBillCycl = ""
'  End If
'  fpZone = QPTrim$(tmpCustRec.ZONE)
'  If tmpCustRec.Seq >= 0 Then
'    fpSeq = QPTrim$(Str$(tmpCustRec.Seq))
'  Else
'    fpSeq = ""
'  End If
'  Select Case tmpCustRec.CASHONLY
'  Case "N", " "
'    fpCashOnly.Value = ValueFalse
'  Case Else
'    fpCashOnly.Value = ValueTrue
'  End Select
'  Select Case tmpCustRec.LATEFEE
'  Case "N", " "
'    fpLateFee.Value = ValueFalse
'  Case Else
'    fpLateFee.Value = ValueTrue
'  End Select
'  Select Case tmpCustRec.CUTOFFYN
'  Case "N", " "
'    fpCutOffYN.Value = ValueFalse
'  Case Else
'    fpCutOffYN.Value = ValueTrue
'  End Select
'  Select Case tmpCustRec.TAXEXPT
'  Case "N", " "
'    fpTaxExpt.Value = ValueFalse
'  Case Else
'    fpTaxExpt.Value = ValueTrue
'  End Select
'  Select Case tmpCustRec.SRCIT
'  Case "N", " "
'    fpSrCit.Value = ValueFalse
'  Case Else
'    fpSrCit.Value = ValueTrue
'  End Select
'  Select Case tmpCustRec.USEDRAFT
'  Case "Y"
'    fpUseDraft.Value = ValueTrue
'  Case Else
'    fpUseDraft.Value = ValueFalse
'  End Select
'
'  fpAcctType = QPTrim$(tmpCustRec.AcctType)
'  fpBankName = QPTrim$(tmpCustRec.BankName)
'  fpBankLoc = QPTrim$(tmpCustRec.BANKLOC)
'  fpTransit = QPTrim$(tmpCustRec.TRANSIT)
'  fpBankAcct = QPTrim$(tmpCustRec.BankAcct)
'  fpBillCmnt = QPTrim$(tmpCustRec.BILLCMNT)
'  fpPayCmnt = QPTrim$(tmpCustRec.PAYCMNT)
'  fpPumpCode = QPTrim$(tmpCustRec.PumpCode)
'  fpUserCode1 = QPTrim$(tmpCustRec.USERCODE1)
'  fpUserCode2 = QPTrim$(tmpCustRec.USERCODE2)
'  fpProRatePCT = QPTrim$(Str$(tmpCustRec.ProRatePCT))
'  fpHHMsg1 = QPTrim$(tmpCustRec.HHMSG1)
'  fpHHMsg2 = QPTrim$(tmpCustRec.HHMSG2)
'  fpHHMsg3 = QPTrim$(tmpCustRec.HHMSG3)
'
'  For cnt = 0 To 14
'    fpServCode(cnt).Text = QPTrim$(tmpCustRec.serv(cnt + 1).Ratecode)
'    fpServMType(cnt).Text = QPTrim$(tmpCustRec.serv(cnt + 1).RMtrType)
'  Next
'
'  For cnt = 0 To 3
'    fpFlatDesc(cnt).Text = QPTrim$(tmpCustRec.FlatRates(cnt + 1).FRDESC)
'    fpFlatAmt(cnt).Text = QPTrim$(Str$(tmpCustRec.FlatRates(cnt + 1).FRAMT))
'    If QPTrim$(tmpCustRec.FlatRates(cnt + 1).FRFREQ) = "R" Then
'      fpFlatFreq(cnt).ListIndex = 0
'    ElseIf QPTrim$(tmpCustRec.FlatRates(cnt + 1).FRFREQ) = "N" Then
'      fpFlatFreq(cnt).ListIndex = 1
'    Else
'      fpFlatFreq(cnt).ListIndex = -1
'    End If
'    fpFlatRevSrc(cnt).Text = QPTrim$(Str$(tmpCustRec.FlatRates(cnt + 1).REVSRC))
'    fpFlatMin(cnt).Text = QPTrim$(Str$(tmpCustRec.FlatRates(cnt + 1).NumMin))
'  Next
'
'  For cnt = 0 To 1
'    'fpMonOwed(Cnt).Text = QPTrim$(Str$(tmpCustRec.Monthly(Cnt + 1).AMTOWED))
'    fpMonOwed(cnt) = tmpCustRec.Monthly(cnt + 1).AMTOWED
'    fpMonPaid(cnt) = tmpCustRec.Monthly(cnt + 1).TotAmtPD
'    fpMonAmt(cnt) = tmpCustRec.Monthly(cnt + 1).PayAmt
'    fpMonRev(cnt) = tmpCustRec.Monthly(cnt + 1).RevSource
'  Next
'  fpMemFee(0) = tmpCustRec.MFEE1
'  fpMemFee(1) = tmpCustRec.MFEE2
'
'  For cnt = 0 To 6
'    fpMtrSerial(cnt) = QPTrim$(tmpCustRec.LocMeters(cnt + 1).MtrNum)
'    If tmpCustRec.LocMeters(cnt + 1).MTRMulti >= 0 Then
'      fpMtrMulti(cnt) = tmpCustRec.LocMeters(cnt + 1).MTRMulti
'    End If
'    fpLocMType(cnt).Text = QPTrim$(tmpCustRec.LocMeters(cnt + 1).MtrType)
'    fpLocUnit(cnt).Text = QPTrim$(tmpCustRec.LocMeters(cnt + 1).MtrUnit)
'    If tmpCustRec.LocMeters(cnt + 1).NumUser > 0 Then
'      fpMtrUser(cnt) = tmpCustRec.LocMeters(cnt + 1).NumUser
'    End If
'    fpLocMtrIns(cnt).Text = Num2Date(tmpCustRec.LocMeters(cnt + 1).InsDate)
'    'If tmpCustRec.LocMeters(cnt + 1).CurRead > 0 Then
'    If tmpCustRec.LocMeters(cnt + 1).CurRead >= 0 Then
'      fpLocMtrCur(cnt) = tmpCustRec.LocMeters(cnt + 1).CurRead
'    End If
'    'If tmpCustRec.LocMeters(cnt + 1).PrevRead > 0 Then
'    If tmpCustRec.LocMeters(cnt + 1).PrevRead >= 0 Then
'      fpLocMtrPre(cnt) = tmpCustRec.LocMeters(cnt + 1).PrevRead
'    End If
'    fpLocMLRDate(cnt).Text = Num2Date(tmpCustRec.LocMeters(cnt + 1).CurDate)
'    fpMtrIDNO(cnt).Text = QPTrim$(tmpCustRec.LocMeters(cnt + 1).MtrIDNO)
'  Next
'
'  DoEvents
'End Sub
'
'Private Function CheckSaveCustFile%()
'  Dim Changed As Boolean
'  Dim chkCustRec As NewUBCustRecType
'  Dim UBHandle As Integer, Enoughtosave As Boolean
'  Dim CustRecLen As Integer
'  CustRecLen = Len(chkCustRec)
'  Enoughtosave = True
'  If UpDateOwner Then 'check owner info
'    Changed = True
'    GoTo DoneCustChk
'  End If
'
'  If RecNo& > 0 Then
'    UBHandle = FreeFile
'    Open UBCustFile For Random Shared As UBHandle Len = CustRecLen
'    Get #UBHandle, RecNo&, chkCustRec
'    Close UBHandle
'
'    If QPTrim$(chkCustRec.Book) <> QPTrim$(fpBook.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.SEQNUMB) <> QPTrim$(fpSeqNumb.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.Status) <> QPTrim$(fpstatus.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If chkCustRec.OPENDATE <> Date2Num(fpOpenDate.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.SEARCH) <> QPTrim$(fpSearch.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.CustName) <> QPTrim$(fpCustName.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.ADDR1) <> QPTrim$(fpAddr1.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.ADDR2) <> QPTrim$(fpAddr2.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.ServAddr) <> QPTrim$(fpServAddr.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.city) <> QPTrim$(fpCity.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.STATE) <> QPTrim$(fpState.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If Val(chkCustRec.ZIPCODE) <> Val(fpZip.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.DPCode) <> QPTrim$(fpDPCode.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPStripStuff$(chkCustRec.HPHONE) <> QPStripStuff$(fpHPhone.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPStripStuff$(chkCustRec.WPHONE) <> QPStripStuff$(fpWPhone.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If Val(chkCustRec.SOSEC) <> Val(fpSoSec.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.DRVLIC) <> QPTrim$(fpDrvLic.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.CUSTTYPE) <> QPTrim$(fpCustType.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.Addr911) <> QPTrim$(fpAddr911.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    fpGroupCde.col = 0
'    If chkCustRec.GroupCodeRec <> Val(fpGroupCde.ColText) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If Len(QPTrim$(chkCustRec.BillTo)) > 0 Then
'    If QPTrim$(chkCustRec.BillTo) <> Mid$(fpBillTo.Text, 1, 1) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    End If
'    If chkCustRec.BILLCOPY <> Val(fpBillCopy) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.POSTRTE) <> QPTrim$(fpPostRte.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If Len(QPTrim$(fpBillCycl.Text)) = 0 Then
'      If Not chkCustRec.BILLCYCL = -32767 Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'    Else
'      If chkCustRec.BILLCYCL <> Val(fpBillCycl) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'    End If
'    If QPTrim$(chkCustRec.ZONE) <> QPTrim$(fpZone.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If Len(QPTrim$(fpSeq.Text)) = 0 Then
'      If Not chkCustRec.Seq = -32767 Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'    Else
'      If chkCustRec.Seq <> Val(fpSeq) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'    End If
'    If QPTrim$(chkCustRec.CASHONLY) <> fpCashOnly.Text Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.LATEFEE) <> fpLateFee.Text Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.CUTOFFYN) <> fpCutOffYN.Text Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.TAXEXPT) <> fpTaxExpt.Text Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.SRCIT) <> fpSrCit.Text Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If Len(QPTrim$(chkCustRec.USEDRAFT)) > 0 Then
'      If QPTrim$(chkCustRec.USEDRAFT) <> fpUseDraft.Text Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'    End If
'    If QPTrim$(chkCustRec.AcctType) <> QPTrim$(fpAcctType.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.BankName) <> QPTrim$(fpBankName.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.BANKLOC) <> QPTrim$(fpBankLoc.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.TRANSIT) <> QPTrim$(fpTransit.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.BankAcct) <> QPTrim$(fpBankAcct.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.BILLCMNT) <> QPTrim$(fpBillCmnt.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.PAYCMNT) <> QPTrim$(fpPayCmnt.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.PumpCode) <> QPTrim$(fpPumpCode.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.USERCODE1) <> QPTrim$(fpUserCode1.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.USERCODE2) <> QPTrim$(fpUserCode2.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If chkCustRec.ProRatePCT <> fpProRatePCT Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.HHMSG1) <> QPTrim$(fpHHMsg1.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.HHMSG2) <> QPTrim$(fpHHMsg2.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    If QPTrim$(chkCustRec.HHMSG3) <> QPTrim$(fpHHMsg3.Text) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'    For cnt = 0 To 14
'      If QPTrim$(chkCustRec.serv(cnt + 1).Ratecode) <> QPTrim$(fpServCode(cnt).Text) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'      If QPTrim$(chkCustRec.serv(cnt + 1).RMtrType) <> QPTrim$(fpServMType(cnt).Text) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'    Next
'
'    For cnt = 0 To 3
'      If QPTrim$(chkCustRec.FlatRates(cnt + 1).FRDESC) <> QPTrim$(fpFlatDesc(cnt).Text) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'      If chkCustRec.FlatRates(cnt + 1).FRAMT <> fpFlatAmt(cnt) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'      If QPTrim$(chkCustRec.FlatRates(cnt + 1).FRFREQ) <> Mid$(fpFlatFreq(cnt).ColText, 1, 1) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'      If chkCustRec.FlatRates(cnt + 1).REVSRC <> fpFlatRevSrc(cnt) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'      If chkCustRec.FlatRates(cnt + 1).NumMin <> fpFlatMin(cnt) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'    Next
'
'    For cnt = 0 To 1
'      If chkCustRec.Monthly(cnt + 1).AMTOWED <> fpMonOwed(cnt) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'      If chkCustRec.Monthly(cnt + 1).TotAmtPD <> fpMonPaid(cnt) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'      If chkCustRec.Monthly(cnt + 1).PayAmt <> fpMonAmt(cnt) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'      If chkCustRec.Monthly(cnt + 1).RevSource <> fpMonRev(cnt) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'    Next
'    If chkCustRec.MFEE1 <> fpMemFee(0) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'
'    If chkCustRec.MFEE2 <> fpMemFee(1) Then
'      Changed = True
'      GoTo DoneCustChk
'    End If
'
'    For cnt = 0 To 6
'      If QPTrim$(chkCustRec.LocMeters(cnt + 1).MtrNum) <> QPTrim$(fpMtrSerial(cnt)) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
''NOTE: DO NOT change this comparsion. Must be done this way to maintain
''      compatibility with old way of storing a blank numeric field. Old
''      method stored the maximum negitive value of the numeric variable type
''      (i.e. integer, double, long etc.) to represent a blank field. Since
''      the a meter multiplier can not be a negitive value, I am storing a
''      -1 (negitive one) to represent this in the new version.
'
'      If chkCustRec.LocMeters(cnt + 1).MTRMulti <= 0 Then
'        If Val(fpMtrMulti(cnt)) > 0 Then
'          Changed = True
'          GoTo DoneCustChk
'        End If
'      ElseIf chkCustRec.LocMeters(cnt + 1).MTRMulti <> Val(fpMtrMulti(cnt)) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'
'      If QPTrim$(chkCustRec.LocMeters(cnt + 1).MtrType) <> QPTrim$(fpLocMType(cnt).Text) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'
'      If QPTrim$(chkCustRec.LocMeters(cnt + 1).MtrUnit) <> QPTrim$(fpLocUnit(cnt).Text) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'      If Len(QPTrim$(fpMtrUser(cnt).Text)) > 0 Then
'        If chkCustRec.LocMeters(cnt + 1).NumUser <> fpMtrUser(cnt) Then
'          Changed = True
'          GoTo DoneCustChk
'        End If
'      Else
'        If chkCustRec.LocMeters(cnt + 1).NumUser <> -1 Then
'          Changed = True
'          GoTo DoneCustChk
'        End If
'      End If
'
'      If chkCustRec.LocMeters(cnt + 1).InsDate <> Date2Num(fpLocMtrIns(cnt).Text) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'      If Len(QPTrim$(fpLocMtrCur(cnt).Text)) > 0 Then
'        If chkCustRec.LocMeters(cnt + 1).CurRead <> fpLocMtrCur(cnt) Then
'         Changed = True
'         GoTo DoneCustChk
'        End If
'      End If
'      If Len(QPTrim$(fpLocMtrPre(cnt).Text)) > 0 Then
'        If chkCustRec.LocMeters(cnt + 1).PrevRead <> fpLocMtrPre(cnt) Then
'          Changed = True
'          GoTo DoneCustChk
'        End If
'      End If
'      If chkCustRec.LocMeters(cnt + 1).CurDate <> Date2Num(fpLocMLRDate(cnt).Text) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'      If QPTrim$(chkCustRec.LocMeters(cnt + 1).MtrIDNO) <> QPTrim$(fpMtrIDNO(cnt).Text) Then
'        Changed = True
'        GoTo DoneCustChk
'      End If
'    Next
'  Else
''    If fpstatus.ListIndex = -1 Then
''      Enoughtosave = False
''      GoTo DoneCustChk
''    End If
''    If Not Len(QPTrim$(fpSearch.Text)) > 0 Then
''      Enoughtosave = False
''      GoTo DoneCustChk
''    End If
''    If Not Len(QPTrim$(fpCustName.Text)) > 0 Then
''      Enoughtosave = False
''      GoTo DoneCustChk
''    End If
'  End If
'
'DoneCustChk:
'  DoEvents
''  ReDim MsgText(0 To 5) As String
''  Dim FntSize As Integer
''  If Enoughtosave = False Then
''    frmMsgDialog.RetLabel = "-2"
''    FntSize = frmMsgDialog.Label(2).FontSize
''    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
''    MsgText(0) = "ERROR:"
''    MsgText(1) = ""
''    MsgText(2) = ""
''    MsgText(3) = "The Name, SearchName and Status"
''    MsgText(4) = "fields are Required."
''    MsgText(5) = ""
''    GetOKorNot MsgText(), True
''    CheckSaveCustFile% = False
''  End If
'
'  If Changed Then
'    frmChangedWarning.Show vbModal, Me
'    Select Case SaveFlag
'    Case False
'      CheckSaveCustFile% = False
'    Case True
'      CheckSaveCustFile% = True
'    Case 1
'      CheckSaveCustFile% = 1
'    End Select
'  Else
'    CheckSaveCustFile% = False
'  End If
'  DoEvents
'End Function
'
'Private Sub ChkFormatBookSeqN()
'  Dim TBook As String
'  Dim TSeqN As String
'  TBook = QPTrim$(Me.fpBook)
'  TSeqN = QPTrim$(Me.fpSeqNumb)
'  Me.fpBook = FmtBook$(Me.fpBook)
'  Me.fpSeqNumb = FmtSeqN$(Me.fpSeqNumb)
'End Sub
'
'Private Function ChkCustInfoOK%()
'Dim Enoughtosave As Boolean
'  Enoughtosave = True
'  ChkCustInfoOK = False   'assume the worst.
'  If RecNo& > 0 Then
' ' If Not FinalFlag Then     'if this account isn't in final
'
'    Call ChkFormatBookSeqN
'    NBook$ = Me.fpBook + "-" + Me.fpSeqNumb
'    If QPTrim$(fpstatus.Text) = "F" And (OldBook$ <> NBook$) Then
'        vaTabPro1.ActiveTab = 0
'        DoEvents
'        MsgBox "   Final Status Does NOT Allow Location #'s!   " + Chr$(13) + Chr$(13) + "   Please enter a new Status or Location ", vbOKOnly, "ERROR!"
'        Me.fpstatus.SetFocus
'        ChkCustInfoOK = False
'    Else
'    If (OldBook$ <> NBook$) And (NBook$ <> "00-000000") Then
'      'If fpStatus.Text = "A" Or fpStatus.Text = "P" Then
'      If Not Val(Me.fpBook) > 0 And Val(Me.fpSeqNumb) > 0 Then
'        vaTabPro1.ActiveTab = 0
'        DoEvents
'        MsgBox "   Invalid Book!   " + Chr$(13) + Chr$(13) + "   Please enter a new Book number   ", vbOKOnly, "ERROR!"
'        Me.fpBook.SetFocus
'        ChkCustInfoOK = False
'      Else
'      'if they changed the book-seq list num
'      If Chk4DupeLocation(Me.fpBook, Me.fpSeqNumb) Then
'        If Len(OldBook$) > 1 Then
'          Me.fpBook = Left$(OldBook$, 2)
'          Me.fpSeqNumb = Mid$(OldBook$, 4)
'        Else
'          Me.fpBook = ""
'          Me.fpSeqNumb = ""
'        End If
'        vaTabPro1.ActiveTab = 0
'        DoEvents
'        MsgBox "   Duplicate Location Number Found!   " + Chr$(13) + Chr$(13) + "   Please enter a new location number   ", vbOKOnly, "ERROR!"
'        If Me.fpBook.Enabled = True Then
'          Me.fpBook.SetFocus
'        Else
'          Me.fpstatus.SetFocus
'        End If
'        ChkCustInfoOK = False
'      Else
'        ChkCustInfoOK = True
'      End If
'      End If
'    Else
'      ChkCustInfoOK = True
'    End If
'   End If
''  Else
''    ChkCustInfoOK = True
''  End If
'  Else
'    If Chk4DupeLocation(Me.fpBook, Me.fpSeqNumb) Then
'      If Len(OldBook$) > 1 Then
'        Me.fpBook = Left$(OldBook$, 2)
'        Me.fpSeqNumb = Mid$(OldBook$, 4)
'      Else
'        Me.fpBook = ""
'        Me.fpSeqNumb = ""
'      End If
'      vaTabPro1.ActiveTab = 0
'      DoEvents
'      MsgBox "   Duplicate Location Number Found!   " + Chr$(13) + Chr$(13) + "   Please enter a new location number   ", vbOKOnly, "ERROR!"
'      If fpBook.Enabled = True Then
'        Me.fpBook.SetFocus
'      Else
'        Me.fpstatus.SetFocus
'      End If
'      ChkCustInfoOK = False
'    Else
'      ChkCustInfoOK = True
'    End If
'
'    If fpstatus.ListIndex = -1 Then
'      Enoughtosave = False
'      GoTo DoneChk
'    End If
'    If Not Len(QPTrim$(fpSearch.Text)) > 0 Then
'      If QPTrim$(fpstatus.Text) = "A" Then
'        Enoughtosave = False
'        GoTo DoneChk
'      End If
'    End If
'    If Not Len(QPTrim$(fpCustName.Text)) > 0 Then
'      Enoughtosave = False
'      GoTo DoneChk
'    End If
'  End If
'Exit Function
'DoneChk:
'  DoEvents
'  ReDim MsgText(0 To 5) As String
'  Dim FntSize As Integer
'  If Enoughtosave = False Then
'    frmMsgDialog.RetLabel = "-2"
'    FntSize = frmMsgDialog.Label(2).FontSize
'    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
'    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
'    frmMsgDialog.Label(4).FontSize = (FntSize + 2)
'    MsgText(0) = "ERROR:"
'    MsgText(1) = ""
'    MsgText(2) = "The Name, Search Name and Status"
'    MsgText(3) = "are Required Fields."
'    MsgText(4) = ""
'    MsgText(5) = "Please Enter This Information."
'    GetOKorNot MsgText(), True
'    ChkCustInfoOK = False
'  Else
'    ChkCustInfoOK = True
'  End If
'
'End Function
'
'Private Function Chk4DupeLocation(Book$, SeqNum$)
'  Dim TBookSeq  As Long, NumBookSeq As Long
'  Dim BookSeqLen As Integer, Handle As Integer
'  Dim DupeFlag As Boolean
'  ReDim UBBookSeq(1) As BookSeqRecType
'  Chk4DupeLocation = False    'assume it's ok
'  TBookSeq = Val(Book$ + SeqNum$)
'  BookSeqLen = Len(UBBookSeq(1))
'  If FileSize(UBPath$ + "UBOOKSEQ.DAT") > 0 Then
'    Handle = FreeFile
'    Open UBPath$ + "UBOOKSEQ.DAT" For Random Shared As Handle Len = BookSeqLen
'    NumBookSeq = LOF(Handle) \ BookSeqLen
'    For CntL = 1 To NumBookSeq
'      Get Handle, CntL, UBBookSeq(1)
'      If UBBookSeq(1).BookSeq = TBookSeq& Then
'        If Not QPTrim$(fpstatus.Text) = "A" Or QPTrim$(fpstatus.Text) = "P" Then
'       ' If Not fpstatus.Text = "A" Or fpstatus.Text = "P" Then
'          If TBookSeq& <= 0 Then
'            Exit For
'          End If
'        End If
'        DupeFlag = True
'        Exit For
'      End If
'    Next
'  End If
'  Close Handle
'  If DupeFlag Then
'    Chk4DupeLocation = True
'  End If
'  'Erase UBBookSeq
'End Function
'
            

