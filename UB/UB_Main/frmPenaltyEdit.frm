VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmPenaltyEdit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penalty Transaction Edit"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12225
   Icon            =   "frmPenaltyEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Recno 
      Height          =   348
      Left            =   672
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1392
      Visible         =   0   'False
      Width           =   1092
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   8340
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   529
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
            TextSave        =   "12:37 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "12/21/2007"
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
   Begin EditLib.fpText fpCustName 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   4656
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2856
      Width           =   5124
      _Version        =   196608
      _ExtentX        =   9038
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
      NoSpecialKeys   =   3
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
   Begin EditLib.fpText fpBook 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   4656
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3228
      Width           =   372
      _Version        =   196608
      _ExtentX        =   656
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
      AlignTextV      =   1
      AllowNull       =   -1  'True
      NoSpecialKeys   =   3
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
      MarginTop       =   0
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   1
      Text            =   ""
      CharValidationText=   "0123456789"
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
   Begin EditLib.fpText fpSeqNumb 
      CausesValidation=   0   'False
      Height          =   324
      Left            =   5244
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3228
      Width           =   996
      _Version        =   196608
      _ExtentX        =   1757
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
      AlignTextV      =   1
      AllowNull       =   -1  'True
      NoSpecialKeys   =   3
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
      MarginTop       =   0
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   1
      Text            =   ""
      CharValidationText=   "0123456789"
      MaxLength       =   6
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
   Begin EditLib.fpText txtTotRec 
      Height          =   300
      Left            =   4068
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5040
      Width           =   996
      _Version        =   196608
      _ExtentX        =   1757
      _ExtentY        =   529
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
      Text            =   ""
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
   Begin EditLib.fpText txtRec 
      Height          =   300
      Left            =   2724
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5040
      Width           =   996
      _Version        =   196608
      _ExtentX        =   1757
      _ExtentY        =   529
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
      Text            =   ""
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdSave 
      Height          =   375
      Left            =   8490
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6795
      Width           =   1410
      _Version        =   131072
      _ExtentX        =   2487
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmPenaltyEdit.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn CmdExit 
      Height          =   375
      Left            =   10035
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6795
      Width           =   1395
      _Version        =   131072
      _ExtentX        =   2461
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmPenaltyEdit.frx":0AA6
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdDelete 
      Height          =   375
      Left            =   3870
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6795
      Width           =   1395
      _Version        =   131072
      _ExtentX        =   2461
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmPenaltyEdit.frx":0C82
   End
   Begin fpBtnAtlLibCtl.fpBtn CmdPageDn 
      Height          =   375
      Left            =   2325
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6795
      Width           =   1395
      _Version        =   131072
      _ExtentX        =   2461
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmPenaltyEdit.frx":1F55
   End
   Begin fpBtnAtlLibCtl.fpBtn CmdPageUp 
      Height          =   375
      Left            =   780
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6795
      Width           =   1410
      _Version        =   131072
      _ExtentX        =   2487
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmPenaltyEdit.frx":2131
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdTranHist 
      Height          =   375
      Left            =   5400
      TabIndex        =   23
      Top             =   6795
      Width           =   1410
      _Version        =   131072
      _ExtentX        =   2487
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmPenaltyEdit.frx":230B
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdList 
      Height          =   375
      Left            =   6945
      TabIndex        =   24
      Top             =   6795
      Width           =   1410
      _Version        =   131072
      _ExtentX        =   2487
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmPenaltyEdit.frx":24E8
   End
   Begin EditLib.fpDoubleSingle fpAmount 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   324
      Left            =   4656
      TabIndex        =   0
      Top             =   4248
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
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
      ButtonDefaultAction=   0   'False
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   -1  'True
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   0   'False
      InvalidColor    =   -2147483637
      InvalidOption   =   2
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   3
      OnFocusAlignV   =   1
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2412
      Left            =   1848
      Top             =   2568
      Width           =   8532
   End
   Begin VB.Label LabelDel 
      BackStyle       =   0  'Transparent
      Caption         =   "Deleted !"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   276
      Left            =   6312
      TabIndex        =   22
      Top             =   4272
      Visible         =   0   'False
      Width           =   1308
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
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
      Height          =   324
      Left            =   3372
      TabIndex        =   21
      Top             =   4272
      Width           =   1140
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   1860
      TabIndex        =   15
      Top             =   5040
      Width           =   828
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "of"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   3660
      TabIndex        =   14
      Top             =   5040
      Width           =   300
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Records"
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
      Left            =   5124
      TabIndex        =   13
      Top             =   5040
      Width           =   996
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name:"
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
      Height          =   324
      Left            =   2580
      TabIndex        =   11
      Top             =   2880
      Width           =   1956
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Location #:"
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
      Height          =   324
      Left            =   3132
      TabIndex        =   10
      Top             =   3252
      Width           =   1404
   End
   Begin VB.Label Label9 
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
      ForeColor       =   &H8000000E&
      Height          =   324
      Left            =   3156
      TabIndex        =   9
      Top             =   3624
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   5076
      TabIndex        =   8
      Top             =   3240
      Width           =   132
   End
   Begin VB.Label LabelAcctNo 
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
      Left            =   4656
      TabIndex        =   7
      Top             =   3612
      Width           =   1140
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   804
      Left            =   2856
      Top             =   984
      Width           =   6492
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty Transaction Edit"
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
      Left            =   3708
      TabIndex        =   3
      Top             =   1224
      Width           =   4812
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   924
      Left            =   2868
      Top             =   888
      Width           =   6492
   End
   Begin VB.Label LabelRec 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Caption         =   "Last Record Displayed."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1836
      TabIndex        =   16
      Top             =   4992
      Visible         =   0   'False
      Width           =   8556
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
Attribute VB_Name = "frmPenaltyEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim PageUp As Boolean, PageDn As Boolean
Dim TransRec As Long, cnt As Long
Dim Answer As Integer '1 for remain on screen,2 for save, 3 for nosave
Private Sub cmdExit_Click()
    Chk4Change
    If Answer = 1 Then
      Exit Sub
    ElseIf Answer = 2 Then
      SavePenTrans (cnt&)
    End If
  
  Load frmUBPenaltyMenu
  DoEvents
  frmUBPenaltyMenu.Show
  Unload frmPenaltyEdit
  DoEvents
End Sub

Private Sub fpAmount_ChangeMode(EditMode As Integer)
  EditMode = 1
  
End Sub


Private Sub fpAmount_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    KeyCode = 0
    DoEvents
    fpCmdSave.SetFocus
  End If
  If KeyCode = vbKeyDelete Then
    fpAmount = 0
    SendKeys "+{Tab}"
    SendKeys "{Tab}"
  End If
End Sub

Private Sub fpCmdDelete_Click()
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
    'frmMsgDialog.RetLabel = "-1"
    FntSize = frmMsgDialog.Label(2).FontSize
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "LAST CHANCE!"
    MsgText(1) = ""
    MsgText(2) = ""
    MsgText(3) = "Are You Sure You Want To Delete"
    MsgText(4) = "This Penalty Transaction?"
    MsgText(5) = ""
    If GetOKorNot(MsgText()) Then
      MarkTransDeleted cnt&
      LabelDel.Visible = True
      fpAmount = 0
    End If
End Sub

Private Sub fpCmdList_Click()
  Chk4Change
  If Answer = 1 Then
    Exit Sub
  ElseIf Answer = 2 Then
    SavePenTrans (cnt&)
  End If
  frmPenaltyList.lstPenalties.ListIndex = txtRec - 1
  frmPenaltyList.Show 1, frmPenaltyEdit
End Sub

Private Sub fpCmdTranHist_Click()
  ReDim MsgText(0 To 5) As String
  Dim FntSize As Integer
  If TransRec > 0 Then
   ' DeActivateControls Me
    DisplayCustTransList RecNo
   ' ActivateControls Me
  Else
    frmMsgDialog.RetLabel = "-2"
    FntSize = frmMsgDialog.Label(2).FontSize
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    MsgText(0) = "ERROR:"
    MsgText(1) = ""
    MsgText(2) = ""
    MsgText(3) = "There are NO transactions to display."
    MsgText(4) = ""
    MsgText(5) = ""
    GetOKorNot MsgText(), True
  End If
End Sub


Private Sub fpCmdSave_Click()
   SavePenTrans (cnt&)
    cnt = cnt + 1
    If cnt > txtTotRec Then
      cnt = txtRec
      Exit Sub
    End If
    If Not (cnt) = 0 Then
      PenaltyRec2Screen cnt
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
    Chk4Change
    If Answer = 1 Then
      Cancel = True
    ElseIf Answer = 2 Then
      SavePenTrans (cnt&)
    End If
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via PenaltyEdit by " + PWUser$
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
      Call cmdExit_Click
    Case vbKeyF3:
      KeyCode = 0
      DoEvents
      Call fpCmdDelete_Click
    Case vbKeyF4:
      KeyCode = 0
      DoEvents
      Call fpCmdTranHist_Click
    Case vbKeyF5:
      KeyCode = 0
      DoEvents
      Call fpCmdList_Click
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      Call fpCmdSave_Click
    Case vbKeyPageDown:
      KeyCode = 0
      'SendKeys "+{Tab}"
      DoEvents
      Call cmdPageDn_Click
      DoEvents
    Case vbKeyPageUp:
      KeyCode = 0
      'SendKeys "+{Tab}"
      DoEvents
      Call cmdPageUp_Click
      DoEvents
    Case Else:
  End Select
End Sub
Private Sub cmdPageDn_Click()
 If cnt < txtTotRec And Not cnt > txtTotRec Then
    Chk4Change
    If Answer = 1 Then
      Exit Sub
    ElseIf Answer = 2 Then
      SavePenTrans (cnt&)
    End If
    cnt = cnt + 1
    If cnt > txtTotRec Then
      cnt = txtRec
      Exit Sub
    End If
    If Not (cnt) < 1 And Not (cnt&) > txtTotRec Then
      PenaltyRec2Screen cnt
'      fpAmount.OnFocusNoSelect = False
      fpAmount.SetFocus
    End If
 End If
End Sub

Private Sub cmdPageUp_Click()
  If cnt > 1 And Not cnt > txtTotRec Then
    Chk4Change
    If Answer = 1 Then
      Exit Sub
    ElseIf Answer = 2 Then
      SavePenTrans (cnt&)
    End If
      cnt = cnt - 1
      If cnt < 1 Or cnt > txtTotRec Then
        cnt = txtRec
        Exit Sub
      End If
      If Not (cnt&) < 1 And Not (cnt&) > txtTotRec Then
        PenaltyRec2Screen cnt
'        fpAmount.OnFocusNoSelect = False
        fpAmount.SetFocus
      End If
 End If
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  UBLog " IN: Edit Penalty File (EPF)"
  cnt = 1
  PenaltyRec2Screen cnt
  Me.HelpContextID = hlpEditPenalty
End Sub

Private Sub Form_Resize()
  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
  End If
End Sub
Public Sub PenaltyRec2Screen(num As Long)
  Dim PenFile As String, UBCustRecLen As Integer, UBTranRecLen As Integer
  Dim NumPTran As Long, PHandle As Integer, CHandle As Integer
  Dim DelFlag As Boolean
'  FrmShowPctComp.Label1 = "Sorting Penalty Records"
'  FrmShowPctComp.Show , Me

'  ReDim PenaltyInfo(1) As PenaltyInfoType
'  'FGetAH "UBPENINF.DAT", PenaltyInfo(1), Len(PenaltyInfo(1)), 1
'  hand2 = FreeFile
'  Open UBPath$ + "UBPENINF.DAT" For Random As hand2
'  Get hand2, 1, PenaltyInfo(1)
'  Close hand2
  cnt = num
  PenFile$ = UBPath$ + "UBPENTRN.DAT"

  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBTranRec(1) As UBTransRecType

  UBCustRecLen = Len(UBCustRec(1))
  UBTranRecLen = Len(UBTranRec(1))

  NumPTran& = FileSize&(PenFile$) / UBTranRecLen

'  EditedFlag = False
  DelFlag = False
'  ShowFlag = True
  PHandle = FreeFile
  Open PenFile$ For Random Shared As PHandle Len = UBTranRecLen
  CHandle = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = UBCustRecLen

  Get PHandle, num&, UBTranRec(1)
  If Not UBTranRec(1).CustAcctNo > 0 Then
    Close
    Exit Sub
  End If
  Get CHandle, UBTranRec(1).CustAcctNo, UBCustRec(1)
  Close
  If UBTranRec(1).ActiveFlag = 0 Then
    DelFlag = True
  End If
  RecNo = UBTranRec(1).CustAcctNo
  LabelAcctNo = Using$("#####", UBTranRec(1).CustAcctNo)
  fpBook = UBCustRec(1).Book
  fpSeqNumb = UBCustRec(1).SEQNUMB
  fpCustName = UBCustRec(1).CustName
  fpAmount = UBTranRec(1).Transamt 'Using$("######.##", )
  'fpAmount.SetFocus
  'fpAmount.OnFocusNoSelect = False
  If DelFlag Then
    LabelDel.Visible = True
  Else
    LabelDel.Visible = False
  End If
  If UBCustRec(1).LastTrans > 0 Then
    TransRec = UBCustRec(1).LastTrans
  End If
  txtRec = num
  txtTotRec = NumPTran&
  If num = 1 Then
    LabelRec.Caption = "First Record Displayed."
    LabelRec.Visible = True
  ElseIf num = txtTotRec Then
    LabelRec.Caption = "Last Record Displayed."
    LabelRec.Visible = True
  Else
    LabelRec.Visible = False
  End If
  Me.Refresh
  DoEvents
End Sub
Private Sub SavePenTrans(Rec&)
  Dim PenFile As String, UBTranRecLen As Integer
  Dim NumPTran As Long, PHandle As Integer, CHandle As Integer
  Dim DelFlag As Boolean, hand2 As Integer, TPenAmt As Double
  Dim EditLog As String
  If Rec& > 0 And Not Rec& > txtTotRec Then
    ReDim PenaltyInfo(1) As PenaltyInfoType
    hand2 = FreeFile
    Open UBPath$ + "UBPENINF.DAT" For Random As hand2
    Get hand2, 1, PenaltyInfo(1)
    Close hand2
    PenFile$ = UBPath$ + "UBPENTRN.DAT"
    ReDim UBTranRec(1) As UBTransRecType
    UBTranRecLen = Len(UBTranRec(1))
  
    NumPTran& = FileSize&(PenFile$) / UBTranRecLen
    
    PHandle = FreeFile
    Open PenFile$ For Random Shared As PHandle Len = UBTranRecLen
    Get PHandle, Rec&, UBTranRec(1)
    TPenAmt# = fpAmount
  
    If TPenAmt# = 0 Then
      UBTranRec(1).ActiveFlag = False
    Else
      UBTranRec(1).ActiveFlag = True
    End If
    If Round#(TPenAmt#) <> Round#(UBTranRec(1).Transamt) Then
      EditLog$ = Str$(UBTranRec(1).CustAcctNo) + "   was " + Using$("#####.##", UBTranRec(1).Transamt)
      EditLog$ = EditLog$ + " to " + Using$("#####.##", TPenAmt#)
      UBLog " EPF: Changed Acct:" + EditLog$
    End If
    UBTranRec(1).RevAmt(PenaltyInfo(1).RevSource) = TPenAmt#
    UBTranRec(1).Transamt = TPenAmt#
  
    Put PHandle, Rec&, UBTranRec(1)
    Close
    MsgBox "Data Updated", vbOKOnly, "Saved"
  Else
    MsgBox "Invalid Record", vbOKOnly
  End If
End Sub
Private Sub MarkTransDeleted(Rec&)
  Dim PenFile As String, UBTranRecLen As Integer
  Dim NumPTran As Long, PHandle As Integer, CHandle As Integer
  Dim DelFlag As Boolean, hand2 As Integer, TPenAmt As Double
  If Rec& > 0 And Not Rec& > txtTotRec Then
    ReDim PenaltyInfo(1) As PenaltyInfoType
    hand2 = FreeFile
    Open UBPath$ + "UBPENINF.DAT" For Random As hand2
    Get hand2, 1, PenaltyInfo(1)
    Close hand2
    PenFile$ = UBPath$ + "UBPENTRN.DAT"
    ReDim UBTranRec(1) As UBTransRecType
    UBTranRecLen = Len(UBTranRec(1))
  
    NumPTran& = FileSize&(PenFile$) / UBTranRecLen
  
    PHandle = FreeFile
    Open PenFile$ For Random Shared As PHandle Len = UBTranRecLen
    Get PHandle, Rec&, UBTranRec(1)
    UBTranRec(1).ActiveFlag = 0
    UBTranRec(1).RevAmt(PenaltyInfo(1).RevSource) = 0
    UBTranRec(1).Transamt = 0
    Put PHandle, Rec&, UBTranRec(1)
    Close
    MsgBox "Data Updated", vbOKOnly, "Deleted"
  Else
    MsgBox "Invalid Record", vbOKOnly
  End If
End Sub
Private Function Chk4Change()
  Dim PenFile As String, UBTranRecLen As Integer, DelFlag As Boolean
  Dim NumPTran As Long, PHandle As Integer, Changed As Boolean
  PenFile$ = UBPath$ + "UBPENTRN.DAT"
  ReDim UBTranRec(1) As UBTransRecType
  UBTranRecLen = Len(UBTranRec(1))
  Answer = 0
  NumPTran& = FileSize&(PenFile$) / UBTranRecLen
  DelFlag = False
  PHandle = FreeFile
  Open PenFile$ For Random Shared As PHandle Len = UBTranRecLen
  Get PHandle, txtRec, UBTranRec(1)
  Close
  If UBTranRec(1).ActiveFlag = 0 Then
    DelFlag = True
  End If
  If UBTranRec(1).Transamt <> fpAmount.Value Then
    Changed = True
  '' MsgBox "tamt = " & UBTranRec(1).Transamt & " and disp = " & fpAmount.Value, vbOKOnly
  Else
    Changed = False
  End If
  If Changed Then
    frmChangedWarning.Show vbModal, Me
    Select Case SaveFlag
    Case False
      Answer = 3
    Case True
      Answer = 2
    Case 1
      Answer = 1
    End Select
  Else
    Answer = 0
  End If
  DoEvents
End Function

