VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmUBSysDraftEdit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ACH Draft Information"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmUBSysDraftEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   635
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
            TextSave        =   "10:32 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "4/21/2005"
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   9120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7656
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
      _ExtentY        =   847
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
      DrawFocusRect   =   1
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
      ButtonDesigner  =   "frmUBSysDraftEdit.frx":030A
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdSave 
      Height          =   480
      Left            =   7536
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7656
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
      _ExtentY        =   847
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
      DrawFocusRect   =   1
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
      ButtonDesigner  =   "frmUBSysDraftEdit.frx":04E5
   End
   Begin fpBtnAtlLibCtl.fpBtn btnPageInfo 
      Height          =   492
      Left            =   2160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   3852
      _Version        =   131072
      _ExtentX        =   6794
      _ExtentY        =   868
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
      Static          =   -1  'True
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
      ButtonDesigner  =   "frmUBSysDraftEdit.frx":06C0
   End
   Begin EditLib.fpText fpBankDest 
      CausesValidation=   0   'False
      Height          =   348
      Left            =   6096
      TabIndex        =   4
      Top             =   2808
      Width           =   1380
      _Version        =   196608
      _ExtentX        =   2434
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "123456789"
      CharValidationText=   "0123456789"
      MaxLength       =   9
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
   Begin EditLib.fpText fpBankOrig 
      CausesValidation=   0   'False
      Height          =   348
      Left            =   6096
      TabIndex        =   5
      Top             =   3336
      Width           =   1380
      _Version        =   196608
      _ExtentX        =   2434
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "123456789"
      CharValidationText=   "0123456789"
      MaxLength       =   9
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
   Begin EditLib.fpText fpBankName 
      CausesValidation=   0   'False
      Height          =   348
      Left            =   6096
      TabIndex        =   6
      Top             =   3840
      Width           =   3036
      _Version        =   196608
      _ExtentX        =   5355
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      Text            =   "12345678901234567890123"
      CharValidationText=   ""
      MaxLength       =   23
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
   Begin EditLib.fpText fpBankLoc 
      CausesValidation=   0   'False
      Height          =   348
      Left            =   6096
      TabIndex        =   7
      Top             =   4320
      Width           =   3036
      _Version        =   196608
      _ExtentX        =   5355
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      Text            =   "12345678901234567890123"
      CharValidationText=   ""
      MaxLength       =   23
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
   Begin EditLib.fpText fpCompAcct 
      CausesValidation=   0   'False
      Height          =   348
      Left            =   6096
      TabIndex        =   8
      Top             =   4800
      Width           =   2652
      _Version        =   196608
      _ExtentX        =   4678
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "12345678901234567890"
      CharValidationText=   "1234567890"
      MaxLength       =   20
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
   Begin EditLib.fpText fpFedID 
      CausesValidation=   0   'False
      Height          =   348
      Left            =   6096
      TabIndex        =   9
      Top             =   5304
      Width           =   1380
      _Version        =   196608
      _ExtentX        =   2434
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "123456789"
      CharValidationText=   "0123456789"
      MaxLength       =   9
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
   Begin EditLib.fpText fpFedPreFix 
      CausesValidation=   0   'False
      Height          =   348
      Left            =   6096
      TabIndex        =   10
      Top             =   5784
      Width           =   396
      _Version        =   196608
      _ExtentX        =   698
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "1"
      CharValidationText=   "0123456789"
      MaxLength       =   1
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
   Begin EditLib.fpText fpFileName 
      CausesValidation=   0   'False
      Height          =   348
      Left            =   6096
      TabIndex        =   11
      Top             =   6264
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
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
      Text            =   "12345678.123"
      CharValidationText=   ""
      MaxLength       =   12
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
   Begin VB.Line Line1 
      X1              =   1944
      X2              =   10260
      Y1              =   2352
      Y2              =   2364
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "*** ALL FIELDS ARE REQUIRED ***"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   2196
      TabIndex        =   21
      Top             =   1848
      Width           =   7836
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Company Information Required for ACH Draft Transmission"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2196
      TabIndex        =   20
      Top             =   1368
      Width           =   7836
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Bank Draft File Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2568
      TabIndex        =   19
      Top             =   6264
      Width           =   3324
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Federal ID Bank Prefix Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2568
      TabIndex        =   18
      Top             =   5784
      Width           =   3324
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Federal ID Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2568
      TabIndex        =   17
      Top             =   5304
      Width           =   3324
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Bank Account Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2568
      TabIndex        =   16
      Top             =   4800
      Width           =   3324
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Originating Bank Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2568
      TabIndex        =   15
      Top             =   4320
      Width           =   3324
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Destination Bank Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2568
      TabIndex        =   14
      Top             =   3840
      Width           =   3324
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Immediate Origin Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2568
      TabIndex        =   13
      Top             =   3336
      Width           =   3324
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Immediate Destination Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2568
      TabIndex        =   12
      Top             =   2808
      Width           =   3324
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      FillColor       =   &H80000004&
      FillStyle       =   0  'Solid
      Height          =   6732
      Left            =   1824
      Top             =   480
      Width           =   8556
   End
End
Attribute VB_Name = "frmUBSysDraftEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim UBSysDraftRecLen As Integer, UBFile As Integer
Dim BeenDone As Boolean
Dim MsgText(0 To 5) As String

Private Sub Form_Load()
  Dim UBSysDraftRec As UBDraftRecType
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
''  If Exist(UBPath + "UBSDRAFT.DAT") Then
''    UBSysDraftRecLen = Len(UBSysDraftRec)
''    UBFile = FreeFile
''    Open UBPath + "UBSDRAFT.DAT" For Random Shared As UBFile Len = UBSysDraftRecLen
''    Get UBFile, 1, UBSysDraftRec
''    Close
''    fpBankDest = QPTrim$(UBSysDraftRec.BANKDEST)
''    fpBankOrig = QPTrim$(UBSysDraftRec.BANKORIG)
''    fpBankName = QPTrim$(UBSysDraftRec.BankName)
''    fpBankLoc = QPTrim$(UBSysDraftRec.BANKLOC)
''    fpCompAcct = QPTrim$(UBSysDraftRec.COMPACCT)
''    fpFedID = QPTrim$(UBSysDraftRec.FEDID)
''    fpFedPreFix = QPTrim$(UBSysDraftRec.FEDPREFX)
''    fpFileName = QPTrim$(UBSysDraftRec.FileName)
''  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If fpCmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via SysDraftEdit by " + PWUser$
        CitiTerminate
      End If
    End If
  End If

End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call fpCmdExit_Click
    Case vbKeyF10
      KeyCode = 0
      DoEvents
      Call fpCmdSave_Click
      Call fpCmdExit_Click
    Case Else:
  End Select
End Sub

Private Sub fpCmdExit_Click()
    Select Case CheckSaveDraftFile
      Case True:  '-1 save chenges
        Call fpCmdSave_Click
        DoEvents
          BeenDone = False
          Load frmUBSetupMenu
          frmUBSetupMenu.Show
          Unload frmUBSysDraftEdit
      Case False:  '0= exit
        DoEvents
          BeenDone = False
          Load frmUBSetupMenu
          frmUBSetupMenu.Show
          Unload frmUBSysDraftEdit
      Case Else     '1 is review
        'continue editing
      End Select
End Sub

Private Sub Form_Activate()
  If Not BeenDone Then 'if we haven't done this yet
    BeenDone = True
    Dim UBSysDraftRec As UBDraftRecType
    UBSysDraftRecLen = Len(UBSysDraftRec)
    UBFile = FreeFile
    Open UBPath + "UBSDRAFT.DAT" For Random Shared As UBFile Len = UBSysDraftRecLen
    Get UBFile, 1, UBSysDraftRec
    Close
    fpBankDest = QPTrim$(UBSysDraftRec.BANKDEST)
    fpBankOrig = QPTrim$(UBSysDraftRec.BANKORIG)
    fpBankName = QPTrim$(UBSysDraftRec.BankName)
    fpBankLoc = QPTrim$(UBSysDraftRec.BANKLOC)
    fpCompAcct = QPTrim$(UBSysDraftRec.COMPACCT)
    fpFedID = QPTrim$(UBSysDraftRec.FEDID)
    fpFedPreFix = QPTrim$(UBSysDraftRec.FEDPREFX)
    fpFileName = QPTrim$(UBSysDraftRec.FileName)
  End If
End Sub

Private Sub fpCmdSave_Click()
  Dim UBSysDraftRec As UBDraftRecType
  UBSysDraftRecLen = Len(UBSysDraftRec)
  If oktosave Then
    UBFile = FreeFile
    Open UBPath + "UBSDRAFT.DAT" For Random Shared As UBFile Len = UBSysDraftRecLen
    Get UBFile, 1, UBSysDraftRec
    UBSysDraftRec.BANKDEST = QPTrim$(fpBankDest.Text)
    UBSysDraftRec.BANKORIG = QPTrim$(fpBankOrig.Text)
    UBSysDraftRec.BankName = QPTrim$(fpBankName.Text)
    UBSysDraftRec.BANKLOC = QPTrim$(fpBankLoc.Text)
    UBSysDraftRec.COMPACCT = QPTrim$(fpCompAcct.Text)
    UBSysDraftRec.FEDID = QPTrim$(fpFedID.Text)
    UBSysDraftRec.FEDPREFX = QPTrim$(fpFedPreFix.Text)
    UBSysDraftRec.FileName = QPTrim$(fpFileName.Text)
    Put UBFile, 1, UBSysDraftRec
    Close UBFile
    frmDataUpdated.Show vbModal   'Display Updated msg.
    DoEvents
  End If
End Sub
Private Function oktosave()
  oktosave = True
  If Len(QPTrim$(fpBankDest.Text)) <= 0 Then
    oktosave = False
    GoTo ExitCheck
  End If
  If Len(QPTrim$(fpBankOrig.Text)) <= 0 Then
    oktosave = False
    GoTo ExitCheck
  End If
  If Len(QPTrim$(fpBankName.Text)) <= 0 Then
    oktosave = False
    GoTo ExitCheck
  End If
  If Len(QPTrim$(fpBankLoc.Text)) <= 0 Then
    oktosave = False
    GoTo ExitCheck
  End If
  If Len(QPTrim$(fpCompAcct.Text)) <= 0 Then
    oktosave = False
    GoTo ExitCheck
  End If
  If Len(QPTrim$(fpFedID.Text)) <= 0 Then
    oktosave = False
    GoTo ExitCheck
  End If
  If Len(QPTrim$(fpFedPreFix.Text)) <= 0 Then
    oktosave = False
    GoTo ExitCheck
  End If
'  If Len(QPTrim$(fpFileName.Text)) <= 0 Then
'    oktosave = False
'    GoTo ExitCheck
'  End If

ExitCheck:
  If oktosave = False Then
    Call BlankInfoError
  End If
End Function
Private Function CheckSaveDraftFile%()

  Dim UBSysDraftRec As UBDraftRecType
  Dim UBFile As Integer
  Dim Changed As Boolean
  Dim UBSysDraftRecLen  As Integer
'  Dim TText As String
  UBSysDraftRecLen = Len(UBSysDraftRec)
  Changed = False
  UBFile = FreeFile
  Open UBPath + "UBSDRAFT.DAT" For Random Shared As UBFile Len = UBSysDraftRecLen
  Get UBFile, 1, UBSysDraftRec
  Close UBFile

  If QPTrim$(UBSysDraftRec.BANKDEST) <> QPTrim$(fpBankDest.Text) Then
    Changed = True
    GoTo ExitCheck
  End If

  If QPTrim$(UBSysDraftRec.BANKORIG) <> QPTrim$(fpBankOrig.Text) Then
    Changed = True
    GoTo ExitCheck
  End If
  
  If QPTrim$(UBSysDraftRec.BankName) <> QPTrim$(fpBankName.Text) Then
    Changed = True
    GoTo ExitCheck
  End If

  If QPTrim$(UBSysDraftRec.BANKLOC) <> QPTrim$(fpBankLoc.Text) Then
    Changed = True
    GoTo ExitCheck
  End If
  If QPTrim$(UBSysDraftRec.COMPACCT) <> QPTrim$(fpCompAcct.Text) Then
    Changed = True
    GoTo ExitCheck
  End If
  If QPTrim$(UBSysDraftRec.FEDID) <> QPTrim$(fpFedID.Text) Then
    Changed = True
    GoTo ExitCheck
  End If
  If QPTrim$(UBSysDraftRec.FEDPREFX) <> QPTrim$(fpFedPreFix.Text) Then
    Changed = True
    GoTo ExitCheck
  End If
  
  If QPTrim$(UBSysDraftRec.FileName) <> QPTrim$(fpFileName.Text) Then
    Changed = True
  End If

ExitCheck:
  If Changed Then
    Load frmChangedWarning
    frmChangedWarning.Show vbModal, Me
    Select Case SaveFlag
    Case False
      CheckSaveDraftFile = False
    Case True
      CheckSaveDraftFile = True
    Case 1
      CheckSaveDraftFile = 1
    End Select
  Else
    CheckSaveDraftFile = False
  End If
End Function
Private Sub BlankInfoError()
  MsgText(0) = "Blank Setup Information"
  MsgText(1) = "ERROR:"
  MsgText(2) = "REQUIRED FIELDS CAN NOT BE BLANK!"
  MsgText(3) = ""
  MsgText(4) = ""
  MsgText(5) = "Enter informtion for Every Field."
  GetOKorNot MsgText(), True
  'fpRateCode.SetFocus
End Sub

Sub DraftStuff()
'    BANKDEST As String * 9
'    BANKORIG As String * 9
'    BANKNAME As String * 23
'    BANKLOC  As String * 23
'    COMPACCT As String * 20
'    FEDID    As String * 9
'    FEDPREFX As String * 1
'    FileName As String * 12

'+-[ ACH Draft Information ]--------------------------------------------+
'¦        Company Information Required for ACH Draft Transmission       ¦
'¦                       ** ALL FIELDS REQUIRED **                      ¦
'¦                                                                      ¦
'¦  Immediate Destination Number: 053100300                             ¦
'¦       Immediate Origin Number: 053100300                             ¦
'¦         Destination Bank Name: FIRST CITIZENS                        ¦
'¦         Originating Bank Name: FIRST CITIZENS                        ¦
'¦   Company Bank Account Number: 003242489815                          ¦
'¦     Company Federal ID Number: 566000311                             ¦
'¦ Federal ID Bank Prefix Number: 1                                     ¦
'¦          Bank Draft File Name: BANKDRFT.ACH                          ¦
'+----------------------------------------------------------------------¦
'¦                                                                      ¦
'¦    Example: For BB&T                                                 ¦
'¦    Immediate Destination Number: 053101121                           ¦
'¦         Immediate Origin Number: 053101121                           ¦
'¦           Destination Bank Name: BRANCH BANKING & TRUST              ¦
'¦           Originating Bank Name: BRANCH BANKING & TRUST              ¦
'¦                                                                      ¦
'¦                                    _  F10=Save  ¦_  Esc=Cancel  ¦    ¦
'+----------------------------------------------------------------------+

End Sub
