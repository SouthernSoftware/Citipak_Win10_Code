VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmFASortBy 
   BackColor       =   &H008F8265&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin EditLib.fpText fptxtSortType 
      Height          =   345
      Left            =   720
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2265
      Visible         =   0   'False
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   609
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
   Begin fpBtnAtlLibCtl.fpBtn cmdDept 
      Height          =   516
      Left            =   1028
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to begin the search using the paramenters entered above."
      Top             =   1500
      Width           =   2145
      _Version        =   131072
      _ExtentX        =   3784
      _ExtentY        =   910
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFASortBy.frx":0000
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdFundAsset 
      Height          =   516
      Left            =   3428
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to begin the search using the paramenters entered above."
      Top             =   1500
      Width           =   2145
      _Version        =   131072
      _ExtentX        =   3784
      _ExtentY        =   910
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFASortBy.frx":0213
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   372
      Left            =   4494
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to begin the search using the paramenters entered above."
      Top             =   2220
      Width           =   1860
      _Version        =   131072
      _ExtentX        =   3281
      _ExtentY        =   656
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmFASortBy.frx":042C
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   2790
      Left            =   60
      Top             =   60
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Sort Type  - By Dept Or By Fund/Asset  "
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
      Height          =   570
      Left            =   1005
      TabIndex        =   1
      Top             =   480
      Width           =   4590
   End
End
Attribute VB_Name = "frmFASortBy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsFATextBoxOverRider
'this form grabs the print out type chosen
'by the user and stays open, but hidden, until
'the calling form can read and process it...it then
'unloads
Private Sub cmdExit_Click()
  fptxtSortType.Text = "Exit"
  Me.Hide
End Sub

Private Sub cmdFundAsset_Click()
  fptxtSortType.Text = "Fund/Asset"
  Me.Hide
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
      SendKeys "X"
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF6:
      SendKeys "%D"
      Call cmdDept_Click
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%F"
      Call cmdFundAsset_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub cmdDept_Click()
  fptxtSortType.Text = "Dept"
  Me.Hide
End Sub

Private Sub Form_Load()
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
End Sub






