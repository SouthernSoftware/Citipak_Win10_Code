VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBankCodeEdit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Code Edit "
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   555
   ClientWidth     =   12225
   Icon            =   "frmBankCodeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo txtBankNum 
      Height          =   384
      Left            =   6120
      TabIndex        =   0
      Top             =   2784
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
      _ExtentY        =   677
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
      BackColor       =   16777215
      ForeColor       =   0
      Text            =   ""
      Columns         =   0
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
      ThreeDInsideHighlightColor=   16777215
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   12632256
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
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
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   13684944
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   14737632
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   12632256
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
      ColDesigner     =   "frmBankCodeEdit.frx":08CA
   End
   Begin EditLib.fpText txtBankAcct 
      Height          =   372
      Left            =   6120
      TabIndex        =   2
      Top             =   4200
      Width           =   2772
      _Version        =   196608
      _ExtentX        =   4890
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   16777215
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   12632256
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   14737632
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   14737632
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   25
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   12632256
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   4210752
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText txtBankName 
      Height          =   372
      Left            =   6120
      TabIndex        =   1
      Top             =   3456
      Width           =   3252
      _Version        =   196608
      _ExtentX        =   5736
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   16777215
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   8421504
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   4
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   14737632
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   14737632
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   30
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   12632256
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   1
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpMask txtGLAcct 
      Height          =   372
      Left            =   6120
      TabIndex        =   3
      Top             =   4920
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   656
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   16777215
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   16777215
      ThreeDTextShadowColor=   12632256
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
      InvalidColor    =   14737632
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   14737632
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   ""
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   -1  'True
      BorderGrayAreaColor=   12632256
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   8385
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   450
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
            TextSave        =   "1:54 PM"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   492
      Left            =   6216
      TabIndex        =   4
      Top             =   7056
      Width           =   1452
      _Version        =   131072
      _ExtentX        =   2561
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmBankCodeEdit.frx":0BD5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   492
      Left            =   7848
      TabIndex        =   11
      Top             =   7056
      Width           =   1452
      _Version        =   131072
      _ExtentX        =   2561
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmBankCodeEdit.frx":0DB5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   9480
      TabIndex        =   12
      Top             =   7056
      Width           =   1452
      _Version        =   131072
      _ExtentX        =   2561
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmBankCodeEdit.frx":0F96
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Code Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   1
      Left            =   3000
      TabIndex        =   10
      Top             =   2760
      Width           =   2772
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   3252
      Left            =   2400
      Top             =   2400
      Width           =   7452
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Account Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   2760
      TabIndex        =   9
      Top             =   4200
      Width           =   3012
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   4080
      TabIndex        =   8
      Top             =   3480
      Width           =   1692
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "G/L Account Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   0
      Left            =   3120
      TabIndex        =   7
      Top             =   4920
      Width           =   2652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Code Edit"
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
      Left            =   3720
      TabIndex        =   5
      Top             =   1200
      Width           =   4812
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   852
      Left            =   2880
      Top             =   960
      Width           =   6492
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   972
      Left            =   2880
      Top             =   840
      Width           =   6492
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
Attribute VB_Name = "frmBankCodeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLAcct As GLAcctRecType
Dim GLAcctidx As GLAcctIndexType
Dim GLBank As GLBankRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Public RecordNum As Integer
Private Sub cmdDelete_Click()
  If MsgBox("Are You Sure You Wish To Delete This Bank Code? OK to Delete, Cancel to Abort.", vbOKCancel, "Delete Bank") = vbOK Then
    Call DeleteBank
    txtBankNum.Clear
    GetBanks
    cmdDelete.Enabled = False
    txtBankNum.SetFocus
  Else
    txtBankNum.SetFocus
    Exit Sub
  End If
End Sub
Private Sub cmdExit_Click()
  frmBankMaintMenu.Show
  Unload frmBankCodeEdit
End Sub
Private Sub cmdSave_Click()
  If ValidBank = False Then
    MsgBox "These Fields May Not Be Left Blank.", vbOKOnly, "Save Canceled"
  Else
    Call SaveBank
    txtBankNum.ListIndex = -1
    txtBankName = ""
    txtBankAcct = ""
    txtGLAcct = ""
    cmdDelete.Enabled = False
  End If
  txtBankNum.SetFocus
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
      KeyCode = 0
      cmdExit_Click
    Case vbKeyF10:
      KeyCode = 0
      cmdSave_Click
    Case vbKeyF3:
      KeyCode = 0
      cmdDelete_Click
    Case Else:
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        ClearInUse PWcnt
      End If
    End If
  End If
End Sub

Private Sub Form_Load()
  Dim AcctFile As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  GetBAcctMsk
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = GLUserName
  GetBanks
  Me.HelpContextID = hlpEditABankMenu
End Sub

Private Function GetBAcctMsk()
  'Dim fundmsk As String
  Dim acctmsk As String
  Dim detmsk As String
  'fundmsk = String(GLFundLen, "#")
  acctmsk = String(GLAcctLen, "#")
  detmsk = String(GLDetLen, "#")
  txtGLAcct.Mask = ("-" & acctmsk & "-" & detmsk)
End Function
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Function ValidBank()
  If txtBankNum.ListIndex <> -1 Then
    If txtBankName <> "" Then
      If txtGLAcct <> "" Then
        ValidBank = True
        Exit Function
      End If
    End If
  End If
End Function


Private Function LoadBank()
  Dim BankFile As Integer, NumBankRecs As Integer
  Dim cnt As Integer
  OpenBankFile BankFile, NumBankRecs
  If txtBankNum.Text <> "" Then
    Get BankFile, txtBankNum.Text, GLBank
    txtBankName = Trim(GLBank.BankName)
    txtBankAcct = Trim(GLBank.BankAcct)
    txtGLAcct = Trim(GLBank.GLAcct)
    cmdDelete.Enabled = True
    'LoadBank = True
  Else
    'LoadBank = False
    txtBankName = ""
    txtBankAcct = ""
    txtGLAcct = ""
    txtBankName.SetFocus
  End If
  Close BankFile
End Function

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub txtBankNum_Click()
  LoadBank
End Sub

Private Sub txtBankNum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    txtBankNum.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    txtBankNum.ListIndex = -1
    txtBankNum.Action = ActionClearSearchBuffer
  End If
  If txtBankNum.ListDown <> True Then
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

Private Sub txtGLAcct_LostFocus()
  If Len(Trim(txtGLAcct)) > 1 Then
    If AcctSearch = False Then
      GetBAcctMsk
      MsgBox "Invalid GL Account", vbOKOnly, "Invalid Entry"
      txtGLAcct.SetFocus
    End If
  End If
End Sub

Private Function AcctSearch()
  Dim FoundAcct As Boolean
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer, AcctFileNum As Integer
  Dim cnt As Integer
  Dim Match As Boolean, LookFor As String
  FoundAcct = False 'assume we can't find it
  If Len(txtGLAcct) <> (Val(GLAcctLen + GLDetLen + 2)) Or InstrCount(txtGLAcct, "-") <> 2 Then
    'MsgBox "Invalid Account Code.", vbOKOnly, "Invalid Data!"
    'GetBAcctMsk
    FoundAcct = False
    'txtGLAcct.SetFocus
  Else
    OpenAcctIdx AcctIdxFileNum, NumAIdxRecs
    For cnt = 1 To NumAIdxRecs
      Get AcctIdxFileNum, cnt, GLAcctidx
      LookFor$ = Mid(GLAcctidx.AcctNum, (GLFundLen + 1))
      If txtGLAcct = Trim(LookFor$) Then
        Match = True
        Close AcctIdxFileNum
        Exit For
      End If
    Next
    If Match Then
      FoundAcct = True
    Else
      FoundAcct = False
    End If
  End If
  Close AcctIdxFileNum
  AcctSearch = FoundAcct
End Function
Private Sub GetBanks()
  Dim BankFile As Integer, NumBankRecs As Integer
  Dim cnt As Integer
  OpenBankFile BankFile, NumBankRecs
  For cnt = 1 To NumBankRecs
    Get BankFile, cnt, GLBank
    If GLBank.Deleted = 0 Then
      txtBankNum.AddItem GLBank.BankNum
    End If
  Next
  txtBankNum.ListIndex = -1
  txtBankName = ""
  txtBankAcct = ""
  txtGLAcct = ""
  Close BankFile
End Sub

Private Sub SaveBank()
  Dim BankFile As Integer, NumBankRecs As Integer
  GLBank.Deleted = 0
  GLBank.BankNum = txtBankNum.Text
  GLBank.BankName = Trim(txtBankName)
  GLBank.BankAcct = Trim(txtBankAcct)
  GLBank.GLAcct = Trim(txtGLAcct)
  OpenBankFile BankFile, NumBankRecs
  Put BankFile, txtBankNum.Text, GLBank
  Close BankFile
  Call MainLog("Bank: " + txtBankNum.Text + " Edited.")
  MsgBox "Your Information Has Been Saved.", vbOKOnly, "Bank Saved"
  txtBankNum.SetFocus
End Sub
Private Sub DeleteBank()
  Dim BankFile As Integer, NumBankRecs As Integer
  If txtBankNum.Text <> "" Then
    OpenBankFile BankFile, NumBankRecs
    Get BankFile, txtBankNum.Text, GLBank
    GLBank.Deleted = -1
    Put BankFile, txtBankNum.Text, GLBank
    Close BankFile
    GLBank.Deleted = 0
    Call MainLog("Bank: " + txtBankNum.Text + " Deleted.")
    MsgBox "Your Information Has Been Deleted.", vbOKOnly, "Bank Deleted"
  End If
End Sub
Private Sub txtGLAcct_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
