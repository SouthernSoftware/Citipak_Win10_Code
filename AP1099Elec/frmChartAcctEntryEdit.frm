VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmChartAcctEntryEdit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chart of Accounts Entry Edit"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   555
   ClientWidth     =   12225
   Icon            =   "frmChartAcctEntryEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboFunction 
      Height          =   405
      Left            =   5640
      TabIndex        =   3
      Top             =   5235
      Width           =   3810
      _Version        =   196608
      _ExtentX        =   6720
      _ExtentY        =   714
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   3
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
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
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
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
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
      ColDesigner     =   "frmChartAcctEntryEdit.frx":08CA
   End
   Begin LpLib.fpCombo txtTyp 
      Height          =   405
      Left            =   5640
      TabIndex        =   2
      Top             =   4560
      Width           =   1560
      _Version        =   196608
      _ExtentX        =   2752
      _ExtentY        =   714
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
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
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
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
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
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
      ColDesigner     =   "frmChartAcctEntryEdit.frx":0C7D
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   6360
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7248
      Width           =   1332
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Esc E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   9480
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7248
      Width           =   1332
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F3 &Delete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   7920
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7248
      Width           =   1332
   End
   Begin EditLib.fpText txtTitle 
      Height          =   372
      Left            =   5646
      TabIndex        =   1
      Top             =   3900
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
      AutoCase        =   4
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
      CharValidationText=   ""
      MaxLength       =   30
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
      ButtonAlign     =   1
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpMask txtNum 
      Height          =   372
      Left            =   5646
      TabIndex        =   0
      Top             =   3240
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
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
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   -2147483634
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483639
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
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
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
            TextSave        =   "4:39 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "6/25/2008"
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Function Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3198
      TabIndex        =   14
      Top             =   5280
      Width           =   2292
   End
   Begin VB.Label lblEditAcct 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Edit Existing Account"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4800
      TabIndex        =   13
      Top             =   6168
      Visible         =   0   'False
      Width           =   2772
   End
   Begin VB.Label lblNewAcct 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "New Account"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5280
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   3228
      Left            =   2286
      Top             =   2760
      Width           =   7644
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3198
      TabIndex        =   11
      Top             =   4600
      Width           =   2292
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Account Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   2598
      TabIndex        =   10
      Top             =   3920
      Width           =   2892
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Account Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3438
      TabIndex        =   9
      Top             =   3240
      Width           =   2052
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      BackStyle       =   0  'Transparent
      Caption         =   "Chart of Accounts Entry/Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   492
      Left            =   3720
      TabIndex        =   7
      Top             =   1440
      Width           =   4812
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   852
      Left            =   2880
      Top             =   1200
      Width           =   6492
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      Height          =   972
      Left            =   2880
      Top             =   1080
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
Attribute VB_Name = "frmChartAcctEntryEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLAcct As GLAcctRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Public RecordNum As Integer
Private Sub cmdDelete_Click()
  If RecordNum > 0 Then
    If MsgBox("Are You Sure You Wish To Delete This Account? OK to Delete, Cancel to Abort.", vbOKCancel, "Delete Account") = vbOK Then
      Call DeleteAcct
    Else
      Exit Sub
    End If
  Else
    MsgBox "This Account Has Not Been Saved And May Not Be Deleted.", vbOKOnly, "Deletion Denied"
  End If
  txtNum.SetFocus
End Sub
Private Sub cmdExit_Click()
  frmChartAcctMaintMenu.Show
  Unload frmChartAcctEntryEdit
End Sub
Private Sub cmdSave_Click()
  If txtNum = "" Or txtTitle = "" Or txtTyp.ListIndex = -1 Then
    MsgBox "A Blank Field May Not Be Saved.", vbOKOnly, "Save Canceled"
  Else
    Call SaveAcct
    lblNewAcct.Visible = False
  End If
  txtNum.SetFocus
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
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%D"
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_Load()
  Dim AcctFile As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctMsk
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpAccountEntryEdit
  'txtTyp is combo box and only first character is stored in file
  txtTyp.AddItem "Asset"
  txtTyp.AddItem "Liability"
  txtTyp.AddItem "Revenue"
  txtTyp.AddItem "Expenditure"
  fpcboFunction.AddItem "None" & Chr$(9) & "No Function Selected" & Chr$(9) & Str(0)
  FillFNCTList fpcboFunction
  fpcboFunction.ListIndex = 0
End Sub
Private Function GetAcctMsk()
'basis for Account mask is from setup file Length setings
  Dim fundmsk As String
  Dim acctmsk As String
  Dim detmsk As String
  fundmsk = String(GLFundLen, "#")
  acctmsk = String(GLAcctLen, "#")
  detmsk = String(GLDetLen, "#")
  txtNum.Mask = (fundmsk & "-" & acctmsk & "-" & detmsk)
End Function
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        ClearInUse PWcnt
      End If
    End If
  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub txtNum_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtNum_LostFocus()
  If Len(Trim(txtNum)) > 0 Then
    If AcctSearch = False Then
      txtNum.SetFocus
    Else
      txtTitle.SetFocus
    End If
  Else
    txtTitle = ""
    txtTyp.ListIndex = -1
    fpcboFunction.ListIndex = 0
    lblNewAcct.Visible = False
    lblEditAcct.Visible = False
    cmdDelete.Enabled = False
  End If
End Sub
Private Sub GetAcct(RecordNum As Integer)
  Dim AcctFile As Integer
  OpenAcctFile AcctFile
  Get AcctFile, RecordNum, GLAcct
  txtNum = Trim(GLAcct.Num)
  txtTitle = Trim(GLAcct.Title)
  Select Case GLAcct.Typ
    Case "A"
      txtTyp.ListIndex = 0
    Case "L"
      txtTyp.ListIndex = 1
    Case "R"
      txtTyp.ListIndex = 2
    Case "E"
      txtTyp.ListIndex = 3
    Case Else
  End Select
  If GLAcct.FNCTRec > 0 Then
    fpcboFunction.ColumnSearch = 2
    fpcboFunction.SearchText = Str(GLAcct.FNCTRec)
    fpcboFunction.Action = 0
    If fpcboFunction.SearchIndex <> -1 Then
      fpcboFunction.ListIndex = fpcboFunction.SearchIndex
    End If
  Else
    fpcboFunction.ListIndex = 0
  End If
  Close AcctFile
End Sub
Private Function AcctSearch()
  Dim FoundAcct As Boolean
  FoundAcct = False 'assume we can't find it
  If Len(txtNum) <> (Val(GLFundLen + GLAcctLen + GLDetLen + 2)) Or InstrCount(txtNum, "-") <> 2 Then
    MsgBox "Invalid Account Code.", vbOKOnly, "Invalid Data!"
    GetAcctMsk
    txtTitle = ""
    txtTyp.ListIndex = -1
    lblEditAcct.Visible = False
    lblNewAcct.Visible = False
    cmdDelete.Enabled = False
    FoundAcct = False
    txtNum.SetFocus
  Else
    RecordNum = AcctFind(txtNum)
    If RecordNum > 0 Then
      FoundAcct = True
      GetAcct RecordNum
      lblEditAcct.Visible = True
      lblNewAcct.Visible = False
      cmdDelete.Enabled = True
    Else
      If FindFund(Left(txtNum, GLFundLen)) > 0 Then
        FoundAcct = True
        lblNewAcct.Visible = True
        lblEditAcct.Visible = False
        txtTitle = ""
        txtTyp.ListIndex = -1
      Else
        MsgBox "Invalid Fund Code.", vbOKOnly
      End If
    End If
  End If
  AcctSearch = FoundAcct
End Function
Private Sub SaveAcct()
  Dim AcctFile As Integer
  'Dim NumAccts As Integer
  GLAcct.Deleted = 0
  GLAcct.Num = txtNum
  GLAcct.Title = Trim(txtTitle)
  GLAcct.Typ = Mid(txtTyp.Text, 1, 1)
  fpcboFunction.Col = 2
  GLAcct.FNCTRec = fpcboFunction.ColText
  OpenAcctFile AcctFile
  If RecordNum = 0 Then
    RecordNum = (LOF(AcctFile) / Len(GLAcct)) + 1
    GLAcct.FrstTran = 0
    GLAcct.LastTran = 0
    GLAcct.PYAct = 0
    GLAcct.BegBal = 0
    GLAcct.Bgt = 0
    GLAcct.Bal = 0
    GLAcct.Encumb = 0
    GLAcct.MTD = 0
    GLAcct.YTD = 0
    GLAcct.NYEst = 0
    GLAcct.NYReq = 0
    GLAcct.NYRec = 0
    GLAcct.NYApp = 0
    GLAcct.FrstBTran = 0
    GLAcct.LastBTran = 0
    GLAcct.FrstPTran = 0
    GLAcct.LastPTran = 0
    GLAcct.Res = ""
    GLAcct.ChkByte = Chr$(1)
    GLAcct.Marked = 0
  End If
  Put AcctFile, RecordNum, GLAcct
  Close AcctFile
  Call MainLog("GLAcct: " + txtNum + " Saved.")
  Getsorted
  MsgBox "Your Information Has Been Saved.", vbOKOnly, "Account Saved"
  txtNum = ""
  txtTitle = ""
  txtTyp.ListIndex = -1
  fpcboFunction.ListIndex = 0
  lblNewAcct.Visible = False
  lblEditAcct.Visible = False
  RecordNum = 0
  txtNum.SetFocus
End Sub
Private Sub DeleteAcct()
  Dim AcctFile As Integer
  'Dim NumAccts as Integer
  If Not Exist("GJEdit.dat") Then
    OpenAcctFile AcctFile
    Get AcctFile, RecordNum, GLAcct
    If GLAcct.LastTran = 0 Then
      GLAcct.Deleted = -1
      Put AcctFile, RecordNum, GLAcct
      Call MainLog("GLAcct: " + txtNum + " Deleted.")
      txtNum = ""
      txtTitle = ""
      txtTyp.ListIndex = -1
      fpcboFunction.ListIndex = 0
      lblEditAcct.Visible = False
      txtNum.SetFocus
    Else
      MsgBox "This Account Has Transactions And May Not Be Deleted.", vbOKOnly, "Deletion Canceled"
    End If
  Close AcctFile
  Getsorted
  GLAcct.Deleted = 0
  Else
    MsgBox "Unposted Transactions Exist," & Chr(13) & "You May Not Delete Any Accounts Until All Entries Are Posted.", vbOKOnly, "Delete Not Allowed"
  End If
End Sub

Private Sub Getsorted()
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdSave.Enabled = False
  Me.cmdDelete.Enabled = False
  Me.mnuOptions.Enabled = False
  QSortAcctIndex frmChartAcctEntryEdit
  Call MainLog("GLAccts Sorted via Chart Accts Enter/Edit.")
  Me.mnuOptions.Enabled = True
  Me.cmdExit.Enabled = True
  Me.cmdSave.Enabled = True
  Me.cmdDelete.Enabled = True
  EnableCloseButton Me.hwnd, True
End Sub

Private Sub txtTitle_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub txtTyp_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    txtTyp.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    txtTyp.ListIndex = -1
    txtTyp.Action = ActionClearSearchBuffer
  End If
  If txtTyp.ListDown <> True Then
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
Private Sub fpcboFunction_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboFunction.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboFunction.ListIndex = 0
    fpcboFunction.Action = ActionClearSearchBuffer
  End If
  If fpcboFunction.ListDown <> True Then
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

