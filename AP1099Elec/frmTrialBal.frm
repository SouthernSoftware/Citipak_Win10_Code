VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmPrnTrialBal 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trial Balance"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmTrialBal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboFund2 
      Height          =   384
      Left            =   5904
      TabIndex        =   2
      Top             =   4608
      Width           =   3084
      _Version        =   196608
      _ExtentX        =   5440
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   2
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   2
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
      AutoSearchFill  =   -1  'True
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
      ColDesigner     =   "frmTrialBal.frx":08CA
   End
   Begin LpLib.fpCombo fpcboFund1 
      Height          =   384
      Left            =   5904
      TabIndex        =   1
      Top             =   3840
      Width           =   3084
      _Version        =   196608
      _ExtentX        =   5440
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
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   2
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   2
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
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   100
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmTrialBal.frx":0CED
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   384
      Left            =   5904
      TabIndex        =   3
      Top             =   5400
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   1
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
      ScrollBarH      =   3
      DataFieldList   =   ""
      ColumnEdit      =   0
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   3504
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
      ColDesigner     =   "frmTrialBal.frx":1110
   End
   Begin VB.CheckBox Chk0Bal 
      BackColor       =   &H008F8265&
      Caption         =   "Exclude Zero Balance Accounts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   636
      Left            =   5880
      TabIndex        =   4
      Top             =   5904
      Width           =   3732
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Print"
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
      Left            =   8400
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7488
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
      Left            =   10080
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7488
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   8508
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7133
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "12:30 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "10/15/2004"
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
   Begin EditLib.fpDateTime txtDate 
      Height          =   372
      Left            =   5904
      TabIndex        =   0
      Top             =   3072
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
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
      ButtonColor     =   14737632
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Type: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   3192
      TabIndex        =   12
      Top             =   5448
      Width           =   2388
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Fund:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   324
      Index           =   1
      Left            =   3888
      TabIndex        =   11
      Top             =   4692
      Width           =   1596
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   3972
      Left            =   2736
      Top             =   2736
      Width           =   7116
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Trial Balance Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4272
      TabIndex        =   10
      Top             =   1368
      Width           =   3852
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3240
      Top             =   1128
      Width           =   5772
   End
   Begin VB.Image Image1 
      Height          =   276
      Left            =   3168
      Picture         =   "frmTrialBal.frx":14E6
      Top             =   3024
      Width           =   288
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   0
      Left            =   3960
      TabIndex        =   9
      Top             =   3108
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Fund:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   1
      Left            =   3840
      TabIndex        =   8
      Top             =   3900
      Width           =   1668
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3240
      Top             =   1008
      Width           =   5772
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
Attribute VB_Name = "frmPrnTrialBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLAcct    As GLAcctRecType
Dim GLFundIdx As GLFundIndexType
Dim GLAcctidx As GLAcctIndexType
Dim GLTrans   As GLTransRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim FY1BegDate As Integer, FY1EndDate As Integer, FY2BegDate As Integer, FY2EndDate As Integer
Dim FirstFund As String, LastFund As String
Dim ActiveYear As Integer

Private Sub cmdExit_Click()
  frmGLReportsMenu.Show
  Unload frmPrnTrialBal
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

Private Function ValidDate()
  Dim TempDate As Integer
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  If CheckValDate(txtDate) = True Then
    TempDate = DateDiff("d", "12/31/1979", txtDate)
    ValidDate = True
    If TempDate >= FY2BegDate Then
      ActiveYear = 2
      
    Else
      ActiveYear = 1
    End If
  Else
    MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    ValidDate = False
    Exit Function
  
  End If
End Function
Private Function ValidFunds()
  If fpcboFund1.Text <> "" And fpcboFund2.Text <> "" Then
    fpcboFund1.Col = 0
    fpcboFund2.Col = 0
    If fpcboFund1.ColText > fpcboFund2.ColText Then
      MsgBox "Invalid Fund Selection, The Beginning Fund Should Be Less or Equal to Ending Fund.", vbOKOnly, "Invalid Selection"
      ValidFunds = False
    Else
      ValidFunds = True
      FirstFund = fpcboFund1.ColText
      LastFund = fpcboFund2.ColText
    End If
  Else
    MsgBox "Fund Selections May Not Be Left Blank.", vbOKOnly, "Invalid Selection"
  End If
End Function
'Private Sub cmdDisplay_Click()
'fpcboFund.Col = 1
'If QPTrim(fpcboFund.ColText) = "" Then
'  Fund = "0"
'Else
'  fpcboFund.Col = 0
'  Fund = fpcboFund.ColText
'End If
'frmBudPrepMaint.SetOptions (Fund)
'frmBudPrepMaint.Show
'Unload frmBudPrepOptions
'End Sub

Private Sub cmdPrint_Click()
  If ValidDate = True Then
    If ValidFunds = True Then
      If fpcboRptType.ListIndex = 0 Then
        rptopt = 1
      ElseIf fpcboRptType.ListIndex = 1 Then
        rptopt = 2
      End If
      If rptopt = 1 Then
        PrnTrialBal
      ElseIf rptopt = 2 Then
        PrnTrialBal2
      End If
    End If
  End If
End Sub
Private Sub fpcboFund1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboFund1.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboFund1.ListIndex = -1
    fpcboFund1.Action = ActionClearSearchBuffer
  End If
  If fpcboFund1.ListDown <> True Then
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
Private Sub fpcboFund2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboFund2.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboFund2.ListIndex = -1
    fpcboFund2.Action = ActionClearSearchBuffer
  End If
  If fpcboFund2.ListDown <> True Then
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
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpTrialBalance
  FundstoList fpcboFund1
  FundstoList fpcboFund2
  txtDate.Text = Format(Now, "mm/dd/yyyy")
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub
Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboFund2.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub PrnTrialBal()
  Dim MaxLines As Integer, LookFor As String, CrLF As String, Newrp As String
  Dim Linecnt As Integer, PRNFile As Integer, FundCnt As Integer
  Dim ReportFile As String, ToPrint As String, SumLine As String
  Dim PRNFile2 As Integer, ReportFile2 As String, ToPrint2 As String
  Dim FF As String, Header As String, StartFund As String, EndFund As String
  Dim PRNFileNum As Integer, cnt As Integer, Howmany As Integer
  Dim FundCode As String, DivLine As String, DivLine2 As String
  Dim CommaFmt As String, TotalFmt As String, FundNumber As String
  Dim TotDr As Double, TotCr As Double, TranCashTot As Double, CalcBal As Double
  ReDim FundList(1) As String
  Dim FundDr As Double, FundCr As Double, FundRecNum As Integer
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer
  Dim FundIdxFileNum As Integer, NumFunds As Integer, EndDate As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctsRecs As Integer, RecNo As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long, PageNum As Integer
  Dim Debit As String, Credit As String, Diff As Double, PYFundBal As Double
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  GetFundList FundList(), NumFunds
  'Define vars used for printing
      'make sure funds are in ascending order
  StartFund$ = QPTrim$(FirstFund$)
  EndFund$ = QPTrim$(LastFund$)
  EndDate = DateDiff("d", "12/31/1979", txtDate)
  Newrp = "TRLBAL"
  GetRPTName Newrp
  ReportFile$ = Newrp   'Report File Name
  ReportFile2$ = "PBal"
  Header$ = "Trial Balance"
  ReDim Desc$(1)
  Desc$(1) = "Acct Number     Title                                   Debit           Credit"
  CommaFmt$ = "###,###,###.##"  'format takes 13 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 14 chars
  SumLine$ = String$(14, "-")   'column summary line
  DivLine$ = String$(80, "-")   'dashed line
  DivLine2$ = String$(80, "=")  'Double Line
  CrLF$ = Chr$(14) + Chr$(10)
  FF$ = Chr$(12)
  If ActiveYear = 2 Then
    ReDim FundList(1)
    GetFundList FundList(), NumFunds
    ReDim PYFundRev#(NumFunds + 1)
    ReDim PYFundExp#(NumFunds + 1)
  End If
  MaxLines = 55
  TotDr# = 0
  TotCr# = 0
  FrmShowPctComp.Label1 = "Printing Trial Balance Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnTrialBal, True
  'ReportFile$ = Unique$(Path$)
  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile
  PRNFile2 = FreeFile
  Open ReportFile2$ For Output As #PRNFile2
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum
  NumGLAcctsRecs = LOF(AcctFileNum) / Len(GLAcct)
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFileNum, NumFunds
  For cnt = 1 To NumGLAccts
    Get AcctIdxFileNum, cnt, GLAcctidx
    Get AcctFileNum, GLAcctidx.RecNum, GLAcct
    FundCode$ = QPTrim$(Left$(GLAcct.Num, GLFundLen))
'*****Problem????
    If FundCode$ >= StartFund$ And FundNumber$ <= EndFund$ Then
      If ActiveYear = 2 Then
        '--Find the fund so we can calc prior year fund balance if necessary
        For FundCnt = 1 To NumFunds
          If FundCode$ = FundList$(FundCnt) Then
            Found = True
            Exit For
          End If
        Next
      End If
      CalcBal# = Round#(GLAcct.BegBal)           'get the beginning balance
      NextTr& = GLAcct.FrstTran   'get the first trans for this acct
      Do Until NextTr& = 0      'keep going 'til we run out of trans
        Get TransFileNum, NextTr&, GLTrans
        If GLTrans.TRDATE <= EndDate Then
          Select Case GLAcct.Typ
          Case "A"              ', "E"
            CalcBal# = Round#(CalcBal# + GLTrans.DrAmt - GLTrans.CrAmt)
          Case "L"              ', "R"
            CalcBal# = Round#(CalcBal# + GLTrans.CrAmt - GLTrans.DrAmt)
          Case "R"
            'IF Trans.TrDate >= FY2BegDate THEN
            '  CalcBal# = CalcBal# + Round#(Trans.CrAmt) - Round#(Trans.DrAmt)
            'ELSE
            '  PYFundRev#(FundCnt) = PYFundRev#(FundCnt) + Round#(Trans.CrAmt)-round(
            'END IF
            Select Case ActiveYear
            Case 1
              CalcBal# = Round#(CalcBal# + GLTrans.CrAmt - GLTrans.DrAmt)
            Case 2
              If GLTrans.TRDATE < FY2BegDate Then
                PYFundRev#(FundCnt) = Round#(PYFundRev#(FundCnt) + GLTrans.CrAmt - GLTrans.DrAmt)
              Else
                CalcBal# = Round#(CalcBal# + GLTrans.CrAmt - GLTrans.DrAmt)
              End If
            End Select
          Case "E"
            'IF Trans.TrDate >= FY2BegDate THEN
            '  CalcBal# = CalcBal# + Round#(Trans.DrAmt) - Round#(Trans.CrAmt)
            'ELSE
            '  PYFundExp#(FundCnt) = PYFundExp#(FundCnt) + Round#(Trans.DrAmt)-
            'END IF
            Select Case ActiveYear
            Case 1
              CalcBal# = Round#(CalcBal# + GLTrans.DrAmt - GLTrans.CrAmt)
            Case 2
              If GLTrans.TRDATE < FY2BegDate Then
                PYFundExp#(FundCnt) = Round#(PYFundExp#(FundCnt) + GLTrans.DrAmt - GLTrans.CrAmt)
              Else
                CalcBal# = Round#(CalcBal# + GLTrans.DrAmt - GLTrans.CrAmt)
              End If
            End Select
          End Select
        End If
        NextTr& = GLTrans.NextTran                'Get the next transaction
      Loop
      GLAcct.Bal = CalcBal#
      Put AcctFileNum, GLAcctidx.RecNum, GLAcct
    End If      'test for account in fund range
    FrmShowPctComp.ShowPctComp cnt, NumGLAccts
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnTrialBal, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next
  ActivateControls frmPrnTrialBal, True
  TotDr# = 0: TotCr# = 0
  For cnt = 1 To NumFunds
    FundDr# = 0: FundCr# = 0
    Get FundIdxFileNum, cnt, GLFundIdx
    FundNumber$ = QPTrim$(GLFundIdx.FundNum)
    If FundNumber$ >= StartFund$ And FundNumber$ <= EndFund$ Then
      FundRecNum = FindFund(FundNumber$)
      FundName$ = QPTrim$(GetFundTitle(FundRecNum))
      If ActiveYear = 2 Then
        'Get the fund count for printing prior year fund balance
        Found = False
        For FundCnt = 1 To NumFunds
          If FundNumber$ = FundList$(FundCnt) Then
            Found = True
            Exit For
          End If
        Next
      End If
      'GoSub PrintPageHeader
      For RecNo = 1 To NumGLAccts               'Active Accts
        Get AcctIdxFileNum, RecNo, GLAcctidx
        FundCode$ = Left$(GLAcctidx.AcctNum, GLFundLen)
        If FundCode$ = FundNumber$ Then
          Get AcctFileNum, GLAcctidx.RecNum, GLAcct
'          Linecnt = Linecnt + 1
'          If Linecnt >= MaxLines Then
'            Print #PRNFile, FF$
'            GoSub PrintPageHeader
'          End If
          'here if you don't want to see zero bal accts
          If Chk0Bal.Value = 1 Then
            If GLAcct.Bal = 0 Then GoTo skip0acct
          End If
          ToPrint$ = Space$(80)
          ToPrint$ = FundCode$ + "~" + FundName$ + "~" + GLAcct.Num
          ToPrint$ = ToPrint$ + "~" + GLAcct.Title
          Select Case GLAcct.Typ
          Case "A", "E"
            If GLAcct.Bal >= 0 Then
              Debit$ = Using$(CommaFmt$, Str$(GLAcct.Bal))
              TotDr# = TotDr# + GLAcct.Bal
              FundDr# = FundDr# + GLAcct.Bal
              Credit$ = "0"
            Else
              Credit$ = Using$(CommaFmt$, Str$(Abs(GLAcct.Bal)))
              TotCr# = TotCr# + Abs(GLAcct.Bal)
              FundCr# = FundCr# + Abs(GLAcct.Bal)
              Debit$ = "0"
            End If
          Case "L", "R"
            If GLAcct.Bal >= 0 Then
              Credit$ = Using$(CommaFmt$, Str$(GLAcct.Bal))
              TotCr# = TotCr# + GLAcct.Bal
              FundCr# = FundCr# + GLAcct.Bal
              Debit$ = "0"
            Else
              Debit$ = Using$(CommaFmt$, Str$(Abs(GLAcct.Bal)))
              TotDr# = TotDr# + Abs(GLAcct.Bal)
              FundDr# = FundDr# + Abs(GLAcct.Bal)
              Credit$ = "0"
            End If
          End Select
          ToPrint$ = ToPrint$ + "~" + Debit$ + "~" + Credit$
          
          Print #PRNFile, ToPrint$
        End If
skip0acct:
      Next
      Diff# = Round#(FundDr# - FundCr#)
'      If Diff# <> 0 Then
'
'        LSet ToPrint$ = "Fund is out of balance :"
'        Mid$(ToPrint$, 25) = Using$(CommaFmt$, Str$(Diff#))
'        Print #PRNFile, ToPrint$
'      End If
      If ActiveYear = 2 Then
        PYFundBal# = 0
        PYFundBal# = Round#(PYFundRev#(FundCnt) - PYFundExp#(FundCnt))
        ToPrint2$ = Space$(80)
        ToPrint2$ = FundNumber$
        ToPrint2$ = ToPrint2$ + "~" + Using$(CommaFmt$, Str$(PYFundBal#))
        Print #PRNFile2, ToPrint2$
      End If
     ' Print #PRNFile,
      'Print #PRNFile, FF$
    End If      'if fund is in range test
  Next          'next fund
 Load frmLoadingRpt
 ' ToPrint$ = Space$(80)
  '--Print a grand total if more than one fund
'  If StartFund$ <> EndFund$ Then
'    ARptTrialBal.ReportFooter.Visible = True
'  Else
'    ARptTrialBal.ReportFooter.Visible = False
'    Print #PRNFile, "Combined totals - All Funds"
'    Print #PRNFile, "Total Debits  : " + Using$(TotalFmt$, Str$(TotDr#))
'    Print #PRNFile, "Total Credits : " + Using$(TotalFmt$, Str$(TotCr#))
'    Print #PRNFile, Chr$(12)
'  End If
  Close
  ARptTrialBal.Label10.Caption = "Reporting For Fund " + StartFund$ + " Thru Fund " + EndFund$
  ARptTrialBal.txtRptDate = "Period Ending: " + txtDate
  ARptTrialBal.gTotDebits = Using$(TotalFmt$, Str$(TotDr#))
  ARptTrialBal.gTotCredits = Using$(TotalFmt$, Str$(TotCr#))
  ARptTrialBal.ActYear = ActiveYear
  ARptTrialBal.txtDate = Now
  ARptTrialBal.txtTown = GLUserName$
  ARptTrialBal.GetName ReportFile$, ReportFile2$
  ARptTrialBal.startrpt

  'End Report Processing
 ' ViewPrint ReportFile$, "Trial Balance"
 ' KillFile ReportFile$
  Exit Sub
'PrintPageHeader:
'  PageNum = PageNum + 1
'  Print #PRNFile, GLUserName$; Tab(45); "Run Date: " + Date$; "        Page "; PageNum
'  Print #PRNFile, FundName$ + " " + Header$
'  Print #PRNFile, "Period Ending: " + txtDate
'  Print #PRNFile,
'  Print #PRNFile, Desc$(1)
'  Print #PRNFile, DivLine$
'  Linecnt = 6
'  Return
CancelExit:
Exit Sub
End Sub
Private Sub PrnTrialBal2()
  Dim MaxLines As Integer, LookFor As String, CrLF As String, Newrp As String
  Dim Linecnt As Integer, PRNFile As Integer, FundCnt As Integer
  Dim ReportFile As String, ToPrint As String, SumLine As String
  Dim FF As String, Header As String, StartFund As String, EndFund As String
  Dim PRNFileNum As Integer, cnt As Integer, Howmany As Integer
  Dim FundCode As String, DivLine As String, DivLine2 As String
  Dim CommaFmt As String, TotalFmt As String, FundNumber As String
  Dim TotDr As Double, TotCr As Double, TranCashTot As Double, CalcBal As Double
  ReDim FundList(1) As String
  Dim FundDr As Double, FundCr As Double, FundRecNum As Integer
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer
  Dim FundIdxFileNum As Integer, NumFunds As Integer, EndDate As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctsRecs As Integer, RecNo As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long, PageNum As Integer
  Dim Debit As String, Credit As String, Diff As Double, PYFundBal As Double
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  GetFundList FundList(), NumFunds
  'Define vars used for printing
      'make sure funds are in ascending order
  StartFund$ = QPTrim$(FirstFund$)
  EndFund$ = QPTrim$(LastFund$)
  EndDate = DateDiff("d", "12/31/1979", txtDate)
  Newrp = "TRLBAL"
  GetRPTName Newrp
  ReportFile$ = Newrp              'Report File Name
  Header$ = "Trial Balance"
  ReDim Desc$(1)
  Desc$(1) = "Acct Number     Title                                   Debit           Credit"
  CommaFmt$ = "###,###,###.##"  'format takes 13 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 14 chars
  SumLine$ = String$(14, "-")   'column summary line
  DivLine$ = String$(80, "-")   'dashed line
  DivLine2$ = String$(80, "=")  'Double Line
  CrLF$ = Chr$(14) + Chr$(10)
  FF$ = Chr$(12)
  If ActiveYear = 2 Then
    ReDim FundList(1)
    GetFundList FundList(), NumFunds
    ReDim PYFundRev#(NumFunds + 1)
    ReDim PYFundExp#(NumFunds + 1)
  End If
  MaxLines = 55
  TotDr# = 0
  TotCr# = 0
  FrmShowPctComp.Label1 = "Printing Trial Balance Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnTrialBal, True
  'ReportFile$ = Unique$(Path$)
  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum
  NumGLAcctsRecs = LOF(AcctFileNum) / Len(GLAcct)
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFileNum, NumFunds
  For cnt = 1 To NumGLAccts
    Get AcctIdxFileNum, cnt, GLAcctidx
    Get AcctFileNum, GLAcctidx.RecNum, GLAcct
    FundCode$ = QPTrim$(Left$(GLAcct.Num, GLFundLen))
'*****Problem????
    If FundCode$ >= StartFund$ And FundNumber$ <= EndFund$ Then
      If ActiveYear = 2 Then
        '--Find the fund so we can calc prior year fund balance if necessary
        For FundCnt = 1 To NumFunds
          If FundCode$ = FundList$(FundCnt) Then
            Found = True
            Exit For
          End If
        Next
      End If
      CalcBal# = Round#(GLAcct.BegBal)           'get the beginning balance
      NextTr& = GLAcct.FrstTran   'get the first trans for this acct
      Do Until NextTr& = 0      'keep going 'til we run out of trans
        Get TransFileNum, NextTr&, GLTrans
        If GLTrans.TRDATE <= EndDate Then
          Select Case GLAcct.Typ
          Case "A"              ', "E"
            CalcBal# = Round#(CalcBal# + GLTrans.DrAmt - GLTrans.CrAmt)
          Case "L"              ', "R"
            CalcBal# = Round#(CalcBal# + GLTrans.CrAmt - GLTrans.DrAmt)
          Case "R"
            'IF Trans.TrDate >= FY2BegDate THEN
            '  CalcBal# = CalcBal# + Round#(Trans.CrAmt) - Round#(Trans.DrAmt)
            'ELSE
            '  PYFundRev#(FundCnt) = PYFundRev#(FundCnt) + Round#(Trans.CrAmt)-round(
            'END IF
            Select Case ActiveYear
            Case 1
              CalcBal# = Round#(CalcBal# + GLTrans.CrAmt - GLTrans.DrAmt)
            Case 2
              If GLTrans.TRDATE < FY2BegDate Then
                PYFundRev#(FundCnt) = Round#(PYFundRev#(FundCnt) + GLTrans.CrAmt - GLTrans.DrAmt)
              Else
                CalcBal# = Round#(CalcBal# + GLTrans.CrAmt - GLTrans.DrAmt)
              End If
            End Select
          Case "E"
            'IF Trans.TrDate >= FY2BegDate THEN
            '  CalcBal# = CalcBal# + Round#(Trans.DrAmt) - Round#(Trans.CrAmt)
            'ELSE
            '  PYFundExp#(FundCnt) = PYFundExp#(FundCnt) + Round#(Trans.DrAmt)-
            'END IF
            Select Case ActiveYear
            Case 1
              CalcBal# = Round#(CalcBal# + GLTrans.DrAmt - GLTrans.CrAmt)
            Case 2
              If GLTrans.TRDATE < FY2BegDate Then
                PYFundExp#(FundCnt) = Round#(PYFundExp#(FundCnt) + GLTrans.DrAmt - GLTrans.CrAmt)
              Else
                CalcBal# = Round#(CalcBal# + GLTrans.DrAmt - GLTrans.CrAmt)
              End If
            End Select
          End Select
        End If
        NextTr& = GLTrans.NextTran                'Get the next transaction
      Loop
      GLAcct.Bal = CalcBal#
      Put AcctFileNum, GLAcctidx.RecNum, GLAcct
    End If      'test for account in fund range
    FrmShowPctComp.ShowPctComp cnt, NumGLAccts
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnTrialBal, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next
  ActivateControls frmPrnTrialBal, True
  TotDr# = 0: TotCr# = 0
  For cnt = 1 To NumFunds
    FundDr# = 0: FundCr# = 0
    Get FundIdxFileNum, cnt, GLFundIdx
    FundNumber$ = QPTrim$(GLFundIdx.FundNum)
    If FundNumber$ >= StartFund$ And FundNumber$ <= EndFund$ Then
      FundRecNum = FindFund(FundNumber$)
      FundName$ = QPTrim$(GetFundTitle(FundRecNum))
      If ActiveYear = 2 Then
        'Get the fund count for printing prior year fund balance
        Found = False
        For FundCnt = 1 To NumFunds
          If FundNumber$ = FundList$(FundCnt) Then
            Found = True
            Exit For
          End If
        Next
      End If
      GoSub PrintPageHeader
      For RecNo = 1 To NumGLAccts               'Active Accts
        Get AcctIdxFileNum, RecNo, GLAcctidx
        FundCode$ = Left$(GLAcctidx.AcctNum, GLFundLen)
        If FundCode$ = FundNumber$ Then
          Get AcctFileNum, GLAcctidx.RecNum, GLAcct
          Linecnt = Linecnt + 1
          If Linecnt >= MaxLines Then
            Print #PRNFile, FF$
            GoSub PrintPageHeader
          End If
          'here if you don't want to see zero bal accts
          If Chk0Bal.Value = 1 Then
            If GLAcct.Bal = 0 Then GoTo skip0acct
          End If
          ToPrint$ = Space$(80)
          LSet ToPrint$ = GLAcct.Num
          Mid$(ToPrint$, 17) = GLAcct.Title
          Select Case GLAcct.Typ
          Case "A", "E"
            If GLAcct.Bal >= 0 Then
              Debit$ = Using$(CommaFmt$, Str$(GLAcct.Bal))
              TotDr# = TotDr# + GLAcct.Bal
              FundDr# = FundDr# + GLAcct.Bal
              Credit$ = ""
            Else
              Credit$ = Using$(CommaFmt$, Str$(Abs(GLAcct.Bal)))
              TotCr# = TotCr# + Abs(GLAcct.Bal)
              FundCr# = FundCr# + Abs(GLAcct.Bal)
              Debit$ = ""
            End If
          Case "L", "R"
            If GLAcct.Bal >= 0 Then
              Credit$ = Using$(CommaFmt$, Str$(GLAcct.Bal))
              TotCr# = TotCr# + GLAcct.Bal
              FundCr# = FundCr# + GLAcct.Bal
              Debit$ = ""
            Else
              Debit$ = Using$(CommaFmt$, Str$(Abs(GLAcct.Bal)))
              TotDr# = TotDr# + Abs(GLAcct.Bal)
              FundDr# = FundDr# + Abs(GLAcct.Bal)
              Credit$ = ""
            End If
          End Select
          Mid$(ToPrint$, 50) = Debit$
          Mid$(ToPrint$, 67) = Credit$
          Print #PRNFile, ToPrint$
        End If
skip0acct:
      Next
      ToPrint$ = Space$(80)
      Mid$(ToPrint$, 50) = SumLine$
      Mid$(ToPrint$, 67) = SumLine$
      Print #PRNFile, ToPrint$
      ToPrint$ = Space$(80)
      LSet ToPrint$ = FundName$ + " " + "Totals"
      Mid$(ToPrint$, 48) = Using$(TotalFmt$, Str$(FundDr#))
      Mid$(ToPrint$, 65) = Using$(TotalFmt$, Str$(FundCr#))
      Print #PRNFile, ToPrint$
      Diff# = Round#(FundDr# - FundCr#)
      If Diff# <> 0 Then
        ToPrint$ = Space$(80)
        LSet ToPrint$ = "Fund is out of balance :"
        Mid$(ToPrint$, 25) = Using$(CommaFmt$, Str$(Diff#))
        Print #PRNFile, ToPrint$
      End If
      If ActiveYear = 2 Then
        PYFundBal# = 0
        PYFundBal# = Round#(PYFundRev#(FundCnt) - PYFundExp#(FundCnt))
        ToPrint$ = Space$(80)
        LSet ToPrint$ = "Prior Year unaudited fund balance adjustment:"
        Mid$(ToPrint$, 41) = Using$(CommaFmt$, Str$(PYFundBal#))
        Print #PRNFile, ToPrint$
      End If
      Print #PRNFile,
      Print #PRNFile, FF$
    End If      'if fund is in range test
  Next          'next fund
  ToPrint$ = Space$(80)
  '--Print a grand total if more than one fund
  If StartFund$ <> EndFund$ Then
    Print #PRNFile, "Combined totals - All Funds"
    Print #PRNFile, "Total Debits  : " + Using$(TotalFmt$, Str$(TotDr#))
    Print #PRNFile, "Total Credits : " + Using$(TotalFmt$, Str$(TotCr#))
    Print #PRNFile, Chr$(12)
  End If
  Close
  'End Report Processing
  ViewPrint ReportFile$, "Trial Balance"
  KillFile ReportFile$
  Exit Sub
PrintPageHeader:
  PageNum = PageNum + 1
  Print #PRNFile, GLUserName$; Tab(45); "Run Date: " + Date$; "        Page "; PageNum
  Print #PRNFile, FundName$ + " " + Header$
  Print #PRNFile, "Period Ending: " + txtDate
  Print #PRNFile,
  Print #PRNFile, Desc$(1)
  Print #PRNFile, DivLine$
  Linecnt = 6
  Return
CancelExit:
Exit Sub
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub
