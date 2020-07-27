VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmPrnBalSheet 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance Sheet"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmPrnBalSheet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   405
      Left            =   5850
      TabIndex        =   3
      Top             =   5310
      Width           =   1920
      _Version        =   196608
      _ExtentX        =   3387
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
      ColDesigner     =   "frmPrnBalSheet.frx":08CA
   End
   Begin LpLib.fpCombo fpcboFund1 
      Height          =   405
      Left            =   5850
      TabIndex        =   1
      Top             =   3915
      Width           =   3090
      _Version        =   196608
      _ExtentX        =   5450
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
      ColDesigner     =   "frmPrnBalSheet.frx":0C68
   End
   Begin LpLib.fpCombo fpcboFund2 
      Height          =   405
      Left            =   5850
      TabIndex        =   2
      Top             =   4605
      Width           =   3090
      _Version        =   196608
      _ExtentX        =   5450
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
      ColDesigner     =   "frmPrnBalSheet.frx":1053
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
      Left            =   10320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7464
      Width           =   1332
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
      Left            =   8640
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7464
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   8508
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "3:15 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "9/21/2010"
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
   Begin EditLib.fpDateTime txtDate 
      Height          =   372
      Left            =   5856
      TabIndex        =   0
      Top             =   3240
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
   Begin VB.Label Label7 
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
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   3120
      TabIndex        =   11
      Top             =   5328
      Width           =   2388
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
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   1
      Left            =   3792
      TabIndex        =   10
      Top             =   3972
      Width           =   1668
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
      ForeColor       =   &H8000000E&
      Height          =   420
      Index           =   0
      Left            =   3882
      TabIndex        =   9
      Top             =   3276
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   3120
      Picture         =   "frmPrnBalSheet.frx":143E
      Top             =   3195
      Width           =   360
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3186
      Top             =   1296
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Balance Sheet"
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
      Left            =   4218
      TabIndex        =   8
      Top             =   1536
      Width           =   3852
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3060
      Left            =   2688
      Top             =   2904
      Width           =   6828
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
      ForeColor       =   &H8000000E&
      Height          =   324
      Index           =   1
      Left            =   3840
      TabIndex        =   7
      Top             =   4668
      Width           =   1596
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      Height          =   972
      Left            =   3186
      Top             =   1176
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
Attribute VB_Name = "frmPrnBalSheet"
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
  Unload frmPrnBalSheet
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

Private Sub cmdPrint_Click()
  If ValidDate = True Then
    If ValidFunds = True Then
      If fpcboRptType.ListIndex = 0 Then
        rptopt = 1
      ElseIf fpcboRptType.ListIndex = 1 Then
        rptopt = 2
      End If
      If rptopt = 1 Then
        PrnBalSheet
      ElseIf rptopt = 2 Then
        PrnBalSheet2
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
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdPrint_Click
      KeyCode = 0
    Case Else:
  End Select
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
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboFund2.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtDate.SetFocus
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
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboFund1.SetFocus
        KeyCode = 0
      End If
    End If
  End If
 
End Sub
Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
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
Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpBalanceSheet
  FundstoList fpcboFund1
  FundstoList fpcboFund2
'  FundList fpcboFund1
'  fpcboFund1.RemoveItem 0
'  FundList fpcboFund2
'  fpcboFund2.RemoveItem 0
  txtDate.Text = Format(Now, "mm/dd/yyyy")
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub
Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboFund1.SetFocus
  End If
End Sub

Private Sub PrnBalSheet()
  Dim TotRev As Double, CrLF As String, Newrp As String
  Dim PRNFile As Integer, FundCnt As Integer
  Dim ReportFile As String, ToPrint As String, SumLine As String
  Dim TotExp As Double, Header As String, StartFund As String, EndFund As String
  Dim PRNFileNum As Integer, cnt As Integer, GlAc As Integer, Rec As Integer
  Dim FundCode As String, DivLine As String, DivLine2 As String
  Dim CommaFmt As String, TotalFmt As String, FundNumber As String
  Dim FundAsset As Double, FundLiab As Double, TotLiab As Double, CalcBal As Double
  ReDim FundList(1) As String
  Dim FundDr As Double, FundCr As Double, FundRecNum As Integer
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer
  Dim FundIdxFile As Integer, NumFunds As Integer, EndDate As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, LCnt As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long
  Dim Debit As String, Credit As String, ACnt As Integer, TotAsset As Double
  Dim PageBreak As String, RptTitle As String, ThisFund As String, UsingFund As Boolean
  Dim FundRec As Integer, AcctType As String, DrCr As String, FundBalAdj As Double
  Dim TotLiabnCap As Double, BalAdj As String, PageNum As Integer
  Dim F As String, N As String
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  GetFundList FundList(), NumFunds
  'Define vars used for printing
      'make sure funds are in ascending order
  StartFund$ = QPTrim$(FirstFund$)
  EndFund$ = QPTrim$(LastFund$)
  EndDate = DateDiff("d", "12/31/1979", txtDate)


  'End of Input
  '=====================================================
  'Start Report Processing
  ReDim Desc$(1)
  CommaFmt$ = "###,###,###.##"  'format takes 13 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 14 chars
  SumLine$ = String$(16, "-")   'column summary line
  DivLine$ = String$(79, "-")   'dashed line
  DivLine2$ = String$(79, "=")  'Double Line
  CrLF$ = Chr$(13) + Chr$(10)
  RptTitle$ = "Balance Sheet"
  Desc$(1) = "Acct Number     Title"
  'Desc$ = "Acct Number     Title
  Newrp = "BalSh"
  GetRPTName Newrp
  ReportFile$ = Newrp  'Report File Name
  'ReportFile$ = Unique$(Path$)
  PRNFile = FreeFile

  Open ReportFile$ For Output As #PRNFile

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum
  NumGLAcctRecs = LOF(AcctFileNum) / Len(GLAcct)
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFile, NumFunds

  ReDim Asset%(1 To NumGLAccts) 'Holds all Asset acct record nums
  ReDim Liab%(1 To NumGLAccts)  'Holds all Liab acct record nums
  ReDim FundList$(1 To NumFunds)                'List of all active Funds
  ReDim FundTotAssets#(1 To NumFunds)           'List of total assets by fund
  ReDim FundTotLiab#(1 To NumFunds)             'list of total liab by fund
  ReDim FundTotRev#(1 To NumFunds)              'List of total revenues by fun
  ReDim FundTotExp#(1 To NumFunds)              'list of tot exp by fund

  For Fund = 1 To NumFunds
    Get FundIdxFile, Fund, GLFundIdx
    FundList$(Fund) = QPTrim$(GLFundIdx.FundNum)
  Next
  FrmShowPctComp.Label1 = "Printing Balance Sheet"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdPrint.Enabled = False
  Me.mnuOptions.Enabled = False
  For GlAc = 1 To NumGLAccts

    Get AcctIdxFileNum, GlAc, GLAcctidx
    Get AcctFileNum, GLAcctidx.RecNum, GLAcct

    FundCode$ = Left$(GLAcct.Num, GLFundLen)
    If FundCode$ >= StartFund$ And FundCode$ <= EndFund$ Then

'''      '--Show the account we're processing
'''      QPrintRC Acct.Num, 25, 14, -1

      '--get the beginning balance and the first transaction rec num
      CalcBal# = Round#(GLAcct.BegBal)
      NextTr& = GLAcct.FrstTran

      '--Loop until we run out of transactions for this account
      Do Until NextTr& = 0

        Get TransFileNum, NextTr&, GLTrans
        If GLTrans.TRDATE <= EndDate Then

          '--Calculate balance depending on account type
          Select Case GLAcct.Typ
          Case "A", "E"
            CalcBal# = Round#(CalcBal# + GLTrans.DrAmt - GLTrans.CrAmt)
          Case "L", "R"
            CalcBal# = Round#(CalcBal# + GLTrans.CrAmt - GLTrans.DrAmt)
          End Select

        End If

        '--Get the next transaction
        NextTr& = GLTrans.NextTran

      Loop

      '--Update account balance
      GLAcct.Bal = Round#(CalcBal#)
      Put AcctFileNum, GLAcctidx.RecNum, GLAcct

      '--Apply balance to proper fund
      For Fund = 1 To NumFunds
        If FundCode$ = FundList$(Fund) Then
          Select Case GLAcct.Typ
          Case "A"
            TotAsset# = TotAsset# + CalcBal#
            ACnt = ACnt + 1
            Asset%(ACnt) = GLAcctidx.RecNum
            FundTotAssets#(Fund) = FundTotAssets#(Fund) + CalcBal#

          Case "L"
            TotLiab# = TotLiab# + CalcBal#
            LCnt = LCnt + 1
            Liab%(LCnt) = GLAcctidx.RecNum
            FundTotLiab#(Fund) = FundTotLiab#(Fund) + CalcBal#

          Case "R"

            TotRev# = TotRev# + CalcBal#
            FundTotRev#(Fund) = FundTotRev#(Fund) + CalcBal#

          Case "E"
            TotExp# = TotExp# + CalcBal#
            FundTotExp#(Fund) = FundTotExp#(Fund) + CalcBal#

          End Select

          Exit For              'acct was a match to this fund so exit loop
        End If  'test for fund in list
      Next      'Summarize next fund's totals
    End If      'End of fund in range test
    FrmShowPctComp.ShowPctComp GlAc, NumGLAccts

  Next          'Process next account


  '--Now write the report to file.
  '--Go thru fund list to see if its one we're reporting on
  For Fund = 1 To NumFunds
    ThisFund$ = FundList$(Fund)
    If ThisFund$ >= StartFund$ And ThisFund$ <= EndFund$ Then
      UsingFund = True
      FundRec = FindFund(ThisFund$)
      FundName$ = GetFundTitle(FundRec)
    Else
      UsingFund = False
    End If

    If UsingFund Then
      '--List Assets for this fund
      AcctType$ = "Assets"
      'GoSub PrintBSHeader

      For cnt = 1 To ACnt
        Rec = Asset%(cnt)
        Get AcctFileNum, Rec, GLAcct
        FundCode$ = Left$(GLAcct.Num, GLFundLen)

        If FundCode$ = ThisFund$ Then
          ToPrint$ = ""
          ToPrint$ = FundCode$ + "~" + FundName$ + "~" + AcctType$
          ToPrint$ = ToPrint$ + "~" + QPTrim(GLAcct.Num)
          ToPrint$ = ToPrint$ + "~" + QPTrim(GLAcct.Title)
          If GLAcct.Bal >= 0 Then
            Debit$ = Using$(CommaFmt$, Str$(GLAcct.Bal))
            'Credit$ = "0"
          Else
            'Debit$ = "0"
            'Credit$ = "(" + Using$(CommaFmt$, Str$(GLAcct.Bal)) + ")"
            Debit$ = "(" + Using$(CommaFmt$, Str$(Abs(GLAcct.Bal))) + ")"
          End If
          ToPrint$ = ToPrint$ + "~" + Debit$
          'ToPrint$ = ToPrint$ + "~" + Credit$
          Print #PRNFile, ToPrint$

        End If

      Next      'AssetAcct

      ToPrint$ = ""
      If FundTotAssets#(Fund) >= 0 Then
        FundAsset# = FundTotAssets#(Fund)
        DrCr$ = " " + Using$(TotalFmt$, Str$(FundAsset#)) + " "
      Else
        FundAsset# = FundTotAssets#(Fund)
        DrCr$ = "(" + Using$(TotalFmt$, Str$(Abs(FundAsset#))) + ")"
        'FundAsset# = Abs(FundTotAssets#(Fund))
      End If

'      Print #PRNFile, DivLine$
'      LSet ToPrint$ = "Total Assets"
'      Mid$(ToPrint$, 61) = DrCr$
'      Print #PRNFile, ToPrint$

      '--List Liabilities
      AcctType$ = "Liabilities"
'      GoSub BreakBSPage
'      GoSub PrintBSHeader

      For cnt = 1 To LCnt
        Rec = Liab%(cnt)
        Get AcctFileNum, Rec, GLAcct

        FundCode$ = Left$(GLAcct.Num, GLFundLen)
        N$ = FundName$
        If FundCode$ = ThisFund$ Then
          ToPrint$ = ""
          ToPrint$ = FundCode$ + "~" + FundName$ + "~" + AcctType$
          ToPrint$ = ToPrint$ + "~" + QPTrim(GLAcct.Num)
          ToPrint$ = ToPrint$ + "~" + QPTrim(GLAcct.Title)
          If GLAcct.Bal >= 0 Then
            Credit$ = Using$(CommaFmt$, Str$(GLAcct.Bal))
            'Debit$ = "0"
          Else
            'Credit$ = "0"
            'Debit$ = "(" + Using$(CommaFmt$, Str$(GLAcct.Bal)) + ")"
            Credit$ = "(" + Using$(CommaFmt$, Str$(Abs(GLAcct.Bal))) + ")"
          End If

          'ToPrint$ = ToPrint$ + "~" + Debit$
          ToPrint$ = ToPrint$ + "~" + Credit$
          Print #PRNFile, ToPrint$

        End If

      Next      'LiabAcct
      
      FundBalAdj# = Round#(FundTotRev#(Fund) - FundTotExp#(Fund))
      If FundBalAdj# >= 0 Then
        BalAdj$ = " " + Using$(CommaFmt$, Str$(FundBalAdj#)) + " "
      Else
        BalAdj$ = "(" + Using$(CommaFmt$, Str$(Abs(FundBalAdj#))) + ")"
      End If
      F$ = FundList$(Fund)
      ToPrint$ = ""
      ToPrint$ = F$ + "~" + N$ + "~" + AcctType$
      ToPrint$ = ToPrint$ + "~" + "Current Fund" + "~" + "Rev Over/Under Exp"
      ToPrint$ = ToPrint$ + "~" + BalAdj$
      Print #PRNFile, ToPrint$
      BalAdj$ = ""
      ToPrint$ = Space$(80)
      If FundTotLiab#(Fund) >= 0 Then
        FundLiab# = FundTotLiab#(Fund)
        DrCr$ = "Cr"
      Else
        'Put this in to replace the one below.
        FundLiab# = FundTotLiab#(Fund)
        'The statement below had Abs and This caused Neg Num not to calc correctly.
        'FundLiab# = Abs(FundTotLiab#(Fund))
        DrCr$ = "Dr"
      End If
      'Print #PRNFile, DivLine$

      TotLiabnCap# = FundLiab# + FundBalAdj#
      If TotLiabnCap# >= 0 Then
        DrCr$ = " " + Using$(TotalFmt$, Str$(TotLiabnCap#)) + " "
      Else
        DrCr$ = "(" + Using$(TotalFmt$, Str$(Abs(TotLiabnCap#))) + ")"
      End If
      'LSet ToPrint$ = "Total Liabilities & Fund Balance"
      'Mid$(ToPrint$, 62) = DrCr$
      'Print #PRNFile, ToPrint$

      'Print #PRNFile, PageBreak$

      '--Now summarize the fund
      'ToPrint$ = SPACE$(80)
      'LSET ToPrint$ = "Fund Summary:"
      'PRINT #PrnFile, ToPrint$
      'ToPrint$ = SPACE$(80)

      'LSET ToPrint$ = "Revenues"
      'MID$(ToPrint$, 17) = FUsing$(STR$(FundTotRev#(Fund)), CommaFmt$)
      'PRINT #PrnFile, ToPrint$
      'ToPrint$ = SPACE$(80)

      'LSET ToPrint$ = "Expenditures"
      'MID$(ToPrint$, 17) = FUsing$(STR$(FundTotExp#(Fund)), CommaFmt$)
      'PRINT #PrnFile, ToPrint$

      'FundBalAdj# = Round#(FundTotRev#(Fund) - FundTotExp#(Fund))
      'IF FundBalAdj# < 0 THEN
      '   DrCr$ = " Debit"
      'ELSE
      '   DrCr$ = " Credit"
      'END IF

      'ToPrint$ = SPACE$(80)
      'LSET ToPrint$ = "Fund Bal Adj"
      'MID$(ToPrint$, 17) = FUsing$(STR$(FundBalAdj#), CommaFmt$) + DrCr$
      'PRINT #PrnFile, ToPrint$

    End If      'using fund test
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdExit.Enabled = True
      Me.cmdPrint.Enabled = True
      Me.mnuOptions.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

  Next          'FundList$()

  Me.cmdExit.Enabled = True
  Me.cmdPrint.Enabled = True
  EnableCloseButton Me.hwnd, True
  Me.mnuOptions.Enabled = True

  '--This summarizes all funds--
  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Assets"
  'MID$(ToPrint$, 17) = STR$(TotAsset#)
  'MID$(ToPrint$, 37) = STR$(ACnt)
  'PRINT #PrnFile, ToPrint$

  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Liabilities"
  'MID$(ToPrint$, 17) = STR$(TotLiab#)
  'MID$(ToPrint$, 37) = STR$(LCnt)
  'PRINT #PrnFile, ToPrint$
  '
  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Revenues"
  'MID$(ToPrint$, 17) = STR$(TotRev#)
  'PRINT #PrnFile, ToPrint$
  '
  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Expenditures"
  'MID$(ToPrint$, 17) = STR$(TotExp#)
  'PRINT #PrnFile, ToPrint$
  'FundBalAdj# = TotRev# - TotExp#
  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Fund Bal Adj"
  'MID$(ToPrint$, 17) = STR$(FundBalAdj#)
  'PRINT #PrnFile, ToPrint$
  '^--All funds summary --

  'PRINT #PrnFile, PageBreak$

  Close

  'End Report Processing
  '===========================================================================
  'Start Report Printing
  'ViewPrint ReportFile$, Header$
  'KillFile ReportFile$
  ARptBalanceSheet.txtDate = Now
  ARptBalanceSheet.txtTown = GLUserName$
  ARptBalanceSheet.Label4.Caption = "Period Ending: " + txtDate
  ARptBalanceSheet.GetName ReportFile$
  ARptBalanceSheet.startrpt
  Exit Sub

CancelExit:
  Exit Sub
End Sub

Private Sub PrnBalSheet2()
  Dim MaxLines As Integer, TotRev As Double, CrLF As String, Newrp As String
  Dim Linecnt As Integer, PRNFile As Integer, FundCnt As Integer
  Dim ReportFile As String, ToPrint As String, SumLine As String
  Dim TotExp As Double, Header As String, StartFund As String, EndFund As String
  Dim PRNFileNum As Integer, cnt As Integer, GlAc As Integer, Rec As Integer
  Dim FundCode As String, DivLine As String, DivLine2 As String
  Dim CommaFmt As String, TotalFmt As String, FundNumber As String
  Dim FundAsset As Double, FundLiab As Double, TotLiab As Double, CalcBal As Double
  ReDim FundList(1) As String
  Dim FundDr As Double, FundCr As Double, FundRecNum As Integer
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer
  Dim FundIdxFile As Integer, NumFunds As Integer, EndDate As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, LCnt As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long
  Dim Debit As String, Credit As String, ACnt As Integer, TotAsset As Double
  Dim PageBreak As String, RptTitle As String, ThisFund As String, UsingFund As Boolean
  Dim FundRec As Integer, AcctType As String, DrCr As String, FundBalAdj As Double
  Dim TotLiabnCap As Double, BalAdj As String, PageNum As Integer
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  GetFundList FundList(), NumFunds
  'Define vars used for printing
      'make sure funds are in ascending order
  StartFund$ = QPTrim$(FirstFund$)
  EndFund$ = QPTrim$(LastFund$)
  EndDate = DateDiff("d", "12/31/1979", txtDate)


  'End of Input
  '=====================================================
  'Start Report Processing
  ReDim Desc$(1)
  CommaFmt$ = "###,###,###.##"  'format takes 13 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 14 chars
  SumLine$ = String$(16, "-")   'column summary line
  DivLine$ = String$(79, "-")   'dashed line
  DivLine2$ = String$(79, "=")  'Double Line
  CrLF$ = Chr$(13) + Chr$(10)
  PageBreak$ = Chr$(12)
  RptTitle$ = "Balance Sheet"
  Desc$(1) = "Acct Number     Title"
  'Desc$ = "Acct Number     Title
  MaxLines = 55
  Newrp = "BalSh"
  GetRPTName Newrp
  ReportFile$ = Newrp  'Report File Name
  'ReportFile$ = Unique$(Path$)
  PRNFile = FreeFile

  Open ReportFile$ For Output As #PRNFile

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum
  NumGLAcctRecs = LOF(AcctFileNum) / Len(GLAcct)
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFile, NumFunds

  ReDim Asset%(1 To NumGLAccts) 'Holds all Asset acct record nums
  ReDim Liab%(1 To NumGLAccts)  'Holds all Liab acct record nums
  ReDim FundList$(1 To NumFunds)                'List of all active Funds
  ReDim FundTotAssets#(1 To NumFunds)           'List of total assets by fund
  ReDim FundTotLiab#(1 To NumFunds)             'list of total liab by fund
  ReDim FundTotRev#(1 To NumFunds)              'List of total revenues by fun
  ReDim FundTotExp#(1 To NumFunds)              'list of tot exp by fund

  For Fund = 1 To NumFunds
    Get FundIdxFile, Fund, GLFundIdx
    FundList$(Fund) = QPTrim$(GLFundIdx.FundNum)
  Next
  FrmShowPctComp.Label1 = "Printing Balance Sheet"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdPrint.Enabled = False
  Me.mnuOptions.Enabled = False
  For GlAc = 1 To NumGLAccts

    Get AcctIdxFileNum, GlAc, GLAcctidx
    Get AcctFileNum, GLAcctidx.RecNum, GLAcct

    FundCode$ = Left$(GLAcct.Num, GLFundLen)
    If FundCode$ >= StartFund$ And FundCode$ <= EndFund$ Then

'''      '--Show the account we're processing
'''      QPrintRC Acct.Num, 25, 14, -1

      '--get the beginning balance and the first transaction rec num
      CalcBal# = Round#(GLAcct.BegBal)
      NextTr& = GLAcct.FrstTran

      '--Loop until we run out of transactions for this account
      Do Until NextTr& = 0

        Get TransFileNum, NextTr&, GLTrans
        If GLTrans.TRDATE <= EndDate Then

          '--Calculate balance depending on account type
          Select Case GLAcct.Typ
          Case "A", "E"
            CalcBal# = Round#(CalcBal# + GLTrans.DrAmt - GLTrans.CrAmt)
          Case "L", "R"
            CalcBal# = Round#(CalcBal# + GLTrans.CrAmt - GLTrans.DrAmt)
          End Select

        End If

        '--Get the next transaction
        NextTr& = GLTrans.NextTran

      Loop

      '--Update account balance
      GLAcct.Bal = Round#(CalcBal#)
      Put AcctFileNum, GLAcctidx.RecNum, GLAcct

      '--Apply balance to proper fund
      For Fund = 1 To NumFunds
        If FundCode$ = FundList$(Fund) Then
          Select Case GLAcct.Typ
          Case "A"
            TotAsset# = TotAsset# + CalcBal#
            ACnt = ACnt + 1
            Asset%(ACnt) = GLAcctidx.RecNum
            FundTotAssets#(Fund) = FundTotAssets#(Fund) + CalcBal#

          Case "L"
            TotLiab# = TotLiab# + CalcBal#
            LCnt = LCnt + 1
            Liab%(LCnt) = GLAcctidx.RecNum
            FundTotLiab#(Fund) = FundTotLiab#(Fund) + CalcBal#

          Case "R"

            TotRev# = TotRev# + CalcBal#
            FundTotRev#(Fund) = FundTotRev#(Fund) + CalcBal#

          Case "E"
            TotExp# = TotExp# + CalcBal#
            FundTotExp#(Fund) = FundTotExp#(Fund) + CalcBal#

          End Select

          Exit For              'acct was a match to this fund so exit loop
        End If  'test for fund in list
      Next      'Summarize next fund's totals
    End If      'End of fund in range test
    FrmShowPctComp.ShowPctComp GlAc, NumGLAccts

  Next          'Process next account


  '--Now write the report to file.
  '--Go thru fund list to see if its one we're reporting on
  For Fund = 1 To NumFunds
    ThisFund$ = FundList$(Fund)
    If ThisFund$ >= StartFund$ And ThisFund$ <= EndFund$ Then
      UsingFund = True
      FundRec = FindFund(ThisFund$)
      FundName$ = GetFundTitle(FundRec)
    Else
      UsingFund = False
    End If

    If UsingFund Then
      '--List Assets for this fund
      AcctType$ = "Assets"
      GoSub PrintBSHeader

      For cnt = 1 To ACnt
        Rec = Asset%(cnt)
        Get AcctFileNum, Rec, GLAcct
        FundCode$ = Left$(GLAcct.Num, GLFundLen)

        If FundCode$ = ThisFund$ Then
          ToPrint$ = Space$(80)
          LSet ToPrint$ = QPTrim$(GLAcct.Num)
          Mid$(ToPrint$, 17) = QPTrim$(GLAcct.Title)
          If GLAcct.Bal >= 0 Then
            Debit$ = Using$(CommaFmt$, Str$(GLAcct.Bal))
            Credit$ = " "
          Else
            Debit$ = " "
            'Credit$ = "(" + Using$(CommaFmt$, Str$(GLAcct.Bal)) + ")"
            Credit$ = "(" + Using$(CommaFmt$, Str$(Abs(GLAcct.Bal))) + ")"
          End If
          Mid$(ToPrint$, 64) = Debit$
          Mid$(ToPrint$, 63) = Credit$
          Print #PRNFile, ToPrint$

          Linecnt = Linecnt + 1
          If Linecnt >= MaxLines Then
            GoSub BreakBSPage
            GoSub PrintBSHeader
          End If

        End If

      Next      'AssetAcct

      ToPrint$ = Space$(80)
      If FundTotAssets#(Fund) >= 0 Then
        FundAsset# = FundTotAssets#(Fund)
        DrCr$ = " " + Using$(TotalFmt$, Str$(FundAsset#)) + " "
      Else
        FundAsset# = FundTotAssets#(Fund)
        DrCr$ = "(" + Using$(TotalFmt$, Str$(Abs(FundAsset#))) + ")"
        'FundAsset# = Abs(FundTotAssets#(Fund))
      End If

      Print #PRNFile, DivLine$
      LSet ToPrint$ = "Total Assets"
      Mid$(ToPrint$, 61) = DrCr$
      Print #PRNFile, ToPrint$

      '--List Liabilities
      AcctType$ = "Liabilities"
      GoSub BreakBSPage
      GoSub PrintBSHeader

      For cnt = 1 To LCnt
        Rec = Liab%(cnt)
        Get AcctFileNum, Rec, GLAcct

        FundCode$ = Left$(GLAcct.Num, GLFundLen)
        If FundCode$ = ThisFund$ Then
          ToPrint$ = Space$(80)
          LSet ToPrint$ = QPTrim$(GLAcct.Num)
          Mid$(ToPrint$, 17) = QPTrim$(GLAcct.Title)
          If GLAcct.Bal >= 0 Then
            Credit$ = Using$(CommaFmt$, Str$(GLAcct.Bal))
            Debit$ = ""
          Else
            Credit$ = ""
            'Debit$ = "(" + Using$(CommaFmt$, Str$(GLAcct.Bal)) + ")"
            Debit$ = "(" + Using$(CommaFmt$, Str$(Abs(GLAcct.Bal))) + ")"
          End If

          Mid$(ToPrint$, 64) = Debit$
          Mid$(ToPrint$, 65) = Credit$
          Print #PRNFile, ToPrint$
          Linecnt = Linecnt + 1
          If Linecnt >= MaxLines Then
            GoSub BreakBSPage
            GoSub PrintBSHeader
          End If

        End If

      Next      'LiabAcct

      FundBalAdj# = Round#(FundTotRev#(Fund) - FundTotExp#(Fund))
      If FundBalAdj# >= 0 Then
        BalAdj$ = " " + Using$(CommaFmt$, Str$(FundBalAdj#)) + " "
      Else
        BalAdj$ = "(" + Using$(CommaFmt$, Str$(Abs(FundBalAdj#))) + ")"
      End If
      LSet ToPrint$ = "Current Fund Rev Over/Under Exp"
      Mid$(ToPrint$, 64) = BalAdj$
      Print #PRNFile, ToPrint$

      ToPrint$ = Space$(80)
      If FundTotLiab#(Fund) >= 0 Then
        FundLiab# = FundTotLiab#(Fund)
        DrCr$ = "Cr"
      Else
        'Put this in to replace the one below.
        FundLiab# = FundTotLiab#(Fund)
        'The statement below had Abs and This caused Neg Num not to calc correctly.
        'FundLiab# = Abs(FundTotLiab#(Fund))
        DrCr$ = "Dr"
      End If
      Print #PRNFile, DivLine$

      TotLiabnCap# = FundLiab# + FundBalAdj#
      If TotLiabnCap# >= 0 Then
        DrCr$ = " " + Using$(TotalFmt$, Str$(TotLiabnCap#)) + " "
      Else
        DrCr$ = "(" + Using$(TotalFmt$, Str$(Abs(TotLiabnCap#))) + ")"
      End If
      LSet ToPrint$ = "Total Liabilities & Fund Balance"
      Mid$(ToPrint$, 62) = DrCr$
      Print #PRNFile, ToPrint$

      Print #PRNFile, PageBreak$

      '--Now summarize the fund
      'ToPrint$ = SPACE$(80)
      'LSET ToPrint$ = "Fund Summary:"
      'PRINT #PrnFile, ToPrint$
      'ToPrint$ = SPACE$(80)

      'LSET ToPrint$ = "Revenues"
      'MID$(ToPrint$, 17) = FUsing$(STR$(FundTotRev#(Fund)), CommaFmt$)
      'PRINT #PrnFile, ToPrint$
      'ToPrint$ = SPACE$(80)

      'LSET ToPrint$ = "Expenditures"
      'MID$(ToPrint$, 17) = FUsing$(STR$(FundTotExp#(Fund)), CommaFmt$)
      'PRINT #PrnFile, ToPrint$

      'FundBalAdj# = Round#(FundTotRev#(Fund) - FundTotExp#(Fund))
      'IF FundBalAdj# < 0 THEN
      '   DrCr$ = " Debit"
      'ELSE
      '   DrCr$ = " Credit"
      'END IF

      'ToPrint$ = SPACE$(80)
      'LSET ToPrint$ = "Fund Bal Adj"
      'MID$(ToPrint$, 17) = FUsing$(STR$(FundBalAdj#), CommaFmt$) + DrCr$
      'PRINT #PrnFile, ToPrint$

    End If      'using fund test
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdExit.Enabled = True
      Me.cmdPrint.Enabled = True
      Me.mnuOptions.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

  Next          'FundList$()

  Me.cmdExit.Enabled = True
  Me.cmdPrint.Enabled = True
  EnableCloseButton Me.hwnd, True
  Me.mnuOptions.Enabled = True

  '--This summarizes all funds--
  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Assets"
  'MID$(ToPrint$, 17) = STR$(TotAsset#)
  'MID$(ToPrint$, 37) = STR$(ACnt)
  'PRINT #PrnFile, ToPrint$

  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Liabilities"
  'MID$(ToPrint$, 17) = STR$(TotLiab#)
  'MID$(ToPrint$, 37) = STR$(LCnt)
  'PRINT #PrnFile, ToPrint$
  '
  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Revenues"
  'MID$(ToPrint$, 17) = STR$(TotRev#)
  'PRINT #PrnFile, ToPrint$
  '
  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Expenditures"
  'MID$(ToPrint$, 17) = STR$(TotExp#)
  'PRINT #PrnFile, ToPrint$
  'FundBalAdj# = TotRev# - TotExp#
  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Fund Bal Adj"
  'MID$(ToPrint$, 17) = STR$(FundBalAdj#)
  'PRINT #PrnFile, ToPrint$
  '^--All funds summary --

  'PRINT #PrnFile, PageBreak$

  Close

  'End Report Processing
  '===========================================================================
  'Start Report Printing
  ViewPrint ReportFile$, Header$
  KillFile ReportFile$

  Exit Sub

PrintBSHeader:
  PageNum = PageNum + 1
  Print #PRNFile, GLUserName; Tab(43); "Run Date: " + Date$; "       Page: "; PageNum
  Print #PRNFile, QPTrim$(FundName$) + " " + RptTitle$
  Print #PRNFile, "Period Ending: " + txtDate
  Print #PRNFile,
  Print #PRNFile, AcctType$
  Print #PRNFile, Desc$(1)
  Print #PRNFile, DivLine$
  Linecnt = 7
  Return

BreakBSPage:
  Print #PRNFile, PageBreak$
  Linecnt = 0
  Return
CancelExit:
  Exit Sub
End Sub

