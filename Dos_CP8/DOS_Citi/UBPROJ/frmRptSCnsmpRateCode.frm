VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptSCnsmpRateCode 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consumption By Rate Code Report"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmRptSCnsmpRateCode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5340
      TabIndex        =   3
      Top             =   4440
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
      _ExtentY        =   614
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      ColDesigner     =   "frmRptSCnsmpRateCode.frx":08CA
   End
   Begin LpLib.fpCombo fpComboRates 
      Height          =   348
      Left            =   5340
      TabIndex        =   2
      Top             =   3924
      Width           =   3924
      _Version        =   196608
      _ExtentX        =   6921
      _ExtentY        =   614
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      AutoSearch      =   0
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
      AutoMenu        =   0   'False
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmRptSCnsmpRateCode.frx":0BF8
   End
   Begin VB.CommandButton cmdPrint 
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
      Left            =   8160
      TabIndex        =   10
      Top             =   7104
      Width           =   1332
   End
   Begin VB.CommandButton cmdExit 
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
      Left            =   9840
      TabIndex        =   9
      Top             =   7104
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   8280
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
            TextSave        =   "3:12 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "6/17/2003"
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
   Begin EditLib.fpDateTime txtDate2 
      Height          =   348
      Left            =   5334
      TabIndex        =   1
      Top             =   3400
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
   Begin EditLib.fpDateTime txtDate1 
      Height          =   348
      Left            =   5334
      TabIndex        =   0
      Top             =   2880
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Code:"
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
      Index           =   1
      Left            =   3882
      TabIndex        =   11
      Top             =   3968
      Width           =   1332
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
      Height          =   324
      Index           =   0
      Left            =   3642
      TabIndex        =   8
      Top             =   3496
      Width           =   1572
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Date:"
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
      Left            =   3546
      TabIndex        =   7
      Top             =   2928
      Width           =   1668
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2652
      Left            =   2514
      Top             =   2496
      Width           =   7164
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
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   2874
      TabIndex        =   6
      Top             =   4488
      Width           =   2340
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3210
      Top             =   840
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Consumption By Rate Code"
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
      Left            =   3228
      TabIndex        =   5
      Top             =   1080
      Width           =   5772
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3210
      Top             =   720
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
Attribute VB_Name = "frmRptSCnsmpRateCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim RateRec As Integer
Dim rptopt As Integer
Public Sub GetRptInfo(RptInfo As Integer)
'Added this to allow for Irrigation Report But Removed from menu as
'an option, leave this just in case use for other report?
  rptopt = RptInfo
End Sub
Private Sub cmdExit_Click()
  frmUBStatReportsMenu.Show
  Unload frmRptSCnsmpRateCode
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        'ClearInUse PWcnt
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

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TownName$
    Dim RCode As String
  Dim Handle As Integer, cnt As Integer
  Dim UBRateTblRecLen As Integer, NumOfRateRecs As Integer
  ReDim UBRateTblRec(1) As UBRateTblRecType
  RCode$ = Space$(10)
  UBRateTblRecLen = Len(UBRateTblRec(1))
  NumOfRateRecs = GetNumRateRecs
  Handle = FreeFile
  Open UBPath$ + "UBRATE.DAT" For Random Shared As Handle Len = UBRateTblRecLen
  For cnt = 1 To NumOfRateRecs
    Get Handle, cnt, UBRateTblRec(1)
    LSet RCode$ = QPTrim$(UBRateTblRec(1).RATECODE)
    fpComboRates.AddItem RCode$ + QPTrim$(UBRateTblRec(1).RATEDESC)
  Next
  Close
  fpComboRates.ListIndex = 0
  txtDate1.Text = Format(Now, "mm/dd/yyyy")
  txtDate2.Text = Format(Now, "mm/dd/yyyy")
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

Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDate2.SetFocus
  End If
End Sub

Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpComboRates.SetFocus
  End If
End Sub


Private Function ValidDate()
  Dim TempDate1 As Integer, TempDate2 As Integer
  If CheckValDate(txtDate1) = False And CheckValDate(txtDate2) = False Then
    MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    ValidDate = False
  Else
    TempDate1 = DateDiff("d", "12/31/1979", txtDate1)
    TempDate2 = DateDiff("d", "12/31/1979", txtDate2)
    If TempDate1 > TempDate2 Then
      ValidDate = False
      MsgBox "The Starting And Ending Dates Must Be In Chronological Order Or Equal", vbOKOnly, "Invalid Date"
    Else
      ValidDate = True
    End If
  End If
End Function

Private Sub cmdPrint_Click()
'RptOpt 1 is Consumption by rate code
'2 is Irrigation Consumption Report
'
  RateRec = fpComboRates.ListIndex + 1
  If ValidDate Then
    DeActivateControls Me, True
    If fpcboRptType.ListIndex = 0 Then
      'do the graphics
      If rptopt = 1 Then
        ConsumpUnitStep2
      Else
        'do other one
        'no other one now
      End If
    ElseIf fpcboRptType.ListIndex = 1 Then
      'do the text
      If rptopt = 1 Then
        ConsumpUnitStep
      Else
      'not using the irrigation report
       ' IrrConSteps
      End If
    End If
    ActivateControls Me, True
  End If
End Sub
Private Sub fpComboRates_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpComboRates.ListDown = True
  End If
  If fpComboRates.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtDate2.SetFocus
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
        fpComboRates.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub ConsumpUnitStep()
  Dim Dash80 As String, IdxName As String, IdxRecLen As Integer
  Dim TblBreak(0 To 10) As Long
  Dim TblUnitVal(0 To 10) As Double
  Dim TotalConsp(0 To 10) As Double
  Dim TotalCust(0 To 10) As Long
  Dim UBCustRecLen As Integer, UBCust As Integer
  Dim IdxFileSize As Long, IdxNumOfRecs As Integer, cnt As Long
  Dim UBTransRecLen As Integer, UBSetupreclen As Integer
  Dim Handle As Integer, UBTrans As Integer, NumOfRecs As Long
  Dim UBRateTblRecLen As Integer, NumOfRates As Integer
  Dim UBRpt As Integer, UBSetUp As Integer, ValidCustomer As Integer
  Dim BegDate As Integer, EndDate As Integer, Snt As Integer
  Dim TownLen As Integer, TabStop As Integer, MeterConsp As Long
  Dim UBSetupLen As Integer, RateFile As Integer, MT As Integer
  Dim RATECODE As String, Greater As Boolean, MaxMeterAmt As Long
  Dim MINAMT As Double, Tnt As Integer, MaxStep As Integer
  Dim CustomerRecord  As Integer, MCnt As Integer, GTMeterConsp As Double
  Dim Multi As Long, Cubic As Boolean, ChkMtr As Boolean
  Dim MTRType As String, MType As String, TMeterConsp As Double
  Dim NonUpdated As Integer, LL As Integer, BigTotal As Double
  Dim ReportFile As String
  MaxLines = 56
  Dash80$ = String$(80, "-")

  NumOfRates = GetNumRateRecs%
  ReDim UBRateTbls(1 To NumOfRates) As UBRateTblRecType

  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  FrmShowPctComp.Label1 = "Creating Consumption Report"
  FrmShowPctComp.Show
  UBRateTblRecLen = Len(UBRateTbls(1))
  RateFile = FreeFile
  Open UBPath$ + "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
  
      Get RateFile, RateRec, UBRateTbls(1)
      RATECODE$ = QPTrim$(UBRateTbls(1).RATECODE)
      MINAMT# = UBRateTbls(1).MINAMT
      For Tnt = 1 To 10
        If UBRateTbls(1).TblBreaks(Tnt).UNITS >= 0 Then
          If UBRateTbls(1).TblBreaks(Tnt).UNITS > 0 Then
            Greater = True
          End If
          If (UBRateTbls(1).TblBreaks(Tnt).UNITS = 0) And (Greater = True) Then
            MaxStep = Tnt
            TblBreak&(Tnt) = 99999999
            TblUnitVal#(Tnt) = UBRateTbls(1).TblBreaks(Tnt - 1).UNITAMT
            Exit For
          End If
          TblBreak&(Tnt) = UBRateTbls(1).TblBreaks(Tnt).UNITS
          TblUnitVal#(Tnt) = UBRateTbls(1).TblBreaks(Tnt).UNITAMT
          MaxStep = Tnt
        Else
          MaxStep = Tnt
          TblBreak&(Tnt) = 99999999
          TblUnitVal#(Tnt) = UBRateTbls(1).TblBreaks(Tnt - 1).UNITAMT
          Exit For
        End If
      Next Tnt
  Close
  BegDate = Date2Num%(txtDate1)
  EndDate = Date2Num%(txtDate2)
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize(BookIndexFile)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

  'IndexName$ = BookIndexFile
  'NumOfRecs = FileSize(IndexName$) \ 4
  'ReDim IndexArray(1 To NumOfRecs) As UBCustIndexRecType
  'FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  Handle = FreeFile
  Open BookIndexFile For Random Shared As Handle Len = IdxRecLen
  For cnt = 1 To IdxNumOfRecs
    Get #Handle, cnt, IdxBuff(cnt)
  Next
  Close Handle

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  ReDim UBSetUpRec(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUpRec(1))

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen

  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  NumOfRecs& = LOF(UBTrans) / UBTransRecLen
  ReportFile$ = UBPath$ + "UBBKCNSP.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  Rem Report Goes Here
  UBSetUp = FreeFile
  Open UBPath$ + "UBSETUP.DAT" For Random Access Read Write Shared As UBSetUp Len = UBSetupreclen
  If LOF(UBSetUp) / UBSetupreclen = 0 Then
    TownName$ = "Undefined"
  Else
    Get UBSetUp, 1, UBSetUpRec(1)
    TownName$ = UBSetUpRec(1).UTILNAME
    TownLen = Len(RTrim$(TownName$))
    TabStop = 40 - (TownLen / 2)
    If TabStop < 1 Then TabStop = 1
  End If
  Close UBSetUp

  GoSub DoUnitStepHeader

  For cnt& = 1 To NumOfRecs&
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs&
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitConsStep
    End If

    Get UBTrans, cnt&, UBTransRec(1)
    If UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate Then
      If UBTransRec(1).TransType = 1 Or UBTransRec(1).TransType = 101 Then
        'Set Valid Flag to Off to Check Each Customer For Valid Rate Code
        ValidCustomer = 0
        If Linecnt > MaxLines Then
          Print #UBRpt, Chr$(12)
          GoSub DoUnitStepHeader
        End If
        'IF Valid Customer Then Get the Customer and Check to See if Rate Appl
        CustomerRecord = UBTransRec(1).CustAcctNo

        If CustomerRecord > 0 Then
          Get UBCust, CustomerRecord, UBCustRec(1)
            For Snt = 1 To 15
              If QPTrim$(UBCustRec(1).Serv(Snt).RATECODE) = RATECODE$ Then
                MTRType$ = UBCustRec(1).Serv(Snt).RMtrType
                Select Case MTRType$
                Case "W"
                  MT = 1
                Case "S"
                  MT = 2
                Case "C"
                  MT = 3
                Case "D"
                  MT = 5
                Case Else
                  MT = 4
                End Select
                ValidCustomer = 1
                Exit For
              End If
            Next Snt
        End If
        'IF Customer Has a Rate Code Match then Analyze Meter Type to Get Usag
        If ValidCustomer = 1 Then
          Multi& = 0
          Cubic = False
          For MCnt = 1 To 7
            MType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)
            If MTRType$ = MType$ Then
            'IF LEN(MType$) > 0 THEN
              Multi& = UBCustRec(1).LocMeters(MCnt).MTRMulti
              If Multi& = 0 Then
                Multi& = 1
              End If
              If UBCustRec(1).LocMeters(MCnt).MTRUnit = "C" Then
                Cubic = True
              End If
              Exit For
            End If
          Next
          If Multi& <= 0 Then Multi& = 1
          'IF WhatRev > 0 THEN
            ChkMtr = True
          'ELSE
          '  ChkMtr = False
          'END IF
          For MCnt = 1 To 7
            If ChkMtr = True Then
              If UBTransRec(1).MtrTypes(MCnt) <> MT Then
                GoTo SkipThisMtr
              End If
            End If
            MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
            If MeterConsp& < 0 Then
              MaxMeterAmt& = 10& ^ (Len(Str$(UBTransRec(1).PrevRead(MCnt))) - 1)
              MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
            End If
            If Cubic Then
              MeterConsp& = MeterConsp& * 7.481
            End If
            MeterConsp& = MeterConsp& * Multi&
            'IF MeterConsp& = 1 THEN STOP
            TMeterConsp# = TMeterConsp# + MeterConsp&
            GTMeterConsp# = GTMeterConsp# + MeterConsp&
            'IF MeterConsp& >= 13000 THEN
            'LPRINT CustomerRecord
            'STOP
            'END IF

            MeterConsp& = 0
            'END IF
SkipThisMtr:
          Next MCnt
        End If
        If (TMeterConsp# >= 0) And (ValidCustomer = 1) Then

          NonUpdated = 1        'Set Flag to Let Me Know When this Cust Cons U
          For LL = 1 To MaxStep
            If TMeterConsp# >= TblBreak&(LL - 1) And TMeterConsp# <= TblBreak&(LL) Then
              TotalConsp#(LL) = TotalConsp#(LL) + TMeterConsp#
              TotalCust(LL) = TotalCust(LL) + 1
              NonUpdated = 0
              Exit For
            End If
          Next LL
          If NonUpdated = 1 Then
            TotalConsp#(MaxStep) = TotalConsp#(MaxStep) + TMeterConsp#
            TotalCust(MaxStep) = TotalCust(MaxStep) + 1
          End If
        End If
        TMeterConsp# = 0
      End If
    End If

  Next

  GoSub DoUnitStepFooter:

  Close

  Erase TblBreak&, TotalConsp#, TotalCust

'  If Not AbortFlag Then
'    PrintRptFile , "UBBKCNSP.RPT", 1, RetCode, EntryPoint
'  End If
  ViewPrint ReportFile$, "Consumption by RateCode"
  'KillFile "UBBKCNSP.RPT"
  Exit Sub

DoUnitStepHeader:
  PageNo = PageNo + 1
  Print #UBRpt, Tab(29); "Consumption by RateCode"; Tab(70); "Page #"; PageNo
  Print #UBRpt, TownName$
  Print #UBRpt, "Report Date: "; Now
  Print #UBRpt, " "
  Print #UBRpt, "    For Rate Code: "; fpComboRates.Text
  Print #UBRpt, " Period Beginning: "; txtDate1
  Print #UBRpt, "    Period Ending: "; txtDate2
  Print #UBRpt, " "
  Print #UBRpt, Dash80$
  Linecnt = 6
Return

DoUnitStepFooter:
  TblBreak&(MaxStep) = 99999999
  'TblBreak&(MaxStep + 1) = 99999999
  For LL = 1 To MaxStep
    Print #UBRpt, "Step # "; LL;
    Print #UBRpt, Tab(12); "From "; TblBreak&(LL - 1); " to "; TblBreak&(LL)
    Print #UBRpt, "Consumption = "; Using("#########,#", TotalConsp#(LL));
    Print #UBRpt, "  # of Cust = "; Using("#####,#", TotalCust(LL));
    BigTotal# = Round#(BigTotal# + (Round#(TotalConsp#(LL) * TblUnitVal#(LL))))
    If TotalCust(LL) > 0 Then
      Print #UBRpt, Using("  #####,#.##", Round#(TotalConsp#(LL) * TblUnitVal#(LL))); Using("  #####,#.##", Round#(MINAMT# * TotalCust(LL)))
      'PRINT #UBRpt, "  Avg Use= "; USING "#####,#.##"; TotalConsp#(LL) / Tota
    Else
      Print #UBRpt, ""
    End If
    Print #UBRpt, Dash80$
  Next LL
  Print #UBRpt, "Grand Total Consumption:"; Using("############", GTMeterConsp)
  'PRINT #UBRpt, "Grand Total Consumption:"; USING "#######,#.##"; BigTotal#
  Print #UBRpt, Chr$(12);
Return
  GoTo ExitConsStep
ExitConsStep:
  Close
Exit Sub
End Sub
Private Sub ConsumpUnitStep2()
  Dim IdxName As String, IdxRecLen As Integer, ToPrint As String
  Dim TblBreak(0 To 10) As Long
  Dim TblUnitVal(0 To 10) As Double
  Dim TotalConsp(0 To 10) As Double
  Dim TotalCust(0 To 10) As Long
  Dim UBCustRecLen As Integer, UBCust As Integer
  Dim IdxFileSize As Long, IdxNumOfRecs As Integer, cnt As Long
  Dim UBTransRecLen As Integer, UBSetupreclen As Integer
  Dim Handle As Integer, UBTrans As Integer, NumOfRecs As Long
  Dim UBRateTblRecLen As Integer, NumOfRates As Integer
  Dim UBRpt As Integer, UBSetUp As Integer, ValidCustomer As Integer
  Dim BegDate As Integer, EndDate As Integer, Snt As Integer
  Dim TownLen As Integer, TabStop As Integer, MeterConsp As Long
  Dim UBSetupLen As Integer, RateFile As Integer, MT As Integer
  Dim RATECODE As String, Greater As Boolean, MaxMeterAmt As Long
  Dim MINAMT As Double, Tnt As Integer, MaxStep As Integer
  Dim CustomerRecord  As Integer, MCnt As Integer, GTMeterConsp As Double
  Dim Multi As Long, Cubic As Boolean, ChkMtr As Boolean
  Dim MTRType As String, MType As String, TMeterConsp As Double
  Dim NonUpdated As Integer, LL As Integer, BigTotal As Double
  Dim ReportFile As String
  NumOfRates = GetNumRateRecs%
  ReDim UBRateTbls(1 To NumOfRates) As UBRateTblRecType

  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  FrmShowPctComp.Label1 = "Creating Consumption Report"
  FrmShowPctComp.Show
  UBRateTblRecLen = Len(UBRateTbls(1))
  RateFile = FreeFile
  Open UBPath$ = "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
  
      Get RateFile, RateRec, UBRateTbls(1)
      RATECODE$ = QPTrim$(UBRateTbls(1).RATECODE)
      MINAMT# = UBRateTbls(1).MINAMT
      For Tnt = 1 To 10
        If UBRateTbls(1).TblBreaks(Tnt).UNITS >= 0 Then
          If UBRateTbls(1).TblBreaks(Tnt).UNITS > 0 Then
            Greater = True
          End If
          If (UBRateTbls(1).TblBreaks(Tnt).UNITS = 0) And (Greater = True) Then
            MaxStep = Tnt
            TblBreak&(Tnt) = 99999999
            TblUnitVal#(Tnt) = UBRateTbls(1).TblBreaks(Tnt - 1).UNITAMT
            Exit For
          End If
          TblBreak&(Tnt) = UBRateTbls(1).TblBreaks(Tnt).UNITS
          TblUnitVal#(Tnt) = UBRateTbls(1).TblBreaks(Tnt).UNITAMT
          MaxStep = Tnt
        Else
          MaxStep = Tnt
          TblBreak&(Tnt) = 99999999
          TblUnitVal#(Tnt) = UBRateTbls(1).TblBreaks(Tnt - 1).UNITAMT
          Exit For
        End If
      Next Tnt
  Close
  BegDate = Date2Num%(txtDate1)
  EndDate = Date2Num%(txtDate2)
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize(BookIndexFile)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

  'IndexName$ = BookIndexFile
  'NumOfRecs = FileSize(IndexName$) \ 4
  'ReDim IndexArray(1 To NumOfRecs) As UBCustIndexRecType
  'FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
  Handle = FreeFile
  Open BookIndexFile For Random Shared As Handle Len = IdxRecLen
  For cnt = 1 To IdxNumOfRecs
    Get #Handle, cnt, IdxBuff(cnt)
  Next
  Close Handle

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))
  ReDim UBSetUpRec(1) As UBSetupRecType
  UBSetupreclen = Len(UBSetUpRec(1))

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen

  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
  NumOfRecs& = LOF(UBTrans) / UBTransRecLen
  ReportFile$ = UBPath$ + "UBBKCNSP.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt
  Rem Report Goes Here
'  UBSetUp = FreeFile
'  Open "UBSETUP.DAT" For Random Access Read Write Shared As UBSetUp Len = UBSetupreclen
'  If LOF(UBSetUp) / UBSetupreclen = 0 Then
'    TownName$ = "Undefined"
'  Else
'    Get UBSetUp, 1, UBSetUpRec(1)
'    TownName$ = UBSetUpRec(1).UTILNAME
'    TownLen = Len(RTrim$(TownName$))
'    TabStop = 40 - (TownLen / 2)
'    If TabStop < 1 Then TabStop = 1
'  End If
'  Close UBSetUp
'
'  GoSub DoUnitStepHeader

  For cnt& = 1 To NumOfRecs&
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs&
    If FrmShowPctComp.Out Then
      Close
      Unload FrmShowPctComp
      GoTo ExitConsStep
    End If

    Get UBTrans, cnt&, UBTransRec(1)
    If UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate Then
      If UBTransRec(1).TransType = 1 Or UBTransRec(1).TransType = 101 Then
        'Set Valid Flag to Off to Check Each Customer For Valid Rate Code
        ValidCustomer = 0
'        If Linecnt > MaxLines Then
'          Print #UBRpt, Chr$(12)
'          GoSub DoUnitStepHeader
'        End If
        'IF Valid Customer Then Get the Customer and Check to See if Rate Appl
        CustomerRecord = UBTransRec(1).CustAcctNo

        If CustomerRecord > 0 Then
          Get UBCust, CustomerRecord, UBCustRec(1)
            For Snt = 1 To 15
              If QPTrim$(UBCustRec(1).Serv(Snt).RATECODE) = RATECODE$ Then
                MTRType$ = UBCustRec(1).Serv(Snt).RMtrType
                Select Case MTRType$
                Case "W"
                  MT = 1
                Case "S"
                  MT = 2
                Case "C"
                  MT = 3
                Case "D"
                  MT = 5
                Case Else
                  MT = 4
                End Select
                ValidCustomer = 1
                Exit For
              End If
            Next Snt
        End If
        'IF Customer Has a Rate Code Match then Analyze Meter Type to Get Usag
        If ValidCustomer = 1 Then
          Multi& = 0
          Cubic = False
          For MCnt = 1 To 7
            MType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)
            If MTRType$ = MType$ Then
            'IF LEN(MType$) > 0 THEN
              Multi& = UBCustRec(1).LocMeters(MCnt).MTRMulti
              If Multi& = 0 Then
                Multi& = 1
              End If
              If UBCustRec(1).LocMeters(MCnt).MTRUnit = "C" Then
                Cubic = True
              End If
              Exit For
            End If
          Next
          If Multi& <= 0 Then Multi& = 1
          'IF WhatRev > 0 THEN
            ChkMtr = True
          'ELSE
          '  ChkMtr = False
          'END IF
          For MCnt = 1 To 7
            If ChkMtr = True Then
              If UBTransRec(1).MtrTypes(MCnt) <> MT Then
                GoTo SkipThisMtr
              End If
            End If
            MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
            If MeterConsp& < 0 Then
              MaxMeterAmt& = 10& ^ (Len(Str$(UBTransRec(1).PrevRead(MCnt))) - 1)
              MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
            End If
            If Cubic Then
              MeterConsp& = MeterConsp& * 7.481
            End If
            MeterConsp& = MeterConsp& * Multi&
            'IF MeterConsp& = 1 THEN STOP
            TMeterConsp# = TMeterConsp# + MeterConsp&
            GTMeterConsp# = GTMeterConsp# + MeterConsp&
            'IF MeterConsp& >= 13000 THEN
            'LPRINT CustomerRecord
            'STOP
            'END IF

            MeterConsp& = 0
            'END IF
SkipThisMtr:
          Next MCnt
        End If
        If (TMeterConsp# >= 0) And (ValidCustomer = 1) Then

          NonUpdated = 1        'Set Flag to Let Me Know When this Cust Cons U
          For LL = 1 To MaxStep
            If TMeterConsp# >= TblBreak&(LL - 1) And TMeterConsp# <= TblBreak&(LL) Then
              TotalConsp#(LL) = TotalConsp#(LL) + TMeterConsp#
              TotalCust(LL) = TotalCust(LL) + 1
              NonUpdated = 0
              Exit For
            End If
          Next LL
          If NonUpdated = 1 Then
            TotalConsp#(MaxStep) = TotalConsp#(MaxStep) + TMeterConsp#
            TotalCust(MaxStep) = TotalCust(MaxStep) + 1
          End If
        End If
        TMeterConsp# = 0
      End If
    End If

  Next

  GoSub DoUnitStepFooter:

  Close

  Erase TblBreak&, TotalConsp#, TotalCust
    Load frmLoadingRpt
    ARptSCnsmpRate.txtDate = Now
    ARptSCnsmpRate.txtTown = TownName$
    ARptSCnsmpRate.Title = "Consumption by Rate Code"
    ARptSCnsmpRate.txtRate = fpComboRates.Text
    ARptSCnsmpRate.txtDate1 = txtDate1
    ARptSCnsmpRate.txtDate2 = txtDate2
    ARptSCnsmpRate.totConsump = Using("###,###,###,###", GTMeterConsp)
    ARptSCnsmpRate.GetName ReportFile$
    ARptSCnsmpRate.startrpt


'  If Not AbortFlag Then
'    PrintRptFile , , 1, RetCode, EntryPoint
'  End If
'  ViewPrint "UBBKCNSP.RPT", "Consumption by Rate Code"
  'KillFile "UBBKCNSP.RPT"
  Exit Sub

DoUnitStepHeader:
'  PageNo = PageNo + 1
'  Print #UBRpt, Tab(29); "Consumption by RateCode"; Tab(70); "Page #"; PageNo
'  Print #UBRpt, TownName$
'  Print #UBRpt, "Report Date: "; Now
'  Print #UBRpt, " "
'  Print #UBRpt, "    For Rate Code: ";
'  Print #UBRpt, " Period Beginning: ";
'  Print #UBRpt, "    Period Ending: ";
'  Print #UBRpt, " "
'  Print #UBRpt, Dash80$
'  Linecnt = 6
Return

DoUnitStepFooter:
  TblBreak&(MaxStep) = 99999999
  'TblBreak&(MaxStep + 1) = 99999999"  # of Cust = " +"Consumption = " +"From " +" to " +
  For LL = 1 To MaxStep
    ToPrint$ = "Step # " + Str(LL)
    ToPrint$ = ToPrint$ + "~" + Str(TblBreak&(LL - 1)) + "~" + Str(TblBreak&(LL))
    ToPrint$ = ToPrint$ + "~" + Using("#,###,###,###", TotalConsp#(LL))
    ToPrint$ = ToPrint$ + "~" + Using("###,###", TotalCust(LL))
    BigTotal# = Round#(BigTotal# + (Round#(TotalConsp#(LL) * TblUnitVal#(LL))))
    If TotalCust(LL) > 0 Then
      ToPrint$ = ToPrint$ + "~" + Using("  ###,###.##", Round#(TotalConsp#(LL) * TblUnitVal#(LL))) + "~" + Using("  ###,###.##", Round#(MINAMT# * TotalCust(LL)))
      'PRINT #UBRpt, "  Avg Use= "; USING "#####,#.##"; TotalConsp#(LL) / Tota
    Else
      ToPrint$ = ToPrint$ + "~ ~ "
    End If
    Print #UBRpt, ToPrint$
    ToPrint$ = ""
  Next LL
  'Print #UBRpt, "Grand Total Consumption:"; Using("############", GTMeterConsp)
  'PRINT #UBRpt, "Grand Total Consumption:"; USING "#######,#.##"; BigTotal#
  'Print #UBRpt, Chr$(12);
Return
  GoTo ExitConsStep
ExitConsStep:
  Close
Exit Sub
End Sub
'
'Did not finigh the IrrConSteps because decided to remove option from
'menu - report only used by MOWASA - no one else uses irrigation
'Private Sub IrrConSteps()
'  Dim Dash80 As String, IdxName As String, IdxRecLen As Integer
'  Dim TblBreak(0 To 10) As Long
'  Dim TblUnitVal(0 To 10) As Double
'  Dim TotalConsp(0 To 10) As Double
'  Dim TotalCust(0 To 10) As Integer
'  Dim UBCustRecLen As Integer, UBCust As Integer
'  Dim IdxFileSize As Long, IdxNumOfRecs As Integer, cnt As Long
'  Dim UBTransRecLen As Integer, UBSetupreclen As Integer
'  Dim Handle As Integer, UBTrans As Integer, NumOfRecs As Long
'  Dim UBRateTblRecLen As Integer, NumOfRates As Integer
'  Dim UBRpt As Integer, UBSetUp As Integer, ValidCustomer As Integer
'  Dim BegDate As Integer, EndDate As Integer, Snt As Integer
'  Dim TownLen As Integer, TabStop As Integer, MeterConsp As Long
'  Dim UBSetupLen As Integer, RateFile As Integer, MT As Integer
'  Dim RATECODE As String, Greater As Boolean, MaxMeterAmt As Long
'  Dim MINAMT As Double, Tnt As Integer, MaxStep As Integer
'  Dim CustomerRecord  As Integer, MCnt As Integer, GTMeterConsp As Double
'  Dim Multi As Long, Cubic As Boolean, ChkMtr As Boolean
'  Dim MTRType As String, MType As String, TMeterConsp As Double
'  Dim NonUpdated As Integer, LL As Integer, BigTotal As Double
'
' ' ReDim TblBreak&(10), TotalConsp#(10), TotalCust(10)
'  MaxLines = 56
'  Dash80$ = String$(80, "-")
'  NumOfRates = GetNumRateRecs%
'  ReDim UBRateTbls(1 To NumOfRates) As UBRateTblRecType
'
'
''  RateFile = FreeFile
''  Open "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
''  For cnt = 1 To NumOfRates
''    Get RateFile, cnt, UBRateTbls(cnt)
''    Choice$(Cnt, 0) = UCASE$(UBRateTbls(Cnt).RATECODE + " " + UBRateTbls(Cnt).
''  Next
''  Close
'  FrmShowPctComp.Label1 = "Creating Irrigation Consumption Report"
'  FrmShowPctComp.Show
'
'  ReDim UBSetUpRec(1) As UBSetupRecType
'  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
'
'
'
'  UBRateTblRecLen = Len(UBRateTbls(1))
'
'  RateFile = FreeFile
'  Open "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
''  For Lnt = 1 To NumOfRates
''    WhatRate$ = QPTrim$(Left$(Form$(3, 0), 4))
''    ThisRate$ = QPTrim$(Left$(Choice$(Lnt, 0), 4))
''    If ThisRate$ = WhatRate$ Then
'      Get RateFile, RateRec, UBRateTbls(1)
'      RATECODE$ = QPTrim$(UBRateTbls(1).RATECODE)
'      For Tnt = 1 To 10
'        If UBRateTbls(1).TblBreaks(Tnt).UNITS >= 0 Then
'          If UBRateTbls(1).TblBreaks(Tnt).UNITS > 0 Then
'            Greater = True
'          End If
'          If (UBRateTbls(1).TblBreaks(Tnt).UNITS = 0) And (Greater = True) Then
'            MaxStep = Tnt
'            TblBreak&(Tnt) = 99999999
'            Exit For
'          End If
'          TblBreak&(Tnt) = UBRateTbls(1).TblBreaks(Tnt).UNITS
'          MaxStep = Tnt
'        Else
'          MaxStep = Tnt
'          TblBreak&(Tnt) = 99999999
'          Exit For
'        End If
'      Next Tnt
''    End If
''  Next Lnt
'  Close
'
'  BegDate = Date2Num%(txtDate1)
'  EndDate = Date2Num%(txtDate2)
'  IdxRecLen = 4 'we are using a long integer
'  IdxFileSize& = FileSize(BookIndexFile)
'  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
'
''  IndexName$ = BookIndexFile
''  NumOfRecs = FileSize(IndexName$) \ 4
''  ReDim IndexArray(1 To NumOfRecs) As UBCustIndexRecType
''  FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
'  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
'  Handle = FreeFile
'  Open BookIndexFile For Random Shared As Handle Len = IdxRecLen
'  For cnt = 1 To IdxNumOfRecs
'    Get #Handle, cnt, IdxBuff(cnt)
'  Next
'  Close Handle
'
'  ReDim UBCustRec(1) As NewUBCustRecType
'  UBCustRecLen = Len(UBCustRec(1))
'
'  ReDim UBTransRec(1) As UBTransRecType
'  UBTransRecLen = Len(UBTransRec(1))
'
'  ReDim UBSetUpRec(1) As UBSetupRecType
'  UBSetupreclen = Len(UBSetUpRec(1))
'
'  UBCust = FreeFile
'  Open "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
'
'  UBTrans = FreeFile
'  Open "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen
'  NumOfRecords! = LOF(UBTrans) / UBTransRecLen
'
'  UBRpt = FreeFile
'  Open "UBBKCNSP.RPT" For Output As UBRpt
'
'
'  Rem Report Goes Here
'  UBSetUp = FreeFile
'  Open "UBSETUP.DAT" For Random Access Read Write Shared As UBSetUp Len = UBSetupreclen
'  If LOF(UBSetUp) / UBSetupreclen = 0 Then
'    TownName$ = "Undefined"
'  Else
'    Get UBSetUp, 1, UBSetUpRec(1)
'    TownName$ = UBSetUpRec(1).UTILNAME
'    TownLen = Len(RTrim$(TownName$))
'    TabStop = 40 - (TownLen / 2)
'    If TabStop < 1 Then TabStop = 1
'  End If
'  Close UBSetUp
'
'
'  GoSub IDoUnitStepHeader
'
'  For cnt = 1 To NumOfRecords
'    FrmShowPctComp.ShowPctComp cnt, NumOfRecords&
'    If FrmShowPctComp.Out Then
'      Close
'      Unload FrmShowPctComp
'      GoTo IExitConRpt
'    End If
'
'    Get UBTrans, cnt, UBTransRec(1)
'    If UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate Then
'      If UBTransRec(1).TransType = 1 Or UBTransRec(1).TransType = 101 Then
'        'Set Valid Flag to Off to Check Each Customer For Valid Rate Code
'        ValidCustomer = 0
'        If Linecnt > MaxLines Then
'          Print #UBRpt, Chr$(12)
'          GoSub IDoUnitStepHeader
'        End If
'
'        'IF Valid Customer Then Get the Customer and Check to See if Rate Appl
'        CustomerRecord = UBTransRec(1).CustAcctNo
'        If CustomerRecord > 0 Then
'          Get UBCust, CustomerRecord, UBCustRec(1)
'          CCode$ = QPTrim$(UBCustRec(1).Serv(WhatRev).RATECODE)
'          If CCode$ = RATECODE$ Then
'            MTRType$ = UBCustRec(1).Serv(WhatRev).RMtrType
'            Select Case MTRType$
'            Case "W"
'              MT = 1
'            Case "S"
'              MT = 2
'            Case "C"
'              MT = 3
'            Case Else
'              MT = 4
'            End Select
'            ValidCustomer = 1
'          End If
'        End If
'        'IF Customer Has a Rate Code Match then Analyze Meter Type to Get Usag
'        If ValidCustomer = 1 Then
'          Multi& = 0
'          Cubic = False
'          For MCnt = 1 To 7
'            MType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)
'            If Len(MType$) > 0 Then
'              Multi& = UBCustRec(1).LocMeters(MCnt).MTRMulti
'              If UBCustRec(1).LocMeters(MCnt).MTRUnit = "C" Then
'                Cubic = True
'              End If
'              Exit For
'            End If
'          Next
'          If Multi& <= 0 Then Multi& = 1
'          For MCnt = 1 To 7
'            If UBTransRec(1).MtrTypes(MCnt) <> MT Then
'              GoTo ISkipThisMtr
'            End If
'            MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
'            If MeterConsp& < 0 Then
'              MaxMeterAmt& = 10& ^ (Len(Str$(UBTransRec(1).PrevRead(MCnt))) - 1)
'              MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
'            End If
'            If Cubic Then
'              MeterConsp& = MeterConsp& * 7.481
'            End If
'            MeterConsp& = MeterConsp& * Multi&
'            TMeterConsp# = TMeterConsp# + MeterConsp&
'            GTMeterConsp# = GTMeterConsp# + MeterConsp&
'            MeterConsp& = 0
'ISkipThisMtr:
'          Next MCnt
'        End If
'
'        If (TMeterConsp# >= 0) And (ValidCustomer = 1) Then
'          NonUpdated = 1        'Set Flag to Let Me Know When this Cust Cons U
'          For LL = 1 To MaxStep
'            If TMeterConsp# >= TblBreak&(LL - 1) And TMeterConsp# <= TblBreak&(LL) Then
'              TotalConsp#(LL) = TotalConsp#(LL) + TMeterConsp#
'              TotalCust(LL) = TotalCust(LL) + 1
'              NonUpdated = 0
'              Exit For
'            End If
'          Next LL
'          If NonUpdated = 1 Then
'            TotalConsp#(MaxStep) = TotalConsp#(MaxStep) + TMeterConsp#
'            TotalCust(MaxStep) = TotalCust(MaxStep) + 1
'          End If
'        End If
'        TMeterConsp# = 0
'
'        If AskAbandonPrint% Then
'          AbortFlag = True
'          Exit For
'        End If
'      End If
'    End If
'    If AskAbandonPrint% Then
'      AbortFlag = True
'      Exit For
'    End If
'
'  '  ShowPctCompL CLng(cnt!), CLng(NumOfRecords!)
'
'  Next
'
'  GoSub IDoUnitStepFooter:
'
'  Close
'
'  Erase TblBreak&, TotalConsp#, TotalCust
'
''  If Not AbortFlag Then
''    PrintRptFile "Consumption by Step", , 1, RetCode, EntryPoint
''  End If
'  ViewPrint "UBBKCNSP.RPT", "Irrigation Consumption by RateCode"
'  'KillFile "UBBKCNSP.RPT"
'  Exit Sub
'IDoUnitStepHeader:
'  PageNo = PageNo + 1
'  'Print #UBRpt, Tab(TabStop); TownName$
'  Print #UBRpt, Tab(24); "Irrigation Consumption by RateCode"; Tab(70); "Page#"; PageNo
'  Print #UBRpt, TownName$
'  Print #UBRpt, "Report Date: "; Now
'  Print #UBRpt, " "
'  Print #UBRpt, "    For Rate Code: "; fpComboRates.Text
'  Print #UBRpt, " Period Beginning: "; txtDate1
'  Print #UBRpt, "    Period Ending: "; txtDate2
'  Print #UBRpt, " "
'  Print #UBRpt, Dash80$
'  Linecnt = 6
'  Return
'
'IDoUnitStepFooter:
'  TblBreak&(MaxStep) = 99999999
'  'TblBreak&(MaxStep + 1) = 99999999
'
'  For LL = 1 To MaxStep
'    Print #UBRpt, "Step # "; LL;
'    Print #UBRpt, Tab(12); "From "; TblBreak&(LL - 1); " to "; TblBreak&(LL)
'    Print #UBRpt, "Consumption = "; Using("#,###,###,###", TotalConsp#(LL));
'    Print #UBRpt, "  # of Cust = "; Using("###,###", TotalCust(LL));
'    If TotalCust(LL) > 0 Then
'      Print #UBRpt, "  Avg Use= "; Using("###,###.##", TotalConsp#(LL) / TotalCust(LL))
'    Else
'      Print #UBRpt, ""
'    End If
'    Print #UBRpt, Dash80$
'  Next LL
'  Print #UBRpt, "Grand Total Consumption:"; Using("###,###,###,###", GTMeterConsp#)
'  Print #UBRpt, Chr$(12);
'  Return
'  GoTo IExitConRpt
'
'
'IExitConRpt:
'  Close
'End Sub
