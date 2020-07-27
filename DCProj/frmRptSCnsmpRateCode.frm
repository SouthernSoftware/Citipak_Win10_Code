VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
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
      Text            =   ""
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
      Text            =   ""
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
      ColDesigner     =   "frmRptSCnsmpRateCode.frx":0CA0
   End
   Begin VB.CheckBox PageBrk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008F8265&
      Caption         =   "Page Break on Rates:"
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
      Left            =   2880
      TabIndex        =   4
      Top             =   4944
      Width           =   2652
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
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   7104
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
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
            TextSave        =   "3:43 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "5/10/2005"
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
      TabIndex        =   12
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   2928
      Width           =   1668
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   3108
      Left            =   2520
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
      TabIndex        =   7
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
      TabIndex        =   6
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
Dim rptopt As Integer, Doall As Boolean
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
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via RptSCnsmpRateCode by " + PWUser$
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
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
    Dim RCode As String
  Dim Handle As Integer, cnt As Integer
  Dim UBRateTblRecLen As Integer, NumOfRateRecs As Integer
  ReDim UBRateTblRec(1) As UBRateTblRecType
  RCode$ = Space$(10)
  UBRateTblRecLen = Len(UBRateTblRec(1))
  NumOfRateRecs = GetNumRateRecs
  Handle = FreeFile
  fpComboRates.AddItem "ALL" + "-Print All Rates"
  Open UBPath$ + "UBRATE.DAT" For Random Shared As Handle Len = UBRateTblRecLen
  For cnt = 1 To NumOfRateRecs
    Get Handle, cnt, UBRateTblRec(1)
    LSet RCode$ = QPTrim$(UBRateTblRec(1).Ratecode)
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
  If CheckValDate(txtDate1) = False Then
    MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    ValidDate = False
  ElseIf CheckValDate(txtDate2) = False Then
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
  RateRec = fpComboRates.ListIndex
  If RateRec = 0 Then
    Doall = True
  Else
    Doall = False
  End If
  If ValidDate Then
    DeActivateControls Me, True
    If fpcboRptType.ListIndex = 0 Then
      'do the graphics
      
        ConsumpUnitStep 0
     
    ElseIf fpcboRptType.ListIndex = 1 Then
      'do the text
        ConsumpUnitStep 1
      ActivateControls Me, True
    Else
      ActivateControls Me, True
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
Private Sub ConsumpUnitStep(RptType As Integer)
  Dim IdxName As String, IdxRecLen As Integer, ToPrint As String
  Dim UBCustRecLen As Integer, UBCust As Integer, RCnt As Integer
  Dim IdxFileSize As Long, IdxNumOfRecs As Long, cnt As Long
  Dim UBTransRecLen As Integer, UBSetupreclen As Integer
  Dim Handle As Integer, UBTrans As Integer, NumofRecs As Long
  Dim UBRateTblRecLen As Integer, NumOfRates As Integer
  Dim UBRpt As Integer, UBSetUp As Integer, ValidCustomer As Integer
  Dim BegDate As Integer, EndDate As Integer, Snt As Integer
  Dim TownLen As Integer, TabStop As Integer, MeterConsp As Long
  Dim UBSetupLen As Integer, RateFile As Integer, MT As Integer
  Dim Greater As Boolean, MaxMeterAmt As Long, RCode As String
  Dim Tnt As Integer, NMinAMT As Double, ToPrintI As String
  Dim CustomerRecord As Long, MCnt As Integer, GTMeterConsp As Double
  Dim Multi As Long, Cubic As Boolean, ChkMtr As Boolean, NTAmt As Double
  Dim MtrType As String, MType As String, TMeterConsp As Double
  Dim NonUpdated As Integer, LL As Integer, BigUTotal As Double
  Dim ReportFile As String, MinGT As Double, GBBigUTotal As Double
  Dim GBMinGT As Double, GBGTMeterConsp As Double, BigTotCust As Long
  Dim GBCustTot As Long, RptInfo As String, Tempcalccnsp As Double
  Dim NewMtrConsp As Double, UNITS As Long, UntPrc As Double
  Dim MinBillAmt As Double, TAmt As Double, MaxFlag As Boolean
  Dim Dash80 As String, RptInfo2 As String
  MaxLines = 56
  Dash80$ = String$(80, "-")

  NumOfRates = GetNumRateRecs%
  ReDim UBRateTbls(1 To NumOfRates) As UBRateTblRecType
  ReDim MINAMT(1 To NumOfRates) As Double
  ReDim minunt(1 To NumOfRates) As Double
  ReDim MaxAmt(1 To NumOfRates) As Double
  ReDim Ratecode(1 To NumOfRates) As String
  ReDim MaxStep(1 To NumOfRates) As Integer
  ReDim TblBreak(1 To NumOfRates, 11) As Long
  ReDim TblUnitVal(1 To NumOfRates, 11) As Double
  ReDim TblBreakfr(1 To NumOfRates, 11) As Long
  ReDim TotalConsp(1 To NumOfRates, 11) As Double
  ReDim tblbrkchg(1 To NumOfRates, 11) As Double
  ReDim TotalCust(1 To NumOfRates) As Long
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  FrmShowPctComp.Label1 = "Creating Consumption Report"
  FrmShowPctComp.Show
  
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
  NumofRecs& = LOF(UBTrans) / UBTransRecLen
  ReportFile$ = UBPath$ + "UBBKCNSP.RPT"
  UBRpt = FreeFile
  Open ReportFile$ For Output As UBRpt

  If Not Doall Then
    UBRateTblRecLen = Len(UBRateTbls(1))
    RateFile = FreeFile
    Open UBPath$ + "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
      Get RateFile, RateRec, UBRateTbls(1)
      Ratecode(1) = QPTrim$(UBRateTbls(1).Ratecode)
      MINAMT(1) = UBRateTbls(1).MINAMT
      minunt(1) = UBRateTbls(1).MINUNITS
      MaxAmt(1) = UBRateTbls(1).MaxAmt
      TblBreak&(1, 1) = minunt(1)
      TblBreakfr&(1, 1) = 0
      TblUnitVal#(1, 1) = 0
      tblbrkchg#(1, 1) = 0
      MaxStep(1) = 1
      For Tnt = 1 To 10
        If UBRateTbls(1).TblBreaks(Tnt).UNITS <= 0 And UBRateTbls(1).TblBreaks(Tnt).UNITAMT <= 0 Then
          Exit For
        End If
        If Tnt = 10 Then
          If UBRateTbls(1).TblBreaks(Tnt).UNITS > 0 And UBRateTbls(1).TblBreaks(Tnt).UNITAMT > 0 Then
            MaxStep(1) = Tnt + 1
            TblBreakfr&(1, Tnt + 1) = UBRateTbls(1).TblBreaks(Tnt).UNITS
            TblBreak&(1, Tnt + 1) = 99999999
            TblUnitVal#(1, Tnt + 1) = UBRateTbls(1).TblBreaks(Tnt).UNITAMT
            tblbrkchg#(1, Tnt + 1) = 0
            Exit For
'          ElseIf UBRateTbls(1).TblBreaks(Tnt).UNITAMT <= 0 Then
'            Exit For
          End If
        Else

        If UBRateTbls(1).TblBreaks(Tnt + 1).UNITS > 0 Then
          If UBRateTbls(1).TblBreaks(Tnt).UNITS >= 0 Then
            If UBRateTbls(1).TblBreaks(Tnt).UNITS = 0 Then
              TblBreak&(1, Tnt + 1) = UBRateTbls(1).TblBreaks(Tnt + 1).UNITS - 1
              TblBreakfr&(1, Tnt + 1) = UBRateTbls(1).TblBreaks(Tnt).UNITS
              TblUnitVal#(1, Tnt + 1) = UBRateTbls(1).TblBreaks(Tnt).UNITAMT
              tblbrkchg#(1, Tnt + 1) = 0
              MaxStep(1) = Tnt + 1
            Else
             ' If Tnt > 1 Then
                TblBreakfr&(1, Tnt + 1) = UBRateTbls(1).TblBreaks(Tnt).UNITS
                TblBreak&(1, Tnt + 1) = UBRateTbls(1).TblBreaks(Tnt + 1).UNITS - 1
             ' Else
             '   TblBreakfr&(1, Tnt + 1) = minunt(1)
             '   TblBreak&(1, Tnt + 1) = UBRateTbls(1).TblBreaks(Tnt).UNITS
             ' End If
              TblUnitVal#(1, Tnt + 1) = UBRateTbls(1).TblBreaks(Tnt).UNITAMT
              tblbrkchg#(1, Tnt + 1) = 0
              MaxStep(1) = Tnt + 1
            End If
          Else
            MaxStep(1) = Tnt + 1
            TblBreakfr&(1, Tnt + 1) = 0
            TblBreak&(1, Tnt + 1) = 99999999
            TblUnitVal#(1, Tnt + 1) = UBRateTbls(1).TblBreaks(Tnt).UNITAMT
            tblbrkchg#(1, Tnt + 1) = 0
          End If
        Else
          MaxStep(1) = Tnt + 1
          TblBreakfr&(1, Tnt + 1) = UBRateTbls(1).TblBreaks(Tnt).UNITS
          TblBreak&(1, Tnt + 1) = 99999999
          TblUnitVal#(1, Tnt + 1) = UBRateTbls(1).TblBreaks(Tnt).UNITAMT
          tblbrkchg#(1, Tnt + 1) = 0
          Exit For
        End If
        End If
      Next Tnt
  Close RateFile
  Else
    
    UBRateTblRecLen = Len(UBRateTbls(1))
    RateFile = FreeFile
    Open UBPath$ + "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
       For RateRec = 1 To NumOfRates
        Greater = False
        Get RateFile, RateRec, UBRateTbls(RateRec)
        Ratecode(RateRec) = QPTrim$(UBRateTbls(RateRec).Ratecode)
        MINAMT(RateRec) = UBRateTbls(RateRec).MINAMT
        minunt(RateRec) = UBRateTbls(RateRec).MINUNITS
        MaxAmt(RateRec) = UBRateTbls(RateRec).MaxAmt
        TblBreakfr&(RateRec, 1) = 0
        TblBreak&(RateRec, 1) = minunt(RateRec)
        TblUnitVal#(RateRec, 1) = 0
        tblbrkchg#(RateRec, 1) = 0
        MaxStep(RateRec) = 1
        For Tnt = 1 To 10
          If UBRateTbls(RateRec).TblBreaks(Tnt).UNITS <= 0 And UBRateTbls(RateRec).TblBreaks(Tnt).UNITAMT <= 0 Then
            Exit For
          End If
          If Tnt = 10 Then
            If UBRateTbls(RateRec).TblBreaks(Tnt).UNITS > 0 And UBRateTbls(RateRec).TblBreaks(Tnt).UNITAMT > 0 Then
              MaxStep(RateRec) = Tnt + 1
              TblBreakfr&(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITS
              TblBreak&(RateRec, Tnt + 1) = 99999999
              TblUnitVal#(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITAMT
              tblbrkchg#(RateRec, Tnt + 1) = 0
              Exit For
'            ElseIf UBRateTbls(RateRec).TblBreaks(Tnt).UNITAMT <= 0 Then
'              Exit For
            End If
          Else
          If UBRateTbls(RateRec).TblBreaks(Tnt + 1).UNITS > 0 Then
            If UBRateTbls(RateRec).TblBreaks(Tnt).UNITS >= 0 Then
              If UBRateTbls(RateRec).TblBreaks(Tnt).UNITS = 0 Then
                TblBreakfr&(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITS
                TblBreak&(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt + 1).UNITS - 1
                TblUnitVal#(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITAMT
                tblbrkchg#(RateRec, Tnt + 1) = 0
                MaxStep(RateRec) = Tnt + 1
              Else
 '               If Tnt > 1 Then
                 TblBreakfr&(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITS
                 TblBreak&(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt + 1).UNITS - 1
 '               Else
 '                TblBreakfr&(RateRec, Tnt + 1) = minunt(RateRec)
 '                TblBreak&(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITS
 '               End If
                TblUnitVal#(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITAMT
                tblbrkchg#(RateRec, Tnt + 1) = 0
                MaxStep(RateRec) = Tnt + 1
              End If
            Else
              MaxStep(RateRec) = Tnt + 1
              TblBreakfr&(RateRec, Tnt + 1) = 0
              TblBreak&(RateRec, Tnt + 1) = 99999999
              TblUnitVal#(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITAMT
              tblbrkchg#(RateRec, Tnt + 1) = 0
            End If
          Else
            MaxStep(RateRec) = Tnt + 1
            TblBreakfr&(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITS
            TblBreak&(RateRec, Tnt + 1) = 99999999
            TblUnitVal#(RateRec, Tnt + 1) = UBRateTbls(RateRec).TblBreaks(Tnt).UNITAMT
            tblbrkchg#(RateRec, Tnt + 1) = 0
            Exit For
          End If
          End If
        Next Tnt
      Next RateRec
    Close RateFile
  End If
  GoSub DoRptHeader
  If Doall Then
    RptInfo$ = " All Rates"
    For RCnt = 1 To NumOfRates
      NMinAMT# = MINAMT(RCnt)
      RCode$ = Ratecode(RCnt)
      RateRec = RCnt
      If RptType = 1 Then
        GoSub DoRateHeader
      End If
      GoSub DoEachRate
      GoSub DoUnitStepFooter
    Next
  Else
    NMinAMT# = MINAMT(1)
    RCode$ = Ratecode(1)
    RptInfo$ = RCode$
    RateRec = 1
    If RptType = 1 Then
      GoSub DoRateHeader
    End If
    GoSub DoEachRate
    GoSub DoUnitStepFooter
  End If
  GoSub DoGrandFooter
  Close

  Erase TblBreak&, TotalConsp#, TotalCust
  Doall = False
  If RptType = 0 Then
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmRptSCnsmpRateCode
    ARptSCnsmpRate.txtDate = Now
    ARptSCnsmpRate.txtTown = TOWNNAME$
    ARptSCnsmpRate.Title = "Consumption by Rate Code"
    ARptSCnsmpRate.txtDate1 = txtDate1
    ARptSCnsmpRate.txtDate2 = txtDate2
    ARptSCnsmpRate.FldRptInfo = RptInfo$
 '   ARptSCnsmpRate.totCust = Using("###,###,###,###", GBCustTot)
    ARptSCnsmpRate.totConsump = Using("###,###,###,###", GBGTMeterConsp#)
    ARptSCnsmpRate.totUsage = Using(" $ ##,###,###.##", GBBigUTotal#)
    ARptSCnsmpRate.totMin = Using(" $ ##,###,###.##", GBMinGT#)
    ARptSCnsmpRate.totcharges = Using(" $ ###,###,###.##", (Round(GBMinGT# + GBBigUTotal#)))
    If PageBrk.Value = 1 Then
      ARptSCnsmpRate.GetName ReportFile$, True
    Else
      ARptSCnsmpRate.GetName ReportFile$, False
    End If
    ARptSCnsmpRate.startrpt
  Else
    ViewPrint ReportFile$, "Consumption by RateCode"
  End If
  Exit Sub
DoEachRate:
  FrmShowPctComp.Label1 = "Processing Rate " + RCode$
  FrmShowPctComp.Show

  For cnt& = 1 To NumofRecs&
    FrmShowPctComp.ShowPctComp cnt, NumofRecs&
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
        'IF Valid Customer Then Get the Customer and Check to See if Rate Appl
        CustomerRecord = UBTransRec(1).CustAcctNo

        If CustomerRecord > 0 Then
          Get UBCust, CustomerRecord, UBCustRec(1)
            For Snt = 1 To 15
              If QPTrim$(UBCustRec(1).serv(Snt).Ratecode) = RCode$ Then
                MtrType$ = UBCustRec(1).serv(Snt).RMtrType
                Select Case MtrType$
                Case "W"
                  MT = 1
                Case "S"
                  MT = 2
                Case "C"
                  MT = 3
                Case "D"
                  MT = 5
                Case "G"
                  MT = 6
                Case "T"
                  MT = 7
                Case "L"
                  MT = 8
                Case "I"
                  MT = 9
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
            MType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrType)
            If MtrType$ = MType$ Then
            'IF LEN(MType$) > 0 THEN
              Multi& = UBCustRec(1).LocMeters(MCnt).MTRMulti
              If Multi& = 0 Then
                Multi& = 1
              End If
              If UBCustRec(1).LocMeters(MCnt).MtrUnit = "C" Then
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
            ''If MeterConsp& > 0 Then Stop
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
'''          For LL = 1 To MaxStep(RateRec)
'''            If NewMtrConsp# > TblBreak&(RateRec, LL - 1) And NewMtrConsp# <= TblBreak&(RateRec, LL) Then
'''              NewMtrConsp# = (TMeterConsp# - TblBreak&(RateRec, LL - 1))
'''              TotalConsp#(RateRec, LL) = TotalConsp#(RateRec, LL) + NewMtrConsp#
'''              TotalConsp#(RateRec, LL - 1) = (TotalConsp#(RateRec, LL - 1) + TblBreak&(RateRec, LL - 1))
'''              TotalCust(RateRec, LL) = TotalCust(RateRec, LL) + 1
'''              NonUpdated = 0
'''              Exit For
'''            End If
'''          Next LL
'''          If NonUpdated = 1 Then
'''            TotalConsp#(RateRec, MaxStep(RateRec)) = TotalConsp#(RateRec, MaxStep(RateRec)) + NewMtrConsp#
'''            TotalCust(RateRec, MaxStep(RateRec)) = TotalCust(RateRec, MaxStep(RateRec)) + 1
'''          End If
        NewMtrConsp# = TMeterConsp#
        TotalCust(RateRec) = TotalCust(RateRec) + 1
        MinBillAmt# = MINAMT#(RateRec)
        TAmt# = 0
        If MaxAmt#(RateRec) > 0 Then
          MaxFlag = True
        Else
          MaxFlag = False
        End If
      If MaxStep(RateRec) >= 2 Then
        If NewMtrConsp# >= TblBreakfr&(RateRec, 1) And NewMtrConsp# <= TblBreakfr&(RateRec, 2) Then
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 1))
          'special patch for cave junction
    '      If UNITS& = 0 Then
    '        UNITS& = 1
    '      End If
          If UNITS& <= 0 Then UNITS& = NewMtrConsp#
          TotalConsp#(RateRec, 1) = TotalConsp#(RateRec, 1) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 1))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + NTAmt#)
          Else
            tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + Round#(UNITS& * TblUnitVal#(RateRec, 1)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 1)))
          GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 1) Then
          UNITS& = (TblBreak&(RateRec, 1) - TblBreakfr&(RateRec, 1))
          TotalConsp#(RateRec, 1) = TotalConsp#(RateRec, 1) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 1))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + NTAmt#)
          Else
            tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + Round#(UNITS& * TblUnitVal#(RateRec, 1)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 1)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 2) Then
          UNITS& = (TblBreak&(RateRec, 1) - TblBreakfr&(RateRec, 1))
          TotalConsp#(RateRec, 1) = TotalConsp#(RateRec, 1) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 1))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + NTAmt#)
          Else
            tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + Round#(UNITS& * TblUnitVal#(RateRec, 1)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 1)))
        End If
      Else          'no other rate breaks
        UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 1))
        TotalConsp#(RateRec, 1) = TotalConsp#(RateRec, 1) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 1))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + NTAmt#)
          Else
            tblbrkchg(RateRec, 1) = (tblbrkchg(RateRec, 1) + Round#(UNITS& * TblUnitVal#(RateRec, 1)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 1)))
        GoTo GOTIT
      End If
    
      'Break 2
      If MaxStep(RateRec) >= 3 Then
        If NewMtrConsp# >= TblBreakfr&(RateRec, 2) And NewMtrConsp# <= TblBreakfr&(RateRec, 3) Then
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 1))
          TotalConsp#(RateRec, 2) = TotalConsp#(RateRec, 2) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 2))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 2) = (tblbrkchg(RateRec, 2) + NTAmt#)
          Else
            tblbrkchg(RateRec, 2) = (tblbrkchg(RateRec, 2) + Round#(UNITS& * TblUnitVal#(RateRec, 2)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 2)))
          GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 2) Then
          If TblBreakfr&(RateRec, 2) < 1 Then
            UNITS& = (TblBreakfr&(RateRec, 3) - 1)
          Else
            UNITS& = (TblBreakfr&(RateRec, 3) - TblBreakfr&(RateRec, 2))
          End If
          TotalConsp#(RateRec, 2) = TotalConsp#(RateRec, 2) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 2))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 2) = (tblbrkchg(RateRec, 2) + NTAmt#)
          Else
            tblbrkchg(RateRec, 2) = (tblbrkchg(RateRec, 2) + Round#(UNITS& * TblUnitVal#(RateRec, 2)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 2)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 3) Then
          If TblBreakfr&(RateRec, 2) < 1 Then
            UNITS& = (TblBreakfr&(RateRec, 3) - 1)
          Else
            UNITS& = (TblBreakfr&(RateRec, 3) - TblBreakfr&(RateRec, 2))
          End If
          TotalConsp#(RateRec, 2) = TotalConsp#(RateRec, 2) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 2))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 2) = (tblbrkchg(RateRec, 2) + NTAmt#)
          Else
            tblbrkchg(RateRec, 2) = (tblbrkchg(RateRec, 2) + Round#(UNITS& * TblUnitVal#(RateRec, 2)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 2)))
        End If
      Else
        UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 1))
        TotalConsp#(RateRec, 2) = TotalConsp#(RateRec, 2) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 2))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 2) = (tblbrkchg(RateRec, 2) + NTAmt#)
          Else
            tblbrkchg(RateRec, 2) = (tblbrkchg(RateRec, 2) + Round#(UNITS& * TblUnitVal#(RateRec, 2)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 2)))
        GoTo GOTIT
      End If
    
      'Break 3
      If MaxStep(RateRec) >= 4 Then
        If NewMtrConsp# >= TblBreakfr&(RateRec, 3) And NewMtrConsp# <= TblBreakfr&(RateRec, 4) Then
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 2))
          TotalConsp#(RateRec, 3) = TotalConsp#(RateRec, 3) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 3))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 3) = (tblbrkchg(RateRec, 3) + NTAmt#)
          Else
            tblbrkchg(RateRec, 3) = (tblbrkchg(RateRec, 3) + Round#(UNITS& * TblUnitVal#(RateRec, 3)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 3)))
          GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 3) Then
          UNITS& = (TblBreakfr&(RateRec, 4) - TblBreakfr&(RateRec, 3))
          TotalConsp#(RateRec, 3) = TotalConsp#(RateRec, 3) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 3))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 3) = (tblbrkchg(RateRec, 3) + NTAmt#)
          Else
            tblbrkchg(RateRec, 3) = (tblbrkchg(RateRec, 3) + Round#(UNITS& * TblUnitVal#(RateRec, 3)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 3)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 4) Then
          UNITS& = (TblBreakfr&(RateRec, 4) - TblBreakfr&(RateRec, 3))
          TotalConsp#(RateRec, 3) = TotalConsp#(RateRec, 3) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 3))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 3) = (tblbrkchg(RateRec, 3) + NTAmt#)
          Else
            tblbrkchg(RateRec, 3) = (tblbrkchg(RateRec, 3) + Round#(UNITS& * TblUnitVal#(RateRec, 3)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 3)))
        End If
      Else
        UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 2))
        TotalConsp#(RateRec, 3) = TotalConsp#(RateRec, 3) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 3))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 3) = (tblbrkchg(RateRec, 3) + NTAmt#)
          Else
            tblbrkchg(RateRec, 3) = (tblbrkchg(RateRec, 3) + Round#(UNITS& * TblUnitVal#(RateRec, 3)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 3)))
        GoTo GOTIT
      End If
    
      'Break 4
     If MaxStep(RateRec) >= 5 Then
       If NewMtrConsp# >= TblBreakfr&(RateRec, 4) And NewMtrConsp# <= TblBreakfr&(RateRec, 5) Then
         UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 3))
         TotalConsp#(RateRec, 4) = TotalConsp#(RateRec, 4) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 4))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 4) = (tblbrkchg(RateRec, 4) + NTAmt#)
          Else
            tblbrkchg(RateRec, 4) = (tblbrkchg(RateRec, 4) + Round#(UNITS& * TblUnitVal#(RateRec, 4)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 4)))
         GoTo GOTIT
       ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 4) Then
         UNITS& = (TblBreakfr&(RateRec, 5) - TblBreakfr&(RateRec, 4))
         TotalConsp#(RateRec, 4) = TotalConsp#(RateRec, 4) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 4))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 4) = (tblbrkchg(RateRec, 4) + NTAmt#)
          Else
            tblbrkchg(RateRec, 4) = (tblbrkchg(RateRec, 4) + Round#(UNITS& * TblUnitVal#(RateRec, 4)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 4)))
         GoTo GOTIT
       ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 5) Then
         UNITS& = (TblBreakfr&(RateRec, 5) - TblBreakfr&(RateRec, 4))
         TotalConsp#(RateRec, 4) = TotalConsp#(RateRec, 4) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 4))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 4) = (tblbrkchg(RateRec, 4) + NTAmt#)
          Else
            tblbrkchg(RateRec, 4) = (tblbrkchg(RateRec, 4) + Round#(UNITS& * TblUnitVal#(RateRec, 4)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 4)))
       End If
     Else
       UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 3))
       TotalConsp#(RateRec, 4) = TotalConsp#(RateRec, 4) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 4))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 4) = (tblbrkchg(RateRec, 4) + NTAmt#)
          Else
            tblbrkchg(RateRec, 4) = (tblbrkchg(RateRec, 4) + Round#(UNITS& * TblUnitVal#(RateRec, 4)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 4)))
       GoTo GOTIT
     End If
    
     'break 5
     If MaxStep(RateRec) >= 6 Then
       If NewMtrConsp# >= TblBreakfr&(RateRec, 5) And NewMtrConsp# <= TblBreakfr&(RateRec, 6) Then
         UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 4))
         TotalConsp#(RateRec, 5) = TotalConsp#(RateRec, 5) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 5))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 5) = (tblbrkchg(RateRec, 5) + NTAmt#)
          Else
            tblbrkchg(RateRec, 5) = (tblbrkchg(RateRec, 5) + Round#(UNITS& * TblUnitVal#(RateRec, 5)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 5)))
         GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 5) Then
          UNITS& = (TblBreakfr&(RateRec, 6) - TblBreakfr&(RateRec, 5))
          TotalConsp#(RateRec, 5) = TotalConsp#(RateRec, 5) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 5))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 5) = (tblbrkchg(RateRec, 5) + NTAmt#)
          Else
            tblbrkchg(RateRec, 5) = (tblbrkchg(RateRec, 5) + Round#(UNITS& * TblUnitVal#(RateRec, 5)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 5)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 6) Then
          UNITS& = (TblBreakfr&(RateRec, 6) - TblBreakfr&(RateRec, 5))
          TotalConsp#(RateRec, 5) = TotalConsp#(RateRec, 5) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 5))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 5) = (tblbrkchg(RateRec, 5) + NTAmt#)
          Else
            tblbrkchg(RateRec, 5) = (tblbrkchg(RateRec, 5) + Round#(UNITS& * TblUnitVal#(RateRec, 5)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 5)))
        End If
      Else
        UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 4))
        TotalConsp#(RateRec, 5) = TotalConsp#(RateRec, 5) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 5))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 5) = (tblbrkchg(RateRec, 5) + NTAmt#)
          Else
            tblbrkchg(RateRec, 5) = (tblbrkchg(RateRec, 5) + Round#(UNITS& * TblUnitVal#(RateRec, 5)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 5)))
        GoTo GOTIT
      End If
    
      'break 6
      If MaxStep(RateRec) >= 7 Then
        If NewMtrConsp# >= TblBreakfr&(RateRec, 6) And NewMtrConsp# <= TblBreakfr&(RateRec, 7) Then
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 5))
          TotalConsp#(RateRec, 6) = TotalConsp#(RateRec, 6) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 6))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 6) = (tblbrkchg(RateRec, 6) + NTAmt#)
          Else
            tblbrkchg(RateRec, 6) = (tblbrkchg(RateRec, 6) + Round#(UNITS& * TblUnitVal#(RateRec, 6)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 6)))
          GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 6) Then
          UNITS& = (TblBreakfr&(RateRec, 7) - TblBreakfr&(RateRec, 6))
          TotalConsp#(RateRec, 6) = TotalConsp#(RateRec, 6) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 6))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 6) = (tblbrkchg(RateRec, 6) + NTAmt#)
          Else
            tblbrkchg(RateRec, 6) = (tblbrkchg(RateRec, 6) + Round#(UNITS& * TblUnitVal#(RateRec, 6)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 6)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 7) Then
          UNITS& = (TblBreakfr&(RateRec, 7) - TblBreakfr&(RateRec, 6))
          TotalConsp#(RateRec, 6) = TotalConsp#(RateRec, 6) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 6))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 6) = (tblbrkchg(RateRec, 6) + NTAmt#)
          Else
            tblbrkchg(RateRec, 6) = (tblbrkchg(RateRec, 6) + Round#(UNITS& * TblUnitVal#(RateRec, 6)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 6)))
        End If
      Else
        UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 5))
        TotalConsp#(RateRec, 6) = TotalConsp#(RateRec, 6) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 6))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 6) = (tblbrkchg(RateRec, 6) + NTAmt#)
          Else
            tblbrkchg(RateRec, 6) = (tblbrkchg(RateRec, 6) + Round#(UNITS& * TblUnitVal#(RateRec, 6)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 6)))
        GoTo GOTIT
      End If
    
      'break 7
      If MaxStep(RateRec) >= 8 Then
        If NewMtrConsp# >= TblBreakfr&(RateRec, 7) And NewMtrConsp# <= TblBreakfr&(RateRec, 8) Then
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 6))
          TotalConsp#(RateRec, 7) = TotalConsp#(RateRec, 7) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 7))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 7) = (tblbrkchg(RateRec, 7) + NTAmt#)
          Else
            tblbrkchg(RateRec, 7) = (tblbrkchg(RateRec, 7) + Round#(UNITS& * TblUnitVal#(RateRec, 7)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 7)))
          GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 7) Then
          UNITS& = (TblBreakfr&(RateRec, 8) - TblBreakfr&(RateRec, 7))
          TotalConsp#(RateRec, 7) = TotalConsp#(RateRec, 7) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 7))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 7) = (tblbrkchg(RateRec, 7) + NTAmt#)
          Else
            tblbrkchg(RateRec, 7) = (tblbrkchg(RateRec, 7) + Round#(UNITS& * TblUnitVal#(RateRec, 7)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 7)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 8) Then
          UNITS& = (TblBreakfr&(RateRec, 8) - TblBreakfr&(RateRec, 7))
          TotalConsp#(RateRec, 7) = TotalConsp#(RateRec, 7) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 7))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 7) = (tblbrkchg(RateRec, 7) + NTAmt#)
          Else
            tblbrkchg(RateRec, 7) = (tblbrkchg(RateRec, 7) + Round#(UNITS& * TblUnitVal#(RateRec, 7)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 7)))
        End If
      Else
        UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 6))
        TotalConsp#(RateRec, 7) = TotalConsp#(RateRec, 7) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 7))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 7) = (tblbrkchg(RateRec, 7) + NTAmt#)
          Else
            tblbrkchg(RateRec, 7) = (tblbrkchg(RateRec, 7) + Round#(UNITS& * TblUnitVal#(RateRec, 7)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 7)))
        GoTo GOTIT
      End If
      'break 8
      If MaxStep(RateRec) >= 9 Then
        If NewMtrConsp# >= TblBreakfr&(RateRec, 8) And NewMtrConsp# <= TblBreakfr&(RateRec, 9) Then
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 7))
          TotalConsp#(RateRec, 8) = TotalConsp#(RateRec, 8) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 8))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 8) = (tblbrkchg(RateRec, 8) + NTAmt#)
          Else
            tblbrkchg(RateRec, 8) = (tblbrkchg(RateRec, 8) + Round#(UNITS& * TblUnitVal#(RateRec, 8)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 8)))
          GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 8) Then
          UNITS& = (TblBreakfr&(RateRec, 9) - TblBreakfr&(RateRec, 8))
          TotalConsp#(RateRec, 8) = TotalConsp#(RateRec, 8) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 8))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 8) = (tblbrkchg(RateRec, 8) + NTAmt#)
          Else
            tblbrkchg(RateRec, 8) = (tblbrkchg(RateRec, 8) + Round#(UNITS& * TblUnitVal#(RateRec, 8)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 8)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 9) Then
          UNITS& = (TblBreakfr&(RateRec, 9) - TblBreakfr&(RateRec, 8))
          TotalConsp#(RateRec, 8) = TotalConsp#(RateRec, 8) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 8))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 8) = (tblbrkchg(RateRec, 8) + NTAmt#)
          Else
            tblbrkchg(RateRec, 8) = (tblbrkchg(RateRec, 8) + Round#(UNITS& * TblUnitVal#(RateRec, 8)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 8)))
        End If
      Else
        UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 7))
        TotalConsp#(RateRec, 8) = TotalConsp#(RateRec, 8) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 8))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 8) = (tblbrkchg(RateRec, 8) + NTAmt#)
          Else
            tblbrkchg(RateRec, 8) = (tblbrkchg(RateRec, 8) + Round#(UNITS& * TblUnitVal#(RateRec, 8)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 8)))
        GoTo GOTIT
      End If
    
      'break 9
      If MaxStep(RateRec) >= 10 Then
        If NewMtrConsp# >= TblBreakfr&(RateRec, 9) And NewMtrConsp# <= TblBreakfr&(RateRec, 10) Then
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 8))
          TotalConsp#(RateRec, 9) = TotalConsp#(RateRec, 9) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 9))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 9) = (tblbrkchg(RateRec, 9) + NTAmt#)
          Else
            tblbrkchg(RateRec, 9) = (tblbrkchg(RateRec, 9) + Round#(UNITS& * TblUnitVal#(RateRec, 9)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 9)))
          GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 9) Then
          UNITS& = (TblBreakfr&(RateRec, 10) - TblBreakfr&(RateRec, 9))
          TotalConsp#(RateRec, 9) = TotalConsp#(RateRec, 9) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 9))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 9) = (tblbrkchg(RateRec, 9) + NTAmt#)
          Else
            tblbrkchg(RateRec, 9) = (tblbrkchg(RateRec, 9) + Round#(UNITS& * TblUnitVal#(RateRec, 9)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 9)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 10) Then
          UNITS& = (TblBreakfr&(RateRec, 10) - TblBreakfr&(RateRec, 9))
          TotalConsp#(RateRec, 9) = TotalConsp#(RateRec, 9) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 9))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 9) = (tblbrkchg(RateRec, 9) + NTAmt#)
          Else
            tblbrkchg(RateRec, 9) = (tblbrkchg(RateRec, 9) + Round#(UNITS& * TblUnitVal#(RateRec, 9)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 9)))
        End If
      Else
        UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 8))
        TotalConsp#(RateRec, 9) = TotalConsp#(RateRec, 9) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 9))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 9) = (tblbrkchg(RateRec, 9) + NTAmt#)
          Else
            tblbrkchg(RateRec, 9) = (tblbrkchg(RateRec, 9) + Round#(UNITS& * TblUnitVal#(RateRec, 9)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 9)))
       GoTo GOTIT
      End If
    
      If MaxStep(RateRec) >= 11 Then
        If NewMtrConsp# >= TblBreakfr&(RateRec, 10) And NewMtrConsp# <= TblBreakfr&(RateRec, 11) Then
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 9))
          TotalConsp#(RateRec, 10) = TotalConsp#(RateRec, 10) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 10))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 10) = (tblbrkchg(RateRec, 10) + NTAmt#)
          Else
            tblbrkchg(RateRec, 10) = (tblbrkchg(RateRec, 10) + Round#(UNITS& * TblUnitVal#(RateRec, 10)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 10)))
          GoTo GOTIT
        ElseIf NewMtrConsp# < TblBreakfr&(RateRec, 10) Then
          UNITS& = (TblBreakfr&(RateRec, 11) - TblBreakfr&(RateRec, 10))
          TotalConsp#(RateRec, 10) = TotalConsp#(RateRec, 10) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 10))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 10) = (tblbrkchg(RateRec, 10) + NTAmt#)
          Else
            tblbrkchg(RateRec, 10) = (tblbrkchg(RateRec, 10) + Round#(UNITS& * TblUnitVal#(RateRec, 10)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 10)))
          GoTo GOTIT
        ElseIf NewMtrConsp# > TblBreakfr&(RateRec, 11) Then
          UNITS& = (TblBreakfr&(RateRec, 11) - TblBreakfr&(RateRec, 10))
          TotalConsp#(RateRec, 10) = TotalConsp#(RateRec, 10) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 10))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 10) = (tblbrkchg(RateRec, 10) + NTAmt#)
          Else
            tblbrkchg(RateRec, 10) = (tblbrkchg(RateRec, 10) + Round#(UNITS& * TblUnitVal#(RateRec, 10)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 10)))
    '      '*****
          UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 10))
          TotalConsp#(RateRec, 11) = TotalConsp#(RateRec, 11) + UNITS&
          If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 11))) >= MaxAmt(RateRec) And MaxFlag Then
            NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
            tblbrkchg(RateRec, 11) = (tblbrkchg(RateRec, 11) + NTAmt#)
          Else
            tblbrkchg(RateRec, 11) = (tblbrkchg(RateRec, 11) + Round#(UNITS& * TblUnitVal#(RateRec, 11)))
          End If
          TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 11)))
        End If
      Else
        UNITS& = (NewMtrConsp# - TblBreak&(RateRec, 9))
        TotalConsp#(RateRec, 10) = TotalConsp#(RateRec, 10) + UNITS&
        If Round#(MinBillAmt# + TAmt# + (UNITS& * TblUnitVal#(RateRec, 10))) >= MaxAmt(RateRec) And MaxFlag Then
          NTAmt# = (MaxAmt(RateRec) - (TAmt# + MinBillAmt#))
          tblbrkchg(RateRec, 10) = (tblbrkchg(RateRec, 10) + NTAmt#)
        Else
          tblbrkchg(RateRec, 10) = (tblbrkchg(RateRec, 10) + Round#(UNITS& * TblUnitVal#(RateRec, 10)))
        End If
        TAmt# = Round#(TAmt# + (UNITS& * TblUnitVal#(RateRec, 10)))
       GoTo GOTIT
      End If
    
      End If
GOTIT:
      TMeterConsp# = 0
      End If
    End If

  Next
Return

DoUnitStepHeader:
Return
DoRptHeader:
If RptType = 1 Then
  PageNo = PageNo + 1
  Print #UBRpt, Tab(29); "Consumption by RateCode"; Tab(70); "Page #"; PageNo
  Print #UBRpt, TOWNNAME$
  Print #UBRpt, "     Report Date: "; Now
  Print #UBRpt, "Period Beginning: "; txtDate1
  Print #UBRpt, "   Period Ending: "; txtDate2
  Print #UBRpt, "      Report Opt: "; "Rates - "; RptInfo$
  Print #UBRpt, " "
  Print #UBRpt, Dash80$
  LineCnt = 5
End If
Return
DoRateHeader:
  'PageNo = PageNo + 1
'  Print #UBRpt, Tab(29); "Consumption by RateCode"; Tab(70); "Page #"; PageNo
'  Print #UBRpt, TOWNNAME$
'  Print #UBRpt, "Report Date: "; Now
 If RptType = 1 Then
  Print #UBRpt, " "
  Print #UBRpt, "    For Rate Code: "; RCode$
  'Print #UBRpt, " Period Beginning: "; txtDate1
  'Print #UBRpt, "    Period Ending: "; txtDate2
  Print #UBRpt, ; Tab(52); "Usage Charge"; Tab(68); "  Min Charge"
  Print #UBRpt, Dash80$
  LineCnt = LineCnt + 6
End If
Return

DoUnitStepFooter:
  UntPrc = 0
  TblBreak&(RateRec, MaxStep(RateRec)) = 99999999
  If RptType = 1 Then
    Print #UBRpt, "MinAmount - "; Using("###,###.##", MINAMT(RateRec)); "  MinUnits - "; minunt(RateRec);
    If MaxAmt(RateRec) > 0 Then
      Print #UBRpt, "  MaxAmount - "; Using("###,###.##", MaxAmt(RateRec))
    Else
      Print #UBRpt,
    End If
    Print #UBRpt, "--------------------------------------------------------------"
  For LL = 1 To MaxStep(RateRec)
    If LL = 1 Then
      Print #UBRpt, "Min -";
    Else
      Print #UBRpt, "Step # "; LL - 1;
    End If
    Print #UBRpt, Tab(12); "From "; TblBreakfr&(RateRec, LL); " to "; TblBreak&(RateRec, LL)
    Print #UBRpt, "Consumption = "; Tab(18); Using("#########,#", TotalConsp#(RateRec, LL));
    'Print #UBRpt, Tab(29); " # of Trans = "; Using("#####,#", TotalCust(RateRec, LL));
    
    If TblUnitVal#(RateRec, LL) > 0 Then
     'Tempcalccnsp# = Round(TotalCust(RateRec, LL) * minunt(RateRec))
    Else
      TblUnitVal#(RateRec, LL) = 0
    End If
    UntPrc# = tblbrkchg#(RateRec, LL)
    BigUTotal# = Round#(BigUTotal# + UntPrc#)
    'MinGT# = Round#(MinGT# + Round#(NMinAMT# * TotalCust(RateRec)))
    If LL = 1 Then
      Print #UBRpt, Tab(51); Using("###,###,###.##", UntPrc#); Tab(68); Using("  ###,###.##", Round#(NMinAMT# * TotalCust(RateRec)))
    Else
      Print #UBRpt, Tab(51); Using("###,###,###.##", UntPrc#)
    End If
    MinGT# = Round#(MinGT# + Round#(NMinAMT# * TotalCust(RateRec)))
    BigTotCust = BigTotCust + TotalCust(RateRec)
    'If TotalCust(RateRec, LL) > 0 Then
      
      'PRINT #UBRpt, "  Avg Use= "; USING "#####,#.##"; TotalConsp#(LL) / Tota
   ' Else
   '   Print #UBRpt, ""
   ' End If
    Print #UBRpt, Dash80$
    LineCnt = LineCnt + 4
  Next LL
  Print #UBRpt, "Rate Totals: "; Using("###,###,###,###", GTMeterConsp#); 'Tab(41); "  "; Using("#####,#", BigTotCust);
  Print #UBRpt, Tab(51); Using("###,###,###.##", BigUTotal#); Tab(67); Using("##,###,###.##", Round#(NMinAMT# * TotalCust(RateRec)))
  Print #UBRpt, "# of Trans - "; Using("###,###,###", TotalCust(RateRec))
  LineCnt = LineCnt + 2
  If PageBrk = 1 Then
    Print #UBRpt, Chr$(12);
  Else
    Print #UBRpt,
    LineCnt = LineCnt + 1
  End If
  BigTotCust = BigTotCust + TotalCust(RateRec)
  GBBigUTotal# = Round#(GBBigUTotal# + BigUTotal#)
  GBMinGT# = Round#(NMinAMT# * TotalCust(RateRec))
  GBGTMeterConsp# = Round#(GBGTMeterConsp# + GTMeterConsp#)
  GBCustTot = GBCustTot + BigTotCust
  BigUTotal# = 0
  MinGT# = 0
  BigTotCust = 0
  GTMeterConsp# = 0

  Else
  If MaxAmt(RateRec) > 0 Then
    ToPrintI$ = "Min Amount - " + Using("###,###.##", MINAMT(RateRec)) + " MinUnits - " + Str(minunt(RateRec)) + "   MaxAmount - " + Using("###,###.##", MaxAmt(RateRec))
  Else
    ToPrintI$ = "Min Amount - " + Using("###,###.##", MINAMT(RateRec)) + " MinUnits - " + Str(minunt(RateRec))
  End If
  For LL = 1 To MaxStep(RateRec)
    If LL = 1 Then
      ToPrint$ = "Min - "
      ToPrint$ = ToPrint$ + "~" + Str(TblBreakfr&(RateRec, LL)) + "~" + Str(TblBreak&(RateRec, LL))
      ToPrint$ = ToPrint$ + "~" + Using("#,###,###,###", TotalConsp#(RateRec, LL))
      ToPrint$ = ToPrint$ + "~" + Using("###,###", TotalCust(RateRec))
    Else
      ToPrint$ = "Step # " + Str(LL - 1)
      ToPrint$ = ToPrint$ + "~" + Str(TblBreakfr&(RateRec, LL)) + "~" + Str(TblBreak&(RateRec, LL))
      ToPrint$ = ToPrint$ + "~" + Using("#,###,###,###", TotalConsp#(RateRec, LL))
      ToPrint$ = ToPrint$ + "~" + " "
    End If
   ' Using("  ###,###.##", Round#(NMinAMT# * Totalcust(RateRec))
    If TblUnitVal#(RateRec, LL) > 0 Then
    Else
      TblUnitVal#(RateRec, LL) = 0
    End If
    UntPrc# = tblbrkchg#(RateRec, LL)
'    If UntPrc# > MaxAmt(RateRec) Then
'      UntPrc# = MaxAmt(RateRec)
'    End If
    BigUTotal# = Round#(BigUTotal# + UntPrc#)
    'MinGT# = Round#(MinGT# + Round#(NMinAMT# * TotalCust(RateRec)))
    If LL = 1 Then
      ToPrint$ = ToPrint$ + "~" + Using("###,###,###.##", UntPrc#) + "~" + Using("  ###,###.##", Round#(NMinAMT# * TotalCust(RateRec)))
    Else
      ToPrint$ = ToPrint$ + "~" + Using("###,###,###.##", UntPrc#) + "~" + " "
    End If
    Print #UBRpt, RCode$ + "~" + ToPrint$ + "~" + ToPrintI$
    ToPrint$ = ""
  Next LL
  BigTotCust = BigTotCust + TotalCust(RateRec)
  GBBigUTotal# = Round#(GBBigUTotal# + BigUTotal#)
  GBMinGT# = Round#(NMinAMT# * TotalCust(RateRec))
  GBGTMeterConsp# = Round#(GBGTMeterConsp# + GTMeterConsp#)
  GBCustTot = GBCustTot + BigTotCust
  BigUTotal# = 0
  MinGT# = 0
  BigTotCust = 0
  GTMeterConsp# = 0
  End If
Return
DoGrandFooter:
If RptType = 1 Then
  If PageBrk = 1 Then
    GoSub DoRptHeader
  End If
  Print #UBRpt,
 ' Print #UBRpt, "Grand Total Customers     : "; Using("###,###,###,###", GBCustTot)
  Print #UBRpt, "Grand Total Consumption   : "; Using("###,###,###,###", GBGTMeterConsp#)
  Print #UBRpt, "Grand Total Usage Charge  : "; Using(" $ ##,###,###.##", GBBigUTotal#)
  Print #UBRpt, "Grand Total Minimum Charge: "; Using(" $ ##,###,###.##", GBMinGT#)
  Print #UBRpt, "Grand Total Charges       : "; Using(" $ ##,###,###.##", (Round(GBMinGT# + GBBigUTotal#)))
  Print #UBRpt,
End If
Return

GoTo ExitConsStep
ExitConsStep:
  
  Close
Exit Sub
End Sub
'Private Function GetCharge#(RateTbl As UBRateTblRecType, TMeterConsp&, MeterMulti&)
'  Dim MinBillAmt As Double, TAmt As Double, LastTblCnt As Integer
'  Dim BCnt As Integer, MeterConsump As Long, UNITS As Long
'  'STOP
'
'  MinBillAmt# = RateTbl.MINAMT
'
'  If MinBillAmt# < -1000000 Then
'    MinBillAmt# = 0
'    TAmt# = -1
'    GoTo GotTAmt
'  End If
'
''SunnyBeech 091701
'  If TMeterConsp& <= RateTbl.MINUNITS Then
'    TAmt# = 0
'    GoTo GotTAmt
'  End If
'
'  LastTblCnt = 10
'  For BCnt = 1 To 10
'    If RateTbl.TblBreaks(BCnt).UNITAMT <= 0 Then
'      LastTblCnt = BCnt - 1
'      Exit For
'    End If
'  Next
'
'  MeterConsump& = TMeterConsp&
'
'  TAmt# = 0
'
'  If LastTblCnt >= 2 Then
'    If MeterConsump& >= RateTbl.TblBreaks(1).UNITS And MeterConsump& <= RateTbl.TblBreaks(2).UNITS Then
'      UNITS& = (MeterConsump& - RateTbl.TblBreaks(1).UNITS)
'      'special patch for cave junction
'      If UNITS& = 0 Then
'        UNITS& = 1
'      End If
'      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
'      GoTo GotTAmt
'    Else
'      UNITS& = (RateTbl.TblBreaks(2).UNITS - RateTbl.TblBreaks(1).UNITS)
'      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
'    End If
'  Else          'no other rate breaks
'    UNITS& = (MeterConsump& - RateTbl.TblBreaks(1).UNITS)
'    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
'    GoTo GotTAmt
'  End If
'
'  'Break 2
'  If LastTblCnt >= 3 Then
'    If MeterConsump& > RateTbl.TblBreaks(2).UNITS And MeterConsump& <= RateTbl.TblBreaks(3).UNITS Then
'      UNITS& = (MeterConsump& - RateTbl.TblBreaks(2).UNITS)
'      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
'      GoTo GotTAmt
'    Else
'      UNITS& = (RateTbl.TblBreaks(3).UNITS - RateTbl.TblBreaks(2).UNITS)
'      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
'    End If
'  Else
'    UNITS& = (MeterConsump& - RateTbl.TblBreaks(2).UNITS)
'    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
'    GoTo GotTAmt
'  End If
'
'  'Break 3
'  If LastTblCnt >= 4 Then
'    If MeterConsump& >= RateTbl.TblBreaks(3).UNITS And MeterConsump& <= RateTbl.TblBreaks(4).UNITS Then
'      UNITS& = (MeterConsump& - RateTbl.TblBreaks(3).UNITS)
'      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
'      GoTo GotTAmt
'    Else
'      UNITS& = (RateTbl.TblBreaks(4).UNITS - RateTbl.TblBreaks(3).UNITS)
'      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
'    End If
'  Else
'    UNITS& = (MeterConsump& - RateTbl.TblBreaks(3).UNITS)
'    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
'    GoTo GotTAmt
'  End If
'
'  'Break 4
' If LastTblCnt >= 5 Then
'   If MeterConsump& >= RateTbl.TblBreaks(4).UNITS And MeterConsump& <= RateTbl.TblBreaks(5).UNITS Then
'     UNITS& = (MeterConsump& - RateTbl.TblBreaks(4).UNITS)
'     TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
'     GoTo GotTAmt
'   Else
'     UNITS& = (RateTbl.TblBreaks(5).UNITS - RateTbl.TblBreaks(4).UNITS)
'     TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
'   End If
' Else
'   UNITS& = (MeterConsump& - RateTbl.TblBreaks(4).UNITS)
'   TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
'   GoTo GotTAmt
' End If
'
' 'break 5
' If LastTblCnt >= 6 Then
'   If MeterConsump& >= RateTbl.TblBreaks(5).UNITS And MeterConsump& <= RateTbl.TblBreaks(6).UNITS Then
'     UNITS& = (MeterConsump& - RateTbl.TblBreaks(5).UNITS)
'     TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
'     GoTo GotTAmt
'    Else
'      UNITS& = (RateTbl.TblBreaks(6).UNITS - RateTbl.TblBreaks(5).UNITS)
'      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
'    End If
'  Else
'    UNITS& = (MeterConsump& - RateTbl.TblBreaks(5).UNITS)
'    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
'    GoTo GotTAmt
'  End If
'
'  'break 6
'  If LastTblCnt >= 7 Then
'    If MeterConsump& >= RateTbl.TblBreaks(6).UNITS And MeterConsump& <= RateTbl.TblBreaks(7).UNITS Then
'      UNITS& = (MeterConsump& - RateTbl.TblBreaks(6).UNITS)
'      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
'      GoTo GotTAmt
'    Else
'      UNITS& = (RateTbl.TblBreaks(7).UNITS - RateTbl.TblBreaks(6).UNITS)
'      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
'    End If
'  Else
'    UNITS& = (MeterConsump& - RateTbl.TblBreaks(6).UNITS)
'    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
'    GoTo GotTAmt
'  End If
'
'  'break 7
'  If LastTblCnt >= 8 Then
'    If MeterConsump& >= RateTbl.TblBreaks(7).UNITS And MeterConsump& <= RateTbl.TblBreaks(8).UNITS Then
'      UNITS& = (MeterConsump& - RateTbl.TblBreaks(7).UNITS)
'      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
'      GoTo GotTAmt
'    Else
'      UNITS& = (RateTbl.TblBreaks(8).UNITS - RateTbl.TblBreaks(7).UNITS)
'      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
'    End If
'  Else
'    UNITS& = (MeterConsump& - RateTbl.TblBreaks(7).UNITS)
'    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
'    GoTo GotTAmt
'  End If
'  'break 8
'  If LastTblCnt >= 9 Then
'    If MeterConsump& >= RateTbl.TblBreaks(8).UNITS And MeterConsump& <= RateTbl.TblBreaks(9).UNITS Then
'      UNITS& = (MeterConsump& - RateTbl.TblBreaks(8).UNITS)
'      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
'      GoTo GotTAmt
'    Else
'      UNITS& = (RateTbl.TblBreaks(9).UNITS - RateTbl.TblBreaks(8).UNITS)
'      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
'    End If
'  Else
'    UNITS& = (MeterConsump& - RateTbl.TblBreaks(8).UNITS)
'    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
'    GoTo GotTAmt
'  End If
'
'  'break 9
'  If LastTblCnt >= 10 Then
'    If MeterConsump& >= RateTbl.TblBreaks(9).UNITS And MeterConsump& <= RateTbl.TblBreaks(10).UNITS Then
'      UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)
'      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
'      GoTo GotTAmt
'    Elseif MeterConsump& < RateTbl.TblBreaks(9).UNITS then
'      UNITS& = (RateTbl.TblBreaks(10).UNITS - RateTbl.TblBreaks(9).UNITS)
'      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
'    Elseif MeterConsump& > RateTbl.TblBreaks(10).UNITS then
'      UNITS& = (MeterConsump&-RateTbl.TblBreaks(10).UNITS)
'      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(10).UNITAMT))
'    End If
'  Else
'    UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)
'    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
'    GoTo GotTAmt
'  End If
'
'GotTAmt:
'  GetRevCharge# = Round#(MinBillAmt# + TAmt#)
'
'End Function
'
