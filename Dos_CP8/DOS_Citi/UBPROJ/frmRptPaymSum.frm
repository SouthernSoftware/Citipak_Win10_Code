VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptPaymSum 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Summary Report"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmRptPaymSum.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5508
      TabIndex        =   3
      Top             =   4908
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6371
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
      ColDesigner     =   "frmRptPaymSum.frx":08CA
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
      Left            =   10080
      TabIndex        =   5
      Top             =   7560
      Width           =   1332
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
      Left            =   8400
      TabIndex        =   4
      Top             =   7560
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
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
            TextSave        =   "11:44 AM"
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
      Left            =   5508
      TabIndex        =   1
      Top             =   3840
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
      Left            =   5508
      TabIndex        =   0
      Top             =   3324
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
   Begin EditLib.fpText fptxtOperator 
      Height          =   348
      Left            =   5508
      TabIndex        =   2
      Top             =   4368
      Width           =   732
      _Version        =   196608
      _ExtentX        =   1291
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
      ThreeDOutsideStyle=   2
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "(0 = All Operators)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   6408
      TabIndex        =   12
      Top             =   4464
      Width           =   1740
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   2772
      Left            =   2640
      Top             =   2928
      Width           =   6972
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
      Height          =   372
      Index           =   0
      Left            =   3846
      TabIndex        =   11
      Top             =   3882
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
      Height          =   372
      Left            =   3750
      TabIndex        =   10
      Top             =   3366
      Width           =   1668
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
      Left            =   3078
      TabIndex        =   9
      Top             =   4950
      Width           =   2388
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator No:"
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
      Index           =   2
      Left            =   3342
      TabIndex        =   8
      Top             =   4434
      Width           =   2076
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3210
      Top             =   1320
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Payment Summary"
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
      Left            =   3618
      TabIndex        =   7
      Top             =   1560
      Width           =   5004
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3210
      Top             =   1200
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
Attribute VB_Name = "frmRptPaymSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Private Sub cmdExit_Click()
  frmUBReportsMenu.Show
  Unload frmRptPaymSum
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
  fptxtOperator = "0"
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

Private Sub fptxtOperator_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
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
Private Sub txtDate1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDate2.SetFocus
  End If
End Sub

Private Sub txtDate2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtOperator.SetFocus
  End If
End Sub
Private Sub fptxtOperator_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboRptType.SetFocus
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
        fptxtOperator.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub cmdPrint_Click()
  DeActivateControls Me, True
  If fpcboRptType.ListIndex = 0 Then
    PaymentSumReport2
  ElseIf fpcboRptType.ListIndex = 1 Then
    PaymentSumReport
  End If
  ActivateControls Me, True
End Sub

Private Sub PaymentSumReport()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim CustomerCnt As Long, UBTransRecLen As Integer, UBTrans As Integer
  Dim IndexName As String, Handle As Integer, Dash80 As String
  Dim IdxRecLen As Integer, IdxFileSize As Long, UBRpt As Integer
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, MaxRevenue As Integer
  Dim cnt As Long, UBCust As Integer, RCnt As Integer, UseType As Boolean
  Dim ThisType As String, CUSTTYPE As String, CustBook As Integer
  Dim Trans As Long, UBTransLen As Integer, BegDate As Integer
  Dim EndDate As Integer, BegOperator As Integer, EndOperator As Integer
  Dim BegTrans As Integer, EndTrans As Integer, TransDesc As String
  Dim Amount As Double, TotalTrans As Double, TransCnt As Long
  Dim Detail As String, Date1 As String, Date2 As String, Operator As String
  Dim TotalRevsAmt As Double, EstCnt As Integer, TrType As String
  Dim TrTyp As Integer, OperatorNum As String, TRevName As String
  Dim FromDate As String, ToDate As String, TransOK As Boolean
  Dim TaxExempt As Boolean, RevCnt As Integer, Diff As Double
  Dim Tax As Double, TaxTotal As Double, Reportfile As String
  ReDim RevenueName$(15)
  ReDim Revenues(1 To 15) As Double
  ReDim TaxRates(1 To 15) As Single
  ReDim TaxAmt(1 To 15) As Double
  FrmShowPctComp.Label1 = "Creating Payment Summary Report"
  FrmShowPctComp.Show , Me

  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))

  ReDim UBSetUp(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetupreclen

  For RCnt = 1 To 15
    TRevName$ = QPTrim$(UBSetUp(1).Revenues(RCnt).REVNAME)
    If Len(TRevName$) > 0 Then
      RevenueName$(RCnt) = TRevName$
      TaxRates(RCnt) = UBSetUp(1).Revenues(RCnt).TAXRATE
    Else
      MaxRevenue = RCnt - 1
      Exit For
    End If
  Next

  BegDate = Date2Num(txtDate1)
  EndDate = Date2Num(txtDate2)

  '***************
  ' Set Up Specifications from Input Screen
  OperatorNum$ = fptxtOperator
  Operator = Val(OperatorNum$)
  FromDate$ = txtDate1
  ToDate$ = txtDate2

  If Operator = 0 Then
    BegOperator = 0
    EndOperator = 99
  Else
    BegOperator = Operator
    EndOperator = Operator
  End If

  '***************
  MaxLines = 55
  PageNo = 0
  Dash80$ = String$(80, "-")
  Reportfile$ = UBPath$ + "UBPAYSUM.RPT"
  UBRpt = FreeFile
  Open Reportfile$ For Output As UBRpt

  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen

  NumOfRecs& = LOF(UBTrans) \ UBTransRecLen

'  BlockClear
'  ShowProcessingScrn "Payment Summary Report."

  GoSub DoDetailedRptHeader3

  For cnt& = 1 To NumOfRecs&
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      GoTo ExitDetailedListing3
    End If

    Get UBTrans, cnt&, UBTransRec(1)
    TransOK = False

   ' IF UBTrans(1).CustAcctNo < 0 THEN
   '   LPRINT "Trans: "; Cnt&
   '   GOTO SkipEm:
   ' END IF
    If (UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate) And (UBTransRec(1).OperatorNumber >= BegOperator And UBTransRec(1).OperatorNumber <= EndOperator) Then
      Select Case UBTransRec(1).TransType
      Case TranBillPayment, TranBillPayment + 100
        TransOK = True
      Case TranDraftPayment, TranDraftPayment + 100
        TransOK = True
      'CASE TranDepositPayment, TranDepositPayment + 100
      '  TransOK = True
      End Select

      If TransOK Then
        If UBTransRec(1).TaxExempt = "Y" Then
          TaxExempt = True
          GoTo SkipEm
          'LPRINT UBTrans(1).CustAcctNo
        Else
          TaxExempt = False
        End If

        For RevCnt = 1 To 15
          If Not TaxExempt Then
            If TaxRates(RevCnt) > 0 Then
              Diff# = Round#(UBTransRec(1).RevAmt(RevCnt) / (1 + TaxRates(RevCnt)))
              Tax# = Round#(UBTransRec(1).RevAmt(RevCnt) - Diff#)
              TaxAmt(RevCnt) = Round#(TaxAmt(RevCnt) + Tax#)
              Revenues(RevCnt) = Round#(Revenues(RevCnt) + (UBTransRec(1).RevAmt(RevCnt) - Tax#))
            Else
              Revenues(RevCnt) = Round#(Revenues(RevCnt) + UBTransRec(1).RevAmt(RevCnt))
            End If
          Else
            Revenues(RevCnt) = Round#(Revenues(RevCnt) + UBTransRec(1).RevAmt(RevCnt))
          End If
        Next
        TransCnt& = TransCnt& + 1
      End If
    End If

SkipEm:
   ' ShowPctCompL cnt&, NumOfRecs&
  Next

  GoSub DoDetailedRptFooter3

  Close

'  If Not AbortFlag Then
'    PrintRptFile , "UBPAYSUM.RPT", 1, RetCode, EntryP
'  End If
  ViewPrint Reportfile$, "Payment Summary Report."
  KillFile Reportfile$

ExitDetailedListing3:

  Exit Sub

DoDetailedRptHeader3:
  Print #UBRpt, TownName$
  Print #UBRpt, Tab(29); "Payment Summary Report"
  Print #UBRpt, "Beginning Date: "; FromDate$;
  If Val(OperatorNum$) = 0 Then
    Print #UBRpt, Tab(65); " Operator #: ALL"
  Else
    Print #UBRpt, Tab(65); " Operator #: "; OperatorNum$
  End If
  Print #UBRpt, "   Ending Date: "; ToDate$
  Print #UBRpt,
  Print #UBRpt, "    Source                           Revenue Amt                 Tax"
Return

DoDetailedRptFooter3:
  Print #UBRpt, Dash80$
  For cnt = 1 To MaxRevenue
    Print #UBRpt, Tab(5); RevenueName$(cnt); Tab(35); Using("$##,###,###.##", Revenues(cnt)); Tab(56); Using("$##,###,###.##", TaxAmt(cnt))
    TotalTrans# = Round#(TotalTrans# + Revenues(cnt))
    TaxTotal# = Round#(TaxTotal# + TaxAmt(cnt))
  Next
  Print #UBRpt, Dash80$
  Print #UBRpt, "Total Payments: "; Tab(20); Using("######", TransCnt&)
  Print #UBRpt, "Revenue Totals: "; Tab(35); Using("$##,###,###.##", TotalTrans#); Tab(56); Using("$##,###,###.##", TaxTotal#)
  Print #UBRpt, FF$
Return
End Sub

Private Sub PaymentSumReport2()
  Dim UBCustRecLen As Integer, UBSetupreclen As Integer
  Dim CustomerCnt As Long, UBTransRecLen As Integer, UBTrans As Integer
  Dim IndexName As String, Handle As Integer
  Dim IdxRecLen As Integer, IdxFileSize As Long, UBRpt As Integer
  Dim IdxNumOfRecs As Long, NumOfRecs As Long, MaxRevenue As Integer
  Dim cnt As Long, UBCust As Integer, RCnt As Integer, UseType As Boolean
  Dim ThisType As String, CUSTTYPE As String, CustBook As Integer
  Dim Trans As Long, UBTransLen As Integer, BegDate As Integer
  Dim EndDate As Integer, BegOperator As Integer, EndOperator As Integer
  Dim BegTrans As Integer, EndTrans As Integer, TransDesc As String
  Dim Amount As Double, TotalTrans As Double, TransCnt As Long
  Dim Detail As String, Date1 As String, Date2 As String, Operator As String
  Dim TotalRevsAmt As Double, EstCnt As Integer, TrType As String
  Dim TrTyp As Integer, OperatorNum As String, TRevName As String
  Dim FromDate As String, ToDate As String, TransOK As Boolean
  Dim TaxExempt As Boolean, RevCnt As Integer, Diff As Double
  Dim Tax As Double, TaxTotal As Double, ToPrint As String, Reportfile As String
  ReDim RevenueName$(15)
  ReDim Revenues(1 To 15) As Double
  ReDim TaxRates(1 To 15) As Single
  ReDim TaxAmt(1 To 15) As Double
  FrmShowPctComp.Label1 = "Creating Payment Summary Report"
  FrmShowPctComp.Show , Me

  ReDim UBTransRec(1) As UBTransRecType
  UBTransRecLen = Len(UBTransRec(1))

  ReDim UBSetUp(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetupreclen

  For RCnt = 1 To 15
    TRevName$ = QPTrim$(UBSetUp(1).Revenues(RCnt).REVNAME)
    If Len(TRevName$) > 0 Then
      RevenueName$(RCnt) = TRevName$
      TaxRates(RCnt) = UBSetUp(1).Revenues(RCnt).TAXRATE
    Else
      MaxRevenue = RCnt - 1
      Exit For
    End If
  Next

  BegDate = Date2Num(txtDate1)
  EndDate = Date2Num(txtDate2)

  '***************
  ' Set Up Specifications from Input Screen
  OperatorNum$ = fptxtOperator
  Operator = Val(OperatorNum$)
  FromDate$ = txtDate1
  ToDate$ = txtDate2

  If Operator = 0 Then
    BegOperator = 0
    EndOperator = 99
  Else
    BegOperator = Operator
    EndOperator = Operator
  End If

  '***************
  Reportfile$ = UBPath$ + "UBPAYSUM.RPT"
  UBRpt = FreeFile
  Open Reportfile$ For Output As UBRpt

  UBTrans = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTrans Len = UBTransRecLen

  NumOfRecs& = LOF(UBTrans) \ UBTransRecLen

  For cnt& = 1 To NumOfRecs&
    FrmShowPctComp.ShowPctComp cnt, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      GoTo ExitDetailedListing3
    End If

    Get UBTrans, cnt&, UBTransRec(1)
    TransOK = False

   ' IF UBTrans(1).CustAcctNo < 0 THEN
   '   LPRINT "Trans: "; Cnt&
   '   GOTO SkipEm:
   ' END IF
    If (UBTransRec(1).TransDate >= BegDate And UBTransRec(1).TransDate <= EndDate) And (UBTransRec(1).OperatorNumber >= BegOperator And UBTransRec(1).OperatorNumber <= EndOperator) Then
      Select Case UBTransRec(1).TransType
      Case TranBillPayment, TranBillPayment + 100
        TransOK = True
      Case TranDraftPayment, TranDraftPayment + 100
        TransOK = True
      'CASE TranDepositPayment, TranDepositPayment + 100
      '  TransOK = True
      End Select

      If TransOK Then
        If UBTransRec(1).TaxExempt = "Y" Then
          TaxExempt = True
          GoTo SkipEm
          'LPRINT UBTrans(1).CustAcctNo
        Else
          TaxExempt = False
        End If

        For RevCnt = 1 To 15
          If Not TaxExempt Then
            If TaxRates(RevCnt) > 0 Then
              Diff# = Round#(UBTransRec(1).RevAmt(RevCnt) / (1 + TaxRates(RevCnt)))
              Tax# = Round#(UBTransRec(1).RevAmt(RevCnt) - Diff#)
              TaxAmt(RevCnt) = Round#(TaxAmt(RevCnt) + Tax#)
              Revenues(RevCnt) = Round#(Revenues(RevCnt) + (UBTransRec(1).RevAmt(RevCnt) - Tax#))
            Else
              Revenues(RevCnt) = Round#(Revenues(RevCnt) + UBTransRec(1).RevAmt(RevCnt))
            End If
          Else
            Revenues(RevCnt) = Round#(Revenues(RevCnt) + UBTransRec(1).RevAmt(RevCnt))
          End If
        Next
        TransCnt& = TransCnt& + 1
      End If
    End If

SkipEm:
   ' ShowPctCompL cnt&, NumOfRecs&
  Next

  GoSub DoDetailedRptFooter3

  Close
  If Val(OperatorNum$) = 0 Then
    OperatorNum$ = "ALL"
  End If
  If TransCnt& > 0 Then
    Load frmLoadingRpt
      ARptPaymSum.txtDate = Now
      ARptPaymSum.txtTown = TownName$
      ARptPaymSum.Title = "Payment Summary Report"
      ARptPaymSum.txtDateBeg = FromDate$
      ARptPaymSum.txtDateEnd = ToDate$
      ARptPaymSum.txtoperator = OperatorNum$
      ARptPaymSum.totTrans = TransCnt&
      ARptPaymSum.GetName Reportfile$
      ARptPaymSum.startrpt
  Else
    MsgBox "No Information to print.", vbOKOnly, "No Information"
  End If

'  If Not AbortFlag Then
'    PrintRptFile , "UBPAYSUM.RPT", 1, RetCode, EntryP
'  End If
'  ViewPrint , "Payment Summary Report."
'  KillFile "UBPAYSUM.RPT"

ExitDetailedListing3:

  Exit Sub

'DoDetailedRptHeader3:
'  Print #UBRpt, TownName$
'  Print #UBRpt, Tab(29); "Payment Summary Report"
'  Print #UBRpt, "Beginning Date: "; FromDate$;
'  If Val(OperatorNum$) = 0 Then
'    Print #UBRpt, Tab(65); " Operator #: ALL"
'  Else
'    Print #UBRpt, Tab(65); " Operator #: "; OperatorNum$
'  End If
'  Print #UBRpt, "   Ending Date: "; ToDate$
''  Print #UBRpt,
''  Print #UBRpt, "    Source                           Revenue Amt                 Tax"
'Return

DoDetailedRptFooter3:
  'ToPrint$ = ""
  For cnt = 1 To MaxRevenue
    Print #UBRpt, RevenueName$(cnt) + "~" + Using("$##,###,###.##", Revenues(cnt)) + "~" + Using("$##,###,###.##", TaxAmt(cnt))
    TotalTrans# = Round#(TotalTrans# + Revenues(cnt))
    TaxTotal# = Round#(TaxTotal# + TaxAmt(cnt))
  Next
'  Print #UBRpt, "Total Payments: "; Tab(20); Using("######", TransCnt&)
'  Print #UBRpt, "Revenue Totals: "; Tab(35); Using("$##,###,###.##", TotalTrans#); Tab(56); Using("$##,###,###.##", TaxTotal#)
Return
End Sub

