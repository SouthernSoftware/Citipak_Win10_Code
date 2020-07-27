VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmRptCustTranHist 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Transaction History"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   2172
   ClientWidth     =   12192
   ClipControls    =   0   'False
   Icon            =   "frmRptCustTranHist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12192
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   348
      Left            =   5340
      TabIndex        =   1
      Top             =   4992
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
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
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
      ColDesigner     =   "frmRptCustTranHist.frx":08CA
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   336
      Left            =   0
      TabIndex        =   4
      Top             =   8532
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   593
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
   Begin EditLib.fpText fpCustName 
      Height          =   348
      Left            =   5340
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3960
      Width           =   3996
      _Version        =   196608
      _ExtentX        =   7048
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
      ButtonWrap      =   0   'False
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
      UserEntry       =   1
      HideSelection   =   0   'False
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
   Begin EditLib.fpBoolean fpDetailFlag 
      Height          =   348
      Left            =   5340
      TabIndex        =   0
      Top             =   4488
      Width           =   324
      _Version        =   196608
      _ExtentX        =   572
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      AutoToggle      =   -1  'True
      BooleanStyle    =   1
      ToggleFalse     =   "Nn"
      TextFalse       =   "N"
      BooleanPicture  =   0
      AlignPictureH   =   3
      AlignPictureV   =   1
      GroupId         =   0
      GroupTag        =   0
      GroupSelect     =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      MultiLine       =   0   'False
      AlignTextH      =   1
      AlignTextV      =   0
      ToggleTrue      =   "Yy"
      TextTrue        =   "Y"
      Value           =   1
      BooleanMode     =   0
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      BorderGrayAreaColor=   -2147483637
      ToggleGrayed    =   ""
      TextGrayed      =   ""
      AllowMnemonic   =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDOnFocusInvert=   0   'False
      Caption         =   "N"
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      BooleanDataType =   0
      OLEDropMode     =   0
   End
   Begin EditLib.fpLongInteger fpCustRecNo 
      Height          =   300
      Left            =   744
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1488
      Visible         =   0   'False
      Width           =   684
      _Version        =   196608
      _ExtentX        =   1206
      _ExtentY        =   529
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
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
      AllowNull       =   -1  'True
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
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
      MaxValue        =   "2147483647"
      MinValue        =   "-2147483648"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   1
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdPrint 
      Height          =   480
      Left            =   9300
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
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
      ButtonDesigner  =   "frmRptCustTranHist.frx":0C68
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdExit 
      Height          =   480
      Left            =   7656
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
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
      ButtonDesigner  =   "frmRptCustTranHist.frx":0E45
   End
   Begin VB.Label Label2 
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
      Left            =   2856
      TabIndex        =   10
      Top             =   5016
      Width           =   2388
   End
   Begin VB.Label DetailLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Detail:"
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
      Left            =   4416
      TabIndex        =   8
      Top             =   4536
      Width           =   756
   End
   Begin VB.Label PromptLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer:"
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
      Left            =   3504
      TabIndex        =   7
      Top             =   4008
      Width           =   1668
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   2316
      Left            =   2580
      Top             =   3456
      Width           =   7044
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3210
      Top             =   1368
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Customer Transaction History"
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
      Left            =   3492
      TabIndex        =   5
      Top             =   1608
      Width           =   5220
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      Height          =   972
      Left            =   3210
      Top             =   1248
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
Attribute VB_Name = "frmRptCustTranHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim RecNo As Long, AcctNum As Long
Dim fromform As Form, toform As Form, codeopt As Integer
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
End Sub

Private Sub Form_Activate()
  GetName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If fpCmdExit.Enabled = False Then
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
Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpCmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpCmdExit.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub
Private Sub mnuExit_Click()
  fpCmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub fpCmdExit_Click()
 ' Load frmUBCustMenu
  DoEvents
  If codeopt = 1 Then
    ActivateControls frmCustEditLookUP
  ElseIf codeopt = 2 Then
    ActivateControls frmDisplayList
  End If

 ' frmUBCustMenu.Show
  Unload frmRptCustTranHist
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      KeyCode = 0
      Call fpCmdExit_Click
    Case vbKeyF10, vbKeyReturn
      KeyCode = 0
      Call fpCmdPrint_Click
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TownName$
  fpDetailFlag = ValueTrue
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  'GetName
End Sub
Private Sub GetName()
  ReDim UBCustRec(1) As NewUBCustRecType
  Dim UBCustRecLen As Integer, UBSetupLen As Integer, UBCust As Integer
  RecNo& = fpCustRecNo
  UBCustRecLen = Len(UBCustRec(1))
  UBCust = FreeFile
  Open UBCustFile For Random Shared As UBCust Len = UBCustRecLen
  Get #UBCust, RecNo&, UBCustRec(1)
  Close UBCust
  fpCustName = UBCustRec(1).CustName
End Sub


Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
  
End Sub
Private Sub fpCmdPrint_Click()
DeActivateControls Me, True
If fpcboRptType.ListIndex = 0 Then
  'do graphics
  CustTransHistoryRpt2
ElseIf fpcboRptType.ListIndex = 1 Then
  'do text
  CustTransHistoryRpt
End If
ActivateControls Me, True
fpCmdExit_Click
End Sub

'***************************************
'consumption hist report moved to frmcustconshist
'**************************************

Private Sub CustTransHistoryRpt() 'Text Report
  Dim t As String, Dash80 As String
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBTranRec(1) As UBTransRecType
  ReDim UBSetUpRec(1) As UBSetupRecType
  
  ReDim TotalConsump(1 To 7) As Long
  ReDim DidCnt(1 To 7) As Integer
    
  Dim UBCustRecLen As Integer, UBTranRecLen As Integer
  Dim UBSetupLen As Integer, RevCnt As Integer, Rev2Flag As Integer
  Dim TempRev As String, ReportFile As String
  Dim RevText1 As String, RevText2 As String
  Dim UBCust As Integer, UBRpt As Integer, UBTran As Integer
  Dim Cubic As Integer, MTCnt As Integer, NumOfRevs As Integer
  Dim ThisTrans As Long, MeterConsp As Long, MaxMeterAmt As Long
  Dim FirstTrans As Integer, TYear As Integer
  Dim LastDate As String, MeterType As String
  Dim PDate As Integer, AbortFlag As Integer
  Dim DidEst As Integer, EstCnt As Integer
  Dim DetailFlag As Integer, DidAMeter As Integer
  Dim MtrCnt As Integer, WhatMtrCNT As Integer
  Dim PrintedOne As Integer, TabStop As Integer
  Dim RevOffset As Integer
  
  ReportFile$ = UBPath$ + "UBTRAHIS.RPT"
  DetailFlag = frmRptCustTranHist.fpDetailFlag.Text = "Y"
  
  t$ = Space$(10)
  MaxLines = 40

  Dash80$ = String$(80, "-")
  UBCustRecLen = Len(UBCustRec(1))
  UBTranRecLen = Len(UBTranRec(1))

  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  For RevCnt = 1 To MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME)
    If Len(TempRev$) = 0 Then
      NumOfRevs = RevCnt - 1
      Exit For
    Else
      RSet t$ = QPTrim$(Left$(TempRev$, 8))
      If RevCnt <= 8 Then
        RevText1$ = RevText1$ + t$
      Else
        RevText2$ = RevText2$ + t$
      End If
    End If
  Next

  If Len(QPTrim$(RevText2$)) > 0 Then
    Rev2Flag = True
  End If
  
  UBCust = FreeFile
  Open UBCustFile For Random Shared As UBCust Len = UBCustRecLen
  Get #UBCust, RecNo&, UBCustRec(1)
  Close UBCust

  For MTCnt = 1 To 7
    If UBCustRec(1).LocMeters(MTCnt).MTRUnit = "C" Then
      Cubic = True
      Exit For
    End If
  Next

  UBRpt = FreeFile
  Open ReportFile For Output As UBRpt

  UBTran = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen

  GoSub DOTranHistHeader

  ThisTrans& = UBCustRec(1).LastTrans
  
  FirstTrans = True

  Do While ThisTrans& > 0
    Get #UBTran, ThisTrans&, UBTranRec(1)
      If FirstTrans Then
        LastDate$ = Num2Date$(UBTranRec(1).TransDate)
        TYear = Val(Right$(LastDate$, 4))
        PDate = Date2Num(Left$(LastDate$, 3) + "01-" + QPTrim$(Str$(TYear - 1)))
        FirstTrans = False
      End If
      
      GoSub DOTransDetail
      Print #UBRpt, Dash80$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #UBRpt, FF$
        GoSub DOTranHistHeader
      End If
    ThisTrans& = UBTranRec(1).PrevTrans
  Loop
  GoSub DOTranHistFooter

  Close
  If Not AbortFlag Then
    ViewPrint ReportFile$, "Customer Transaction Report."
  End If

ExitTransHist:
Exit Sub

DOTransDetail:
  Print #UBRpt, Num2Date(UBTranRec(1).TransDate);

  Select Case UBTranRec(1).TransType
    Case TranUtilityBill, TranUtilityBill + 100
      DidEst = False
      For EstCnt = 1 To 7
        If UBTranRec(1).ESTREAD(EstCnt) = "Y" Then
          DidEst = True
          Exit For
        End If
      Next

      Print #UBRpt, Tab(16); "Utility Bill";
      If DidEst Then
        Print #UBRpt, "*e";
      End If
      If DetailFlag Then
        Print #UBRpt, Tab(31); Num2Date$(UBTranRec(1).ReadDate); Tab(43); Num2Date$(UBTranRec(1).PrevDate);
      End If
      Print #UBRpt, Tab(57); Using("$#####.##", UBTranRec(1).TransAmt); Tab(71); Using("$#####.##", UBTranRec(1).RunBalance);
      If DetailFlag Then
        GoSub DoMtrDetail
        GoSub PrintRevDetail
      Else
        Print #UBRpt,
      End If
    Case TranLateCharge, TranLateCharge + 100
      Print #UBRpt, Tab(16); "Late Charge"; Tab(57); Using$("$#####.##", UBTranRec(1).TransAmt); Tab(71); Using("$#####.##", UBTranRec(1).RunBalance)
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranReconnectFee, TranReconnectFee + 100
      Print #UBRpt, Tab(16); "Reconnect Fee";
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranBillPayment, TranBillPayment + 100
      Print #UBRpt, Tab(16); "Bill Payment"; Tab(57); Using$("$#####.##", UBTranRec(1).TransAmt); Tab(71); Using("$#####.##", UBTranRec(1).RunBalance)
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranAppliedDeposit, TranAppliedDeposit + 100
      Print #UBRpt, Tab(16); "Applied Deposit"; Tab(57); Using$("$#####.##", UBTranRec(1).TransAmt); Tab(67); Using("$#####.##", UBTranRec(1).RunBalance)
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranPenaltyCharge, TranPenaltyCharge + 100
      Print #UBRpt, Tab(16); "Penalty Charge"; Tab(57); Using$("$#####.##", UBTranRec(1).TransAmt); Tab(71); Using("$#####.##", UBTranRec(1).RunBalance)
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranDepositPayment, TranDepositPayment + 100
      Print #UBRpt, Tab(16); "Deposit Payment"; Tab(57); Using$("$#####.##", UBTranRec(1).TransAmt); Tab(71); Using("$#####.##", UBTranRec(1).RunBalance)
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranDraftPayment, TranDraftPayment + 100
      Print #UBRpt, Tab(16); "Draft Payment";
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranRefundDeposit, TranRefundDeposit + 100
      Print #UBRpt, Tab(16); "Refund Deposit"; Tab(57); Using$("$#####.##", UBTranRec(1).TransAmt); Tab(71); Using("$#####.##", UBTranRec(1).RunBalance)
    Case TranBeginBalance, TranBeginBalance + 100
      Print #UBRpt, Tab(16); "Beginning Balance";
    Case TranUpwardAdjustment, TranUpwardAdjustment + 100
      Print #UBRpt, Tab(16); "UP Adjustment  " + Left$(UBTranRec(1).BillMsg, 25); Tab(57); Using$("$#####.##", UBTranRec(1).TransAmt); Tab(71); Using("$#####.##", UBTranRec(1).RunBalance)
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranDownwardAdjustment, TranDownwardAdjustment + 100
      Print #UBRpt, Tab(16); "DN Adjustment  " + Left$(UBTranRec(1).BillMsg, 25); Tab(57); Using$("$#####.##", UBTranRec(1).TransAmt); Tab(71); Using("$#####.##", UBTranRec(1).RunBalance)
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
    Case TranMiscPayment, TranMiscPayment + 100
      Print #UBRpt, Tab(16); "Misc Payment"
      If DetailFlag Then
        GoSub PrintRevDetail
      End If
  End Select
skipit:
Return

DoMtrDetail:
  DidAMeter = False
  For MtrCnt = 1 To 7
    If UBTranRec(1).MtrTypes(MtrCnt) > 0 Then
      DidAMeter = True
      Select Case UBTranRec(1).MtrTypes(MtrCnt)
      Case MtrWaterOnly
        MeterType$ = "      Water"
      Case MtrSewerOnly
        MeterType$ = "      Sewer"
      Case MtrCombined
        MeterType$ = "Water/Sewer"
      Case MtrElectric
        MeterType$ = "   Electric"
      Case MtrDemand
        MeterType$ = " D Electric"
      Case MtrGas
        MeterType$ = "  Gas Meter"
      Case MtrTouchRead
        MeterType$ = " Touch Read"
      Case MtrLightsService
        MeterType$ = "  L Service"
      End Select
      WhatMtrCNT = UBTranRec(1).MtrTypes(MtrCnt)
      If WhatMtrCNT = 0 Then
        WhatMtrCNT = 1
      End If
      GoSub PrintMtrDetail
    End If
  Next
  If Not DidAMeter Then
    MeterType$ = "        "
    MtrCnt = 1
    GoSub PrintMtrDetail
  End If
Return

PrintMtrDetail:
  Print #UBRpt, Tab(16); MeterType$;
  Print #UBRpt, Tab(31); Using("##########", UBTranRec(1).CurRead(MtrCnt));
  Print #UBRpt, Tab(43); Using("##########", UBTranRec(1).PrevRead(MtrCnt));
  MeterConsp& = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
  If MeterConsp& < 0 Then
    MaxMeterAmt& = 10& ^ (Len(Str$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
    MeterConsp& = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
  End If
  If Cubic Then
    MeterConsp& = MeterConsp& * 7.481
  End If
  Print #UBRpt, Tab(57); Using$("##########", MeterConsp&)
  If DidAMeter Then
    TotalConsump(WhatMtrCNT) = TotalConsump(WhatMtrCNT) + MeterConsp&
    DidCnt(WhatMtrCNT) = DidCnt(WhatMtrCNT) + 1
  End If
  Linecnt = Linecnt + 1
Return

PrintRevDetail:
    PrintedOne = False
    For RevCnt = 0 To 7
      If UBTranRec(1).RevAmt(RevCnt + 1) <> 0 Then
        PrintedOne = True
        TabStop = (RevCnt * 10) + 1
        Print #UBRpt, Tab(TabStop); Using$("#######.##", UBTranRec(1).RevAmt(RevCnt + 1));
      End If
    Next
    If PrintedOne Then
      Print #UBRpt,
      Linecnt = Linecnt + 1
    End If
    RevOffset = 7
    PrintedOne = False
    For RevCnt = 0 To 6
      If UBTranRec(1).RevAmt(RevCnt + 1 + RevOffset) <> 0 Then
        PrintedOne = True
        TabStop = (RevCnt * 10) + 1
        Print #UBRpt, Tab(TabStop); Using$("#######.##", UBTranRec(1).RevAmt(RevCnt + 1 + RevOffset));
      End If
    Next
    If PrintedOne Then
      Print #UBRpt,
      Linecnt = Linecnt + 1
    End If

Return

DOTranHistHeader:
  Linecnt = 7
  Print #UBRpt, Tab(28); "Transaction History Report. "
  Print #UBRpt, "Customer: "; UBCustRec(1).CustName; Tab(57); "Report Date: "; Date$
  If DetailFlag Then
    Print #UBRpt, " Account:"; RecNo&
    Print #UBRpt, "Ser Addr: "; UBCustRec(1).SERVADDR
    Print #UBRpt, "Location: "; QPTrim$(UBCustRec(1).Book); "-"; QPTrim$(UBCustRec(1).SEQNUMB)
    Linecnt = Linecnt + 2
    For MtrCnt = 1 To 7
      If Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)) > 0 Then
        Print #UBRpt, Tab(6); "Mtr# "; QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
        Linecnt = Linecnt + 1
      End If
    Next
  End If
  Print #UBRpt,
  If DetailFlag Then
    Print #UBRpt, "Trans Date     Trans Type     Cur.Date     Pre.Date      TR Amount      Balance"
  Else
    Print #UBRpt, "Trans Date     Trans Type                                TR Amount      Balance"
  End If
  If DetailFlag Then
    Print #UBRpt, "               Meter Type     Cur.Read     Pre.Read       Usage"
  End If
  If DetailFlag Then
    Print #UBRpt, RevText1$
    If Rev2Flag Then
      Print #UBRpt, RevText2$
      Linecnt = 8
    End If
  Else
    Linecnt = 5
  End If
  Print #UBRpt, Dash80$
Return

DOTranHistFooter:
  If FirstTrans Then
    Print #UBRpt, "NO TRANSACTIONS!!!"
    Print #UBRpt, Dash80$
  End If
  For MtrCnt = 1 To 7
    If DidCnt(MtrCnt) > 0 Then
      Print #UBRpt, "Average Consumption: "; Using$("#########", TotalConsump(MtrCnt) / DidCnt(MtrCnt))
    End If
  Next
  Print #UBRpt, FF$
Return

End Sub
Private Sub CustTransHistoryRpt2() 'Graphics Report
  Dim t As String, ToPrintM As String, ToPrintD As String
  Dim PrnMtr(1 To 7) As String, PrnRev(1 To 15) As String, ToPrintH As String
  Dim PrnLoc As String, PrnSvcAddr As String, ToPrintR As String
  Dim UBCustRecLen As Integer, UBTranRecLen As Integer
  Dim UBSetupLen As Integer, RevCnt As Integer, Rev2Flag As Integer
  Dim TempRev As String, ReportFile As String, PrnAvg As String
  Dim RevText1 As String, RevText2 As String
  Dim UBCust As Integer, UBRpt As Integer, UBTran As Integer
  Dim Cubic As Integer, MTCnt As Integer, NumOfRevs As Integer
  Dim ThisTrans As Long, MeterConsp As Long, MaxMeterAmt As Long
  Dim FirstTrans As Integer, TYear As Integer
  Dim LastDate As String, MeterType As String
  Dim PDate As Integer, AbortFlag As Integer
  Dim DidEst As Integer, EstCnt As Integer
  Dim DetailFlag As Integer, DidAMeter As Integer
  Dim MtrCnt As Integer, WhatMtrCNT As Integer
  Dim PrintedOne As Integer, TabStop As Integer
  Dim RevOffset As Integer, Detail As Boolean
   ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBTranRec(1) As UBTransRecType
  ReDim UBSetUpRec(1) As UBSetupRecType
  
  ReDim TotalConsump(1 To 7) As Long
  ReDim DidCnt(1 To 7) As Integer
 
  ReportFile$ = UBPath$ + "UBTRAHIS.RPT"
  DetailFlag = frmRptCustTranHist.fpDetailFlag.Text = "Y"
  If DetailFlag Then
    Detail = True
  End If
  t$ = Space$(10)
  UBCustRecLen = Len(UBCustRec(1))
  UBTranRecLen = Len(UBTranRec(1))

  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  For RevCnt = 1 To MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME)
    If Len(TempRev$) = 0 Then
      NumOfRevs = RevCnt - 1
      Exit For
    Else
'      RSet t$ = QPTrim$(Left$(TempRev$, 8))
'      If RevCnt <= 8 Then
'        RevText1$ = RevText1$ + t$
'      Else
'        RevText2$ = RevText2$ + t$
'      End If
      PrnRev(RevCnt) = QPTrim$(Left$(TempRev$, 8))
    End If
  Next

'  If Len(QPTrim$(RevText2$)) > 0 Then
'    Rev2Flag = True
'  End If
  
  UBCust = FreeFile
  Open UBCustFile For Random Shared As UBCust Len = UBCustRecLen
  Get #UBCust, RecNo&, UBCustRec(1)
  Close UBCust

  For MTCnt = 1 To 7
    If UBCustRec(1).LocMeters(MTCnt).MTRUnit = "C" Then
      Cubic = True
      Exit For
    End If
  Next

  UBRpt = FreeFile
  Open ReportFile For Output As UBRpt

  UBTran = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen

  GoSub DOTranHistHeader

  ThisTrans& = UBCustRec(1).LastTrans
  
  FirstTrans = True

  Do While ThisTrans& > 0
    Get #UBTran, ThisTrans&, UBTranRec(1)
      If FirstTrans Then
        LastDate$ = Num2Date$(UBTranRec(1).TransDate)
        TYear = Val(Right$(LastDate$, 4))
        PDate = Date2Num(Left$(LastDate$, 3) + "01-" + QPTrim$(Str$(TYear - 1)))
        FirstTrans = False
      End If
      
      GoSub DOTransDetail
      Print #UBRpt, ToPrintH$ + "~" + ToPrintD$ + "~" + ToPrintM$ + ToPrintR$
      ToPrintD$ = ""
      ToPrintM$ = ""
      ToPrintR$ = ""
 '     Linecnt = Linecnt + 1
'      If Linecnt > MaxLines Then
'        Print #UBRpt, FF$
'        GoSub DOTranHistHeader
'      End If
    ThisTrans& = UBTranRec(1).PrevTrans
  Loop
  GoSub DOTranHistFooter

  Close
 ' If Not AbortFlag Then
'   ViewPrint ReportFile$, "Customer Transaction Report."
'  End If
  Load frmLoadingRpt
  ARptCustTranHist.txtDate = Now
  ARptCustTranHist.txtTown = TownName$
  ARptCustTranHist.Title = "Customer Transaction History"
  ARptCustTranHist.txtLocation = PrnLoc$
  ARptCustTranHist.lblSerAddr = PrnSvcAddr$
    ARptCustTranHist.Lab1.Caption = PrnRev(1)
    ARptCustTranHist.Lab2.Caption = PrnRev(2)
    ARptCustTranHist.Lab3.Caption = PrnRev(3)
    ARptCustTranHist.Lab4.Caption = PrnRev(4)
    ARptCustTranHist.Lab5.Caption = PrnRev(5)
    ARptCustTranHist.Lab6.Caption = PrnRev(6)
    ARptCustTranHist.Lab7.Caption = PrnRev(7)
    ARptCustTranHist.Lab8.Caption = PrnRev(8)
    ARptCustTranHist.Lab9.Caption = PrnRev(9)
    ARptCustTranHist.Lab10.Caption = PrnRev(10)
    ARptCustTranHist.Lab11.Caption = PrnRev(11)
    ARptCustTranHist.Lab12.Caption = PrnRev(12)
    ARptCustTranHist.Lab13.Caption = PrnRev(13)
    ARptCustTranHist.Lab14.Caption = PrnRev(14)
    ARptCustTranHist.Lab15.Caption = PrnRev(15)
  
  ARptCustTranHist.totAvg = PrnAvg$
  ARptCustTranHist.GetName ReportFile$, Detail, NumOfRevs
  ARptCustTranHist.startrpt

ExitTransHist:
Exit Sub

DOTransDetail:
  ToPrintD$ = Num2Date(UBTranRec(1).TransDate)

  Select Case UBTranRec(1).TransType
    Case TranUtilityBill, TranUtilityBill + 100
      DidEst = False
      For EstCnt = 1 To 7
        If UBTranRec(1).ESTREAD(EstCnt) = "Y" Then
          DidEst = True
          Exit For
        End If
      Next

      ToPrintD$ = ToPrintD$ + "~" + "Utility Bill"
      If DidEst Then
        ToPrintD$ = ToPrintD$ + "*e"
      End If
      'If DetailFlag Then
        ToPrintD$ = ToPrintD$ + "~" + Num2Date$(UBTranRec(1).ReadDate) + "~" + Num2Date$(UBTranRec(1).PrevDate)
      'End If
      ToPrintD$ = ToPrintD$ + "~" + Using("$#####.##", UBTranRec(1).TransAmt) + "~" + Using("$#####.##", UBTranRec(1).RunBalance)
      'If DetailFlag Then
        GoSub DoMtrDetail
        GoSub PrintRevDetail
      'Else
       ' Print #UBRpt,
      'End If
    Case TranLateCharge, TranLateCharge + 100
      ToPrintD$ = ToPrintD$ + "~" + "Late Charge"
      ToPrintD$ = ToPrintD$ + "~ ~ ~" + Using$("$#####.##", UBTranRec(1).TransAmt) + "~" + Using("$#####.##", UBTranRec(1).RunBalance)
      'If DetailFlag Then
        GoSub DoMtrDetail
        GoSub PrintRevDetail
      'End If
    Case TranReconnectFee, TranReconnectFee + 100
      ToPrintD$ = ToPrintD$ + "~" + "Reconnect Fee" + "~ ~ ~ ~ ~"
      'If DetailFlag Then
        GoSub DoMtrDetail
        GoSub PrintRevDetail
      'End If
    Case TranBillPayment, TranBillPayment + 100
      ToPrintD$ = ToPrintD$ + "~" + "Bill Payment" + "~ ~ ~" + Using$("$#####.##", UBTranRec(1).TransAmt) + "~" + Using("$#####.##", UBTranRec(1).RunBalance)
      'If DetailFlag Then
        GoSub DoMtrDetail
        GoSub PrintRevDetail
      'End If
    Case TranAppliedDeposit, TranAppliedDeposit + 100
      ToPrintD$ = ToPrintD$ + "~" + "Applied Deposit" + "~ ~ ~" + Using$("$#####.##", UBTranRec(1).TransAmt) + "~" + Using("$#####.##", UBTranRec(1).RunBalance)
      'If DetailFlag Then
        GoSub DoMtrDetail
        GoSub PrintRevDetail
      'End If
    Case TranPenaltyCharge, TranPenaltyCharge + 100
      ToPrintD$ = ToPrintD$ + "~" + "Penalty Charge" + "~ ~ ~" + Using$("$#####.##", UBTranRec(1).TransAmt) + "~" + Using("$#####.##", UBTranRec(1).RunBalance)
      'If DetailFlag Then
        GoSub DoMtrDetail
        GoSub PrintRevDetail
      'End If
    Case TranDepositPayment, TranDepositPayment + 100
      ToPrintD$ = ToPrintD$ + "~" + "Deposit Payment" + "~ ~ ~" + Using$("$#####.##", UBTranRec(1).TransAmt) + "~" + Using("$#####.##", UBTranRec(1).RunBalance)
      'If DetailFlag Then
        GoSub DoMtrDetail
        GoSub PrintRevDetail
      'End If
    Case TranDraftPayment, TranDraftPayment + 100
      ToPrintD$ = ToPrintD$ + "~" + "Draft Payment ~ ~ ~ ~ ~"
      'If DetailFlag Then
        GoSub DoMtrDetail
        GoSub PrintRevDetail
      'End If
    Case TranRefundDeposit, TranRefundDeposit + 100
      ToPrintD$ = ToPrintD$ + "~" + "Refund Deposit" + "~ ~ ~" + Using$("$#####.##", UBTranRec(1).TransAmt) + "~" + Using("$#####.##", UBTranRec(1).RunBalance)
        GoSub DoMtrDetail
        GoSub PrintRevDetail

    Case TranBeginBalance, TranBeginBalance + 100
      ToPrintD$ = ToPrintD$ + "~" + "Beginning Balance ~ ~ ~ ~ ~"
        GoSub DoMtrDetail
        GoSub PrintRevDetail

    Case TranUpwardAdjustment, TranUpwardAdjustment + 100
      ToPrintD$ = ToPrintD$ + "~" + "UP Adjustment  ~ ~ " + Left$(UBTranRec(1).BillMsg, 25) + "~" + Using$("$#####.##", UBTranRec(1).TransAmt) + "~" + Using("$#####.##", UBTranRec(1).RunBalance)
      'If DetailFlag Then
        GoSub DoMtrDetail
        GoSub PrintRevDetail
     'End If
    Case TranDownwardAdjustment, TranDownwardAdjustment + 100
      ToPrintD$ = ToPrintD$ + "~" + "DN Adjustment  ~ ~" + Left$(UBTranRec(1).BillMsg, 25) + "~" + Using$("$#####.##", UBTranRec(1).TransAmt) + "~" + Using("$#####.##", UBTranRec(1).RunBalance)
      'If DetailFlag Then
        GoSub DoMtrDetail
        GoSub PrintRevDetail
      'End If
    Case TranMiscPayment, TranMiscPayment + 100
      ToPrintD$ = ToPrintD$ + "~" + "Misc Payment ~ ~ ~ ~ ~"
      'If DetailFlag Then
        GoSub DoMtrDetail
        GoSub PrintRevDetail
      'End If
  End Select
skipit:
Return

DoMtrDetail:
  DidAMeter = False
  For MtrCnt = 1 To 7
    If UBTranRec(1).MtrTypes(MtrCnt) > 0 Then
      DidAMeter = True
      Select Case UBTranRec(1).MtrTypes(MtrCnt)
      Case MtrWaterOnly
        MeterType$ = "Water"
      Case MtrSewerOnly
        MeterType$ = "Sewer"
      Case MtrCombined
        MeterType$ = "Water/Sewer"
      Case MtrElectric
        MeterType$ = "Electric"
      Case MtrDemand
        MeterType$ = "D Electric"
      Case MtrGas
        MeterType$ = "Gas Meter"
      Case MtrTouchRead
        MeterType$ = "Touch Read"
      Case MtrLightsService
        MeterType$ = "L Service"
      End Select
      WhatMtrCNT = UBTranRec(1).MtrTypes(MtrCnt)
      If WhatMtrCNT = 0 Then
        WhatMtrCNT = 1
      End If
      GoSub PrintMtrDetail
    Else
     ToPrintM$ = ToPrintM$ + " ~ ~ ~ ~"
    End If
  Next
'  If Not DidAMeter Then
'    MeterType$ = "        "
'    MtrCnt = 1
'    GoSub PrintMtrDetail
'  End If
Return

PrintMtrDetail:
  ToPrintM$ = ToPrintM$ + MeterType$
  ToPrintM$ = ToPrintM$ + "~" + Using("##########", UBTranRec(1).CurRead(MtrCnt))
  ToPrintM$ = ToPrintM$ + "~" + Using("##########", UBTranRec(1).PrevRead(MtrCnt))
  MeterConsp& = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
  If MeterConsp& < 0 Then
    MaxMeterAmt& = 10& ^ (Len(Str$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
    MeterConsp& = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
  End If
  If Cubic Then
    MeterConsp& = MeterConsp& * 7.481
  End If
  ToPrintM$ = ToPrintM$ + "~" + Using$("##########", MeterConsp&) + "~"
  If DidAMeter Then
    TotalConsump(WhatMtrCNT) = TotalConsump(WhatMtrCNT) + MeterConsp&
    DidCnt(WhatMtrCNT) = DidCnt(WhatMtrCNT) + 1
  End If
  'Linecnt = Linecnt + 1
Return

PrintRevDetail:
    PrintedOne = False
    For RevCnt = 1 To 15
      If UBTranRec(1).RevAmt(RevCnt) <> 0 Then
        PrintedOne = True
        'TabStop = (RevCnt * 10) + 1
        ToPrintR$ = ToPrintR$ + Using$("#######.##", UBTranRec(1).RevAmt(RevCnt)) + "~"
      Else
        ToPrintR$ = ToPrintR$ + "~ "
      End If
    Next
'    If PrintedOne Then
''      Print #UBRpt,
''      Linecnt = Linecnt + 1
'    End If
'    RevOffset = 7
'    PrintedOne = False
'    For RevCnt = 0 To 6
'      If UBTranRec(1).RevAmt(RevCnt + 1 + RevOffset) <> 0 Then
'        PrintedOne = True
'        'TabStop = (RevCnt * 10) + 1
'        ToPrintR$ = ToPrintR$ + "~" + Using$("#######.##", UBTranRec(1).RevAmt(RevCnt + 1 + RevOffset))
'      Else
'        ToPrintR$ = ToPrintR$ + "~ "
'      End If
'    Next
    If PrintedOne Then
'      Print #UBRpt,
'      Linecnt = Linecnt + 1
    End If

Return

DOTranHistHeader:
  'Linecnt = 7
  'Print #UBRpt, Tab(28); "Transaction History Report. "
  ToPrintH$ = QPTrim(UBCustRec(1).CustName)
  'If DetailFlag Then
    ToPrintH$ = ToPrintH$ + "~" + Str(RecNo&)
    PrnSvcAddr$ = QPTrim(UBCustRec(1).SERVADDR)
    PrnLoc$ = QPTrim$(UBCustRec(1).Book) + "-" + QPTrim$(UBCustRec(1).SEQNUMB)
    'Linecnt = Linecnt + 2
    For MtrCnt = 1 To 7
      If Len(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)) > 0 Then
        PrnMtr(MtrCnt) = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
        'Linecnt = Linecnt + 1
      End If
    Next
  'End If
'  Print #UBRpt,
'  If DetailFlag Then
'    Print #UBRpt, "Trans Date     Trans Type     Cur.Date     Pre.Date      TR Amount      Balance"
'  Else
'    Print #UBRpt, "Trans Date     Trans Type                                TR Amount      Balance"
'  End If
'  If DetailFlag Then
'    Print #UBRpt, "               Meter Type     Cur.Read     Pre.Read       Usage"
'  End If
'  If DetailFlag Then
'    Print #UBRpt, RevText1$
'    If Rev2Flag Then
'      Print #UBRpt, RevText2$
'      Linecnt = 8
'    End If
'  Else
'    Linecnt = 5
'  End If
  'Print #UBRpt, Dash80$
Return

DOTranHistFooter:
  If FirstTrans Then
    'Print #UBRpt, "NO TRANSACTIONS!!!"
    'Print #UBRpt, Dash80$
  End If
  For MtrCnt = 1 To 7
    If DidCnt(MtrCnt) > 0 Then
      PrnAvg$ = Using$("#########", TotalConsump(MtrCnt) / DidCnt(MtrCnt))
    End If
  Next
 'Print #UBRpt, FF$
Return

End Sub

'all this from original form save till make sure do not need
'Private Sub fpCmdExit_Click()
'  Load frmUBCustMenu
'  DoEvents
'  frmUBCustMenu.Show
'  Unload frmRptCustHistory
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'  Select Case KeyCode
'    Case vbKeyEscape:
'      KeyCode = 0
'      Call fpCmdExit_Click
'    Case vbKeyF10, vbKeyReturn
'      KeyCode = 0
'      Call fpCmdSearch_Click
'    Case vbKeyF7:
'      KeyCode = 0
'      Call fpCmdChoice_Click
'    Case Else:
'  End Select
'End Sub
'
'Private Sub Form_Load()
'  Set Temp_Class = New Resize_Class
'  Temp_Class.InitResizeClass Me
'  Set Over = New clsTextBoxOverRider
'  Over.OverRide Me
'  StatusBar1.Panels.Item(1).Text = TownName$
'  DefLookUp = GetDefaultLookUP    'get the user default lookup
'  Call SetPromptLabel             'set lookup prompt
'End Sub
'
'Private Sub Form_Resize()
'  If Me.WindowState <> vbMinimized Then
'    Me.Visible = False
'    Temp_Class.ResizeControls Me
'    Me.Visible = True
'    Me.SetFocus
'  End If
'
'End Sub
'
'Private Sub SetPromptLabel()
'
'  If DefLookUp > 6 Or DefLookUp < 1 Then
'    DefLookUp = 1
'  End If
'  Select Case DefLookUp
'  Case 1:
'    Me.PromptLabel = "Account Number:"
'  Case 2:
'    Me.PromptLabel = "Search Name:"
'  Case 3:
'    Me.PromptLabel = "Meter Number:"
'  Case 4:
'    Me.PromptLabel = "Service Address:"
'  Case 5:
'    Me.PromptLabel = "Location Number:"
'  Case 6:
'    Me.PromptLabel = "911/Other:"
'  End Select
'
'End Sub

'Private Sub fpCmdSearch_Click()
'  Dim LookFor As String
'  LookFor$ = QPTrim$(Me.fpSearchText)
'  'DeActivateControls Me
'  RecNo& = LookUp(LookFor$, DefLookUp, False, False, Me)
''  ActivateControls Me
'  If RecNo& > 0 Then
'    frmLoadingRpt.Show
'    DoEvents
'    If ConsumpFlag Then
'      Call CustConsumpHistRpt
'    Else
'      Call CustTransHistoryRpt
'    End If
'    Unload frmLoadingRpt
'  Else
'    Me.fpSearchText.SetFocus
'  End If
'
'End Sub

