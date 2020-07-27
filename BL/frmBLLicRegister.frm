VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLLicRegister 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Register"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLLicRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4836
      Left            =   1932
      TabIndex        =   4
      Top             =   2016
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   8530
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmBLLicRegister.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   405
         Left            =   2925
         TabIndex        =   0
         Tag             =   $"frmBLLicRegister.frx":08E6
         Top             =   1830
         Width           =   3570
         _Version        =   196608
         _ExtentX        =   6297
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
         AutoSearch      =   2
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
         MaxEditLen      =   5
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
         AutoSearchFillDelay=   200
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmBLLicRegister.frx":0992
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   405
         Left            =   2925
         TabIndex        =   1
         Tag             =   $"frmBLLicRegister.frx":0C89
         Top             =   2490
         Width           =   3570
         _Version        =   196608
         _ExtentX        =   6297
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
         Columns         =   0
         Sorted          =   0
         SelDrawFocusRect=   -1  'True
         ColumnSeparatorChar=   9
         ColumnSearch    =   -1
         ColumnWidthScale=   2
         RowHeight       =   -1
         WrapList        =   0   'False
         WrapWidth       =   0
         AutoSearch      =   2
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
         AutoSearchFill  =   -1  'True
         AutoSearchFillDelay=   200
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmBLLicRegister.frx":0D42
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   645
         Left            =   3120
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the 'License Processing' menu."
         Top             =   3645
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1138
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         ButtonDesigner  =   "frmBLLicRegister.frx":1039
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   645
         Left            =   5235
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   $"frmBLLicRegister.frx":1217
         Top             =   3645
         Width           =   1875
         _Version        =   131072
         _ExtentX        =   3307
         _ExtentY        =   1138
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         ButtonDesigner  =   "frmBLLicRegister.frx":14CC
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
         Height          =   645
         Left            =   720
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   $"frmBLLicRegister.frx":16AB
         Top             =   3645
         Width           =   2175
         _Version        =   131072
         _ExtentX        =   3836
         _ExtentY        =   1138
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         ButtonDesigner  =   "frmBLLicRegister.frx":177B
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   1788
         Left            =   912
         Top             =   1488
         Width           =   5964
      End
      Begin VB.Label lblBalloon 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "HELP BALLOONS ON"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   252
         Left            =   768
         TabIndex        =   9
         Top             =   4320
         Width           =   2100
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Report Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   1152
         TabIndex        =   7
         Top             =   2592
         Width           =   1500
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1440
         Top             =   432
         Width           =   4908
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "License Register"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   396
         Left            =   1728
         TabIndex        =   6
         Top             =   576
         Width           =   4332
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Print Order:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1392
         TabIndex        =   5
         Top             =   1920
         Width           =   1308
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   2016
      TabIndex        =   10
      Top             =   7104
      Width           =   780
      _Version        =   131072
      _ExtentX        =   1376
      _ExtentY        =   783
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   9405029
      ForeColor       =   8454143
      Text            =   ""
      Shape           =   0
      ShapeRoundWidth =   195
      ShapeRoundHeight=   195
      BorderWidth     =   -1
      BorderColor     =   -2147483630
      ThreeDWidth     =   -1
      ThreeDShadowColor=   -2147483632
      ThreeDHighlightColor=   16777215
      ThreeDText      =   0
      ThreeDTextHighlightColor=   16777215
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignPictureH   =   0
      AlignPictureV   =   0
      PictureStyle    =   0
      WordWrap        =   -1  'True
      ScaleMode       =   1
      ThreeDStyle     =   2
      Position        =   0
      PosBaseX        =   0
      PosBaseY        =   0
      PosOffsetX      =   -100
      PosOffsetY      =   300
      MaxWidth        =   3000
      CloudInset      =   100
      CloudMinWidth   =   600
      TailShape       =   2
      TailType        =   2
      TailBaseOffsetOutside=   300
      TailBaseOffsetInside=   100
      TailBaseAxisOutside=   0
      TailBaseAxisInside=   0
      TailBubbleCount =   3
      AlignTextH      =   1
      AlignTextV      =   1
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      AutoScan        =   0
      ScanAllDescendants=   -1  'True
      Interval        =   500
      IntervalNext    =   200
      AutoSize        =   -1  'True
      UseTagProp      =   -1  'True
      HideOnInactiveApp=   0   'False
      HideOnMouseDown =   2
      HideOnKeyDown   =   2
      HideOnFocus     =   0   'False
      ScanDisabledControls=   -1  'True
      ThreeDAppearance=   0
      FollowFocus     =   0   'False
      TemplateName    =   ""
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   5100
      Left            =   1800
      Top             =   1872
      Width           =   8052
   End
End
Attribute VB_Name = "frmBLLicRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim IdxRecs() As Double
  Dim NumOfCustRecs As Double
  Dim CreditFlag As Boolean
  
Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    cmdExit.ToolTipText = ""
    cmdProcess.ToolTipText = ""
    fpcmbPrintOrder.ToolTipText = ""
    fpcmbPrintOpt.ToolTipText = ""
    cmdHelp.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    cmdExit.ToolTipText = "Press to return to the 'License Processing' menu."
'    cmdProcess.ToolTipText = "Press the 'Process' button to calculate fees for all customers earmarked for renewal."
'    fpcmbPrintOrder.ToolTipText = "This report can be printed in alphabetical order or in numerical order."
'    fpcmbPrintOpt.ToolTipText = "Select graphical to print on a laser printer or choose text to print on a dot matrix printer."
'    cmdHelp.ToolTipText = "Press 'Turn Help On' to activate informational balloons for each field. Press 'Turn Help Off' to deactivate the informational balloons."
  End If
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  CreditFlag = False
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%T"
      Call cmdHelp_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLLicRegister.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  lblBalloon.Visible = False
'  cmdExit.ToolTipText = "Press to return to the 'License Processing' menu."
'  cmdProcess.ToolTipText = "Press the 'Process' button to calculate fees for all customers earmarked for renewal."
'  fpcmbPrintOrder.ToolTipText = "This report can be printed in alphabetical order or in numerical order."
'  fpcmbPrintOpt.ToolTipText = "Select graphical to print on a laser printer or choose text to print on a dot matrix printer."
'  cmdHelp.ToolTipText = "Press 'Turn Help On' to activate informational balloons for each field. Press 'Turn Help Off' to deactivate the informational balloons."
  fpcmbPrintOrder.Text = "Billing Name Order"
  fpcmbPrintOrder.AddItem "Billing Name Order"
  fpcmbPrintOrder.AddItem "Account Number Order"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  fpcmbPrintOpt.Text = "Graphical"
End Sub

Private Sub fpcmbPrintOpt_Change()
  If QPTrim$(fpcmbPrintOpt.Text) = "" Then
    fpcmbPrintOpt.Text = "Graphical"
  End If
End Sub

Private Sub fpcmbPrintOrder_Change()
  If QPTrim$(fpcmbPrintOrder.Text) = "" Then
    fpcmbPrintOrder.Text = "Billing Name Order"
  End If
End Sub

Private Sub fpcmbPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbPrintOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOrder.ListIndex = -1
  End If
  If fpcmbPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPrintOpt.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOpt.ListIndex = -1
  End If
  If fpcmbPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdExit.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub cmdExit_Click()
  frmBLPrintLicMenu.Show
  DoEvents
  Unload frmBLLicRegister
End Sub

Private Sub cmdProcess_Click()
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim cnt As Integer, x As Integer
  Dim OHandle As Integer
  Dim OperRec As CitiPassType 'CMOperRecType
  Dim NumOperRecs As Integer
  Dim Operator$
  Dim y As Integer
  Dim PayHandle As Integer
  Dim EditPayRec As AREditPaymentRecType
  Dim NumOfPayRecs As Integer
  Dim OCnt As Integer
  
  OpenCustFile CHandle
  NumOfCustRecs = LOF(CHandle) / Len(CustRec)
  
  OmitCnt = 0
  
  If Exist("artmppst.dat") Then 'registers have been run and not posted
    If Exist("artmplic.dat") Or Exist("licprnOK.dat") Then
      frmBLWarnRegReprint.Show vbModal
      If frmBLWarnRegReprint.fptxtChoice.Text = "continue" Then
        KillFile "artmppst.dat"
        KillFile "artmplic.dat" 'added 2/22/05
        KillFile "licprnOK.dat" 'added 2/22/05
        MainLog ("License registers processing has already occurred and user was warned that rerunning license registers will delete any existing temporary posting data saved. User opted to continue anyway.")
      Else
        Close
        Exit Sub
      End If
    End If
  End If
  
  If Exist("artmppen.dat") Then
    For cnt = 1 To NumOfCustRecs
      Get #CHandle, cnt, CustRec
      If QPTrim$(CustRec.IssueLicense) = "Y" Then
        If EmpInPenProcess(QPTrim$(CustRec.CustNumb)) = True Then
          OmitCnt = OmitCnt + 1
          ReDim Preserve OmitList(1 To OmitCnt) As Long
          OmitList(OmitCnt) = cnt
        End If
      End If
SkipThisOne:
    Next cnt
  End If
  
  If OmitCnt > 0 Then
    frmBLOmitList.Label1.Alignment = 0
    frmBLOmitList.Label1.Caption = "The following is a list of all customers who qualify for a license fee but are currently included in an unposted penalty fee calculation. Press F10: Exclude this list from license fees and reset their 'Print                  Next' flag to 'No'.                                                      Press F5: Include this list in license fees and delete penalty file.     Press ESC: Abort license fee processing.                                Press F3: Print list."
    frmBLOmitList.Show vbModal
    If frmBLOmitList.fptxtChoice.Text = "delete" Then
      Unload frmBLOmitList
      KillFile "artmppen.dat"
      MainLog ("User elected to delete 'artmppen.dat' file in order to allow all license fees to be calculated.")
    ElseIf frmBLOmitList.fptxtChoice.Text = "continue" Then
      'here the user has elected to continue with the license register
      'processing but without the customers involved in the penalty
      'process...these customers will get there Set Renewal Flag (Y/N)? flags
      'reset to 'N' (Set Renewal Flag (Y/N)? = .IssueLicense)
      For x = 1 To OmitCnt
        Get #CHandle, OmitList(x), CustRec
          CustRec.IssueLicense = "N" 'License processing will skip any customer
          'whose flag is set to 'N'
        Put #CHandle, OmitList(x), CustRec
      Next x
      Unload frmBLOmitList
      MainLog ("User presented with a list of all customers who would not be processed for license fees because they were involved in an unposted penalty fee process. User elected to continue license fee process excluding customers on list.")
    ElseIf frmBLOmitList.fptxtChoice.Text = "abort" Then
      Unload frmBLOmitList
      MainLog ("User elected to abort license fee calculations after being shown a list of all customers ineligible because they were already involved in a penalty fee file.")
      fpcmbPrintOrder.SetFocus
      Close
      Exit Sub
    End If
  End If
'  Close
  
'_______________________________________________________________
  'Checking for customers who are in an unposted pay file
  'go to the password file and get the
  'operator numbers...
  'Scenario: Customer pays his business
  'license fee in advance-> that payment is included in an
  'unposted payment file-> user runs license register including
  'customer that just made payment-> user does not post license
  'fees-> unposted payment file gets posted...customer's balance
  'reduced by amount of his payment->license fees now get posted
  'using the total balance saved in the temporary file that was
  'amount BEFORE the payment was made-> Total balance now reverts
  'back to the amount before the payment was made->customer's
  'payment gets erased
  'THIS IS WHY WE CANNOT ALLOW CUSTOMERS INVOLVED IN AN UNPOSTED
  'PAYMENT FILE TO ALSO BE ADDED TO AN UNPOSTED LICENSE FILE
  OpenCitiPassFile OHandle, NumOperRecs
  If NumOperRecs = 0 Then
    Close OHandle
    Return
  End If

  ReDim OpIdx(1 To NumOperRecs) As Integer
  For x = 1 To NumOperRecs
    Get OHandle, x, OperRec
      'load an array with the operator numbers
      OpIdx(x) = OperRec.PassNum
  Next x
  Close OHandle
  
  OCnt = 0
  ReDim InPayCnt(1 To 1) As String
  For x = 1 To NumOperRecs
    Operator = Str(OpIdx(x))
    If Exist(BLPayFileName + Operator$ + ".DAT") Then
      'if the file above exists then this operator has
      'saved at least one transaction
      OpenPayFile PayHandle, OpIdx(x) 'look thru all operator files
      NumOfPayRecs = LOF(PayHandle) / Len(EditPayRec)
      For y = 1 To NumOfPayRecs
        Get PayHandle, y, EditPayRec
        If QPTrim$(EditPayRec.CustNumber) = "" Then GoTo Deleted
        OCnt = OCnt + 1
        ReDim Preserve InPayCnt(1 To OCnt) As String
        InPayCnt(OCnt) = QPTrim$(EditPayRec.CustNumber)
Deleted:
      Next y
    End If
SkipIt:
  Next x
  Close PayHandle
  
  If OCnt = 0 Then GoTo EmptyPayQueue
  
  ReDim InPayOmit(1 To 1) As Long
  PayOmitCnt = 0
  For x = 1 To NumOfCustRecs
    Get #CHandle, x, CustRec
      For y = 1 To OCnt
        If QPTrim$(CustRec.CustNumb) = InPayCnt(y) Then
          If QPTrim$(CustRec.IssueLicense) = "Y" Then
            PayOmitCnt = PayOmitCnt + 1
            ReDim Preserve InPayOmit(1 To PayOmitCnt) As Long
            InPayOmit(PayOmitCnt) = x
            Exit For
          End If
        End If
      Next y
  Next x
  
  If PayOmitCnt = 0 Then GoTo EmptyPayQueue
  frmBLPayOmitList.Label1.Alignment = 0
  frmBLPayOmitList.Label1.Caption = "The following is a list of all customers who qualify for a license fee but are currently included in an unposted payment file.            Press F10: Exclude this list from license fees and reset their 'Print                  Next' flag to 'No'.                                                      Press ESC: Abort license fee processing.                                Press F3: Print list."
  frmBLPayOmitList.Show vbModal
  If frmBLPayOmitList.fptxtChoice.Text = "continue" Then
    'here the user has elected to continue with the license register
    'processing but without the customers involved in the penalty
    'process...these customers will get there Set Renewal Flag (Y/N)? flags
    'reset to 'N' (Set Renewal Flag (Y/N)? = .IssueLicense)
    For x = 1 To PayOmitCnt
      Get #CHandle, InPayOmit(x), CustRec
        CustRec.IssueLicense = "N" 'License processing will skip any customer
        'whose flag is set to 'N'
      Put #CHandle, InPayOmit(x), CustRec
    Next x
    Unload frmBLPayOmitList
    MainLog ("User presented with a list of all customers who would not be processed for license fees because they were involved in an unposted payment process. User elected to continue license fee process excluding customers on list.")
  ElseIf frmBLPayOmitList.fptxtChoice.Text = "abort" Then
    Unload frmBLPayOmitList
    MainLog ("User elected to abort license fee calculations after being shown a list of all customers ineligible because they were already involved in an unposted payment file.")
    fpcmbPrintOrder.SetFocus
    Close
    Exit Sub
  End If
  
EmptyPayQueue:
  Close
'_______________________________________________________________
  
  If fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  ElseIf fpcmbPrintOpt.Text = "Text" Then
    frmBLMessageBoxJr.Label1.Caption = "Pitch 12 is recommended for this report."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Call PrintText
  Else
    Exit Sub
  End If
End Sub

Private Sub PrintText()
  Dim ReportFile$
  Dim FF$, x As Double
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim CustCnt As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim NumOfCodeRecs As Integer
  Dim CHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim CustNameIdxRec As CustSearchNameIdxType
  Dim CustNumIdxRec As CustNumIdxType
  Dim IdxHandle As Integer
  Dim NumOfCustIdxRecs As Double
  Dim NameFlag As Boolean
  Dim NumFlag As Boolean
  Dim RptHandle As Integer
  Dim Page As Integer
  Dim cnt&, DelFlag$, CustFee#
  Dim CCode$
  Dim TC$, TotalFee#
  Dim CatIdxRecs As CatCodeIdxType
  Dim CatIdxHandle As Integer
  Dim NumOfCatIdxRecs As Integer
  Dim ThisCode$
  Dim ThisFee As Double
  Dim CustNum As Integer
  Dim Code As String * 5
  Dim TempRec As TempTransPostType
  Dim TempHandle As Integer
  Dim NumOfTempRecs As Integer
  Dim NumOfCustRecs As Integer
  Dim IssFeeTot As Double
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim y As Integer
  
  On Error GoTo ERRORSTUFF
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  Call SetFee
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  OpenCustFile CustHandle
  
  OpenCatCodeIdxFile CatIdxHandle
  NumOfCatIdxRecs = LOF(CatIdxHandle) / Len(CatIdxRecs)
  
  OpenTempPostFile TempHandle
  NumOfTempRecs = LOF(TempHandle) / Len(TempRec)
  If NumOfTempRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for a new business license. If a customer is involved in a penalty file they won't qualify. If you wish to set customer flags go to the 'Set License To Print' screen or go to the Customer Edit screen to reset a specific customer's 'Set Renewal Flag (Y/N)?' field to Yes."
    frmBLMessageBoxJr.Label1.Top = 400
    frmBLMessageBoxJr.Label1.Height = 1500
    frmBLMessageBoxJr.Show vbModal
    Close
    cmdExit.Enabled = True
    cmdProcess.Enabled = True
    cmdHelp.Enabled = True
    Exit Sub
  End If
  
  ReDim CatIdx(1 To NumOfCatIdxRecs) As String
  ReDim CatDesc(1 To NumOfCatIdxRecs) As String
  ReDim CatFeeAmt(1 To NumOfCatIdxRecs) As Double
  ReDim CatCnt(1 To NumOfCatIdxRecs) As Integer
  
  OpenCatCodeFile CHandle
  NumOfCodeRecs = LOF(CHandle) / Len(CodeRec)
  frmBLLoadReport.Label1.Caption = "Linking To Report"
  frmBLLoadReport.Label2.Visible = False
  frmBLLoadReport.Show
  DoEvents
  For x = 1 To NumOfCatIdxRecs
    Get CatIdxHandle, x, CatIdxRecs
    CatIdx(x) = QPTrim$(CatIdxRecs.CatCodeNum)
    For y = 1 To NumOfCodeRecs
      Get CHandle, y, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatIdx(x) Then
        If NumOfCodeRecs = NumOfCatIdxRecs Then
          CatDesc(x) = QPTrim$(CodeRec.CODEDESC)
        Else
          CatDesc(x) = ""
        End If
        Exit For
      End If
    Next y
  Next x
  Close CatIdxHandle
  Close CHandle
  
  ReportFile$ = "ARLICREG.PRN"  'Report File Name
  FF$ = Chr$(12)
  MaxLines = 53
  LineCnt = 0

  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  GoSub PrintLicRegRptHeader
  
  For cnt& = 1 To NumOfTempRecs 'NumOfCustRecs 'NumOfARRecs
    Get TempHandle, cnt, TempRec
    Get CustHandle, Val(TempRec.CustomerNumber), CustRec
    
    If QPTrim$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SortName) = "DELETED" Then
      GoTo DelSkip
    End If
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo DelSkip
    
    If CustRec.IssueLicense = "Y" Then
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintLicRegRptHeader
      End If
      CustNum = CustNum + 1
'      IssFeeTot = IssFeeTot + TempRec.IssFee
      IssFeeTot = IssFeeTot + TownRec.IssFee
'      CustFee# = OldRound(TempRec.CatFee1 + TempRec.CatFee2 + TempRec.CatFee3 + TempRec.CatFee4 + TempRec.CatFee5 + TempRec.IssFee)
      CustFee# = OldRound(TempRec.CatFee1 + TempRec.CatFee2 + TempRec.CatFee3 + TempRec.CatFee4 + TempRec.CatFee5 + TownRec.IssFee)
      CCode$ = QPTrim$(CustRec.BILLCAT1)
      ThisCode = CCode
      ThisFee = TempRec.CatFee1
      GoSub CollectTotals
      TC$ = QPTrim$(CustRec.BILLCAT2)
      If Len(TC$) > 0 Then
        ThisCode = TC$
        ThisFee = TempRec.CatFee2
        GoSub CollectTotals
        CCode$ = CCode$ + "/" + TC$
      End If
      TC$ = QPTrim$(CustRec.BILLCAT3)
      If Len(TC$) > 0 Then
        ThisCode = TC$
        ThisFee = TempRec.CatFee3
        GoSub CollectTotals
        CCode$ = CCode$ + "/" + TC$
      End If
      TC$ = QPTrim$(CustRec.BILLCAT4)
      If Len(TC$) > 0 Then
        ThisCode = TC$
        ThisFee = TempRec.CatFee4
        GoSub CollectTotals
        CCode$ = CCode$ + "/" + TC$
      End If
      TC$ = QPTrim$(CustRec.BILLCAT5)
      If Len(TC$) > 0 Then
        ThisCode = TC$
        ThisFee = TempRec.CatFee5
        GoSub CollectTotals
        CCode$ = CCode$ + "/" + TC$
      End If
      If CustRec.Prorate < 100 Then
        If TempRec.CreditUsed = False Then
          Print #RptHandle, Using("#####0", TempRec.CustomerNumber); Tab(10); Left$(CustRec.BillName, 30); Tab(42); CCode$; Tab(70); Using("$##,##0.00", CustFee#); Tab(83); Using("#0%", CustRec.Prorate / 100)
        Else
          Print #RptHandle, Using("#####0", TempRec.CustomerNumber); Tab(10); Left$(CustRec.BillName, 30); Tab(42); CCode$ + " *"; Tab(70); Using("$##,##0.00", CustFee#); Tab(83); Using("#0%", CustRec.Prorate / 100)
        End If
      Else
        If TempRec.CreditUsed = False Then
          Print #RptHandle, Using("#####0", TempRec.CustomerNumber); Tab(10); Left$(CustRec.BillName, 30); Tab(42); CCode$; Tab(70); Using("$##,##0.00", CustFee#)
        Else
          Print #RptHandle, Using("#####0", TempRec.CustomerNumber); Tab(10); Left$(CustRec.BillName, 30); Tab(42); CCode$ + " *"; Tab(70); Using("$##,##0.00", CustFee#)
        End If
      End If
      TotalFee# = OldRound(TotalFee# + CustFee#)
      
      LineCnt = LineCnt + 1
    End If

DelSkip:
  Next

  GoSub PrintLicRegRptEnding
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now
  
  Unload frmBLLoadReport
  
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True

  ViewPrint ReportFile$, "Licenses to Print Register", True
  
  KillFile ReportFile$
  MainLog ("Business license registers processed in text format.")

  Exit Sub
  
CollectTotals:
  For x = 1 To NumOfCatIdxRecs
    If QPTrim$(ThisCode) = QPTrim$(CatIdx(x)) Then
      CatFeeAmt(x) = CatFeeAmt(x) + ThisFee
      CatCnt(x) = CatCnt(x) + 1
      Exit For
    End If
  Next x
  
  Return

PrintLicRegRptHeader:
  Page = Page + 1
  Print #RptHandle, Tab(18); "Business License System : Licenses to Print Register"
  Print #RptHandle, QPTrim$(TownRec.TownName)
  Print #RptHandle, "Report Date: "; Date$; Tab(72); "Page #"; Page
  Print #RptHandle, "XX% = Prorated Amount If Less Than 100%"
  Print #RptHandle, "* = Credit applied to fee"
  If TownRec.IssFee > 0 Then
    Print #RptHandle, QPTrim$(Using("$#,##0.00", TownRec.IssFee)) + " issuance fee charged to each customer."
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, ""
  Print #RptHandle, "Cust #"; Tab(10); "Billing Name"; Tab(40); "Category Codes"; Tab(70); "Fee Amount"
  Print #RptHandle, String$(80, "=")
  LineCnt = 8
  Return

PrintLicRegRptEnding:
  Print #RptHandle, FF$
  Print #RptHandle, Tab(18); "Business License System : Category Fee Totals"
  Print #RptHandle,
  Print #RptHandle, String$(80, "=")
  Print #RptHandle, Tab(3); "Category"; Tab(21); "Count"; Tab(40); "Description"; Tab(71); "Total Fees"
  Print #RptHandle, String$(80, "-")
  For x = 1 To NumOfCatIdxRecs
    If CatFeeAmt(x) > 0 Then
      RSet Code = CatIdx(x)
      Print #RptHandle, Tab(5); Code; Tab(23); Using("####", (CatCnt(x))); Tab(30); CatDesc(x); Tab(71); Using("$##,##0.00", CatFeeAmt(x))
    End If
  Next x
  If IssFeeTot > 0 Then
      Print #RptHandle, Tab(3); "Issuance Fee"; Tab(71); Using("$##,##0.00", IssFeeTot)
  End If
  Print #RptHandle, String$(80, "-")
  Print #RptHandle, Tab(3); "License Count:"; Tab(30); Str(CustNum); Tab(68); Using("$#,###,##0.00", TotalFee#)
  Print #RptHandle, FF$
  Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLLicRegister", "PrintText", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
  
  
End Sub

Private Sub SetFee()

  Dim CodeHandle As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim NumOfARCatRecs As Integer
  Dim NameFlag As Boolean
  Dim NumFlag As Boolean
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim CustNameIdxRec As CustNameIdxType 'CustSearchNameIdxType
  Dim CustNumIdxRec As CustNumIdxType
  Dim IdxHandle As Integer
  Dim NumOfCustIdxRecs As Double
  Dim RptHandle As Integer
  Dim Page As Integer, cnt As Long
  Dim ProrateFlag As Boolean
  Dim ProAt#, CatCode$, Snt&
  Dim FeeAmt#, Mult#, Revenue#
  Dim x As Double, Prorate#, Nextx As Integer
  Dim TempRec As TempTransPostType
  Dim TempRec2 As TempTransPostType
  Dim TempHandle As Integer
  Dim OverLic As Double
  Dim OverPen As Double
  Dim OverTot As Double
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim IssFee As Double
  Dim TempIssFee As Double
  Dim Towncnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenTownFile TownHandle
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  OpenCatCodeFile CodeHandle
  NumOfARCatRecs = LOF(CodeHandle) \ Len(CodeRec)
  If NumOfARCatRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Error: The list of category codes cannot be indexed. Please make sure there are category codes saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Billing Name Order" Then
    NameFlag = True
    NumFlag = False
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Account Number Order" Then
    NumFlag = True
    NameFlag = False
  Else
    fpcmbPrintOrder.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "Please make a selection for Print Order."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbPrintOrder.BackColor = &HFFFFFF
    fpcmbPrintOrder.SetFocus
    Close
    Exit Sub
  End If

  If NameFlag = True Then
'    OpenSrchNameIdxFile IdxHandle
    OpenCustNameIdxFile IdxHandle
    NumOfCustIdxRecs = LOF(IdxHandle) / Len(CustNameIdxRec)
  Else
    OpenCustNumIdxFile IdxHandle
    NumOfCustIdxRecs = LOF(IdxHandle) / Len(CustNumIdxRec)
  End If

  If NumOfCustIdxRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no business customers indexed."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  OpenCustFile CustHandle

  ReDim IdxRecs(1 To NumOfCustIdxRecs) As Double
  If NameFlag = True Then
    For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNameIdxRec
      IdxRecs(x) = CustNameIdxRec.CustRec
    Next x
  Else
      For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNumIdxRec
      IdxRecs(x) = CustNumIdxRec.CustRec
    Next x
  End If
  Close IdxHandle

  frmBLShowPctComp.cmdCancel.Visible = False
  frmBLShowPctComp.Label1 = "Loading Licenses to Print Register"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdHelp.Enabled = False
  
  OpenTempPostFile TempHandle
  Nextx = 1
  For cnt = 1 To NumOfCustIdxRecs
    Get CustHandle, IdxRecs(cnt), CustRec
    IssFee = TownRec.IssFee
    ProrateFlag = False
    If CustRec.IssueLicense <> "Y" Then GoTo SkipEm
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo SkipEm
    If CustRec.Deleted = "Y" Or QPTrim$(CustRec.SortName) = "DELETED" Or CustRec.Inactive = "Y" Then
      GoTo SkipEm
    End If
    
    If Exist("artmppen.dat") And EmpInPenProcess(CStr(IdxRecs(cnt))) = True Then
      GoTo SkipEm
    End If
    
    'Clear OverPen & OverLic
    OverPen = 0
    OverLic = 0
    'Clear TempRec from previous calculations
    TempRec = TempRec2
    TempRec.CreditUsed = False
    'assign the temp fields currrent amounts...temp will
    'become permanent during posting
    TempRec.PenBal = CustRec.PenBal
    TempRec.LicBal = CustRec.LicBal
    
    TempRec.CatFeeBal1 = CustRec.FeeLicBal1
    TempRec.CatFeeBal2 = CustRec.FeeLicBal2
    TempRec.CatFeeBal3 = CustRec.FeeLicBal3
    TempRec.CatFeeBal4 = CustRec.FeeLicBal4
    TempRec.CatFeeBal5 = CustRec.FeeLicBal5
    
    'if a negative (credit) amount exists figure assign credit amount to
    'OverLic as a positive amount
    If TempRec.LicBal < 0 Then OverLic = Abs(TempRec.LicBal)
    
    If IssFee > 0 And OverLic > 0 Then 'we need to reduce the
    'the credit license balance to reflect the cost of the Issue Fee first
    'so we pay for the issuance fee out of whatever credit amount exists
      GoSub CalcLicBal 'if we are going to reduce the issuance fee
      'by the amount of the credit then we must also go into the
      'license balances and bring them closer to 0 because at
      'this point we know we have at least one license balance credit
    End If
    
    TempRec.LicBal = TempRec.CatFeeBal1 + TempRec.CatFeeBal2 + TempRec.CatFeeBal3 + TempRec.CatFeeBal4 + TempRec.CatFeeBal5
    If CustRec.PenBal < 0 Then OverPen = Abs(CustRec.PenBal)
    
    If IssFee > 0 And OverPen > 0 Then
      TempRec.PenBal = TempRec.PenBal + OverPen
      If OverPen >= IssFee Then
        IssFee = 0
        OverPen = OverPen - IssFee
      ElseIf OverPen < IssFee Then
        IssFee = IssFee - OverPen
        OverPen = 0
      End If
    End If
    
    TempRec.IssFeeBal = CustRec.IssuanceBal + IssFee
    TempRec.IssFee = IssFee 'TownRec.IssFee
    
    Prorate# = CustRec.Prorate
    
    If Prorate# >= 100 Or Prorate# <= 0 Then
      Prorate# = 100
    Else
      ProrateFlag = True
      Prorate# = OldRound(Prorate# * 0.01)
    End If

    CatCode$ = QPTrim$(CustRec.BILLCAT1)
    If Len(CatCode$) > 0 Then
      For Snt& = 1 To NumOfARCatRecs
        Get CodeHandle, Snt&, CodeRec
        'find the code that matches this category
        If QPTrim$(CodeRec.CatCode) = CatCode$ Then
          TempRec.CatCodeRec1 = Snt&
          'if it's a flat fate then do this
          If CodeRec.CodeType = "F" Then
            'get fee and assign accordingly
            If ProrateFlag = True Then
              TempRec.CatFee1 = OldRound(Prorate# * CodeRec.Fee)
            Else
              TempRec.CatFee1 = CodeRec.Fee
            End If
            If TempRec.CatFee1 < 0 Then TempRec.CatFee1 = 0
            'if there remains a negative penalty balance or a negative
            'category balance then these negative amounts are applied
            'to the new fee calculations
            If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
              Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
              TempRec.CreditUsed = True
              TempRec.PenBal = OverPen
            Else
              TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
            End If
            GoTo C2
          End If
          'if it's a multiplier then do this
          If CodeRec.CodeType = "M" Then
            'get number of multipliers for this customer
            Mult = CustRec.REV1
            If ProrateFlag = True Then
              TempRec.CatFee1 = OldRound(Mult * CodeRec.Fee)
              TempRec.CatFee1 = OldRound(TempRec.CatFee1 * Prorate#)
            Else
              TempRec.CatFee1 = OldRound(Mult * CodeRec.Fee)
            End If
            If TempRec.CatFee1 < 0 Then TempRec.CatFee1 = 0
            'now apply any credits
            If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
              Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
              TempRec.CreditUsed = True
              TempRec.PenBal = OverPen
            Else
              TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
            End If
            GoTo C2
          End If
          If CodeRec.CodeType = "S" Then
            'if it's a step rate then find the level
            'that applies to this customer's revenue
            Revenue# = CustRec.REV1
            If ProrateFlag = True Then
              If Revenue# <= CodeRec.Recpt1 Then
                TempRec.CatFee1 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
                If TempRec.CatFee1 < CodeRec.BaseAmt1 Then TempRec.CatFee1 = CodeRec.BaseAmt1
                TempRec.CatFee1 = OldRound(Prorate# * TempRec.CatFee1)
                If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
                End If
                GoTo C2
              End If
              If Revenue# <= CodeRec.Recpt2 Then
                TempRec.CatFee1 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
                If TempRec.CatFee1 < CodeRec.BaseAmt2 Then TempRec.CatFee1 = CodeRec.BaseAmt2
                TempRec.CatFee1 = OldRound(Prorate# * TempRec.CatFee1)
                If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
                End If
                GoTo C2
              End If
              If Revenue# <= CodeRec.Recpt3 Then
                TempRec.CatFee1 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
                If TempRec.CatFee1 < CodeRec.BaseAmt3 Then TempRec.CatFee1 = CodeRec.BaseAmt3
                TempRec.CatFee1 = OldRound(Prorate# * TempRec.CatFee1)
                If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
                End If
                GoTo C2
              End If
              If Revenue# <= CodeRec.Recpt4 Then
                TempRec.CatFee1 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
                If TempRec.CatFee1 < CodeRec.BaseAmt4 Then TempRec.CatFee1 = CodeRec.BaseAmt4
                TempRec.CatFee1 = OldRound(Prorate# * TempRec.CatFee1)
                If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
                End If
                GoTo C2
              End If
              If Revenue# <= CodeRec.Recpt5 Then
                TempRec.CatFee1 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
                If TempRec.CatFee1 < CodeRec.BaseAmt5 Then TempRec.CatFee1 = CodeRec.BaseAmt5
                TempRec.CatFee1 = OldRound(Prorate# * TempRec.CatFee1)
                If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
                End If
                GoTo C2
              End If
              If Revenue# <= CodeRec.Recpt6 Then
                TempRec.CatFee1 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
                If TempRec.CatFee1 < CodeRec.BaseAmt6 Then TempRec.CatFee1 = CodeRec.BaseAmt6
                TempRec.CatFee1 = OldRound(Prorate# * TempRec.CatFee1)
                If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
                End If
                GoTo C2
              End If
            Else 'ProrateFlag = False
              If Revenue# <= CodeRec.Recpt1 Then
                TempRec.CatFee1 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
                If TempRec.CatFee1 < CodeRec.BaseAmt1 Then TempRec.CatFee1 = CodeRec.BaseAmt1
                If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
                End If
                GoTo C2
              End If
              If Revenue# <= CodeRec.Recpt2 Then
                TempRec.CatFee1 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
                If TempRec.CatFee1 < CodeRec.BaseAmt2 Then TempRec.CatFee1 = CodeRec.BaseAmt2
                If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
                End If
                GoTo C2
              End If
              If Revenue# <= CodeRec.Recpt3 Then
                TempRec.CatFee1 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
                If TempRec.CatFee1 < CodeRec.BaseAmt3 Then TempRec.CatFee1 = CodeRec.BaseAmt3
                If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
                End If
                GoTo C2
              End If
              If Revenue# <= CodeRec.Recpt4 Then
                TempRec.CatFee1 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
                If TempRec.CatFee1 < CodeRec.BaseAmt4 Then TempRec.CatFee1 = CodeRec.BaseAmt4
                If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
                End If
                GoTo C2
              End If
              If Revenue# <= CodeRec.Recpt5 Then
                TempRec.CatFee1 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
                If TempRec.CatFee1 < CodeRec.BaseAmt5 Then TempRec.CatFee1 = CodeRec.BaseAmt5
                If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
                End If
                GoTo C2
              End If
              If Revenue# <= CodeRec.Recpt6 Then
                TempRec.CatFee1 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
                If TempRec.CatFee1 < CodeRec.BaseAmt6 Then TempRec.CatFee1 = CodeRec.BaseAmt6
                If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
                End If
                GoTo C2
              End If
            End If
          End If
        End If  'End Test for Code
      Next Snt&
    Else
      TempRec.CatFee1 = 0
    End If      'End Test for Cat 1


C2:             'Category #2
    CatCode$ = QPTrim$(CustRec.BILLCAT2)
    If Len(CatCode$) > 0 Then
      For Snt& = 1 To NumOfARCatRecs
        Get CodeHandle, Snt&, CodeRec
        If QPTrim$(CodeRec.CatCode) = CatCode$ Then
          TempRec.CatCodeRec2 = Snt&
          If CodeRec.CodeType = "F" Then
            If ProrateFlag = True Then
              TempRec.CatFee2 = OldRound(Prorate# * CodeRec.Fee)
            Else
              TempRec.CatFee2 = CodeRec.Fee
            End If
            If TempRec.CatFee2 < 0 Then TempRec.CatFee2 = 0
            If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
              Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
              TempRec.CreditUsed = True
              TempRec.PenBal = OverPen
            Else
              TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
            End If
            GoTo C3
          End If
          If CodeRec.CodeType = "M" Then
            Mult = CustRec.REV2
            If ProrateFlag = True Then
              TempRec.CatFee2 = OldRound(Mult * CodeRec.Fee)
              TempRec.CatFee2 = OldRound(TempRec.CatFee2 * Prorate#)
            Else
              TempRec.CatFee2 = OldRound(Mult * CodeRec.Fee)
            End If
            If TempRec.CatFee2 < 0 Then TempRec.CatFee2 = 0
            If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
              Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
              TempRec.CreditUsed = True
              TempRec.PenBal = OverPen
            Else
              TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
            End If
            GoTo C3
          End If
          If CodeRec.CodeType = "S" Then
            Revenue# = CustRec.REV2
            If ProrateFlag = True Then
              If Revenue# <= CodeRec.Recpt1 Then
                TempRec.CatFee2 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
                If TempRec.CatFee2 < CodeRec.BaseAmt1 Then TempRec.CatFee2 = CodeRec.BaseAmt1
                TempRec.CatFee2 = OldRound(Prorate# * TempRec.CatFee2)
                If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
                End If
                GoTo C3
              End If
              If Revenue# <= CodeRec.Recpt2 Then
                TempRec.CatFee2 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
                If TempRec.CatFee2 < CodeRec.BaseAmt2 Then TempRec.CatFee2 = CodeRec.BaseAmt2
                TempRec.CatFee2 = OldRound(Prorate# * TempRec.CatFee2)
                If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
                End If
                GoTo C3
              End If
              If Revenue# <= CodeRec.Recpt3 Then
                TempRec.CatFee2 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
                If TempRec.CatFee2 < CodeRec.BaseAmt3 Then TempRec.CatFee2 = CodeRec.BaseAmt3
                TempRec.CatFee2 = OldRound(Prorate# * TempRec.CatFee2)
                If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
                End If
                GoTo C3
              End If
              If Revenue# <= CodeRec.Recpt4 Then
                TempRec.CatFee2 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
                If TempRec.CatFee2 < CodeRec.BaseAmt4 Then TempRec.CatFee2 = CodeRec.BaseAmt4
                TempRec.CatFee2 = OldRound(Prorate# * TempRec.CatFee2)
                If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
                End If
                GoTo C3
              End If
              If Revenue# <= CodeRec.Recpt5 Then
                TempRec.CatFee2 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
                If TempRec.CatFee2 < CodeRec.BaseAmt5 Then TempRec.CatFee2 = CodeRec.BaseAmt5
                TempRec.CatFee2 = OldRound(Prorate# * TempRec.CatFee2)
                If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
                End If
                GoTo C3
              End If
              If Revenue# <= CodeRec.Recpt6 Then
                TempRec.CatFee2 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
                If TempRec.CatFee2 < CodeRec.BaseAmt6 Then TempRec.CatFee2 = CodeRec.BaseAmt6
                TempRec.CatFee2 = OldRound(Prorate# * TempRec.CatFee2)
                If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
                End If
                GoTo C3
              End If
            Else
              If Revenue# <= CodeRec.Recpt1 Then
                TempRec.CatFee2 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
                If TempRec.CatFee2 < CodeRec.BaseAmt1 Then TempRec.CatFee2 = CodeRec.BaseAmt1
                If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
                End If
                GoTo C3
              End If
              If Revenue# <= CodeRec.Recpt2 Then
                TempRec.CatFee2 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
                If TempRec.CatFee2 < CodeRec.BaseAmt2 Then TempRec.CatFee2 = CodeRec.BaseAmt2
                If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
                End If
                GoTo C3
              End If
              If Revenue# <= CodeRec.Recpt3 Then
                TempRec.CatFee2 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
                If TempRec.CatFee2 < CodeRec.BaseAmt3 Then TempRec.CatFee2 = CodeRec.BaseAmt3
                If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
                End If
                GoTo C3
              End If
              If Revenue# <= CodeRec.Recpt4 Then
                TempRec.CatFee2 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
                If TempRec.CatFee2 < CodeRec.BaseAmt4 Then TempRec.CatFee2 = CodeRec.BaseAmt4
                If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
                End If
                GoTo C3
              End If
              If Revenue# <= CodeRec.Recpt5 Then
                TempRec.CatFee2 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
                If TempRec.CatFee2 < CodeRec.BaseAmt5 Then TempRec.CatFee2 = CodeRec.BaseAmt5
                If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
                End If
                GoTo C3
              End If
              If Revenue# <= CodeRec.Recpt6 Then
                TempRec.CatFee2 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
                If TempRec.CatFee2 < CodeRec.BaseAmt6 Then TempRec.CatFee2 = CodeRec.BaseAmt6
                If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
                End If
                GoTo C3
              End If
            End If
          End If
        End If  'End Test for Code
      Next Snt&
    Else
      TempRec.CatFee2 = 0
    End If      'End Test for Cat 1


C3:
    CatCode$ = QPTrim$(CustRec.BILLCAT3)
    If Len(CatCode$) > 0 Then
      For Snt& = 1 To NumOfARCatRecs
        Get CodeHandle, Snt&, CodeRec
        If QPTrim$(CodeRec.CatCode) = CatCode$ Then
          TempRec.CatCodeRec3 = Snt&
          If CodeRec.CodeType = "F" Then
            If ProrateFlag = True Then
              TempRec.CatFee3 = OldRound(Prorate# * CodeRec.Fee)
            Else
              TempRec.CatFee3 = CodeRec.Fee
            End If
            If TempRec.CatFee3 < 0 Then TempRec.CatFee3 = 0
            If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
              Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
              TempRec.CreditUsed = True
              TempRec.PenBal = OverPen
            Else
              TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
            End If
            GoTo c4
          End If
          If CodeRec.CodeType = "M" Then
            Mult = CustRec.REV3
            If ProrateFlag = True Then
              TempRec.CatFee3 = OldRound(Mult * CodeRec.Fee)
              TempRec.CatFee3 = OldRound(TempRec.CatFee3 * Prorate#)
            Else
              TempRec.CatFee3 = OldRound(Mult * CodeRec.Fee)
            End If
            If TempRec.CatFee3 < 0 Then TempRec.CatFee3 = 0
            If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
              Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
              TempRec.CreditUsed = True
              TempRec.PenBal = OverPen
            Else
              TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
            End If
            GoTo c4
          End If
          If CodeRec.CodeType = "S" Then
            Revenue# = CustRec.REV3
            If ProrateFlag = True Then
              If Revenue# <= CodeRec.Recpt1 Then
                TempRec.CatFee3 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
                If TempRec.CatFee3 < CodeRec.BaseAmt1 Then TempRec.CatFee3 = CodeRec.BaseAmt1
                TempRec.CatFee3 = OldRound(Prorate# * TempRec.CatFee3)
                If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
                End If
                GoTo c4
              End If
              If Revenue# <= CodeRec.Recpt2 Then
                TempRec.CatFee3 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
                If TempRec.CatFee3 < CodeRec.BaseAmt2 Then TempRec.CatFee3 = CodeRec.BaseAmt2
                TempRec.CatFee3 = OldRound(Prorate# * TempRec.CatFee3)
                If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
                End If
                GoTo c4
              End If
              If Revenue# <= CodeRec.Recpt3 Then
                TempRec.CatFee3 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
                If TempRec.CatFee3 < CodeRec.BaseAmt3 Then TempRec.CatFee3 = CodeRec.BaseAmt3
                TempRec.CatFee3 = OldRound(Prorate# * TempRec.CatFee3)
                If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
                End If
                GoTo c4
              End If
              If Revenue# <= CodeRec.Recpt4 Then
                TempRec.CatFee3 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
                If TempRec.CatFee3 < CodeRec.BaseAmt4 Then TempRec.CatFee3 = CodeRec.BaseAmt4
                TempRec.CatFee3 = OldRound(Prorate# * TempRec.CatFee3)
                If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
                End If
                GoTo c4
              End If
              If Revenue# <= CodeRec.Recpt5 Then
                TempRec.CatFee3 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
                If TempRec.CatFee3 < CodeRec.BaseAmt5 Then TempRec.CatFee3 = CodeRec.BaseAmt5
                TempRec.CatFee3 = OldRound(Prorate# * TempRec.CatFee3)
                If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
                End If
                GoTo c4
              End If
              If Revenue# <= CodeRec.Recpt6 Then
                TempRec.CatFee3 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
                If TempRec.CatFee3 < CodeRec.BaseAmt6 Then TempRec.CatFee3 = CodeRec.BaseAmt6
                TempRec.CatFee3 = OldRound(Prorate# * TempRec.CatFee3)
                If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
                End If
                GoTo c4
              End If
            Else 'prorateflag = false
              If Revenue# <= CodeRec.Recpt1 Then
                TempRec.CatFee3 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
                If TempRec.CatFee3 < CodeRec.BaseAmt1 Then TempRec.CatFee3 = CodeRec.BaseAmt1
                If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
                End If
                GoTo c4
              End If
              If Revenue# <= CodeRec.Recpt2 Then
                TempRec.CatFee3 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
                If TempRec.CatFee3 < CodeRec.BaseAmt2 Then TempRec.CatFee3 = CodeRec.BaseAmt2
                If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
                End If
                GoTo c4
              End If
              If Revenue# <= CodeRec.Recpt3 Then
                TempRec.CatFee3 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
                If TempRec.CatFee3 < CodeRec.BaseAmt3 Then TempRec.CatFee3 = CodeRec.BaseAmt3
                If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
                End If
                GoTo c4
              End If
              If Revenue# <= CodeRec.Recpt4 Then
                TempRec.CatFee3 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
                If TempRec.CatFee3 < CodeRec.BaseAmt4 Then TempRec.CatFee3 = CodeRec.BaseAmt4
                If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
                End If
                GoTo c4
              End If
              If Revenue# <= CodeRec.Recpt5 Then
                TempRec.CatFee3 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
                If TempRec.CatFee3 < CodeRec.BaseAmt5 Then TempRec.CatFee3 = CodeRec.BaseAmt5
                If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
                End If
                GoTo c4
              End If
              If Revenue# <= CodeRec.Recpt6 Then
                TempRec.CatFee3 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
                If TempRec.CatFee3 < CodeRec.BaseAmt6 Then TempRec.CatFee3 = CodeRec.BaseAmt6
                If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
                End If
                GoTo c4
              End If
            End If
          End If
        End If  'End Test for Code
      Next Snt&
    Else
      TempRec.CatFee3 = 0
    End If      'End Test for Cat 3

c4:
    CatCode$ = QPTrim$(CustRec.BILLCAT4)
    If Len(CatCode$) > 0 Then
      For Snt& = 1 To NumOfARCatRecs
        Get CodeHandle, Snt&, CodeRec
        If QPTrim$(CodeRec.CatCode) = CatCode$ Then
          TempRec.CatCodeRec4 = Snt&
          If CodeRec.CodeType = "F" Then
            If ProrateFlag = True Then
              TempRec.CatFee4 = OldRound(Prorate# * CodeRec.Fee)
            Else
              TempRec.CatFee4 = CodeRec.Fee
            End If
            If TempRec.CatFee4 < 0 Then TempRec.CatFee4 = 0
            If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
              Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
              TempRec.CreditUsed = True
              TempRec.PenBal = OverPen
            Else
              TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
            End If
            GoTo c5
          End If
          If CodeRec.CodeType = "M" Then
            Mult = CustRec.REV4
            If ProrateFlag = True Then
              TempRec.CatFee4 = OldRound(Mult * CodeRec.Fee)
              TempRec.CatFee4 = OldRound(TempRec.CatFee4 * Prorate#)
            Else
              TempRec.CatFee4 = OldRound(Mult * CodeRec.Fee)
            End If
            If TempRec.CatFee4 < 0 Then TempRec.CatFee4 = 0
            If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
              Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
              TempRec.CreditUsed = True
              TempRec.PenBal = OverPen
            Else
              TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
            End If
            GoTo c5
          End If
          If CodeRec.CodeType = "S" Then
            Revenue# = CustRec.REV4
            If ProrateFlag = True Then
              If Revenue# <= CodeRec.Recpt1 Then
                TempRec.CatFee4 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
                If TempRec.CatFee4 < CodeRec.BaseAmt1 Then TempRec.CatFee4 = CodeRec.BaseAmt1
                TempRec.CatFee4 = OldRound(Prorate# * TempRec.CatFee4)
                If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
                End If
                GoTo c5
              End If
              If Revenue# <= CodeRec.Recpt2 Then
                TempRec.CatFee4 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
                If TempRec.CatFee4 < CodeRec.BaseAmt2 Then TempRec.CatFee4 = CodeRec.BaseAmt2
                TempRec.CatFee4 = OldRound(Prorate# * TempRec.CatFee4)
                If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
                End If
                GoTo c5
              End If
              If Revenue# <= CodeRec.Recpt3 Then
                TempRec.CatFee4 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
                If TempRec.CatFee4 < CodeRec.BaseAmt3 Then TempRec.CatFee4 = CodeRec.BaseAmt3
                TempRec.CatFee4 = OldRound(Prorate# * TempRec.CatFee4)
                If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
                End If
                GoTo c5
              End If
              If Revenue# <= CodeRec.Recpt4 Then
                TempRec.CatFee4 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
                If TempRec.CatFee4 < CodeRec.BaseAmt4 Then TempRec.CatFee4 = CodeRec.BaseAmt4
                TempRec.CatFee4 = OldRound(Prorate# * TempRec.CatFee4)
                If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
                End If
                GoTo c5
              End If
              If Revenue# <= CodeRec.Recpt5 Then
                TempRec.CatFee4 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
                If TempRec.CatFee4 < CodeRec.BaseAmt5 Then TempRec.CatFee4 = CodeRec.BaseAmt5
                TempRec.CatFee4 = OldRound(Prorate# * TempRec.CatFee4)
                If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
                End If
                GoTo c5
              End If
              If Revenue# <= CodeRec.Recpt6 Then
                TempRec.CatFee4 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
                If TempRec.CatFee4 < CodeRec.BaseAmt6 Then TempRec.CatFee4 = CodeRec.BaseAmt6
                TempRec.CatFee4 = OldRound(Prorate# * TempRec.CatFee4)
                If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
                End If
                GoTo c5
              End If
            Else 'ProrateFlag = False
              If Revenue# <= CodeRec.Recpt1 Then
                TempRec.CatFee4 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
                If TempRec.CatFee4 < CodeRec.BaseAmt1 Then TempRec.CatFee4 = CodeRec.BaseAmt1
                If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
                End If
                GoTo c5
              End If
              If Revenue# <= CodeRec.Recpt2 Then
                TempRec.CatFee4 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
                If TempRec.CatFee4 < CodeRec.BaseAmt2 Then TempRec.CatFee4 = CodeRec.BaseAmt2
                If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
                End If
                GoTo c5
              End If
              If Revenue# <= CodeRec.Recpt3 Then
                TempRec.CatFee4 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
                If TempRec.CatFee4 < CodeRec.BaseAmt3 Then TempRec.CatFee4 = CodeRec.BaseAmt3
                If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
                End If
                GoTo c5
              End If
              If Revenue# <= CodeRec.Recpt4 Then
                TempRec.CatFee4 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
                If TempRec.CatFee4 < CodeRec.BaseAmt4 Then TempRec.CatFee4 = CodeRec.BaseAmt4
                If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
                End If
                GoTo c5
              End If
              If Revenue# <= CodeRec.Recpt5 Then
                TempRec.CatFee4 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
                If TempRec.CatFee4 < CodeRec.BaseAmt5 Then TempRec.CatFee4 = CodeRec.BaseAmt5
                If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
                End If
                GoTo c5
              End If
              If Revenue# <= CodeRec.Recpt6 Then
                TempRec.CatFee4 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
                If TempRec.CatFee4 < CodeRec.BaseAmt6 Then TempRec.CatFee4 = CodeRec.BaseAmt6
                If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
                End If
                GoTo c5
              End If
            End If
          End If
        End If  'End Test for Code
      Next Snt&
    Else
      TempRec.CatFee4 = 0
    End If      'End Test for Cat 1

c5:
    CatCode$ = QPTrim$(CustRec.BILLCAT5)
    If Len(CatCode$) > 0 Then
      For Snt& = 1 To NumOfARCatRecs
        Get CodeHandle, Snt&, CodeRec
        If QPTrim$(CodeRec.CatCode) = CatCode$ Then
          TempRec.CatCodeRec5 = Snt&
          If CodeRec.CodeType = "F" Then
            If ProrateFlag = True Then
              TempRec.CatFee5 = OldRound(Prorate# * CodeRec.Fee)
            Else
              TempRec.CatFee5 = CodeRec.Fee
            End If
            If TempRec.CatFee5 < 0 Then TempRec.CatFee5 = 0
            If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
              Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
              TempRec.CreditUsed = True
              TempRec.PenBal = OverPen
            Else
              TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
            End If
            GoTo FinishSaving
          End If
          If CodeRec.CodeType = "M" Then
            Mult = CustRec.REV5
            If ProrateFlag = True Then
              TempRec.CatFee5 = OldRound(Mult * CodeRec.Fee)
              TempRec.CatFee5 = OldRound(TempRec.CatFee5 * Prorate#)
            Else
              TempRec.CatFee5 = OldRound(Mult * CodeRec.Fee)
            End If
            If TempRec.CatFee5 < 0 Then TempRec.CatFee5 = 0
            If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
              Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
              TempRec.CreditUsed = True
              TempRec.PenBal = OverPen
            Else
              TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
            End If
            GoTo FinishSaving
          End If
          If CodeRec.CodeType = "S" Then
            Revenue# = CustRec.REV5
            If ProrateFlag = True Then
              If Revenue# <= CodeRec.Recpt1 Then
                TempRec.CatFee5 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
                If TempRec.CatFee5 < CodeRec.BaseAmt1 Then TempRec.CatFee5 = CodeRec.BaseAmt1
                TempRec.CatFee5 = OldRound(Prorate# * TempRec.CatFee5)
                If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
                End If
                GoTo FinishSaving
              End If
              If Revenue# <= CodeRec.Recpt2 Then
                TempRec.CatFee5 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
                If TempRec.CatFee5 < CodeRec.BaseAmt2 Then TempRec.CatFee5 = CodeRec.BaseAmt2
                TempRec.CatFee5 = OldRound(Prorate# * TempRec.CatFee5)
                If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
                End If
                GoTo FinishSaving
              End If
              If Revenue# <= CodeRec.Recpt3 Then
                TempRec.CatFee5 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
                If TempRec.CatFee5 < CodeRec.BaseAmt3 Then TempRec.CatFee5 = CodeRec.BaseAmt3
                TempRec.CatFee5 = OldRound(Prorate# * TempRec.CatFee5)
                If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
                End If
                GoTo FinishSaving
              End If
              If Revenue# <= CodeRec.Recpt4 Then
                TempRec.CatFee5 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
                If TempRec.CatFee5 < CodeRec.BaseAmt4 Then TempRec.CatFee5 = CodeRec.BaseAmt4
                TempRec.CatFee5 = OldRound(Prorate# * TempRec.CatFee5)
                If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
                End If
                GoTo FinishSaving
              End If
              If Revenue# <= CodeRec.Recpt5 Then
                TempRec.CatFee5 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
                If TempRec.CatFee5 < CodeRec.BaseAmt5 Then TempRec.CatFee5 = CodeRec.BaseAmt5
                TempRec.CatFee5 = OldRound(Prorate# * TempRec.CatFee5)
                If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
                End If
                GoTo FinishSaving
              End If
              If Revenue# <= CodeRec.Recpt6 Then
                TempRec.CatFee5 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
                If TempRec.CatFee5 < CodeRec.BaseAmt6 Then TempRec.CatFee5 = CodeRec.BaseAmt6
                TempRec.CatFee5 = OldRound(Prorate# * TempRec.CatFee5)
                If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
                End If
                GoTo FinishSaving
              End If
            Else 'ProrateFlag = False
              If Revenue# <= CodeRec.Recpt1 Then
                TempRec.CatFee5 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
                If TempRec.CatFee5 < CodeRec.BaseAmt1 Then TempRec.CatFee5 = CodeRec.BaseAmt1
                If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
                End If
                GoTo FinishSaving
              End If
              If Revenue# <= CodeRec.Recpt2 Then
                TempRec.CatFee5 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
                If TempRec.CatFee5 < CodeRec.BaseAmt2 Then TempRec.CatFee5 = CodeRec.BaseAmt2
                If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
                End If
                GoTo FinishSaving
              End If
              If Revenue# <= CodeRec.Recpt3 Then
                TempRec.CatFee5 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
                If TempRec.CatFee5 < CodeRec.BaseAmt3 Then TempRec.CatFee5 = CodeRec.BaseAmt3
                If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
                End If
                GoTo FinishSaving
              End If
              If Revenue# <= CodeRec.Recpt4 Then
                TempRec.CatFee5 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
                If TempRec.CatFee5 < CodeRec.BaseAmt4 Then TempRec.CatFee5 = CodeRec.BaseAmt4
                If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
                End If
                GoTo FinishSaving
              End If
              If Revenue# <= CodeRec.Recpt5 Then
                TempRec.CatFee5 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
                If TempRec.CatFee5 < CodeRec.BaseAmt5 Then TempRec.CatFee5 = CodeRec.BaseAmt5
                If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
                End If
                GoTo FinishSaving
              End If
              If Revenue# <= CodeRec.Recpt6 Then
                TempRec.CatFee5 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
                If TempRec.CatFee5 < CodeRec.BaseAmt6 Then TempRec.CatFee5 = CodeRec.BaseAmt6
                If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                  Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                  TempRec.CreditUsed = True
                  TempRec.PenBal = OverPen
                Else
                  TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
                End If
                GoTo FinishSaving
              End If
            End If
          End If
        End If  'End Test for Code
      Next Snt&
    Else
      TempRec.CatFee5 = 0
    End If      'End Test for Cat 1
FinishSaving:
    TempRec.CustomerNumber = CStr(IdxRecs(cnt))
'    TempRec.ChargeAccount = False
    TempRec.LICENSE = QPTrim$(CustRec.LICENSE)
    TempRec.Posted2GL = "N"
    TempRec.Prev = 0
    TempRec.TransDesc = "LICENSE"
    TempRec.TransType = 1
    TempRec.VALID = 0
    'each category fee is re-calculated above taking into account
    'any negative outstanding balances
    TempRec.LicBal = TempRec.CatFeeBal1 + TempRec.CatFeeBal2 + TempRec.CatFeeBal3 + TempRec.CatFeeBal4 + TempRec.CatFeeBal5
    TempRec.AcctBal = TempRec.PenBal + TempRec.LicBal + TempRec.IssFeeBal
    TempRec.BalanceAfterTrans = TempRec.AcctBal
'    TempRec.TransAmount = TempRec.CatFee1 + TempRec.CatFee2 + TempRec.CatFee3 + TempRec.CatFee4 + TempRec.CatFee5 + TempRec.IssFee
    TempRec.TransAmount = TempRec.CatFee1 + TempRec.CatFee2 + TempRec.CatFee3 + TempRec.CatFee4 + TempRec.CatFee5 + TownRec.IssFee
    TempRec.TransDate = 0

    Put TempHandle, Nextx, TempRec
    Nextx = Nextx + 1
SkipEm:
    frmBLShowPctComp.ShowPctComp cnt, NumOfCustIdxRecs 'NumOfCustRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      cmdHelp.Enabled = True
      Exit Sub
    End If
  Next cnt

  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True

  Close         'Close all open files now

  Exit Sub
  
CalcLicBal:
  
  'we only want to bring negative balances closer to zero here
  'since the overall license balance is negative then at least one
  'of the individual license balances has to be negative
  
  If TempRec.CatFeeBal1 >= 0 Then GoTo NextOne 'this isn't negative so move on
  'If the issue fee reduces the bal1 amount to zero but leaves a positive
  'amount in iss fee then carry the iss fee balance to the next category.
  'Otherwise bring the bal1 amount closer to zero and make iss fee zero and
  'then you're done
  If Abs(TempRec.CatFeeBal1) >= IssFee Then
    TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + IssFee 'adding IssFee
    OverLic = OverLic - IssFee
    'brings the negative Bal1 closer to zero
    TempRec.CreditUsed = True
    IssFee = 0
    GoTo DoneHere
  Else
    'there is a balance in IssFee so reduce the
    'IssFee balance by the amount added to Bal1
    'and go to the next category
    IssFee = IssFee - Abs(TempRec.CatFeeBal1)
    TempRec.CreditUsed = True
    TempRec.CatFeeBal1 = 0
    OverLic = 0
  End If
  
NextOne:
  If TempRec.CatFeeBal2 >= 0 Then GoTo NextTwo
  If Abs(TempRec.CatFeeBal2) >= IssFee Then
    TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + IssFee
    OverLic = OverLic - IssFee
    TempRec.CreditUsed = True
    IssFee = 0
    GoTo DoneHere
  Else
    IssFee = IssFee - Abs(TempRec.CatFeeBal2)
    TempRec.CreditUsed = True
    TempRec.CatFeeBal2 = 0
    OverLic = 0
  End If
  
NextTwo:
  If TempRec.CatFeeBal3 >= 0 Then GoTo NextThree
  If Abs(TempRec.CatFeeBal3) >= IssFee Then
    TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + IssFee
    OverLic = OverLic - IssFee
    TempRec.CreditUsed = True
    IssFee = 0
    GoTo DoneHere
  Else
    IssFee = IssFee - Abs(TempRec.CatFeeBal3)
    TempRec.CreditUsed = True
    TempRec.CatFeeBal3 = 0
    OverLic = 0
  End If
 
NextThree:
  If TempRec.CatFeeBal4 >= 0 Then GoTo NextFour
  If Abs(TempRec.CatFeeBal4) >= IssFee Then
    TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + IssFee
    OverLic = OverLic - IssFee
    TempRec.CreditUsed = True
    IssFee = 0
    GoTo DoneHere
  Else
    IssFee = IssFee - Abs(TempRec.CatFeeBal4)
    TempRec.CatFeeBal4 = 0
    TempRec.CreditUsed = True
    OverLic = 0
  End If
  
NextFour:
  If TempRec.CatFeeBal5 >= 0 Then GoTo DoneHere
  If Abs(TempRec.CatFeeBal5) >= IssFee Then
    TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + IssFee
    OverLic = OverLic - IssFee
    TempRec.CreditUsed = True
    IssFee = 0
    GoTo DoneHere
  Else
    IssFee = IssFee - Abs(TempRec.CatFeeBal5)
    TempRec.CreditUsed = True
    TempRec.CatFeeBal5 = 0
    OverLic = 0
  End If

DoneHere:
  Return
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLLicRegister", "SetFee", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
  
  
End Sub

Private Sub PrintGraphics()
  Dim ReportFile$
  Dim SubReportFile$
  Dim x As Double
  Dim CodeRec As ARNewCatCodeRecType
  Dim CHandle As Integer
  Dim NumOfCodeRecs As Integer
  Dim NumOfARCatRecs As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim CustNameIdxRec As CustSearchNameIdxType
  Dim CustIdxRec As CustNumIdxType
  Dim IdxHandle As Integer
  Dim NumOfCustIdxRecs As Double
  Dim NameFlag As Boolean
  Dim NumFlag As Boolean
  Dim RptHandle As Integer
  Dim SubRptHandle As Integer
  Dim cnt&, CustFee#
  Dim CCode$
  Dim TC$, TotalFee#
  Dim CatIdxRecs As CatCodeIdxType
  Dim CatIdxHandle As Integer
  Dim NumOfCatIdxRecs As Integer
  Dim ThisCode$
  Dim ThisFee As Double
  Dim CustCnt As Integer
  Dim dlm$, TownRec As TownSetUpType
  Dim TownName$, TownHandle As Integer
  Dim TempRec As TempTransPostType
  Dim TempHandle As Integer
  Dim NumOfTempRecs As Integer
  Dim NumOfCustRecs As Integer
  Dim IssFeeTot As Double
  Dim y As Integer
  
  On Error GoTo ERRORSTUFF
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  Call SetFee
  
  OpenCustFile CustHandle
  
  OpenCatCodeIdxFile CatIdxHandle
  NumOfCatIdxRecs = LOF(CatIdxHandle) / Len(CatIdxRecs)
  
  OpenTempPostFile TempHandle
  NumOfTempRecs = LOF(TempHandle) / Len(TempRec)
  If NumOfTempRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No businesses qualify for a new business license. If a customer is involved in a penalty file they won't qualify. If you wish to set customer flags go to the 'Set License To Print' screen or go to the Customer Edit screen to reset a specific customer's 'Set Renewal Flag (Y/N)?' field to Yes."
    frmBLMessageBoxJr.Label1.Top = 400
    frmBLMessageBoxJr.Label1.Height = 1500
    frmBLMessageBoxJr.Show vbModal
    Close
    cmdExit.Enabled = True
    cmdProcess.Enabled = True
    cmdHelp.Enabled = True
    Exit Sub
  End If
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  TownName = QPTrim$(TownRec.TownName) + ", " + QPTrim$(TownRec.State)
  dlm = "~"

  ReDim CatIdx(1 To NumOfCatIdxRecs) As String
  ReDim CatDesc(1 To NumOfCatIdxRecs) As String
  ReDim CatFeeAmt(1 To NumOfCatIdxRecs) As Double
  ReDim CatCnt(1 To NumOfCatIdxRecs) As Integer
  
  frmBLLoadReport.Label1.Caption = "Linking To Report"
  frmBLLoadReport.Label2.Visible = False
  frmBLLoadReport.Show
  DoEvents
  
  OpenCatCodeFile CHandle
  NumOfCodeRecs = LOF(CHandle) / Len(CodeRec)
  For x = 1 To NumOfCatIdxRecs
    Get CatIdxHandle, x, CatIdxRecs
    CatIdx(x) = QPTrim$(CatIdxRecs.CatCodeNum)
    For y = 1 To NumOfCodeRecs
      Get CHandle, y, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatIdx(x) Then
        If NumOfCodeRecs = NumOfCatIdxRecs Then
          CatDesc(x) = QPTrim$(CodeRec.CODEDESC)
        Else
          CatDesc(x) = ""
        End If
        Exit For
      End If
    Next y
  Next x
  Close CatIdxHandle
  Close CHandle
  
  ReportFile$ = "BLRPTS\ARLICREG.RPT"

  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  For cnt& = 1 To NumOfTempRecs ' NumOfCustRecs 'NumOfARRecs
    Get TempHandle, cnt, TempRec
    Get CustHandle, Val(TempRec.CustomerNumber), CustRec
    
    If QPTrim$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SortName) = "DELETED" Then
      GoTo DelSkip
    End If
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo DelSkip
    
    CustCnt = CustCnt + 1
    IssFeeTot = IssFeeTot + TownRec.IssFee
    
    CustFee# = OldRound(TempRec.CatFee1 + TempRec.CatFee2 + TempRec.CatFee3 + TempRec.CatFee4 + TempRec.CatFee5 + TownRec.IssFee)
    'CCode accumulates category descriptions to print on
    'the register
    CCode$ = QPTrim$(CustRec.BILLCAT1)
    ThisCode = CCode
    ThisFee = TempRec.CatFee1
    GoSub CollectTotals
    TC$ = QPTrim$(CustRec.BILLCAT2)
    If Len(TC$) > 0 Then
      ThisCode = TC$
      ThisFee = TempRec.CatFee2
      GoSub CollectTotals
      CCode$ = CCode$ + "/" + TC$
    End If
    TC$ = QPTrim$(CustRec.BILLCAT3)
    If Len(TC$) > 0 Then
      ThisCode = TC$
      ThisFee = TempRec.CatFee3
      GoSub CollectTotals
      CCode$ = CCode$ + "/" + TC$
    End If
    TC$ = QPTrim$(CustRec.BILLCAT4)
    If Len(TC$) > 0 Then
      ThisCode = TC$
      ThisFee = TempRec.CatFee4
      GoSub CollectTotals
      CCode$ = CCode$ + "/" + TC$
    End If
    TC$ = QPTrim$(CustRec.BILLCAT5)
    If Len(TC$) > 0 Then
      ThisCode = TC$
      ThisFee = TempRec.CatFee5
      GoSub CollectTotals
      CCode$ = CCode$ + "/" + TC$
    End If
    If TempRec.CreditUsed = False Then
      '                                   0                                 1                   2             3              4                     5
      Print #RptHandle, QPTrim$(TempRec.CustomerNumber); dlm; QPTrim$(CustRec.BillName); dlm; CCode$; dlm; CustFee#; dlm; TownName; dlm; CStr(CustRec.Prorate); dlm;
    Else
      '                                   0                                 1                   2                     3              4                   5
      Print #RptHandle, QPTrim$(TempRec.CustomerNumber); dlm; QPTrim$(CustRec.BillName); dlm; CCode$ + " *"; dlm; CustFee#; dlm; TownName; dlm; CStr(CustRec.Prorate); dlm;
    End If
    
    If TownRec.IssFee > 0 Then
      '                                         6
      Print #RptHandle, QPTrim$(Using("$#,##0.00", TownRec.IssFee)) + " issuance fee charged to each customer."; dlm; 1
    Else
      '                 6
      Print #RptHandle, ""; dlm; 1
    End If
    
    TotalFee# = OldRound(TotalFee# + CustFee#)
      

DelSkip:
  Next
  GoSub PrintLicRegRptEnding
  Close         'Close all open files now
  
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True
  
  DoEvents
  
  arBLLicRegister.Show
  frmBLLoadReport.Label1.Caption = "Loading ......"
  frmBLLoadReport.Label2.Visible = True
  frmBLLoadReport.Show
  MainLog ("Business license registers processed in graphics format.")
  Exit Sub
  
CollectTotals:
  'collect totals by category code
  For x = 1 To NumOfCatIdxRecs
    If QPTrim$(ThisCode) = QPTrim$(CatIdx(x)) Then
      CatFeeAmt(x) = CatFeeAmt(x) + ThisFee
      CatCnt(x) = CatCnt(x) + 1
      Exit For
    End If
  Next x
  
  Return

PrintLicRegRptEnding:

  SubReportFile$ = "BLRPTS\ARSUBLICREG.RPT"
  SubRptHandle = FreeFile
  Open SubReportFile$ For Output As #SubRptHandle

  For x = 1 To NumOfCatIdxRecs
    If CatFeeAmt(x) > 0 Then
      '                         0               1                2              3               4              5
      Print #SubRptHandle, CatIdx(x); dlm; CatFeeAmt(x); dlm; CustCnt; dlm; TotalFee#; dlm; CatDesc(x); dlm; CatCnt(x)
    End If
  Next x
  If IssFeeTot > 0 Then
    '                           0                 1              2              3            4        5
    Print #SubRptHandle, "Issuance Fee"; dlm; IssFeeTot; dlm; CustCnt; dlm; TotalFee#; dlm; ""; dlm; ""
  End If
    
  Close SubRptHandle
  
  Return


ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLLicRegister", "PrintGraphics", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
  

End Sub

Private Sub ApplyCredits2ThisFee(ByRef ThisBal, ByVal ThisTFee As Double, ByRef OverPen As Double)
  
  On Error GoTo ERRORSTUFF
  
  'either ThisBal or OverPen has a negative value, possibly both
  If OverPen > 0 Then 'OverPen = a negative penalty balance
    If ThisTFee >= OverPen Then 'reduce fee by the credit in penalty and bring penalty balance up to 0
      ThisTFee = ThisTFee - OverPen
      OverPen = 0
      CreditFlag = True
    ElseIf ThisTFee < OverPen Then 'reduce fee to 0 then bring penalty credit closer to 0
      OverPen = OverPen - ThisTFee
      ThisTFee = 0
    End If
  End If
  
  'bring any negative outstanding balance closer to zero
  'while reducing this license fee
  If ThisBal < 0 Then 'ThisBal is a negative license balance
    If ThisTFee >= Abs(ThisBal) Then
      ThisTFee = ThisTFee + ThisBal
      ThisBal = ThisTFee 'ThisBal now becomes whatever this category's fee is
    Else
      ThisBal = ThisBal + ThisTFee
      ThisTFee = 0
    End If
  End If
   
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLLicRegister", "ApplyCredits2ThisFee", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
  
  
End Sub

'Private Function ApplyCredits2ThisFeeNew(ByRef ThisBal As Double, ByRef ThisFee As Double, ByRef OverLic As Double, ByRef OverPen As Double) As Boolean
'  ApplyCredits2ThisFee = False
'  If OverPen > 0 Then
'    If ThisFee >= OverPen Then 'reduce fee by the credit in penalty and bring penalty balance up to 0
'      ThisFee = ThisFee - OverPen
'      OverPen = 0
'      CreditFlag = True
'      ApplyCredits2ThisFee = True
'    ElseIf ThisFee < OverPen Then 'reduce fee to 0 then bring penalty credit closer to 0
'      OverPen = OverPen - ThisFee
'      ThisFee = 0
'      CreditFlag = True
'      ApplyCredits2ThisFee = True
'    End If
'  End If
'
'  If OverLic > 0 Then
'    If ThisFee >= OverLic Then
'      ThisFee = ThisFee - OverLic
'      OverLic = 0
'      CreditFlag = True
'      ApplyCredits2ThisFee = True
'    ElseIf ThisFee < OverLic Then
'      OverLic = OverLic - ThisFee
'      ThisFee = 0
'      CreditFlag = True
'      ApplyCredits2ThisFee = True
'    End If
'  End If
'
'End Function
'
