VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "EDT32X30.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmFAEditAssetCode 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbStatus 
      Height          =   384
      Left            =   5328
      TabIndex        =   1
      Top             =   4224
      Width           =   2124
      _Version        =   196608
      _ExtentX        =   3746
      _ExtentY        =   677
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
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmFAEditAssetCode.frx":0000
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ESC &Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   5328
      TabIndex        =   4
      Top             =   6720
      Width           =   1884
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "F10 &SAVE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   7908
      TabIndex        =   5
      Top             =   6720
      Width           =   1884
   End
   Begin VB.CommandButton cmdAssetList 
      Caption         =   "F11 Code &List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   6816
      TabIndex        =   3
      Top             =   5184
      Width           =   1836
   End
   Begin EditLib.fpText fptxtDesc 
      Height          =   396
      Left            =   4368
      TabIndex        =   0
      Top             =   3312
      Width           =   4428
      _Version        =   196608
      _ExtentX        =   7810
      _ExtentY        =   698
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
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
      MaxLength       =   20
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
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtGroupCode 
      Height          =   396
      Left            =   4656
      TabIndex        =   2
      Top             =   5184
      Width           =   2076
      _Version        =   196608
      _ExtentX        =   3662
      _ExtentY        =   698
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   8454143
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
      CharValidationText=   "1, 2, 3, 4, 5, 6, 7, 8, 9, 0, ."
      MaxLength       =   4
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
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Group Code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   2688
      TabIndex        =   9
      Top             =   5280
      Width           =   1788
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   2736
      TabIndex        =   8
      Top             =   3408
      Width           =   1452
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   4224
      TabIndex        =   7
      Top             =   4320
      Width           =   924
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2940
      TabIndex        =   6
      Top             =   1212
      Width           =   6012
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   1068
      Width           =   8652
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3996
      Left            =   1632
      Top             =   2460
      Width           =   8412
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   1020
      Width           =   8652
   End
End
Attribute VB_Name = "frmFAEditAssetCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdAssetList_Click()
  frmFAAssetCodeList.Show vbModal
End Sub

Private Sub cmdExit_Click()
  frmFAAssetsCodesmenu.Show
  DoEvents
  Unload frmFAEditAssetCode
End Sub
Private Function Check4Dups() As Boolean
  Dim CodeHandle As Integer
  Dim CodeRec As FAAssetCodeRecType
  Dim x As Integer
  Dim NumOfRecs As Integer
  Dim CompareThis$
  
  Check4Dups = False
  OpenFACodeNameFile CodeHandle
  NumOfRecs = LOF(CodeHandle) \ Len(CodeRec)
  
  CompareThis = QPTrim$(fptxtDesc.Text)
  If CodeNum = 0 Then
    For x = 1 To NumOfRecs
      Get CodeHandle, x, CodeRec
      If CompareThis = QPTrim$(CodeRec.AssetDesc) Then
        MsgBox "You have entered a description that is already in use. Please choose another description."
        fptxtDesc.SetFocus
        Check4Dups = True
        Exit For
      End If
    Next x
  Else
    For x = 1 To NumOfRecs
      If x <> CodeNum Then
        Get CodeHandle, x, CodeRec
        If CompareThis = QPTrim$(CodeRec.AssetDesc) Then
          MsgBox "You have entered a description that is already in use. Please choose another description."
          fptxtDesc.SetFocus
          Check4Dups = True
          Exit For
        End If
      End If
    Next x
  End If
  
  CompareThis = QPTrim$(fptxtGroupCode.Text)
  If CodeNum = 0 Then
    For x = 1 To NumOfRecs
      Get CodeHandle, x, CodeRec
      If CompareThis = QPTrim$(CodeRec.ASSETCODE) Then
        MsgBox "You have entered a code number that is already in use. Please choose another 4 digit code number."
        fptxtGroupCode.SetFocus
        Check4Dups = True
        Exit For
      End If
    Next x
  Else
    For x = 1 To NumOfRecs
      If x <> CodeNum Then
        Get CodeHandle, x, CodeRec
        If CompareThis = QPTrim$(CodeRec.ASSETCODE) Then
          MsgBox "You have entered a code number that is already in use. Please choose another 4 digit code number."
          fptxtGroupCode.SetFocus
          Check4Dups = True
          Exit For
        End If
      End If
    Next x
  End If
  
End Function
Private Sub cmdSave_Click()
  Dim CodeHandle As Integer
  Dim CodeRec As FAAssetCodeRecType
  Dim NumOfRecs As Integer
  Dim DoWhatFlag As WarnOption
  Dim ChangeFlag As Boolean
  
  If QPTrim$(fptxtDesc.Text) = "" Then
    MsgBox "Please enter a description for this asset code"
    Exit Sub
  End If
  ChangeFlag = False
  
  If Len(QPTrim$(fptxtGroupCode.Text)) <> 4 Then
    MsgBox "Please enter a 4 digit number for the code."
    fptxtGroupCode.SetFocus
    Close CodeHandle
    Exit Sub
  End If
  
  If Check4Dups = True Then Exit Sub
  
  OpenFACodeNameFile CodeHandle
  
  If CodeNum > 0 Then
    Get CodeHandle, CodeNum, CodeRec
    If QPTrim$(fptxtDesc.Text) <> QPTrim$(CodeRec.AssetDesc) Then
      ChangeFlag = True
      fptxtDesc.SetFocus
    ElseIf QPTrim$(fpcmbStatus.Text) <> QPTrim$(CodeRec.AssetStatus) Then
      ChangeFlag = True
      fpcmbStatus.SetFocus
    ElseIf QPTrim$(fptxtGroupCode.Text) <> QPTrim$(CodeRec.ASSETCODE) Then
      ChangeFlag = True
      fptxtGroupCode.SetFocus
    End If
    If ChangeFlag = True Then
      DoWhatFlag = PromptWarnOverWrite(Me)
      Select Case DoWhatFlag
        Case WarnOption.wSave
        Case WarnOption.wExit
          Close CodeHandle
          frmFAAssetsCodesmenu.Show
          DoEvents
          Unload frmFAEditAssetCode
          Exit Sub
        Case WarnOption.wReturn
          Close CodeHandle
          Exit Sub
        Case WarnOption.wGo2Add
          Close CodeHandle
          CodeNum = 0
          Call LoadMe
          Exit Sub
        Case Else
          Close CodeHandle
          MsgBox "Please make a valid selection"
          Exit Sub
      End Select
    End If
  End If
  NumOfRecs = LOF(CodeHandle) \ Len(CodeRec)
  If CodeNum = 0 Then
    CodeRec.ASSETCODE = QPTrim$(fptxtGroupCode.Text)
    CodeRec.AssetDesc = QPTrim$(fptxtDesc.Text)
    CodeRec.AssetStatus = QPTrim$(fpcmbStatus.Text)
    Put CodeHandle, NumOfRecs + 1, CodeRec
    Close CodeHandle
  Else
    CodeRec.ASSETCODE = QPTrim$(fptxtGroupCode.Text)
    CodeRec.AssetDesc = QPTrim$(fptxtDesc.Text)
    CodeRec.AssetStatus = QPTrim$(fpcmbStatus.Text)
    Put CodeHandle, CodeNum, CodeRec
    Close CodeHandle
  End If
  
  MsgBox "Your information has been saved"
  frmFAAssetsCodesmenu.Show
  DoEvents
  Unload frmFAEditAssetCode
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%D"
      KeyCode = 0
    Case vbKeyF9:
      SendKeys "%T"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      KeyCode = 0
    Case vbKeyF11:
      SendKeys "%L"
      KeyCode = 0
    Case vbKeyF12:
      SendKeys "%G"
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
      Call UnloadAllFormsAndOpn
'      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmPayrollMainMenu.")
      End
    End If
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub LoadMe()
  Dim CodeHandle As Integer
  Dim CodeRec As FAAssetCodeRecType
  
  If CodeNum = 0 Then
    Me.Caption = "Adding Asset Code"
    Me.Label2 = "Adding Fixed Asset Code"
    fpcmbStatus.Text = "Inactive"
    fptxtGroupCode.Text = ""
    fptxtDesc.Text = ""
  Else
    Me.Caption = "Editing Asset Code"
    Me.Label2 = "Editing Fixed Asset Code"
    OpenFACodeNameFile CodeHandle
    Get CodeHandle, CodeNum, CodeRec
    fpcmbStatus.Text = CodeRec.AssetStatus
    fptxtGroupCode.Text = CodeRec.ASSETCODE
    fptxtDesc.Text = CodeRec.AssetDesc
    Close CodeHandle
  End If
  
  fpcmbStatus.AddItem "Active"
  fpcmbStatus.AddItem "Inactive"
End Sub

Private Sub fpcmbStatus_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbStatus.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbStatus.ListIndex = -1
  End If
  If fpcmbStatus.ListDown <> True Then
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
