VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFASystemSetup 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets System Setup"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmSystemSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbDepType 
      Height          =   405
      Left            =   5010
      TabIndex        =   1
      ToolTipText     =   "Select the Depreciation Type this system will use to calculate fixed asset depreciation."
      Top             =   3690
      Width           =   4035
      _Version        =   196608
      _ExtentX        =   7117
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
      ColDesigner     =   "frmSystemSetup.frx":08CA
   End
   Begin EditLib.fpText fptxtTownName 
      Height          =   396
      Left            =   4356
      TabIndex        =   0
      ToolTipText     =   "Enter the name of the town which owns the fixed assets controlled in this module."
      Top             =   2832
      Width           =   4620
      _Version        =   196608
      _ExtentX        =   8149
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
      MaxLength       =   150
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
   Begin EditLib.fpText fptxtFirstYearPct 
      Height          =   396
      Left            =   6324
      TabIndex        =   2
      ToolTipText     =   "If Fixed 1st year is selected as the depreciation type then enter the desired first year percentage. "
      Top             =   4512
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 ."
      MaxLength       =   14
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
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   675
      Left            =   1710
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Click on this button to bring up an explanation on how each type of depreciation works."
      Top             =   6150
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   1191
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmSystemSetup.frx":0BC1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   690
      Left            =   4884
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6144
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   1217
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmSystemSetup.frx":0D9C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   690
      Left            =   8052
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to commit the data entered above to memory."
      Top             =   6144
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   1217
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   13684944
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
      ButtonDesigner  =   "frmSystemSetup.frx":0F78
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   2940
      Left            =   1008
      Top             =   2400
      Width           =   9804
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Depreciation Type:"
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
      Left            =   2016
      TabIndex        =   6
      Top             =   3792
      Width           =   2700
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "First Year Percentage:"
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
      Left            =   3264
      TabIndex        =   5
      Top             =   4608
      Width           =   2700
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Town Name:"
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
      Left            =   2532
      TabIndex        =   4
      Top             =   2976
      Width           =   1548
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FIXED ASSETS SYSTEM SETUP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   2940
      TabIndex        =   3
      Top             =   1008
      Width           =   6012
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   816
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   768
      Width           =   8652
   End
End
Attribute VB_Name = "frmFASystemSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim TempTownName$
  Dim TempPct1St As Double
  Dim TempDeprType$

Private Sub cmdExit_Click()
  Dim FASetup As FASetupRecType
  Dim SetupHandle As Integer
  Dim ChangeFlag As Boolean
  Dim DoWhatFlag As SaveChangeOptions1
  
  On Error GoTo ERRORSTUFF
  'exit procedure examines all the screen's fields and if
  'a change is detected than the user is given a chance
  'to save this change before exiting
  If Not Exist("FASETUP.DAT") Then
    Close
    GoTo RecNumIsZero
  End If
  
  ChangeFlag = False
  OpenFASetUpFile SetupHandle
  Get SetupHandle, 1, FASetup
  Close
  
  If QPTrim$(fptxtTownName) <> QPTrim$(FASetup.TownName) Then
    ChangeFlag = True
    fptxtTownName.SetFocus
    GoTo ChangeFound
  End If
  
  If Val(fptxtFirstYearPct) <> FASetup.Pct1St Then
    ChangeFlag = True
    fptxtFirstYearPct.SetFocus
    GoTo ChangeFound
  End If

  If QPTrim$(fpcmbDepType.Text) = "NOT SAVED" Then
    If MsgBox("The Depreciation Type is not valid. Do you wish to make a selection from the three choices now?", vbYesNo) = vbYes Then
      fpcmbDepType.SetFocus
      Exit Sub
    End If
  End If
  
  If QPTrim$(fpcmbDepType.Text) <> QPTrim$(FASetup.DeprType) Then
    ChangeFlag = True
    fpcmbDepType.SetFocus
  End If

ChangeFound:
  If ChangeFlag = True Then
    ChangeFlag = False
    DoWhatFlag = PromptSaveChanges(Me)
    Select Case DoWhatFlag
    Case SaveChangeOptions1.scoSaveChanges 'save changes
      Call cmdSave_Click
      Exit Sub
    Case SaveChangeOptions1.scoReviewChanges 'review is just bringing back the current form
      Exit Sub
    Case SaveChangeOptions1.scoAbandonChanges 'abandon
      frmFAMaintMenu.Show
      DoEvents
      Unload frmFASystemSetup
      Exit Sub
    Case Else:
    'Do nothing because we don't know about any options except
    'save, review or abandon...used as a placeholder for adding
    'other options at a later date
    End Select
  End If
RecNumIsZero:
  If Exist("fromBuildDep.dat") Then
    frmFABuildYrEndDep.Show
    KillFile "fromBuildDep.dat"
  Else
    frmFAMaintMenu.Show
  End If
  Close
  DoEvents
  Unload frmFASystemSetup
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFASystemSetup", "cmdExit_Click", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Unload Me
End Sub

Private Sub cmdHelp_Click()
  frmDeprHelp.Show vbModal
End Sub

Private Sub cmdSave_Click()
  Dim FASetup As FASetupRecType
  Dim SetupHandle As Integer
  Dim NewFlag As Boolean
  
  On Error GoTo ERRORSTUFF
  NewFlag = False
  
  If Not Exist("FASETUP.DAT") Then
    NewFlag = True
  End If
  
  If QPTrim$(fpcmbDepType.Text) = "NOT SAVED" Or QPTrim$(fpcmbDepType.Text) = "" Then
    MsgBox "Please save one of the three depreciation types listed."
    fpcmbDepType.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtTownName.Text) = "" Then
    MsgBox "Please enter a town name."
    fptxtTownName.SetFocus
    Exit Sub
  End If
  
  FASetup.TownName = QPTrim$(fptxtTownName)
  FASetup.DeprType = QPTrim$(fpcmbDepType.Text)
  
  If FASetup.DeprType = "Fixed 1st year percentage" Then
    FASetup.Pct1St = Val(fptxtFirstYearPct)
  Else
    FASetup.Pct1St = 100
  End If
  
  If FASetup.DeprType = "Prorate 1st year" Then
    FASetup.PRate1St = "Y"
  Else
    FASetup.PRate1St = "N"
  End If
  FASetup.Filler1 = ""
  OpenFASetUpFile SetupHandle
  Put SetupHandle, 1, FASetup
  Close SetupHandle
  
  If NewFlag = False Then
    Call LogSaves
  End If
  
  MsgBox "Your information has been saved"
  If Exist("fromBuildDep.dat") Then
    frmFABuildYrEndDep.Show
    KillFile "fromBuildDep.dat"
  Else
    frmFAMaintMenu.Show
  End If
  DoEvents
  Unload frmFASystemSetup
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFASystemSetup", "cmdSave_Click", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Unload Me
  
End Sub

Private Sub Command1_Click()

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
    'Me.Visible = False
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
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%H"
      Call cmdHelp_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
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
      MainLog ("FixedAssets.exe terminated via menu bar on frmFASystemSetup.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim FASetup As FASetupRecType
  Dim SetupHandle As Integer
  
  fpcmbDepType.AddItem "Prorate 1st year"
  fpcmbDepType.AddItem "Fixed 1st year percentage"
  fpcmbDepType.AddItem "Whole year"
  If Exist("FASETUP.DAT") Then
    OpenFASetUpFile SetupHandle
    Get SetupHandle, 1, FASetup
    Close SetupHandle
    fptxtTownName = QPTrim$(FASetup.TownName)
    TempTownName$ = QPTrim$(FASetup.TownName) 'global
    fptxtFirstYearPct = Using$("##0.00", FASetup.Pct1St)
    TempPct1St = FASetup.Pct1St 'global
    fpcmbDepType.Text = QPTrim$(FASetup.DeprType)
    TempDeprType$ = QPTrim$(FASetup.DeprType) 'global
  Else
    fptxtTownName.Text = ""
    fptxtFirstYearPct.Text = "0.00"
    fpcmbDepType.Text = ""
  End If
  
  If QPTrim$(FASetup.DeprType) = "Prorate 1st year" Then
    fptxtFirstYearPct.Enabled = False
  ElseIf QPTrim$(FASetup.DeprType) = "Whole year" Then
    fptxtFirstYearPct.Enabled = False
  End If
  
End Sub

Private Sub fpcmbDepType_Change()
  'this routine disables the percentage field if it
  'is not needed and it sets the type field default
  'as Prorate 1st year
  
  If QPTrim$(fpcmbDepType.Text) = "Prorate 1st year" Then
    fptxtFirstYearPct.Enabled = False
  ElseIf QPTrim$(fpcmbDepType.Text) = "Fixed 1st year percentage" Then
    fptxtFirstYearPct.Enabled = True
  ElseIf QPTrim$(fpcmbDepType.Text) = "Whole year" Then
    fptxtFirstYearPct.Enabled = False
  ElseIf QPTrim$(fpcmbDepType.Text) = "" Then
    fpcmbDepType.Text = "Prorate 1st year"
    fptxtFirstYearPct.Enabled = False
  End If

End Sub

Private Sub fpcmbDepType_KeyDown(KeyCode As Integer, Shift As Integer)
  'this prevents the user from inadvertently changing the data
  'in this combo box while tabbing through the form
  If KeyCode = vbKeySpace Then
    fpcmbDepType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDepType.ListIndex = -1
  End If
  If fpcmbDepType.ListDown <> True Then
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
Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fptxtFirstYearPct_Change()
  '100 is the default for this field
  If QPTrim$(fptxtFirstYearPct.Text) = "" Then
    fptxtFirstYearPct.Text = "100.00"
  End If
End Sub

Private Sub fptxtFirstYearPct_LostFocus()
  Dim Number As Double
  'formats the field
  Number = CDbl(fptxtFirstYearPct.Text)
  fptxtFirstYearPct.Text = Using$("##0.00", Number)
End Sub

Private Sub LogSaves()
  Dim FASetup As FASetupRecType
  Dim SetupHandle As Integer
  
  On Error Resume Next
  'records all saves into the mainlog
  OpenFASetUpFile SetupHandle
  Get SetupHandle, 1, FASetup
  Close SetupHandle
   
  If QPTrim$(TempTownName$) <> QPTrim$(FASetup.TownName) Then
    MainLog ("The fixed asset town name has been changed from " + QPTrim$(TempTownName$) + " and saved as " + QPTrim$(FASetup.TownName) + ".")
  End If
  
  If QPTrim$(TempDeprType$) <> QPTrim$(FASetup.DeprType) Then
    MainLog ("The fixed asset depreciation type has been changed from " + QPTrim$(TempDeprType$) + " and saved as " + QPTrim$(FASetup.DeprType) + ".")
  End If
  
  If TempPct1St <> FASetup.Pct1St Then
    MainLog ("The fixed asset first year percentage has been changed from " + CStr(TempPct1St) + " and saved as " + CStr(FASetup.Pct1St) + ".")
  End If

End Sub
