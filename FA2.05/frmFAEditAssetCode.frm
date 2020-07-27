VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFAEditAssetCode 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAEditAssetCode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbStatus 
      Height          =   405
      Left            =   5085
      TabIndex        =   1
      ToolTipText     =   "Enter the active status of this asset code."
      Top             =   4230
      Width           =   2130
      _Version        =   196608
      _ExtentX        =   3757
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
      MaxEditLen      =   10
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
      ColDesigner     =   "frmFAEditAssetCode.frx":08CA
   End
   Begin EditLib.fpText fptxtDesc 
      Height          =   396
      Left            =   4368
      TabIndex        =   2
      ToolTipText     =   "Enter the description of this new asset code."
      Top             =   5094
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
      Left            =   4560
      TabIndex        =   0
      ToolTipText     =   "Enter the group code number of this asset code."
      Top             =   3366
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
   Begin fpBtnAtlLibCtl.fpBtn cmdAssetList 
      Height          =   405
      Left            =   6765
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to bring up a list of all asset codes."
      Top             =   3360
      Width           =   1845
      _Version        =   131072
      _ExtentX        =   3254
      _ExtentY        =   714
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
      ButtonDesigner  =   "frmFAEditAssetCode.frx":0BC1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   690
      Left            =   3609
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to bring up a list of all asset codes."
      Top             =   6918
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
      ButtonDesigner  =   "frmFAEditAssetCode.frx":0DA1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   690
      Left            =   6159
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to save the data entered above."
      Top             =   6918
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
      ButtonDesigner  =   "frmFAEditAssetCode.frx":0F7D
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
      Left            =   2592
      TabIndex        =   6
      Top             =   3462
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
      TabIndex        =   5
      Top             =   5190
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
      Left            =   3984
      TabIndex        =   4
      Top             =   4326
      Width           =   924
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2940
      TabIndex        =   3
      Top             =   1458
      Width           =   6012
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   1314
      Width           =   8652
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3468
      Left            =   1680
      Top             =   2706
      Width           =   8412
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   1266
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
  Dim Go2Flag As Boolean
  Dim TempASSETCode$
  Dim TempAssetStatus$
  Dim TempAssetDesc$
  Dim FirstTime As Boolean
  
Private Sub cmdAssetList_Click()
  frmFAAssetCodeList.Show vbModal
End Sub

Private Sub cmdExit_Click()
  Dim CodeHandle As Integer
  Dim CodeRec As FAAssetCodeRecType
  Dim NumOfRecs As Integer
  Dim ChangeFlag As Boolean
  Dim DoWhatFlag As SaveChangeOptions1
  
  On Error GoTo ERRORSTUFF
  OpenFACodeNameFile CodeHandle
  NumOfRecs = LOF(CodeHandle) \ Len(CodeRec)
  If NumOfRecs = 0 Then
    Close
    frmFAAssetCodesMenu.Show
    DoEvents
    KillFile ("editassetopen.dat") 'this file only good when this form
    'is operative
    Unload frmFAEditAssetCode
    Exit Sub
  End If
  
  If GCodeNum = 0 Then 'user probably arrived here by mistake...
  'anyway there is no business to be done so go back to menu
    Close
    frmFAAssetCodesMenu.Show
    DoEvents
    KillFile ("editassetopen.dat")
    Unload frmFAEditAssetCode
    Exit Sub
  End If
  
  Get CodeHandle, GCodeNum, CodeRec 'begin checking for unsaved changes
  Close CodeHandle
  
  If QPTrim$(CodeRec.AssetDesc) <> QPTrim$(fptxtDesc.Text) Then
    ChangeFlag = True
    fptxtDesc.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(CodeRec.AssetStatus) <> QPTrim$(fpcmbStatus.Text) Then
    ChangeFlag = True
    fpcmbStatus.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(CodeRec.ASSETCODE) <> QPTrim$(fptxtGroupCode.Text) Then
    ChangeFlag = True
    fptxtGroupCode.SetFocus
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
      frmFAAssetCodesMenu.Show
      DoEvents
      KillFile ("editassetopen.dat")
      Unload frmFAEditAssetCode
      Exit Sub
    Case Else:
    'Do nothing because we don't know about any options except
    'save, review or abandon...used as a placeholder for adding
    'other options at a later date
    End Select
  End If
  
  frmFAAssetCodesMenu.Show
  Close
  DoEvents
  GCodeNum = 0
  KillFile ("editassetopen.dat")
  Unload frmFAEditAssetCode
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditAssetCode", "cmdExit_Click", Erl)
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
    Unload Me
  
End Sub
Private Function Check4Dups() As Boolean
  Dim CodeHandle As Integer
  Dim CodeRec As FAAssetCodeRecType
  Dim x As Integer
  Dim NumOfRecs As Integer
  Dim CompareThis$
  
  On Error GoTo ERRORSTUFF
  Check4Dups = False
  OpenFACodeNameFile CodeHandle
  NumOfRecs = LOF(CodeHandle) \ Len(CodeRec)
  
  CompareThis = QPTrim$(fptxtDesc.Text)
  If GCodeNum = 0 Or Go2Flag = True Then 'Go2Flag is set if the user
  'was warned after he opted to save this data that he was overwriting
  'existing data and the user opted to add this data as a new record
    For x = 1 To NumOfRecs
      Get CodeHandle, x, CodeRec
      If CompareThis = QPTrim$(CodeRec.AssetDesc) Then
        Check4Dups = True
        'can't save a new record with an existing description
'        MsgBox "You have entered a description that is already in use. Please choose another description."
        frmFAEditDFACMess.Label1.Caption = "The asset code description entered is already being used for another asset code. Unique asset code descriptions are important in the proper operation of this program. Press F10 to bring up a list of all asset codes saved. Otherwise press ESC to return to the screen."
        frmFAEditDFACMess.Label1.Top = 900
        frmFAEditDFACMess.cmdCont.Text = "F10 &Open Asset Code List"
        frmFAEditDFACMess.cmdExit.Text = "ESC &Return to Screen"
        frmFAEditDFACMess.Show vbModal
        If frmFAEditDFACMess.fptxtChoice.Text = "continue" Then
          Unload frmFAEditDFACMess
          Call cmdAssetList_Click
          Exit For
        Else
          Unload frmFAEditDFACMess
          fptxtDesc.SetFocus
          Exit For
        End If
      End If
    Next x
  Else
    For x = 1 To NumOfRecs
      If x <> GCodeNum Then
        Get CodeHandle, x, CodeRec
        If CompareThis = QPTrim$(CodeRec.AssetDesc) Then
          Check4Dups = True
        'can't save an existing record with an existing description that isn't it's own
          frmFAEditDFACMess.Label1.Caption = "The asset code description entered is already being used for another asset code. Unique asset code descriptions are important in the proper operation of this program. Press F10 to bring up a list of all asset codes saved. Otherwise press ESC to return to the screen."
          frmFAEditDFACMess.Label1.Top = 900
          frmFAEditDFACMess.cmdCont.Text = "F10 &Open Asset Code List"
          frmFAEditDFACMess.cmdExit.Text = "ESC &Return to Screen"
          frmFAEditDFACMess.Show vbModal
          If frmFAEditDFACMess.fptxtChoice.Text = "continue" Then
            Unload frmFAEditDFACMess
            Call cmdAssetList_Click
            Exit For
          Else
            Unload frmFAEditDFACMess
            fptxtDesc.SetFocus
            Exit For
          End If
        End If
      End If
    Next x
  End If
  
  CompareThis = QPTrim$(fptxtGroupCode.Text)
  If GCodeNum = 0 Or Go2Flag = True Then
    For x = 1 To NumOfRecs
      Get CodeHandle, x, CodeRec
      If CompareThis = QPTrim$(CodeRec.ASSETCODE) Then
        'can't save a new record with an existing code number
        Check4Dups = True
        frmFAEditDFACMess.Label1.Caption = "The asset code number entered is already being used for another asset code. Unique asset code numbers are important in the proper operation of this program. Press F10 to bring up a list of all asset codes saved. Otherwise press ESC to return to the screen."
        frmFAEditDFACMess.Label1.Top = 900
        frmFAEditDFACMess.cmdCont.Text = "F10 &Open Asset Code List"
        frmFAEditDFACMess.cmdExit.Text = "ESC &Return to Screen"
        frmFAEditDFACMess.Show vbModal
        If frmFAEditDFACMess.fptxtChoice.Text = "continue" Then
          Unload frmFAEditDFACMess
          Call cmdAssetList_Click
          Exit For
        Else
          Unload frmFAEditDFACMess
          fptxtGroupCode.SetFocus
          Exit For
        End If
      End If
    Next x
  Else
    For x = 1 To NumOfRecs
      If x <> GCodeNum Then
        Get CodeHandle, x, CodeRec
        If CompareThis = QPTrim$(CodeRec.ASSETCODE) Then
        'can't save an existing record with an existing code number that isn't it's own
          Check4Dups = True
          Check4Dups = True
          frmFAEditDFACMess.Label1.Caption = "The asset code number entered is already being used for another asset code. Unique asset code numbers are important in the proper operation of this program. Press F10 to bring up a list of all asset codes saved. Otherwise press ESC to return to the screen."
          frmFAEditDFACMess.Label1.Top = 900
          frmFAEditDFACMess.cmdCont.Text = "F10 &Open Asset Code List"
          frmFAEditDFACMess.cmdExit.Text = "ESC &Return to Screen"
          frmFAEditDFACMess.Show vbModal
          If frmFAEditDFACMess.fptxtChoice.Text = "continue" Then
            Unload frmFAEditDFACMess
            Call cmdAssetList_Click
            Exit For
          Else
            Unload frmFAEditDFACMess
            fptxtGroupCode.SetFocus
            Exit For
          End If
        End If
      End If
    Next x
  End If
  Close CodeHandle
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditAssetCode", "Check4Dups", Erl)
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
    Unload Me
  
End Function
Private Sub cmdSave_Click()
  Dim CodeHandle As Integer
  Dim CodeRec As FAAssetCodeRecType
  Dim NumOfRecs As Integer
  Dim DoWhatFlag As WarnOption
  Dim ChangeFlag As Boolean
  Dim NumOfCodes As Integer
  Dim NewFlag As Boolean
  
  On Error GoTo ERRORSTUFF
  NewFlag = False
  Go2Flag = False
  If QPTrim$(fptxtDesc.Text) = "" Then
    MsgBox "Please enter a description for this asset code"
    Exit Sub
  End If
  ChangeFlag = False
  
'  If Len(QPTrim$(fptxtGroupCode.Text)) <> 4 Then 'commented 7/17/09
'    MsgBox "Please enter a 4 digit number for the code."
'    fptxtGroupCode.SetFocus
'    Close CodeHandle
'    Exit Sub
'  End If
  
  If Len(QPTrim$(fptxtGroupCode.Text)) > 4 Then 'added 7/17/09
    MsgBox "Please enter up to a 4 digit number for the code."
    fptxtGroupCode.SetFocus
    Close CodeHandle
    Exit Sub
  End If
  
  
  If Check4Dups = True Then Exit Sub
  
  OpenFACodeNameFile CodeHandle
  NumOfCodes = LOF(CodeHandle) / Len(CodeRec)
  If GCodeNum > 0 Then 'not a new entry
    Get CodeHandle, GCodeNum, CodeRec
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
    If ChangeFlag = True Then 'existing data is being altered...this could have
    'a negative affect on existing data files when reporting takes place
      DoWhatFlag = PromptWarnOverWrite(Me) 'user warned that existing data is
      'being overwritten and forced to make a choice
      Select Case DoWhatFlag
        Case WarnOption.wSave
          MainLog ("Overwrite warning issued for " + fptxtGroupCode.Text + ". Save option selected in frmFAEditAssetCode.")
        Case WarnOption.wExit
          MainLog ("Overwrite warning issued for " + fptxtGroupCode.Text + ". Exit option selected in frmFAEditAssetCode.")
          Close CodeHandle
          frmFAAssetCodesMenu.Show
          DoEvents
          KillFile ("editassetopen.dat")
          Unload frmFAEditAssetCode
          Exit Sub
        Case WarnOption.wReturn
          MainLog ("Overwrite warning issued for " + fptxtGroupCode.Text + ". Return option selected in frmFAEditAssetCode.")
          Close CodeHandle
          Exit Sub
        Case WarnOption.wGo2Add
          MainLog ("Overwrite warning issued for " + fptxtGroupCode.Text + ". Add to list option selected in frmFAEditAssetCode.")
          Go2Flag = True
          If Check4Dups = True Then 'data entered has already been entered for
          'a different asset and warning was issued in Check4Dups
            Go2Flag = False
            Exit Sub
          End If
          'add this data to the records as a new entry
          Go2Flag = False
          CodeRec.ASSETCODE = QPTrim$(fptxtGroupCode.Text)
          CodeRec.AssetDesc = QPTrim$(fptxtDesc.Text)
          CodeRec.AssetStatus = QPTrim$(fpcmbStatus.Text)
          Put CodeHandle, NumOfCodes + 1, CodeRec
          Close CodeHandle
          GoTo Go2
        Case Else
          Close CodeHandle
          MsgBox "Please make a valid selection"
          Exit Sub
      End Select
    End If
  End If
  NumOfRecs = LOF(CodeHandle) \ Len(CodeRec)
  'not being overwritten or user opted to save after overwrite order issued
  If GCodeNum = 0 Then
    NewFlag = True
    CodeRec.ASSETCODE = QPTrim$(fptxtGroupCode.Text)
    CodeRec.AssetDesc = QPTrim$(fptxtDesc.Text)
    CodeRec.AssetStatus = QPTrim$(fpcmbStatus.Text)
    Put CodeHandle, NumOfRecs + 1, CodeRec 'save as next record
    Close CodeHandle
  Else
    CodeRec.ASSETCODE = QPTrim$(fptxtGroupCode.Text)
    CodeRec.AssetDesc = QPTrim$(fptxtDesc.Text)
    CodeRec.AssetStatus = QPTrim$(fpcmbStatus.Text)
    Put CodeHandle, GCodeNum, CodeRec 'overwrite existing data
    Close CodeHandle
  End If
  
Go2:
  Close
  MsgBox "Your information has been saved"
  Call CreateAssetIdx
  If NewFlag = True Then 'tell user the save procedure was successful
    MainLog ("Asset code number " + QPTrim$(fptxtGroupCode.Text) + "was saved in frmFAEditAssetCode.")
  Else
    Call LogSaves
  End If
  
  If NewFlag = True Then 'keeps adding assets procedure more fluid
    If MsgBox("Do you want to add another new asset code?", vbYesNo) = vbYes Then
      GCodeNum = 0
      Call LoadMe
      fptxtGroupCode.SetFocus
      Exit Sub
    End If
  End If
  
  GCodeNum = 0
  
  frmFAAssetCodesMenu.Show
  DoEvents
  KillFile ("editassetopen.dat")
  Unload frmFAEditAssetCode
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditAssetCode", "cmdSave_Click", Erl)
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
    Unload Me
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  FirstTime = True
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
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%L"
      Call cmdAssetList_Click
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
      KillFile ("editassetopen.dat")
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAEditAssetCode.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Public Sub LoadMe()
  Dim CodeHandle As Integer
  Dim CodeRec As FAAssetCodeRecType
  Dim One As Integer
  Dim FileHandle As Integer
  
  One = 1
  FileHandle = FreeFile
  Open "editassetopen.dat" For Output As FileHandle Len = 2
  'editassetopen.dat is used to let other programs know that this
  'form is opened...these .dat files only exist when the current
  'form is operative
  Print #FileHandle, One
  Close FileHandle
  
  'next if statement assigns labels to this form depending on
  'what is being done
  If GCodeNum = 0 Then
    Me.Caption = "Adding Asset Code"
    Me.Label2 = "Adding Fixed Asset Code"
    fpcmbStatus.Text = "Active"
    fptxtGroupCode.Text = ""
    fptxtDesc.Text = ""
  Else
    Me.Caption = "Editing Asset Code"
    Me.Label2 = "Editing Fixed Asset Code"
    OpenFACodeNameFile CodeHandle
    Get CodeHandle, GCodeNum, CodeRec
    fpcmbStatus.Text = QPTrim$(CodeRec.AssetStatus)
    TempAssetStatus$ = QPTrim$(CodeRec.AssetStatus) 'global
    fptxtGroupCode.Text = QPTrim$(CodeRec.ASSETCODE)
    TempASSETCode$ = QPTrim$(CodeRec.ASSETCODE) 'global
    fptxtDesc.Text = QPTrim$(CodeRec.AssetDesc)
    TempAssetDesc$ = QPTrim$(CodeRec.AssetDesc) 'global
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

Private Sub LogSaves()
  Dim CodeHandle As Integer
  Dim CodeRec As FAAssetCodeRecType
  
  On Error GoTo ERRORSTUFF
  'anytime data is saved it gets recorded in the FALog
  OpenFACodeNameFile CodeHandle
  Get CodeHandle, GCodeNum, CodeRec
  Close CodeHandle
  
  If QPTrim$(TempASSETCode$) <> QPTrim$(CodeRec.ASSETCODE) Then
    MainLog ("Asset code number " + QPTrim$(TempASSETCode$) + " has been changed to " + QPTrim$(CodeRec.ASSETCODE) + " in frmFAEditAssetCode.")
  End If

  If QPTrim$(TempAssetStatus$) <> QPTrim$(CodeRec.AssetStatus) Then
    MainLog ("For asset code number " + QPTrim$(CodeRec.ASSETCODE) + ": status changed from " + QPTrim$(TempAssetStatus$) + " and saved to " + QPTrim$(CodeRec.AssetStatus) + " in frmFAEditAssetCode.")
  End If

  If QPTrim$(TempAssetDesc$) <> QPTrim$(CodeRec.AssetDesc) Then
    MainLog ("For asset code number " + QPTrim$(CodeRec.ASSETCODE) + ": description changed from " + TempAssetDesc$ + " and saved to " + QPTrim$(CodeRec.AssetDesc) + " in frmFAEditAssetCode.")
  End If
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditAssetCode", "LogSaves", Erl)
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
    Unload Me
End Sub

Private Sub fptxtGroupCode_LostFocus()
'  Dim x As Integer
'  Dim CodeHandle As Integer
'  Dim CodeRec As FAAssetCodeRecType
'  Dim NumOfCodes As Integer
'  Dim Found As Boolean
'  Dim Number$
'
'  If QPTrim$(fptxtGroupCode.Text) = "" Then Exit Sub
'  Number$ = QPTrim$(fptxtGroupCode.Text)
'  OpenFACodeNameFile CodeHandle
'  NumOfCodes = LOF(CodeHandle) / Len(CodeRec)
'
'  If NumOfCodes = 0 Then Exit Sub
'
'  If GCodeNum = 0 Then 'start with blank screen
'  'and enter a tag number...if the tag number entered
'  'is already in use then pop screen with its data
'  '...if not this number is a new one
'    For x = 1 To NumOfCodes
'      Get CodeHandle, x, CodeRec
'      If Number = QPTrim$(CodeRec.ASSETCODE) Then
'        GoTo EditIt
'      End If
'    Next x
'    If x = NumOfCodes + 1 Then
'      Close
'      Exit Sub
'    End If
'  End If
'
'EditIt:
'
'  For x = 1 To NumOfCodes
'    Get CodeHandle, x, CodeRec
'    If Number = QPTrim$(CodeRec.ASSETCODE) Then 'match the selected
'    'row with the right code
'      Found = True
'      GCodeNum = x 'now you can assign the correct global
'      Exit For
'    Else
'      Found = False
'      GoTo NotAMatch
'    End If
'
'NotAMatch:
'  Next x
'  Close CodeHandle
'
'  If Found = False Then
'    If MsgBox("The asset code entered does not match any of those saved. Would you like to see the asset code list?", vbYesNo) = vbYes Then
'      Call cmdAssetList_Click
'    Else
'      fptxtGroupCode.SetFocus
'    End If
'  Else
'    Call LoadMe
'    If FirstTime = True Then
'      FirstTime = False
'    Else
'      fpcmbStatus.SetFocus
'    End If
'  End If

End Sub
