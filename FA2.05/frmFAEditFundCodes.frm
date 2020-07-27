VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmFAEditFundCodes 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAEditFundCodes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EditLib.fpText fptxtDesc 
      Height          =   396
      Left            =   4368
      TabIndex        =   1
      ToolTipText     =   "Enter the description of the new fund code."
      Top             =   4824
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
      MaxLength       =   50
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
   Begin EditLib.fpText fptxtFundNum 
      Height          =   396
      Left            =   5304
      TabIndex        =   0
      ToolTipText     =   "Enter the fund number of the new fund."
      Top             =   3936
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
      CharValidationText=   "0 1 2 3 4 5 6 7 8 9 "
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
   Begin fpBtnAtlLibCtl.fpBtn cmdFundList 
      Height          =   405
      Left            =   6930
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to bring up a list of all current fund numbers."
      Top             =   3936
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
      ButtonDesigner  =   "frmFAEditFundCodes.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   675
      Left            =   3690
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to bring up a list of all current fund numbers."
      Top             =   5790
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
      ButtonDesigner  =   "frmFAEditFundCodes.frx":0AAA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   690
      Left            =   6228
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to save the data entered above."
      Top             =   5784
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
      ButtonDesigner  =   "frmFAEditFundCodes.frx":0C86
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fund Code Number:"
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
      Left            =   2892
      TabIndex        =   4
      Top             =   4032
      Width           =   2220
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
      TabIndex        =   3
      Top             =   4920
      Width           =   1452
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
      Left            =   2820
      TabIndex        =   2
      Top             =   1632
      Width           =   6012
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   1473
      Width           =   8652
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3756
      Left            =   1620
      Top             =   3168
      Width           =   8412
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   1425
      Width           =   8652
   End
End
Attribute VB_Name = "frmFAEditFundCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim Go2Flag As Boolean
  Dim TempFundNum As Integer
  Dim TempFundDesc$
  Dim FirstTime As Boolean

Private Sub cmdExit_Click()
  Dim FHandle As Integer
  Dim FundRec As FAFundCodeType
  Dim NumOfRecs As Integer
  Dim ChangeFlag As Boolean
  Dim DoWhatFlag As SaveChangeOptions1
  
  On Error GoTo ERRORSTUFF
  OpenFAFundCodeFile FHandle
  NumOfRecs = LOF(FHandle) \ Len(FundRec)
  
  If NumOfRecs = 0 Then 'no asset records saved
    Close FHandle
    frmFAFundCodeMenu.Show
    DoEvents
    KillFile ("editfundopen.dat")
    Unload frmFAEditFundCodes
    Exit Sub
  End If
  
  If GFundNum = 0 Then 'user opened screen but did not want to
  'save any entries made
    Close
    frmFAFundCodeMenu.Show
    DoEvents
    KillFile ("editfundopen.dat")
    Unload frmFAEditFundCodes
    Exit Sub
  End If
  
  Get FHandle, GFundNum, FundRec
  
  Close FHandle
  'now check to see if any changes have been made that will
  'be deleted upon exit
  
  If QPTrim$(fptxtDesc.Text) <> QPTrim$(FundRec.FundDesc) Then
    ChangeFlag = True
    fptxtDesc.SetFocus
    GoTo ChangeFound
  End If
  
  If Val(fptxtFundNum.Text) <> FundRec.FundNum Then
    ChangeFlag = True
    fptxtFundNum.SetFocus
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
      frmFAFundCodeMenu.Show
      DoEvents
      KillFile ("editfundopen.dat")
      Unload frmFAEditFundCodes
      Exit Sub
    Case Else:
    'Do nothing because we don't know about any options except
    'save, review or abandon...used as a placeholder for adding
    'other options at a later date
    End Select
  End If
  
  frmFAFundCodeMenu.Show
  Close
  DoEvents
  GFundNum = 0
  KillFile ("editfundopen.dat")
  Unload frmFAEditFundCodes
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditFundCodes", "cmdExit_Click", Erl)
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
  Dim FHandle As Integer
  Dim FundRec As FAFundCodeType
  Dim x As Integer
  Dim NumOfRecs As Integer
  Dim CompareThis$
  
  On Error GoTo ERRORSTUFF
  'Check4Dups is commented in frmFAEditDeptCodes
  Check4Dups = False
  
  OpenFAFundCodeFile FHandle
  NumOfRecs = LOF(FHandle) \ Len(FundRec)
  If NumOfRecs = 0 Then
    Close FHandle
    Exit Function
  End If
  
  CompareThis = QPTrim$(fptxtDesc.Text)
  If GFundNum = 0 Or Go2Flag = True Then
    For x = 1 To NumOfRecs
      Get FHandle, x, FundRec
      If CompareThis = QPTrim$(FundRec.FundDesc) Then
        Check4Dups = True
        frmFAEditDFACMess.Label1.Top = 900
        frmFAEditDFACMess.Label1.Caption = "The description entered is already being used for another fund. Unique fund descriptions are important in accurately tracking fixed assets. Please revise the fund description entered. Press F10 to bring up a complete list of all fund numbers. Press ESC to return to the screen."
        frmFAEditDFACMess.cmdCont.Text = "F10 &Open Fund List"
        frmFAEditDFACMess.cmdExit.Text = "ESC &Return to Screen"
        If frmFAEditDFACMess.fptxtChoice.Text = "continue" Then
          Unload frmFAEditDFACMess
          Call cmdFundList_Click
        Else
          Unload frmFAEditDFACMess
          fptxtDesc.SetFocus
        End If
        Exit For
      End If
    Next x
  Else
    For x = 1 To NumOfRecs
      If x <> GFundNum Then
        Get FHandle, x, FundRec
        If CompareThis = QPTrim$(FundRec.FundDesc) Then
          'D = Department   F = Funds      AC = Asset Code
          Check4Dups = True
          frmFAEditDFACMess.Label1.Top = 900
          frmFAEditDFACMess.Label1.Caption = "The description entered is already being used for another fund. Unique fund descriptions are important in accurately tracking fixed assets. Please revise the fund description entered. Press F10 to bring up a complete list of all fund numbers. Press ESC to return to the screen."
          frmFAEditDFACMess.cmdCont.Text = "F10 &Open Fund List"
          frmFAEditDFACMess.cmdExit.Text = "ESC &Return to Screen"
          frmFAEditDFACMess.Show vbModal
          If frmFAEditDFACMess.fptxtChoice.Text = "continue" Then
            Unload frmFAEditDFACMess
            Call cmdFundList_Click
          Else
            Unload frmFAEditDFACMess
            fptxtDesc.SetFocus
          End If
          Exit For
        End If
      End If
    Next x
  End If
  
  CompareThis = QPTrim$(fptxtFundNum.Text)
  If GFundNum = 0 Or Go2Flag = True Then
    For x = 1 To NumOfRecs
      Get FHandle, x, FundRec
      If CompareThis = FundRec.FundNum Then
        Check4Dups = True
        frmFAEditDFACMess.Label1.Top = 900
        frmFAEditDFACMess.Label1.Caption = "The fund number entered is already being used for another fund. Unique fund numbers are important in accurately tracking fixed assets. Please revise the fund number entered. Press F10 to bring up a complete list of all fund numbers. Press ESC to return to the screen."
        frmFAEditDFACMess.cmdCont.Text = "F10 &Open Fund List"
        frmFAEditDFACMess.cmdExit.Text = "ESC &Return to Screen"
        frmFAEditDFACMess.Show vbModal
        If frmFAEditDFACMess.fptxtChoice.Text = "continue" Then
          Unload frmFAEditDFACMess
          Call cmdFundList_Click
          Exit For
        Else
          Unload frmFAEditDFACMess
          fptxtFundNum.SetFocus
          Exit For
        End If
      End If
    Next x
  Else
    For x = 1 To NumOfRecs
      If x <> GFundNum Then
        Get FHandle, x, FundRec
        If CompareThis = FundRec.FundNum Then
          Check4Dups = True
          frmFAEditDFACMess.Label1.Top = 900
          frmFAEditDFACMess.Label1.Caption = "The fund number entered is already being used for another fund. Unique fund numbers are important in accurately tracking fixed assets. Please revise the fund number entered. Press F10 to bring up a complete list of all fund numbers. Press ESC to return to the screen."
          frmFAEditDFACMess.cmdCont.Text = "F10 &Open Fund List"
          frmFAEditDFACMess.cmdExit.Text = "ESC &Return to Screen"
          frmFAEditDFACMess.Show vbModal
          If frmFAEditDFACMess.fptxtChoice.Text = "continue" Then
            Unload frmFAEditDFACMess
            Call cmdFundList_Click
            Exit For
          Else
            Unload frmFAEditDFACMess
            fptxtFundNum.SetFocus
            Exit For
          End If
        End If
      End If
    Next x
  End If
  Close FHandle
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditFundCodes", "Check4Dups", Erl)
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

Private Sub cmdFundList_Click()
  frmFAFundList.Show vbModal
End Sub

Private Sub cmdSave_Click()
  Dim FHandle As Integer
  Dim FundRec As FAFundCodeType
  Dim NumOfRecs As Integer
  Dim DoWhatFlag As WarnOption
  Dim ChangeFlag As Boolean
  Dim NumOfFunds As Integer
  Dim NewFlag As Boolean
  
  On Error GoTo ERRORSTUFF
  'this save routine works exactly like the save routine
  'on frmFAEditDeptCodes and is commented there
  NewFlag = False
  Go2Flag = False
  If QPTrim$(fptxtDesc.Text) = "" Then
    MsgBox "Please enter a description for Fund Code"
    fptxtDesc.SetFocus
    Close
    Exit Sub
  End If
  If QPTrim$(fptxtFundNum.Text) = "" Then
    MsgBox "Please enter a number for Fund Code"
    fptxtFundNum.SetFocus
    Close
    Exit Sub
  End If
  
  ChangeFlag = False
  
  If Check4Dups = True Then Exit Sub
  
  OpenFAFundCodeFile FHandle
  NumOfFunds = LOF(FHandle) / Len(FundRec)
  
  If GFundNum > 0 Then
    Get FHandle, GFundNum, FundRec
    If QPTrim$(fptxtDesc.Text) <> QPTrim$(FundRec.FundDesc) Then
      ChangeFlag = True
      fptxtDesc.SetFocus
    ElseIf QPTrim$(fptxtFundNum.Text) <> FundRec.FundNum Then
      ChangeFlag = True
      fptxtFundNum.SetFocus
    End If
    If ChangeFlag = True Then
      DoWhatFlag = PromptWarnOverWrite(Me)
      Select Case DoWhatFlag
        Case WarnOption.wSave
          MainLog ("Overwrite warning issued for " + QPTrim$(fptxtFundNum.Text) + ". Save option selected in frmFAEditFundCodes.")
        Case WarnOption.wExit
          MainLog ("Overwrite warning issued for " + QPTrim$(fptxtFundNum.Text) + ". Exit option selected in frmFAEditFundCodes.")
          Close FHandle
          frmFADeptCodeMenu.Show
          DoEvents
          KillFile ("editfundopen.dat")
          Unload frmFAEditFundCodes
          Exit Sub
        Case WarnOption.wReturn
          MainLog ("Overwrite warning issued for " + QPTrim$(fptxtFundNum.Text) + ". Return option selected in frmFAEditFundCodes.")
          Close FHandle
          Exit Sub
        Case WarnOption.wGo2Add
          MainLog ("Overwrite warning issued for " + QPTrim$(fptxtFundNum.Text) + ". Add to list option selected in frmFAEditFundCodes.")
          Go2Flag = True
          If Check4Dups = True Then
            Go2Flag = False
            Exit Sub
          End If
          FundRec.FundNum = Val(fptxtFundNum.Text)
          FundRec.FundDesc = QPTrim$(fptxtDesc.Text)
          Put FHandle, NumOfFunds + 1, FundRec
          Close FHandle
          GoTo Go2
        Case Else
          Close FHandle
          MsgBox "Please make a valid selection"
          Exit Sub
      End Select
    End If
  End If
  NumOfRecs = LOF(FHandle) \ Len(FundRec)
  If GFundNum = 0 Then
    NewFlag = True
    FundRec.FundNum = Val(fptxtFundNum.Text)
    FundRec.FundDesc = QPTrim$(fptxtDesc.Text)
    Put FHandle, NumOfRecs + 1, FundRec
    Close FHandle
  Else
    FundRec.FundNum = Val(fptxtFundNum.Text)
    FundRec.FundDesc = QPTrim$(fptxtDesc.Text)
    Put FHandle, GFundNum, FundRec
    Close FHandle
  End If
  
Go2:
  Call CreateFundIdx
  
  MsgBox "Your information has been saved."
  
  If NewFlag = True Then
    MainLog ("Fund code number " + QPTrim$(fptxtFundNum.Text) + " was saved in frmFAEditFundCodes.")
  Else
    Call LogSaves
  End If
  Close
  
  If NewFlag = True Then
    If MsgBox("Do you want to add another new fund code?", vbYesNo) = vbYes Then
      GFundNum = 0
      Call LoadMe
      fptxtFundNum.SetFocus
      Exit Sub
    End If
  End If
  
  frmFAFundCodeMenu.Show
  DoEvents
  GFundNum = 0
  KillFile ("editfundopen.dat")
  Unload frmFAEditFundCodes
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditFundCodes", "cmdSave_Click", Erl)
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
'    'Me.Visible = False
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
      Call cmdFundList_Click
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
      KillFile ("editfundopen.dat")
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAEditFundCodes.")
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
  Dim FHandle As Integer
  Dim FundRec As FAFundCodeType
  Dim One As Integer
  Dim FileHandle As Integer
  
  One = 1
  FileHandle = FreeFile
  Open "editfundopen.dat" For Output As FileHandle Len = 2
  
  Print #FileHandle, One
  Close FileHandle
  
  OpenFAFundCodeFile FHandle
  
  If GFundNum = 0 Then
    Me.Caption = "Adding Fund Code"
    Me.Label2 = "Adding Fixed Asset Fund Code"
    fptxtFundNum.Text = ""
    fptxtDesc.Text = ""
  Else
    Me.Caption = "Editing Fund Code"
    Me.Label2 = "Editing Fixed Asset Fund Code"
    Get FHandle, GFundNum, FundRec
    fptxtFundNum.Text = FundRec.FundNum
    TempFundNum = FundRec.FundNum 'global
    fptxtDesc.Text = QPTrim$(FundRec.FundDesc)
    TempFundDesc = QPTrim$(FundRec.FundDesc) 'global
    Close FHandle
  End If
  
End Sub

Private Sub LogSaves()
  Dim FAFundRec As FAFundCodeType
  Dim FHandle As Integer
  
  OpenFAFundCodeFile FHandle
  Get FHandle, GFundNum, FAFundRec
  Close FHandle
  
  If TempFundNum <> FAFundRec.FundNum Then
    MainLog ("Fund Number " + CStr(TempFundNum) + " changed and saved as " + CStr(FAFundRec.FundNum) + " in frmFAEditFundCodes.")
  End If
  
  If QPTrim$(TempFundDesc) <> QPTrim$(FAFundRec.FundDesc) Then
    MainLog ("For Fund Number: " + CStr(FAFundRec.FundNum) + " fund description changed from " + QPTrim$(TempFundDesc) + " to " + QPTrim$(FAFundRec.FundDesc) + " in frmFAEditFundCodes.")
  End If
  
End Sub

Private Sub fptxtFundNum_LostFocus()
'  Dim x As Integer
'  Dim FHandle As Integer
'  Dim FundRec As FAFundCodeType
'  Dim NumOfFunds As Integer
'  Dim Found As Boolean
'  Dim Number As Integer
'
'  If QPTrim$(fptxtFundNum.Text) = "" Then Exit Sub
'  Number = Val(fptxtFundNum.Text)
'  OpenFAFundCodeFile FHandle
'  NumOfFunds = LOF(FHandle) / Len(FundRec)
'
'  If NumOfFunds = 0 Then Exit Sub
'
'  If GFundNum = 0 Then 'start with blank screen
'  'and enter a tag number...if the tag number entered
'  'is already in use then pop screen with its data
'  '...if not this number is a new one
'    For x = 1 To NumOfFunds
'      Get FHandle, x, FundRec
'      If Number = FundRec.FundNum Then
'        GoTo EditIt
'      End If
'    Next x
'    If x = NumOfFunds + 1 Then
'      Close
'      Exit Sub
'    End If
'  End If
'
'EditIt:
'
'  For x = 1 To NumOfFunds
'    Get FHandle, x, FundRec
'    If Number = FundRec.FundNum Then 'match the selected
'    'row with the right code
'      Found = True
'      GFundNum = x 'now you can assign the correct global
'      Exit For
'    Else
'      Found = False
'      GoTo NotAMatch
'    End If
'
'NotAMatch:
'  Next x
'  Close FHandle
'
'  If Found = False Then
'    If MsgBox("The fund code entered does not match any of those saved. Would you like to see the fund code list?", vbYesNo) = vbYes Then
'      Call cmdFundList_Click
'    Else
'      fptxtFundNum.SetFocus
'    End If
'  Else
'    Call LoadMe
'    If FirstTime = True Then
'      FirstTime = False
'    Else
'      fptxtDesc.SetFocus
'    End If
'  End If

End Sub
