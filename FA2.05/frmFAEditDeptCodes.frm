VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmFAEditDeptCodes 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAEditDeptCodes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EditLib.fpText fptxtDesc 
      Height          =   390
      Left            =   4410
      TabIndex        =   1
      ToolTipText     =   "Enter the description of the new department code."
      Top             =   4530
      Width           =   4425
      _Version        =   196608
      _ExtentX        =   7805
      _ExtentY        =   688
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
      MaxLength       =   25
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
   Begin EditLib.fpText fptxtDeptCode 
      Height          =   390
      Left            =   5355
      TabIndex        =   0
      ToolTipText     =   "Enter the new department number."
      Top             =   3630
      Width           =   2070
      _Version        =   196608
      _ExtentX        =   3651
      _ExtentY        =   688
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
   Begin fpBtnAtlLibCtl.fpBtn cmdDeptList 
      Height          =   405
      Left            =   7515
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to bring up a list of all departments."
      Top             =   3630
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
      ButtonDesigner  =   "frmFAEditDeptCodes.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   690
      Left            =   3690
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5550
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
      ButtonDesigner  =   "frmFAEditDeptCodes.frx":0AAA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   690
      Left            =   6225
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to save the data entered above."
      Top             =   5550
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
      ButtonDesigner  =   "frmFAEditDeptCodes.frx":0C86
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3750
      Left            =   1620
      Top             =   2970
      Width           =   8415
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   750
      Index           =   1
      Left            =   1500
      Top             =   1470
      Width           =   8655
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
      Height          =   495
      Left            =   2940
      TabIndex        =   4
      Top             =   1620
      Width           =   6015
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
      Left            =   2730
      TabIndex        =   3
      Top             =   4635
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Department Code Number:"
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
      Left            =   2220
      TabIndex        =   2
      Top             =   3720
      Width           =   2985
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   1425
      Width           =   8655
   End
End
Attribute VB_Name = "frmFAEditDeptCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim Go2Flag As Boolean
  Dim TempDeptDesc$
  Dim TempDeptNum As Integer
  Dim FirstFlag As Boolean

Private Sub cmdDeptList_Click()
  frmFADeptList.Show vbModal
End Sub

Private Sub cmdExit_Click()
  Dim DHandle As Integer
  Dim DeptRec As FADeptCodeType
  Dim NumOfRecs As Integer
  Dim ChangeFlag As Boolean
  Dim DoWhatFlag As SaveChangeOptions1
  
  On Error GoTo ERRORSTUFF
  OpenFADeptCodeFile DHandle
  NumOfRecs = LOF(DHandle) \ Len(DeptRec)
  
  If NumOfRecs = 0 Then 'no codes have been saved
    frmFADeptCodeMenu.Show
    DoEvents
    Close
    KillFile ("editdeptopen.dat")
    Unload frmFAEditDeptCodes
    Exit Sub
  End If
  
  If GDeptNum = 0 Then 'user opened this form and then exited
  'without doing anything of consequence
    frmFADeptCodeMenu.Show
    DoEvents
    KillFile ("editdeptopen.dat")
    Unload frmFAEditDeptCodes
    Close
    Exit Sub
  End If
  
  Get DHandle, GDeptNum, DeptRec
  Close DHandle
  'begin checking for unsaved changes
  If QPTrim$(fptxtDesc.Text) <> QPTrim$(DeptRec.DeptDesc) Then
    ChangeFlag = True
    fptxtDesc.SetFocus
    GoTo ChangeFound
  End If
  
  If Val(fptxtDeptCode.Text) <> DeptRec.DeptNum Then
    ChangeFlag = True
    fptxtDeptCode.SetFocus
  End If
  
ChangeFound: 'user warned that data has been changed and not saved
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
      frmFADeptCodeMenu.Show
      Close
      DoEvents
      KillFile ("editdeptopen.dat")
      Unload frmFAEditDeptCodes
      Exit Sub
    Case Else:
    'Do nothing because we don't know about any options except
    'save, review or abandon...used as a placeholder for adding
    'other options at a later date
    End Select
  End If
  
  frmFADeptCodeMenu.Show
  Close
  DoEvents
  GDeptNum = 0
  KillFile ("editdeptopen.dat")
  Unload frmFAEditDeptCodes
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditDeptCodes", "cmdExit_Click", Erl)
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
  Dim DHandle As Integer
  Dim DeptRec As FADeptCodeType
  Dim x As Integer
  Dim NumOfRecs As Integer
  Dim CompareThis$
  
  On Error GoTo ERRORSTUFF
  Check4Dups = False
  
  OpenFADeptCodeFile DHandle
  NumOfRecs = LOF(DHandle) \ Len(DeptRec)
  If NumOfRecs = 0 Then
    Close DHandle
    Exit Function
  End If
  
  CompareThis = QPTrim$(fptxtDesc.Text) 'look for existing descriptions
  'that match what has been entered
  If GDeptNum = 0 Or Go2Flag = True Then
    For x = 1 To NumOfRecs
      Get DHandle, x, DeptRec 'used for either a new addition or
      'an existing record that is being overwritten...with existing assets
      'this would not be relevent because the Go2Flag becomes True
      'only if a change was made to an existing record
      If CompareThis = QPTrim$(DeptRec.DeptDesc) Then
        Check4Dups = True
        frmFAEditDFACMess.Label1.Top = 900
        frmFAEditDFACMess.Label1.Caption = "The description entered is already being used for another department. Unique department descriptions are important in accurately tracking fixed assets. Please revise the department description entered. Press F10 to bring up a complete list of all department numbers. Press ESC to return to the screen."
        frmFAEditDFACMess.cmdCont.Text = "F10 &Open department List"
        frmFAEditDFACMess.cmdExit.Text = "ESC &Return to Screen"
        If frmFAEditDFACMess.fptxtChoice.Text = "continue" Then
          Unload frmFAEditDFACMess
          Call cmdDeptList_Click
          Exit For
        Else
          Unload frmFAEditDFACMess
          fptxtDesc.SetFocus
          Exit For
        End If
      End If
    Next x
  Else 'looks for duplicate data before it is known if the current
  'procedure is overwriting existing data
    For x = 1 To NumOfRecs
      If x <> GDeptNum Then
        Get DHandle, x, DeptRec
        If CompareThis = QPTrim$(DeptRec.DeptDesc) Then
          Check4Dups = True
          frmFAEditDFACMess.Label1.Top = 900
          frmFAEditDFACMess.Label1.Caption = "The description entered is already being used for another department. Unique department descriptions are important in accurately tracking fixed assets. Please revise the department description entered. Press F10 to bring up a complete list of all department numbers. Press ESC to return to the screen."
          frmFAEditDFACMess.cmdCont.Text = "F10 &Open Department List"
          frmFAEditDFACMess.cmdExit.Text = "ESC &Return to Screen"
          frmFAEditDFACMess.Show vbModal
          If frmFAEditDFACMess.fptxtChoice.Text = "continue" Then
            Unload frmFAEditDFACMess
            Call cmdDeptList_Click
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
  
  CompareThis = QPTrim$(fptxtDeptCode.Text)
  If GDeptNum = 0 Or Go2Flag = True Then
    For x = 1 To NumOfRecs
      Get DHandle, x, DeptRec
      If CompareThis = DeptRec.DeptNum Then
        Check4Dups = True
        frmFAEditDFACMess.Label1.Caption = "The department code number entered has already been assigned to another department. Unique department numbers are important to the proper operation of this program. To bring up a complete department number list press F10. Otherwise press ESC to return to the screen."
        frmFAEditDFACMess.Label1.Top = 900
        frmFAEditDFACMess.cmdCont.Text = "F10 &Open Dept List"
        frmFAEditDFACMess.cmdExit.Text = "ESC &Return and Edit"
        frmFAEditDFACMess.Show vbModal
        If frmFAEditDFACMess.fptxtChoice.Text = "continue" Then
          Unload frmFAEditDFACMess
          Call cmdDeptList_Click
          Exit For
        Else
          Unload frmFAEditDFACMess
          fptxtDeptCode.SetFocus
          Exit For
        End If
      End If
    Next x
  Else
    For x = 1 To NumOfRecs
      If x <> GDeptNum Then
        Get DHandle, x, DeptRec
        If CompareThis = DeptRec.DeptNum Then
          Check4Dups = True
          frmFAEditDFACMess.Label1.Caption = "The department code number entered has already been assigned to another department. Unique department numbers are important to the proper operation of this program. To bring up a complete department number list press F10. Otherwise press ESC to return to the screen."
          frmFAEditDFACMess.Label1.Top = 900
          frmFAEditDFACMess.cmdCont.Text = "F10 &Open Dept List"
          frmFAEditDFACMess.cmdExit.Text = "ESC &Return and Edit"
          frmFAEditDFACMess.Show vbModal
          If frmFAEditDFACMess.fptxtChoice.Text = "continue" Then
            Unload frmFAEditDFACMess
            Call cmdDeptList_Click
            Exit For
          Else
            Unload frmFAEditDFACMess
            fptxtDeptCode.SetFocus
            Exit For
          End If
        End If
      End If
    Next x
  End If
  Close DHandle
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditDeptCodes", "Check4Dups", Erl)
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
  Dim DHandle As Integer
  Dim DeptRec As FADeptCodeType
  Dim NumOfRecs As Integer
  Dim DoWhatFlag As WarnOption
  Dim ChangeFlag As Boolean
  Dim NumOfDepts As Integer
  Dim NewFlag As Boolean
  
  On Error GoTo ERRORSTUFF
  NewFlag = False
  Go2Flag = False
  If QPTrim$(fptxtDesc.Text) = "" Then
    MsgBox "Please enter a description for Department Code"
    fptxtDesc.SetFocus
    Close
    Exit Sub
  End If
  If QPTrim$(fptxtDeptCode.Text) = "" Then
    MsgBox "Please enter a number for Department Code"
    fptxtDeptCode.SetFocus
    Close
    Exit Sub
  End If
  
  ChangeFlag = False
  
  If Check4Dups = True Then Exit Sub
  
  OpenFADeptCodeFile DHandle
  NumOfDepts = LOF(DHandle) / Len(DeptRec)
  
  If GDeptNum > 0 Then 'check for changes made to existing data
    Get DHandle, GDeptNum, DeptRec
    If QPTrim$(fptxtDesc.Text) <> QPTrim$(DeptRec.DeptDesc) Then
      ChangeFlag = True
      fptxtDesc.SetFocus
    ElseIf QPTrim$(fptxtDeptCode.Text) <> DeptRec.DeptNum Then
      ChangeFlag = True
      fptxtDeptCode.SetFocus
    End If
    
    If ChangeFlag = True Then 'ok...we found a change that, if saved,
    'deletes current data and replaces it with new data...this could
    'have an impact on depreciation history that will not be updated
    'with this new data
      DoWhatFlag = PromptWarnOverWrite(Me)
      Select Case DoWhatFlag
        Case WarnOption.wSave
          MainLog ("Overwrite warning issued for " + QPTrim$(fptxtDeptCode.Text) + ". Save option selected in frmFAEditDeptCodes.")
        Case WarnOption.wExit
          MainLog ("Overwrite warning issued for " + QPTrim$(fptxtDeptCode.Text) + ". Exit option selected in frmFAEditDeptCodes.")
          Close DHandle
          frmFADeptCodeMenu.Show
          DoEvents
          KillFile ("editdeptopen.dat")
          Unload frmFAEditDeptCodes
          Exit Sub
        Case WarnOption.wReturn
          MainLog ("Overwrite warning issued for " + QPTrim$(fptxtDeptCode.Text) + ". Return option selected in frmFAEditDeptCodes.")
          Close DHandle
          Exit Sub
        Case WarnOption.wGo2Add
          MainLog ("Overwrite warning issued for " + QPTrim$(fptxtDeptCode.Text) + ". Add to list option selected in frmFAEditDeptCodes.")
          Go2Flag = True
          If Check4Dups = True Then
            Go2Flag = False
            Exit Sub
          End If
          DeptRec.DeptNum = Val(fptxtDeptCode.Text)
          DeptRec.DeptDesc = QPTrim$(fptxtDesc.Text)
          Put DHandle, NumOfDepts + 1, DeptRec
          Close DHandle
          GoTo Go2
        Case Else
          Close DHandle
          MsgBox "Please make a valid selection"
          Exit Sub
      End Select
    End If
  End If
  NumOfRecs = LOF(DHandle) \ Len(DeptRec)
  
  'save data if it is a new record with no duplications or
  'if it is an existing record that has been cleared for overwriting
  If GDeptNum = 0 Then 'new record
    NewFlag = True
    DeptRec.DeptNum = Val(fptxtDeptCode.Text)
    DeptRec.DeptDesc = QPTrim$(fptxtDesc.Text)
    Put DHandle, NumOfRecs + 1, DeptRec
    Close DHandle
  Else
    DeptRec.DeptNum = Val(fptxtDeptCode.Text)
    DeptRec.DeptDesc = QPTrim$(fptxtDesc.Text)
    Put DHandle, GDeptNum, DeptRec
    Close DHandle
  End If
  
Go2:
  Call CreateDeptIdx 'since new data is being saved the
  'department index needs updating
  If NewFlag = True Then
    MainLog ("Department number " + QPTrim$(fptxtDeptCode.Text) + " was saved in frmFAEditDeptCodes.")
  Else
   'record overwriting of existing records
    Call LogSaves
  End If
  Close
  
  MsgBox "Your information has been saved"
  
  If NewFlag = True Then 'when entering a list of new departments
  'this feature speeds up the entry process
    If MsgBox("Do you want to add another new department?", vbYesNo) = vbYes Then
      GDeptNum = 0
      Call LoadMe
      fptxtDeptCode.SetFocus
      Exit Sub
    End If
  End If
  
  GDeptNum = 0
  frmFADeptCodeMenu.Show
  DoEvents
  KillFile ("editdeptopen.dat")
  Unload frmFAEditDeptCodes
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditDeptCodes", "cmdSave_Click", Erl)
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
  FirstFlag = True
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
    Case vbKeyF8
      SendKeys "%L"
      Call cmdDeptList_Click
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
      KillFile ("editdeptopen.dat")
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAEditDeptCodes.")
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
  Dim DHandle As Integer
  Dim DeptRec As FADeptCodeType
  Dim One As Integer
  Dim FileHandle As Integer
  
  One = 1
  FileHandle = FreeFile
  Open "editdeptopen.dat" For Output As FileHandle Len = 2
  
  Print #FileHandle, One
  Close FileHandle
  
  'this screen is used for adding or for editing existing data
  'so this if statement loads labeling according to what is happening
  If GDeptNum = 0 Then
    Me.Caption = "Adding Department Code"
    Me.Label2 = "Adding Fixed Asset Dept Code"
    fptxtDeptCode.Text = ""
    fptxtDesc.Text = ""
  Else
    Me.Caption = "Editing Department Code"
    Me.Label2 = "Editing Fixed Asset Dept Code"
    OpenFADeptCodeFile DHandle
    Get DHandle, GDeptNum, DeptRec
    fptxtDeptCode.Text = DeptRec.DeptNum
    TempDeptNum = DeptRec.DeptNum 'global
    fptxtDesc.Text = QPTrim$(DeptRec.DeptDesc)
    TempDeptDesc = QPTrim$(DeptRec.DeptDesc) 'global
    Close DHandle
  End If
  
End Sub

Private Sub LogSaves()
  Dim DeptRec As FADeptCodeType
  Dim DHandle As Integer
  
  OpenFADeptCodeFile DHandle
  Get DHandle, GDeptNum, DeptRec
  Close DHandle
  'record all changes made to existing data
  If TempDeptNum <> DeptRec.DeptNum Then
    MainLog ("Department number " + CStr(TempDeptNum) + " has been changed and saved to " + CStr(DeptRec.DeptNum) + " in frmFAEditDeptCodes.")
  End If
  
  If QPTrim$(TempDeptDesc) <> QPTrim$(DeptRec.DeptDesc) Then
    MainLog ("For department number: " + CStr(DeptRec.DeptNum) + ", the description has been changed from " + QPTrim$(TempDeptDesc) + " and saved as " + QPTrim$(DeptRec.DeptDesc) + " in frmFAEditDeptCodes.")
  End If

End Sub

Private Sub fptxtDeptCode_LostFocus()
'  Dim x As Integer
'  Dim Number As Integer
'  Dim DHandle As Integer
'  Dim DeptRec As FADeptCodeType
'  Dim NumOfDepts As Integer
'  Dim Found As Boolean
'
'  If QPTrim$(fptxtDeptCode.Text) = "" Then Exit Sub
'  Number = Val(fptxtDeptCode.Text)
'  OpenFADeptCodeFile DHandle
'  NumOfDepts = LOF(DHandle) / Len(DeptRec)
'
'  If NumOfDepts = 0 Then Exit Sub
'
'  If GDeptNum = 0 Then 'start with blank screen
'  'and enter a tag number...if the tag number entered
'  'is already in use then pop screen with its data
'  '...if not this number is a new one
'    For x = 1 To NumOfDepts
'      Get DHandle, x, DeptRec
'      If Number = DeptRec.DeptNum Then
'        GoTo EditIt
'      End If
'    Next x
'    If x = NumOfDepts + 1 Then
'      Close
'      Exit Sub
'    End If
'  End If
'
'
'EditIt:
'
'  For x = 1 To NumOfDepts
'    Get DHandle, x, DeptRec
'    If Number = DeptRec.DeptNum Then 'match the selected
'    'row with the right code
'      Found = True
'      GDeptNum = x 'now you can assign the correct global
'      Exit For
'    Else
'      Found = False
'      GoTo NotAMatch
'    End If
'
'NotAMatch:
'  Next x
'  Close DHandle
'
'  If Found = False Then
'    If MsgBox("The department code entered does not match any of those saved. Would you like to see the department code list?", vbYesNo) = vbYes Then
'      Call cmdDeptList_Click
'    Else
'      fptxtDeptCode.SetFocus
'    End If
'  Else
'    Call LoadMe
'    If FirstFlag = True Then
'      FirstFlag = False
'    Else
'      fptxtDesc.SetFocus
'    End If
'  End If

End Sub

