VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVoidListOfChecks 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Of Checks"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmVoidListOfChecks.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList 
      Height          =   4680
      Left            =   2745
      TabIndex        =   0
      Top             =   2670
      Width           =   6165
      _Version        =   196608
      _ExtentX        =   10874
      _ExtentY        =   8255
      TextAlias       =   ""
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
      Columns         =   4
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
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
      BorderStyle     =   1
      BorderColor     =   8454143
      BorderWidth     =   2
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
      AllowColResize  =   2
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   3
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
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
      ColumnHeaderShow=   -1  'True
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      DataField       =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "frmVoidListOfChecks.frx":08CA
   End
   Begin EditLib.fpText fptxtEName 
      Height          =   450
      Left            =   3165
      TabIndex        =   3
      Top             =   1995
      Width           =   5340
      _Version        =   196608
      _ExtentX        =   9419
      _ExtentY        =   783
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
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fpText2 
      Height          =   735
      Left            =   3060
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   780
      Width           =   5415
      _Version        =   196608
      _ExtentX        =   9546
      _ExtentY        =   1291
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13684944
      ForeColor       =   65535
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   1
      BorderColor     =   -2147483643
      BorderWidth     =   3
      ButtonDisable   =   0   'False
      ButtonHide      =   -1  'True
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
      AlignTextV      =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   0
      ControlType     =   1
      Text            =   "Select Check To Void"
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483643
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtNothing 
      Height          =   1110
      Left            =   2025
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   585
      Width           =   7425
      _Version        =   196608
      _ExtentX        =   13102
      _ExtentY        =   1968
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13684944
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   3
      ThreeDOutsideStyle=   2
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   3
      ThreeDFrameWidth=   0
      BorderStyle     =   1
      BorderColor     =   -2147483630
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
      ThreeDText      =   4
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483640
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   -1  'True
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   1
      BorderDropShadowColor=   -2147483634
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
      Height          =   690
      Left            =   4560
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   7560
      Width           =   2565
      _Version        =   131072
      _ExtentX        =   4524
      _ExtentY        =   1217
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
      ButtonDesigner  =   "frmVoidListOfChecks.frx":0C33
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   1350
      Left            =   1890
      Top             =   495
      Width           =   7695
   End
End
Attribute VB_Name = "frmVoidListOfChecks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim MatchCnt As Integer
Dim PrevTransNum As Integer
Public WhichDate As Integer '2/1/05
Private Sub cmdEscape_Click()
  GlobalName = ""
  GlobalTransNum = 0
'  frmVoidChkEmpList.Show
'  DoEvents
  Unload frmVoidListOfChecks
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpList.Col = 1
    If QPTrim$(fpList.ColText) = "" Then
      MsgBox "No employee has been selected"
      Exit Sub
    Else
      Call fpList_DblClick
      KeyCode = 0
      Exit Sub
    End If
  End If
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdEscape_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call ProcessList
  Me.HelpContextID = hlpVoidAPostedPayroll
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      Call Terminate
      MainLog ("Payroll.exe terminated via menu bar on frmVoidListOfChecks.")
      End
    End If
  End If
End Sub

Private Sub fpList_DblClick()
  Dim TransRec(1) As TransRecType
  Dim x As Long, y As Long
  Dim ChkDate As String
  Dim NetPay As String
  Dim TransNum As String
  Dim TransRecCnt As Long
  Dim Chknum As String
  Dim CheckDate As Integer
  Dim NegChkNum As String
  Dim NegFlag As Boolean
  Dim NegPTN As Integer
  Dim THandle As Integer
  Dim DoWhatFlag As VoidYN
  
  fpList.Col = 0
  Chknum = QPTrim$(fpList.ColText)
  If Val(Chknum) < 0 Then
    frmWarnVoid2.Show vbModal, Me
    Close
    Exit Sub
  End If
  
  fpList.Col = 1
  CheckDate = Date2Num(fpList.ColText)
  
  fpList.Col = 3
  TransNum = QPTrim$(fpList.ColText)
  GlobalTransNum = CDbl(TransNum)
  If Len(QPTrim$(TransNum)) = 0 Then
    MsgBox "No transaction to delete."
    Close
    Exit Sub
  End If
  
  NegFlag = False
  OpenTransHistFile THandle
  Get THandle, TransNum, TransRec(1)
  
  GoSub CheckListForNegChk
  
  If NegFlag = True Then
    frmWarnVoid1.Show vbModal, Me
    Close
    Exit Sub
  End If
  
  frmPRVoidChkPrintout.Show vbModal
  
  If frmPRVoidChkPrintout.fptxtChoice.Text = "abort" Then
    Unload frmPRVoidChkPrintout
    Close
    Exit Sub
  Else
    Unload frmPRVoidChkPrintout
  End If
  
  DoWhatFlag = PromptVoid(Me)
  Select Case DoWhatFlag
  Case VoidYN.vynNo
     Exit Sub
  End Select
  'we've gotten this far so we can void now
  Call VoidThisCheck(Val(TransNum), THandle)
  'at this point the void transaction has already happened
EndThis:
  Close THandle
'  fpList.Clear
'  frmVoidChkEmpList.Show
  DoEvents
  Unload frmVoidListOfChecks
  
'---------------------------
'  OpenTransHistFile THandle
'  Get THandle, TransNum, TransRec(1)
'  TransRec(1).Voided = TransRec(1).Voided
'  Get THandle, TransRec(1).VoidRec, TransRec(1)
'  TransRec(1).Voided = TransRec(1).Voided
'  TransRec(1).VoidRec = TransRec(1).VoidRec
'  Close THandle
'---------------------------
  GlobalTransNum = 0
  
  Exit Sub
  
CheckListForNegChk:
  If Chknum = 0 Then Return
  
  NegPTN = PrevTransNum 'start here looking for the same
  'check number but negated
  NegChkNum = "-" & Chknum
  Do
    Get THandle, NegPTN, TransRec(1)
    If Val(NegChkNum) = TransRec(1).CheckNum Then
      If CheckDate = TransRec(1).CheckDate Then
        NegFlag = True 'OK, we found a negative check
        Exit Do
      End If
    End If
    NegPTN = TransRec(1).PrevTransRec
  Loop Until NegPTN <= 0
Return
  
End Sub

Private Sub ProcessList()
  Dim TEmp2Rec(1) As EmpData2Type
  Dim TTransRec(1) As TransRecType
  Dim EHandle As Integer
  Dim THandle As Integer
  Dim PrevTransRecNum&
  
  OpenEmpData2File EHandle
  Get EHandle, RecNum, TEmp2Rec(1)
  Close EHandle
  PrevTransRecNum& = TEmp2Rec(1).LastTransRec
  PrevTransNum = TEmp2Rec(1).LastTransRec 'global
'  If PrevTransRecNum& <= 0 Then'remmed on 12/17/04
'    MsgBox "No transaction records are on file for this employee."
'    Call cmdEscape_Click
'    Exit Sub
'  End If
  MatchCnt = 0
  fptxtEName.Text = "Employee: " + QPTrim$(TEmp2Rec(1).EmpFName) + " " + QPTrim$(TEmp2Rec(1).EmpLName)
  GlobalName = QPTrim$(TEmp2Rec(1).EmpFName) + " " + QPTrim$(TEmp2Rec(1).EmpLName)
  OpenTransHistFile THandle
  Do 'load list with transactions
    MatchCnt = MatchCnt + 1
    Get THandle, PrevTransRecNum&, TTransRec(1)
'    If TTransRec(1).CheckNum = -105376 Or TTransRec(1).CheckNum = 105376 Then Stop
    If TTransRec(1).Voided = "Y" Then GoTo AlreadyVoided
    fpList.InsertRow = "      " & Using$("#######0", TTransRec(1).CheckNum) & Chr(9) & "   " & MakeRegDate(TTransRec(1).CheckDate) & Chr(9) & "   " & Using("$###,##0.00", TTransRec(1).NetPay) & Chr(9) & PrevTransRecNum&
AlreadyVoided:
    PrevTransRecNum& = TTransRec(1).PrevTransRec
  Loop While PrevTransRecNum& And MatchCnt < 150
  Close THandle
  
  fpList.ListIndex = 0
End Sub

Private Sub VoidThisCheck(TransNum As Integer, THandle As Integer)

  ReDim EmpRec2(1) As EmpData2Type
  Dim E2Handle As Integer
  ReDim EmpRec3(1) As EmpData3Type
  Dim E3Handle As Integer
  Dim TransRec As TransRecType
  Dim TransCnt As Integer
  Dim NextHistRec&
  Dim cnt As Integer
  Dim Check As Long
  Dim VoidChkRec As VoidCheckType
  Dim VHandle As Integer
  Dim VCnt As Double, x As Double
  Dim CheckDate As Integer '3/25/08
  
  InFileNames(1) = "PRDATA\PREMP2.DAT"
  InFileNames(2) = "PRDATA\PREMP3.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  'we're going to go thru all transaction variables and
  'emp3rec variables (those that hold employee totals)
  'affected by negating this check and negate them
  NextHistRec& = (LOF(THandle) \ Len(TransRec)) + 1
  
  OpenEmpData3File E3Handle
  OpenEmpData2File E2Handle
  Get E2Handle, RecNum, EmpRec2(1)
  Get E3Handle, RecNum, EmpRec3(1)
  Get THandle, TransNum, TransRec
  Check = TransRec.CheckNum
  'added 12/17/04-----
  TransRec.Voided = "Y"
  TransRec.VoidRec = NextHistRec&
  Put THandle, TransNum, TransRec
  TransRec.VoidRec = TransNum
  'added 12/17/04  ^
  TransRec.CheckNum = -TransRec.CheckNum
  CheckDate = TransRec.CheckDate '3/25/08
  TransRec.CheckDate = WhichDate 'added 3/25/08
  TransRec.RegHrsWork = -TransRec.RegHrsWork
  EmpRec2(1).EMPVUSED = OldRound(EmpRec2(1).EMPVUSED - TransRec.VacUsed)
  EmpRec2(1).EMPVBAL = OldRound(EmpRec2(1).EMPVBAL + TransRec.VacUsed)
  TransRec.VacUsed = -TransRec.VacUsed
  EmpRec2(1).EMPSLUSE = OldRound(EmpRec2(1).EMPSLUSE - TransRec.SickUsed)
  EmpRec2(1).EMPSLBAL = OldRound(EmpRec2(1).EMPSLBAL + TransRec.SickUsed)
  TransRec.SickUsed = -TransRec.SickUsed
  EmpRec2(1).HolUsed = OldRound(EmpRec2(1).HolUsed - TransRec.HOLHOURS)
  EmpRec2(1).HOLBAL = OldRound(EmpRec2(1).HOLBAL + TransRec.HOLHOURS)
  TransRec.HOLHOURS = -TransRec.HOLHOURS
  EmpRec2(1).PerUsed = OldRound(EmpRec2(1).PerUsed - TransRec.PerHours)
  EmpRec2(1).PERBAL = OldRound(EmpRec2(1).PERBAL + TransRec.PerHours)
  TransRec.PerHours = -TransRec.PerHours
  EmpRec2(1).EMPCTUSE = OldRound(EmpRec2(1).EMPCTUSE - TransRec.CompUsed)
  EmpRec2(1).EMPCTBAL = OldRound(EmpRec2(1).EMPCTBAL + TransRec.CompUsed)
  TransRec.CompUsed = -TransRec.CompUsed
  TransRec.RegHrsPaid = -TransRec.RegHrsPaid
  TransRec.OTHours = -TransRec.OTHours
  TransRec.OTHrsPaid = -TransRec.OTHrsPaid
  EmpRec2(1).EMPCTE = OldRound(EmpRec2(1).EMPCTE - TransRec.OT2Comp)
  EmpRec2(1).EMPCTBAL = OldRound(EmpRec2(1).EMPCTBAL - TransRec.OT2Comp)
  TransRec.OT2Comp = -TransRec.OT2Comp
  For cnt = 1 To 8
    TransRec.TDist(cnt).DRHrs = -TransRec.TDist(cnt).DRHrs
    TransRec.TDist(cnt).DOHrs = -TransRec.TDist(cnt).DOHrs
    TransRec.TDist(cnt).DRWage = -TransRec.TDist(cnt).DRWage
    TransRec.TDist(cnt).DOWage = -TransRec.TDist(cnt).DOWage
  Next
  TransRec.TotRegWage = -TransRec.TotRegWage
  TransRec.TotOTWage = -TransRec.TotOTWage
  TransRec.GrossWage = -TransRec.GrossWage
  For cnt = 1 To 3
    TransRec.EAmt(cnt) = -TransRec.EAmt(cnt)
  Next
  For cnt = 1 To 6
    TransRec.EDist(cnt).EAmt = -TransRec.EDist(cnt).EAmt
  Next
  TransRec.TotAdditEarn = -TransRec.TotAdditEarn
  TransRec.GrossPay = -TransRec.GrossPay
  TransRec.SocGrossPay = -TransRec.SocGrossPay
  TransRec.MedGrossPay = -TransRec.MedGrossPay
  TransRec.FedGrossPay = -TransRec.FedGrossPay
  TransRec.StaGrossPay = -TransRec.StaGrossPay
  TransRec.SocTaxAmt = -TransRec.SocTaxAmt
  TransRec.MedTaxAmt = -TransRec.MedTaxAmt
  TransRec.FedTaxAmt = -TransRec.FedTaxAmt
  TransRec.StaTaxAmt = -TransRec.StaTaxAmt
  TransRec.TotTaxAmt = -TransRec.TotTaxAmt
  TransRec.RetireAmt = -TransRec.RetireAmt
  
  For cnt = 1 To 50
    TransRec.DAmt(cnt) = -TransRec.DAmt(cnt)
  Next
  TransRec.TotDedAmt = -TransRec.TotDedAmt
  TransRec.EICAmt = -TransRec.EICAmt
  TransRec.NetPay = -TransRec.NetPay
  TransRec.MatchRetAmt = -TransRec.MatchRetAmt
  TransRec.MatchSocAmt = -TransRec.MatchSocAmt
  TransRec.MatchMedAmt = -TransRec.MatchMedAmt
  TransRec.RetGrossPay = -TransRec.RetGrossPay
  TransRec.TaxFring = -TransRec.TaxFring
  
  'this void procedure is a transaction, so update TransRec.PrevTransRec
  'with the next transaction number
  If EmpRec2(1).LastTransRec >= 0 Then
    TransRec.PrevTransRec = EmpRec2(1).LastTransRec
  Else
    TransRec.PrevTransRec = 0 'this should never happen
  End If

  EmpRec3(1).YTDGrossPay = OldRound#(EmpRec3(1).YTDGrossPay + TransRec.GrossPay)
  EmpRec3(1).YTDFedGrossPay = OldRound#(EmpRec3(1).YTDFedGrossPay + TransRec.FedGrossPay)
  EmpRec3(1).YTDStaGrossPay = OldRound#(EmpRec3(1).YTDStaGrossPay + TransRec.StaGrossPay)
  EmpRec3(1).YTDSocGrossPay = OldRound#(EmpRec3(1).YTDSocGrossPay + TransRec.SocGrossPay)
  EmpRec3(1).YTDMedGrossPay = OldRound#(EmpRec3(1).YTDMedGrossPay + TransRec.MedGrossPay)

  EmpRec3(1).YTDRegPay = OldRound#(EmpRec3(1).YTDRegPay + TransRec.TotRegWage)
  EmpRec3(1).YTDOTPay = OldRound#(EmpRec3(1).YTDOTPay + TransRec.TotOTWage)
  EmpRec3(1).YTDNet = OldRound#(EmpRec3(1).YTDNet + TransRec.NetPay)

  EmpRec3(1).YTDFederal = OldRound#(EmpRec3(1).YTDFederal + TransRec.FedTaxAmt)
  EmpRec3(1).YTDState = OldRound#(EmpRec3(1).YTDState + TransRec.StaTaxAmt)
  EmpRec3(1).YTDSocial = OldRound#(EmpRec3(1).YTDSocial + TransRec.SocTaxAmt)
  EmpRec3(1).YTDMedicare = OldRound#(EmpRec3(1).YTDMedicare + TransRec.MedTaxAmt)
  EmpRec3(1).YTDRetire = OldRound#(EmpRec3(1).YTDRetire + TransRec.RetireAmt)

  'year to date totals on deductions
  For cnt = 1 To 50
    EmpRec3(1).YTDDAmt(cnt) = OldRound#(EmpRec3(1).YTDDAmt(cnt) + TransRec.DAmt(cnt))
    EmpRec3(1).YTDDAmtT = OldRound#(EmpRec3(1).YTDDAmtT + TransRec.DAmt(cnt))
  Next

  'year to date totals on alt earnings
  EmpRec3(1).YTDEarn1 = OldRound#(EmpRec3(1).YTDEarn1 + TransRec.EAmt(1))
  EmpRec3(1).YTDEarn2 = OldRound#(EmpRec3(1).YTDEarn2 + TransRec.EAmt(2))
  EmpRec3(1).YTDEarn3 = OldRound#(EmpRec3(1).YTDEarn3 + TransRec.EAmt(3))
  EmpRec3(1).YTDEarnT = OldRound#(EmpRec3(1).YTDEarn1 + EmpRec3(1).YTDEarn2 + EmpRec3(1).YTDEarn3)

  '** Added Update EIC year to date.  6/06/94
  If TransRec.EICAmt > 0 Then
    EmpRec3(1).YTDEIC = OldRound#(EmpRec3(1).YTDEIC + TransRec.EICAmt)
  End If


  EmpRec2(1).LastTransRec = CInt(NextHistRec&)
  Put THandle, NextHistRec&, TransRec
  Put E2Handle, RecNum, EmpRec2(1)
  Put E3Handle, RecNum, EmpRec3(1)
  Close E2Handle
  Close E3Handle
  
  'THandle is closed after sub is closed
  MsgBox "Check number #" & Check & " has been deleted."
'  If PostVoidChkToGL(TransRec.CheckDate, QPTrim$(EmpRec2(1).EmpNo), Check) = True Then
  If PostVoidChkToGL(CheckDate, QPTrim$(EmpRec2(1).EmpNo), Check) = True Then '3/25/08 changed from line above
    MsgBox "The General Ledger has been updated."
  Else
    MsgBox "The General Ledger is not updated in this procedure."
  End If
  
  MainLog ("Check #" + QPTrim(Str(Check)) + " for " + QPTrim(EmpRec2(1).EmpFName) + " " + QPTrim(EmpRec2(1).EmpLName) + " was voided.") '8/12

End Sub

