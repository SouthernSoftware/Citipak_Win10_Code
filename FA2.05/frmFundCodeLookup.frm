VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFAFundCodeLookup 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets Fund Code Lookup"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFundCodeLookup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fpList1 
      Height          =   3435
      Left            =   3705
      TabIndex        =   0
      ToolTipText     =   "Double click a fund number for it to appear on the edit screen."
      Top             =   2565
      Width           =   4380
      _Version        =   196608
      _ExtentX        =   7726
      _ExtentY        =   6059
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
      Columns         =   2
      Sorted          =   0
      LineWidth       =   1
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
      ColumnWidthScale=   2
      RowHeight       =   -1
      MultiSelect     =   0
      WrapList        =   0   'False
      WrapWidth       =   0
      SelMax          =   -1
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
      DataField       =   ""
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      ColDesigner     =   "frmFundCodeLookup.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   690
      Left            =   3360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7302
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
      ButtonDesigner  =   "frmFundCodeLookup.frx":0BAE
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   690
      Left            =   6540
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Click this button after selecting a fund number and it will appear on the edit screen."
      Top             =   7305
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
      ButtonDesigner  =   "frmFundCodeLookup.frx":0D8A
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed Asset Fund Codes Listing"
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
      TabIndex        =   1
      Top             =   1074
      Width           =   6012
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   930
      Width           =   8652
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   4620
      Left            =   2916
      Top             =   2136
      Width           =   5964
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   882
      Width           =   8652
   End
End
Attribute VB_Name = "frmFAFundCodeLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmFAFundCodeMenu.Show
  Close
  DoEvents
  Unload frmFAFundCodeLookup
End Sub

Private Sub cmdProcess_Click()
  Call fpList1_DblClick
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
'    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  'designed to allow the user to scroll past this field without
  'inadvertently changing the value
  
  If KeyCode = vbKeyReturn Then
    If fpList1.ListIndex <> -1 Then GoTo FundAlreadySelected '8/6
    KeyCode = 0
    Exit Sub
FundAlreadySelected:
    fpList1.Col = 1
    If QPTrim$(fpList1.ColText) = "" Then
      MsgBox "No fund has been selected"
      Exit Sub
    Else
      Call fpList1_DblClick
      KeyCode = 0
      Exit Sub
    End If
  End If
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
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
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAFundCodeLookup.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim FHandle As Integer
  Dim FundRec As FAFundCodeType
  Dim FundIdx As FundNumbSortIdxType
  Dim FIdxHandle As Integer
  Dim FIdxRecNums As Integer
  Dim x As Integer
  Dim NumOfFunds As Integer
   
  On Error GoTo ERRORSTUFF
  
  If Not Exist("FAFNDIDX.DAT") Then 'no file exists
    MsgBox "No fund index saved."
    Exit Sub
  End If
  
  OpenFundIdxFile FIdxHandle
  FIdxRecNums = LOF(FIdxHandle) \ Len(FundIdx)
  If FIdxRecNums = 0 Then 'file exists but no records saved
    MsgBox "No funds saved in index."
    Close
    Exit Sub
  End If
  
  ReDim FIdx(1 To FIdxRecNums) As Integer
  
  For x = 1 To FIdxRecNums
    Get FIdxHandle, x, FundIdx
    FIdx(x) = FundIdx.FundRecNum 'load array with funds in numeric order
  Next x
  Close FIdxHandle
  
  If Not Exist("FAFUNDCD.DAT") Then
    MsgBox "Path to FAFUNDCD.DAT could not be found"
    Exit Sub
  End If

  OpenFAFundCodeFile FHandle
   
  For x = 1 To FIdxRecNums
    Get FHandle, FIdx(x), FundRec
    If FundRec.FundNum = 0 Then GoTo BadCode
    fpList1.InsertRow = FundRec.FundNum & " " & Chr$(9) & QPTrim$(FundRec.FundDesc)
BadCode:
  Next x
  Close FHandle
  fpList1.Row = 0
  fpList1.Selected = True 'set focus on first row
  Exit Sub
   

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAFundCodeLookup", "Form Load", Erl)
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
    Close
    Unload Me
End Sub

Private Sub fpList1_DblClick()
  Dim FundRec As FAFundCodeType
  Dim FHandle As Integer
  Dim NumOfFunds As Integer
  Dim x As Integer
  Dim Desc$
  Dim code As Integer
  Dim Found As Boolean
   
  fpList1.Col = 0
  code = Val(fpList1.ColText) 'assign variable selected in list
  fpList1.Col = 1
  Desc$ = QPTrim$(fpList1.ColText) 'assign variable selected in list
   
  OpenFAFundCodeFile FHandle
  NumOfFunds = LOF(FHandle) \ Len(FundRec)
  'no need to check for zero funds because the list has already been loaded
  'with existing funds
  For x = 1 To NumOfFunds
    Get FHandle, x, FundRec 'start looing for a match
    If code = FundRec.FundNum And Desc$ = QPTrim$(FundRec.FundDesc) Then
      Found = True 'OK...a match is found
      fpList1.Row = -1
      GFundNum = x 'assign global
      Exit For 'so exit this loop
    Else
      Found = False 'no match is found so keep going
      GoTo NotAMatch
    End If
      
NotAMatch:
  Next x
  Close FHandle
  
  If Found = True Then 'looks good so continue with editing
    frmFAEditFundCodes.Show
    DoEvents
    Unload frmFAFundCodeLookup
  Else
    MsgBox "No match found."
    Exit Sub
  End If
End Sub


