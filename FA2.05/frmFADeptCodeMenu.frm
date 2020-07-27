VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmFADeptCodeMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets Department Code Maintenance"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFADeptCodeMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdAddNewDeptCode 
      Height          =   495
      Left            =   3984
      TabIndex        =   1
      ToolTipText     =   "Create a brand new department."
      Top             =   3312
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmFADeptCodeMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditExistingCode 
      Height          =   495
      Left            =   3984
      TabIndex        =   2
      ToolTipText     =   "Select an existing department for making any desired changes to its data."
      Top             =   4080
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmFADeptCodeMenu.frx":0AAF
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintCodeListing 
      Height          =   495
      Left            =   3984
      TabIndex        =   3
      ToolTipText     =   "Brings up a list of departments in a format suitable for printing."
      Top             =   4848
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmFADeptCodeMenu.frx":0C95
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   4005
      TabIndex        =   4
      ToolTipText     =   "Click this button to return to the fixed assets maintenance menu."
      Top             =   5616
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmFADeptCodeMenu.frx":0E80
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   120
      Index           =   3
      Left            =   8600
      Top             =   2100
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   120
      Index           =   4
      Left            =   2100
      Top             =   2100
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT CODES MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2820
      TabIndex        =   0
      Top             =   1246
      Width           =   6012
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8700
      X2              =   8700
      Y1              =   2206
      Y2              =   8088
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220
      X2              =   2220
      Y1              =   2203
      Y2              =   8085
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2200
      X2              =   2929
      Y1              =   8090
      Y2              =   8090
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8683
      X2              =   9403
      Y1              =   8090
      Y2              =   8090
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1095
      Index           =   1
      Left            =   1500
      Top             =   895
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1500
      Top             =   766
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2100
      Top             =   1966
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2220
      Top             =   2196
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8592
      Top             =   1966
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8700
      Top             =   2194
      Width           =   732
   End
End
Attribute VB_Name = "frmFADeptCodeMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdAddNewDeptCode_Click()
  frmFAEditDeptCodes.Show
  DoEvents
  Unload frmFADeptCodeMenu
End Sub

Private Sub cmdEditExistingCode_Click()
  On Error Resume Next
  If Not Exist("FADEPIDX.DAT") Then
    MsgBox "No department index saved."
    Exit Sub
  End If
  frmFADeptCodeLookup.Show
  DoEvents
  Unload frmFADeptCodeMenu
End Sub

Private Sub cmdExit_Click()
  frmFAMaintMenu.Show
  Close
  DoEvents
  Unload frmFADeptCodeMenu
End Sub

Private Sub cmdPrintCodeListing_Click()
  Dim PrintType$
  If Not Exist("FADEPIDX.DAT") Then
    MsgBox "No department index saved."
    Exit Sub
  End If
  
  frmFAReportOpt.Show vbModal 'opens small screen from which the
  'user selects the printing method
  PrintType$ = frmFAReportOpt.fptxtPrintType
  Select Case PrintType$
    Case "Graphical"
      Call PrintGraphics
    Case "Text"
      Call PrintText
    Case "Exit"
  End Select
  Unload frmFAReportOpt
End Sub

Private Sub PrintText()

  Dim DHandle As Integer
  Dim DeptRec As FADeptCodeType
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxRecNums As Integer
  Dim x As Integer
  Dim Dash80$, FF$
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim ItemCnt As Integer
  Dim RptHandle As Integer
  Dim ReportFile$
  Dim cnt&
  Dim Page As Integer

  On Error GoTo ERRORSTUFF
  
  If Not Exist("FADEPIDX.DAT") Then
    MsgBox "No department index saved."
    Exit Sub
  End If
  
  OpenDeptIdxFile DIdxHandle
  DIdxRecNums = LOF(DIdxHandle) \ Len(DeptIdx)
  If DIdxRecNums = 0 Then
    MsgBox "No departments saved in index."
    Close
    Exit Sub
  End If
  
  ReDim DIdx(1 To DIdxRecNums) As Integer
  
  For x = 1 To DIdxRecNums
    Get DIdxHandle, x, DeptIdx
    DIdx(x) = DeptIdx.DeptRecNum 'load up array with record number pointers
  Next x
  Close DIdxHandle

  ReportFile$ = "FADEPT.PRN"   'Report File Name
  Dash80$ = String$(60, "=")
  FF$ = Chr$(12)

  MaxLines = 50
  LineCnt = 0
  ItemCnt = 0

  RptHandle = FreeFile

  Open ReportFile$ For Output As #RptHandle

  GoSub PrintHeader


  If Not Exist("FADEPTCD.DAT") Then
    MsgBox "Path to FADEPTCD.DAT could not be found"
    Exit Sub
  End If
  
  OpenFADeptCodeFile DHandle

  For cnt = 1 To DIdxRecNums
    Get DHandle, DIdx(cnt), DeptRec 'retrieve data in dept numeric order
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    'Check For Disposed Of
    Print #RptHandle, Tab(25); DeptRec.DeptNum;
    Print #RptHandle, Tab(50); DeptRec.DeptDesc
    LineCnt = LineCnt + 1
SkipEm:
  Next cnt

  GoSub PrintEnding
  Close         'Close all open files now

  ViewPrint ReportFile$, "Code Listing", False

  KillFile (ReportFile$)
  Exit Sub

PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(29); "Master Department Listing"
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle,
  Print #RptHandle, Tab(20); "Department Number"; Tab(50); "Description"

  Print #RptHandle, Tab(15); Dash80$
  LineCnt = 5
  Return

PrintEnding:
  Print #RptHandle, FF$
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFADeptCodeMenu", "PrintText", Erl)
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
  GCodeNum = 0
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
      SendKeys "%X"
      Call cmdExit_Click
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
      MainLog ("FixedAssets.exe terminated via menu bar on frmFADeptCodeMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintGraphics()

  Dim DHandle As Integer
  Dim DeptRec As FADeptCodeType
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxRecNums As Integer
  Dim x As Integer
  Dim ItemCnt As Integer
  Dim RptHandle As Integer
  Dim ReportFile$
  Dim cnt&
  Dim dlm$
  Dim FASetUpRec As FASetupRecType
  Dim FASHandle As Integer
  Dim FASetupCnt As Integer
  Dim Employer$
  
  On Error GoTo ERRORSTUFF
  OpenFASetUpFile FASHandle
  FASetupCnt = LOF(FASHandle) / Len(FASetUpRec)
  If FASetupCnt = 0 Then
    Close
    MsgBox "No set up record has been saved."
    Exit Sub
  End If
  
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  
  Employer$ = FASetUpRec.TownName
  
  dlm$ = "~"

  If Not Exist("FADEPIDX.DAT") Then
    MsgBox "No department index saved."
    Exit Sub
  End If
  
  OpenDeptIdxFile DIdxHandle
  DIdxRecNums = LOF(DIdxHandle) \ Len(DeptIdx)
  If DIdxRecNums = 0 Then
    MsgBox "No departments saved in index."
    Close
    Exit Sub
  End If
  
  ReDim DIdx(1 To DIdxRecNums) As Integer
  
  For x = 1 To DIdxRecNums
    Get DIdxHandle, x, DeptIdx
    DIdx(x) = DeptIdx.DeptRecNum 'load array with record number pointers
  Next x
  Close DIdxHandle

  ReportFile$ = "FARPTS\FADEPT.RPT"   'Report File Name
  ItemCnt = 0

  RptHandle = FreeFile

  Open ReportFile$ For Output As #RptHandle

  If Not Exist("FADEPTCD.DAT") Then
    Close
    MsgBox "Path to FADEPTCD.DAT could not be found"
    Exit Sub
  End If
  
  OpenFADeptCodeFile DHandle

  For cnt = 1 To DIdxRecNums
    Get DHandle, DIdx(cnt), DeptRec 'retrieve data in dept number order
    'Check For Disposed Of
    Print #RptHandle, QPTrim$(Employer); dlm;
    Print #RptHandle, DeptRec.DeptNum; dlm;
    Print #RptHandle, QPTrim$(DeptRec.DeptDesc)
SkipEm:
  Next cnt

  Close         'Close all open files now

  arFADeptOutline.Show
  frmFALoadReport.Show
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFADeptCodeMenu", "PrintGraphics", Erl)
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

