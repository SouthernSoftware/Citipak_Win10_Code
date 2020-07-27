VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmFAAssetCodesMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assets Codes Menu"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAAssetsCodesMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdAddNewAssetCode 
      Height          =   495
      Left            =   4005
      TabIndex        =   1
      ToolTipText     =   "Create a brand new asset code record."
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
      ButtonDesigner  =   "frmFAAssetsCodesMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditExistingCode 
      Height          =   495
      Left            =   4005
      TabIndex        =   2
      ToolTipText     =   "Select an existing asset code record to make changes."
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
      ButtonDesigner  =   "frmFAAssetsCodesMenu.frx":0AB0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintCodeListing 
      Height          =   495
      Left            =   4005
      TabIndex        =   3
      ToolTipText     =   "Creates a list of existing asset codes that can be printed."
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
      ButtonDesigner  =   "frmFAAssetsCodesMenu.frx":0C96
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   4005
      TabIndex        =   4
      ToolTipText     =   "Exit back to main menu."
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
      ButtonDesigner  =   "frmFAAssetsCodesMenu.frx":0E7C
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   120
      Index           =   4
      Left            =   8602
      Top             =   2102
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   120
      Index           =   3
      Left            =   2101
      Top             =   2102
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1092
      Index           =   1
      Left            =   1500
      Top             =   896
      Width           =   8652
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8682.765
      X2              =   9402.579
      Y1              =   7881.747
      Y2              =   7881.747
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2200.433
      X2              =   2929.246
      Y1              =   7881.747
      Y2              =   7881.747
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220.428
      X2              =   2220.428
      Y1              =   2146.112
      Y2              =   7876.874
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8699.76
      X2              =   8699.76
      Y1              =   2149.036
      Y2              =   7876.874
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   8579.791
      X2              =   9539.544
      Y1              =   1912.203
      Y2              =   1912.203
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   2100.459
      X2              =   3060.212
      Y1              =   1912.203
      Y2              =   1912.203
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ASSET CODES MENU"
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
Attribute VB_Name = "frmFAAssetCodesMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdAddNewAssetCode_Click()
  frmFAEditAssetCode.Show
  DoEvents
  Unload frmFAAssetCodesMenu
End Sub

Private Sub cmdEditExistingCode_Click()
  On Error Resume Next
  If Not Exist("FAASSIDX.DAT") Then
    MsgBox "No Asset Codes saved in index."
    Exit Sub
  End If
  frmFACodeLookUp.Show
  DoEvents
  Unload frmFAAssetCodesMenu
End Sub

Private Sub cmdExit_Click()
  frmFAMaintMenu.Show
  Close
  DoEvents
  Unload frmFAAssetCodesMenu
End Sub

Private Sub cmdPrintCodeListing_Click()
  Dim PrintType$
  On Error Resume Next
  If Not Exist("FAASSIDX.DAT") Then
    MsgBox "No Asset Codes saved in index."
    Exit Sub
  End If
  
  frmFAReportOpt.Show vbModal
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
  Dim CodeRec As FAAssetCodeRecType
  Dim CodeRecLen As Integer
  Dim CodeIdxRec As ACNumbSortIdxType
  Dim CodeIdxHandle As Integer
  Dim CodeIdxRecNum As Integer
  Dim Dash80$, FF$
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim ItemCnt As Integer
  Dim RptHandle As Integer
  Dim ReportFile$
  Dim FAFile As Integer
  Dim NumOfFARecs As Integer
  Dim cnt&
  Dim Page As Integer
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  If Not Exist("FAASSIDX.DAT") Then
    MsgBox "No Asset Codes saved in index."
    Exit Sub
  End If

  OpenAssIdxFile CodeIdxHandle
  CodeIdxRecNum = LOF(CodeIdxHandle) \ Len(CodeIdxRec)
  If CodeIdxRecNum = 0 Then
    MsgBox "No Asset Codes saved in index."
    Close
    Exit Sub
  End If
  
  ReDim ACIdx(1 To CodeIdxRecNum) As Integer
  For x = 1 To CodeIdxRecNum
    Get CodeIdxHandle, x, CodeIdxRec
      ACIdx(x) = CodeIdxRec.AssRecNum 'load up array with
      'references to record numbers
  Next x
  Close CodeIdxHandle
  
  CodeRecLen = Len(CodeRec)

  ReportFile$ = "FACODE.PRN"   'Report File Name
  Dash80$ = String$(73, "=")
  FF$ = Chr$(12)

  MaxLines = 50
  LineCnt = 0
  ItemCnt = 0

  RptHandle = FreeFile

  Open ReportFile$ For Output As #RptHandle

  GoSub PrintHeader

  OpenFACodeNameFile FAFile
  NumOfFARecs = LOF(FAFile) / Len(CodeRec)

  For cnt = 1 To NumOfFARecs
    Get FAFile, ACIdx(cnt), CodeRec
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    'Check For Disposed Of
    Print #RptHandle, Tab(10); CodeRec.ASSETCODE;
    Print #RptHandle, Tab(30); CodeRec.AssetDesc;
    Print #RptHandle, Tab(70); CodeRec.AssetStatus
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
  Print #RptHandle, Tab(34); "Master Code Listing"
  Print #RptHandle, Tab(5); "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle,
  Print #RptHandle, Tab(5); "Asset Catagory Code"; Tab(30); "Description"; Tab(70); "Status"
  Print #RptHandle, Tab(5); Dash80$
  LineCnt = 5
  Return

PrintEnding:
  Print #RptHandle, FF$
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAAssetCodesMenu", "PrintText", Erl)
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
  GCodeNum = 0 'clear this global before any processing
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
      ClearInUse PWcnt 'sets all password files back to 0
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAAssetCodesMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintGraphics()
  Dim CodeRec As FAAssetCodeRecType
  Dim CodeRecLen As Integer
  Dim CodeIdxRec As ACNumbSortIdxType
  Dim CodeIdxHandle As Integer
  Dim CodeIdxRecNum As Integer
  Dim ItemCnt As Integer
  Dim RptHandle As Integer
  Dim ReportFile$
  Dim FAFile As Integer
  Dim NumOfFARecs As Integer
  Dim cnt&
  Dim x As Integer
  Dim dlm$
  Dim FASetUpRec As FASetupRecType
  Dim FASHandle As Integer
  Dim Employer$
  
  On Error GoTo ERRORSTUFF
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  
  Employer$ = FASetUpRec.TownName
  
  dlm$ = "~" 'delimiter
  If Not Exist("FAASSIDX.DAT") Then
    MsgBox "No Asset Codes saved in index."
    Exit Sub
  End If

  OpenAssIdxFile CodeIdxHandle
  CodeIdxRecNum = LOF(CodeIdxHandle) \ Len(CodeIdxRec)
  If CodeIdxRecNum = 0 Then Exit Sub
  
  ReDim ACIdx(1 To CodeIdxRecNum) As Integer
  For x = 1 To CodeIdxRecNum
    Get CodeIdxHandle, x, CodeIdxRec
      ACIdx(x) = CodeIdxRec.AssRecNum
  Next x
  Close CodeIdxHandle
  
  CodeRecLen = Len(CodeRec)

  ReportFile$ = "FARPTS\FACODE.RPT" 'Report File Name used by AR
  'reports as its source of data from which to display a report
  ItemCnt = 0

  RptHandle = FreeFile

  Open ReportFile$ For Output As #RptHandle

  OpenFACodeNameFile FAFile
  NumOfFARecs = LOF(FAFile) / Len(CodeRec)

  For cnt = 1 To NumOfFARecs
    Get FAFile, ACIdx(cnt), CodeRec
    'Check For Disposed Of
    '                     0
    Print #RptHandle, QPTrim$(Employer); dlm;
    '                     1
    Print #RptHandle, QPTrim$(CodeRec.ASSETCODE); dlm;
    '                     2
    Print #RptHandle, QPTrim$(CodeRec.AssetDesc); dlm;
    '                     3
    Print #RptHandle, QPTrim$(CodeRec.AssetStatus)
  Next cnt

  Close         'Close all open files now

  arFAAssCodeList.Show
  frmFALoadReport.Show
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAAssetCodesMenu", "PrintGraphics", Erl)
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

