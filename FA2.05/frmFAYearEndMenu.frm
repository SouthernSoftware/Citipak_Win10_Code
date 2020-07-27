VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmFAYearEndMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets Year End Menu"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAYearEndMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintDepBuildFile 
      Height          =   495
      Left            =   4005
      TabIndex        =   2
      ToolTipText     =   "Click this button to create a printout of the temporary file created in the Build Depreciation File section."
      Top             =   3975
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
      ButtonDesigner  =   "frmFAYearEndMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdBuildDepFile 
      Height          =   492
      Left            =   4005
      TabIndex        =   1
      ToolTipText     =   $"frmFAYearEndMenu.frx":0ABB
      Top             =   3216
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmFAYearEndMenu.frx":0B46
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPostDepToAssets 
      Height          =   495
      Left            =   4005
      TabIndex        =   3
      ToolTipText     =   "Click this button to commit the temporary file created in Build Depreciation File to memory permanently."
      Top             =   4728
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
      ButtonDesigner  =   "frmFAYearEndMenu.frx":0D31
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDeleteTempDpr 
      Height          =   495
      Left            =   4005
      TabIndex        =   4
      ToolTipText     =   "Click this button to undo the temporary depreciation file created in the Build Depreciation File section."
      Top             =   5472
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
      ButtonDesigner  =   "frmFAYearEndMenu.frx":0F20
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   4005
      TabIndex        =   5
      ToolTipText     =   "Click this button to return to the fixed assets main menu."
      Top             =   6240
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
      ButtonDesigner  =   "frmFAYearEndMenu.frx":110F
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   135
      Index           =   4
      Left            =   8610
      Top             =   2091
      Width           =   960
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   135
      Index           =   3
      Left            =   2110
      Top             =   2092
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FIXED ASSETS YEAR END MENU"
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
      X1              =   8699.76
      X2              =   8699.76
      Y1              =   2149.036
      Y2              =   7870.051
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220.428
      X2              =   2220.428
      Y1              =   2146.112
      Y2              =   7867.127
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2199.434
      X2              =   2929.246
      Y1              =   7881.747
      Y2              =   7881.747
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8682.765
      X2              =   9402.579
      Y1              =   7881.747
      Y2              =   7881.747
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
Attribute VB_Name = "frmFAYearEndMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class
Private Sub cmdBuildDepFile_Click()
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxCnt As Long
  
  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) / Len(TagIdx)
  Close TagIdxHandle
  
  If TagIdxCnt = 0 Then
    MsgBox "There are no fixed assets saved."
    Exit Sub
  End If
  
  frmFABuildYrEndDep.Show
  DoEvents
  Unload frmFAYearEndMenu
End Sub

Private Sub cmdDeleteTempDpr_Click()
  Dim DepFile As Integer
  Dim Nextx As Integer
  Dim FADep(1) As FADepFileType
  Dim NumOfDprRecs As Long
  Dim DprHistRec As DprHistType
  
  DepFile = FreeFile
  Open "FADPREDT.DAT" For Random Access Read Write Shared As #DepFile Len = Len(FADep(1))
  NumOfDprRecs = LOF(DepFile) / Len(FADep(1))
  If NumOfDprRecs = 0 Then
    Close
    MsgBox "There are no pending depreciation files to delete."
    KillFile ("FADPREDT.DAT")
    Exit Sub
  End If
  
  Close
  
  frmFADeletePendingDpr.Show
  DoEvents
  Unload frmFAYearEndMenu
End Sub

Private Sub cmdExit_Click()
  frmFAMainMenu.Show
  Close
  DoEvents
  Unload frmFAYearEndMenu
End Sub

Private Sub cmdPostDepToAssets_Click()
  Dim DepFile As Integer
  Dim Nextx As Integer
  Dim FADep(1) As FADepFileType
  Dim NumOfDprRecs As Long
  Dim DprHistRec As DprHistType
  
  DepFile = FreeFile
  Open "FADPREDT.DAT" For Random Access Read Write Shared As #DepFile Len = Len(FADep(1))
  NumOfDprRecs = LOF(DepFile) / Len(FADep(1))
  If NumOfDprRecs = 0 Then
    If MsgBox("Building depreciation files has not taken place. Do you wish to jump to the Build Depreciation File screen?", vbYesNo) = vbYes Then
      frmFABuildYrEndDep.Show
      DoEvents
      Unload frmFAYearEndPost
      Close
      Exit Sub
    Else
      Close 'added 03/15/2004
      Exit Sub
    End If
  End If
  
  Close
  
  frmFAYearEndPost.Show
  DoEvents
  Unload frmFAYearEndMenu
End Sub

Private Sub cmdPrintDepBuildFile_Click()
  Dim DepFile As Integer
  Dim FADep(1) As FADepFileType
  Dim NumOfDepRecs As Integer
  
  OpenDeprEditFile DepFile
  NumOfDepRecs = LOF(DepFile) / Len(FADep(1))
  Close DepFile

  If NumOfDepRecs = 0 Then
    MsgBox "No temporary depreciation records have been saved. Use the build depreciation feature to create these records."
    Exit Sub
  End If
  
  frmFAPrePostPrint.Show
  DoEvents
  Unload Me
End Sub

Private Sub PrintTextByDept()
  Dim DOrigCost#, DBookTotal#, DCDep#, DYDep#, OrigCost#, BookTotal#, CDep#, YDep#, TAccuDpr#
  Dim YrFile As Integer
  Dim FAYear(1) As FAYearEndType
  Dim YearRecNum As Integer
  Dim LastYr$
  Dim ReportFile$
  Dim Dash80$
  Dim FF$, CurDep#
  Dim MaxLines As Integer
  Dim LineCnt&, ItemCnt&
  Dim RptHandle As Integer
  Dim FAFile As Integer
  Dim FAItemRec As FAItemRecType
  Dim DepFile As Integer
  Dim FADep(1) As FADepFileType
  Dim NumOfDepRecs As Integer
  Dim cnt&, Page As Integer
  Dim ItemRecNo As Long
  Dim DeptNumber As Integer
  Dim DCurDep#
  Dim YTDDep#
  Dim NumOfFARecs As Integer
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxRecNums As Integer
  Dim DeptDesc$, x As Integer
  Dim DItemCnt&
  Dim AccuDpr As Double
  Dim DAccuDpr As Double
  Dim Employer$
  Dim FASetUpRec As FASetupRecType
  Dim FASHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenDeptIdxFile DIdxHandle
  DIdxRecNums = LOF(DIdxHandle) / Len(DeptIdx)
  If DIdxRecNums > 0 Then
    ReDim DeptIndx(1 To DIdxRecNums) As String
    ReDim DeptNum(1 To DIdxRecNums) As Integer
    For x = 1 To DIdxRecNums
      Get DIdxHandle, x, DeptIdx
      DeptIndx(x) = QPTrim$(DeptIdx.DeptIdxDesc)
      DeptNum(x) = DeptIdx.DeptNumb
    Next x
    Close DIdxHandle
  End If
  
  OpenYearFile YrFile
  YearRecNum = LOF(YrFile) / Len(FAYear(1))
  If YearRecNum = 0 Then
    LastYr$ = "N/A"
  Else
    Get YrFile, 1, FAYear(1)
    LastYr$ = FAYear(1).CurYear
  End If
  Close YrFile
  
  ReportFile$ = "FADEPEDT.PRN"  'Report File Name
  Dash80$ = String$(80, "=")
  FF$ = Chr$(12)
  
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  
  Employer$ = FASetUpRec.TownName
  
  MaxLines = 53
  LineCnt& = 0
  ItemCnt& = 0
  DItemCnt& = 0
  
  RptHandle = FreeFile
  
  Open ReportFile$ For Output As #RptHandle
  
  OpenFAItemFile FAFile
  NumOfFARecs = LOF(FAFile) / Len(FAItemRec)
  
  'Open Deprec Edit File
  OpenDeprEditFile DepFile
  NumOfDepRecs = LOF(DepFile) / Len(FADep(1))
  Get DepFile, 1, FADep(1)
  GoSub PrintMasterHeader3
  If NumOfDepRecs = 0 Then
    Close
    MsgBox "No temporary depreciation records have been saved. Use the build depreciation feature to create these records."
    Exit Sub
  End If
  
  For cnt& = 1 To NumOfDepRecs
    Get DepFile, cnt&, FADep(1)
    ItemRecNo = FADep(1).AssetRecord
    Get FAFile, ItemRecNo, FAItemRec
    If cnt& = 1 Then
      DeptNumber = FAItemRec.IDEPT
    End If
    
    If DIdxRecNums > 0 Then
      For x = 1 To DIdxRecNums
        If DeptNum(x) = FAItemRec.IDEPT Then
          DeptDesc = QPTrim$(DeptIndx(x))
          Exit For
        End If
      Next x
    End If
    
    If LineCnt& >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintMasterHeader3
    End If
    If DeptNumber <> FAItemRec.IDEPT Then 'data is being read in dept order
      'Print Subtotals and Clear
      Print #RptHandle, String$(122, "-")
      Print #RptHandle, "Totals for Dept Number: "; DeptNumber; ; "  "; DeptDesc; "  "; "#Items:"; DItemCnt;
      Print #RptHandle, Tab(64); Using("###,###,##0.00", DOrigCost#);
      Print #RptHandle, Tab(79); Using("###,###,##0.00", DYDep#);
      Print #RptHandle, Tab(93); Using("###,###,##0.00", DCurDep#);
      Print #RptHandle, Tab(109); Using("###,###,##0.00", DAccuDpr#)
      LineCnt& = LineCnt& + 2
      
      Print #RptHandle, "": LineCnt& = LineCnt& + 1
      Print #RptHandle, "": LineCnt& = LineCnt& + 1
      
      DeptNumber = FAItemRec.IDEPT
      DOrigCost# = 0
      DCurDep# = 0
      DYDep# = 0
      DItemCnt& = 0
      DAccuDpr = 0
    End If
    
    'Figure Values
    'Calc Depreciation for This Period
'SkipThisDeptTotal:
    YTDDep# = FAItemRec.DEP2DATE
    AccuDpr = 0
    AccuDpr = OldRound(FADep(1).CurYrDep + YTDDep#)
    Print #RptHandle, FAItemRec.ItemTag; Tab(22); Left$(FAItemRec.IDESC1, 28);
    Print #RptHandle, Tab(51); FAItemRec.IDEPT;
    Print #RptHandle, Tab(58); Using("###", FAItemRec.ILIFE);
    Print #RptHandle, Tab(64); Using("###,###,##0.00", FAItemRec.ORGCOST);
    Print #RptHandle, Tab(79); Using("###,###,##0.00", YTDDep#);
    Print #RptHandle, Tab(93); Using("###,###,##0.00", FADep(1).CurYrDep);
    If FADep(1).PctFlag Then
      Print #RptHandle, "*";
    End If
    Print #RptHandle, Tab(108); Using("###,###,##0.00#", AccuDpr#)
    'SubTotal Here
    LineCnt& = LineCnt& + 1
    ItemCnt& = ItemCnt& + 1
    DItemCnt& = DItemCnt& + 1
    'Grand Totals Here
    OrigCost# = OrigCost# + FAItemRec.ORGCOST
    CurDep# = CurDep# + FADep(1).CurYrDep
    YDep# = YDep# + YTDDep#
    TAccuDpr = TAccuDpr + AccuDpr
    'Dept Totals Here
    DOrigCost# = DOrigCost# + FAItemRec.ORGCOST
    DCurDep# = DCurDep# + FADep(1).CurYrDep
    DYDep# = DYDep# + YTDDep#
    DAccuDpr# = DAccuDpr# + AccuDpr#
    
SkipEm3:
  Next cnt&
  'First Print Subtotals
  
'  Print #RptHandle, String$(105, "-")
  Print #RptHandle, String$(122, "-")
  Print #RptHandle, "Totals for Dept Number: "; DeptNumber; ; "  "; DeptDesc; "  "; "#Items:"; DItemCnt;
  Print #RptHandle, Tab(64); Using("###,###,##0.00", DOrigCost#);
  Print #RptHandle, Tab(79); Using("###,###,##0.00", DYDep#);
  Print #RptHandle, Tab(93); Using("###,###,##0.00", DCurDep#);
  Print #RptHandle, Tab(109); Using("###,###,##0.00", DAccuDpr#)
  LineCnt& = LineCnt& + 2
  
  Print #RptHandle, "": LineCnt& = LineCnt& + 1
  Print #RptHandle, "": LineCnt& = LineCnt& + 1
  
  GoSub PrintDepRepEnding1
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now
  
  ViewPrint ReportFile$, "Current Depreciation Report", True
  KillFile (ReportFile$)
  
  Exit Sub
  
PrintMasterHeader3:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Master Asset Listing : Depreciation Edit Report For "; FADep(1).CurrYear
  Print #RptHandle, Employer
  Print #RptHandle, "Report Date: "; Date$; Tab(68); "Page #"; Page
  Print #RptHandle, "Asset Number"; Tab(22); "Description"; Tab(51); "Dept"; Tab(58); "Life"; Tab(65); "Original Cost"; Tab(81); "Dprc To Date"; Tab(94); "Cur Yr Deprec"; Tab(113); "Accum Dprc"
'  Print #RptHandle, String$(105, "=")
  Print #RptHandle, String$(122, "=")
  LineCnt& = 5
  Return
  
PrintDepRepEnding1:
'  Print #RptHandle, String$(105, "-")
  Print #RptHandle, String$(122, "-")
  Print #RptHandle, "Grand Totals: "; Tab(15); "# Items: "; Tab(26); Using("######0", ItemCnt);
  Print #RptHandle, Tab(64); Using("###,###,##0.00", OrigCost#);
  Print #RptHandle, Tab(79); Using("###,###,##0.00", YDep#);
  Print #RptHandle, Tab(93); Using("###,###,##0.00", CurDep#);
  Print #RptHandle, Tab(109); Using("###,###,##0.00", TAccuDpr#)
  Print #RptHandle, FF$
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAYearEndMenu", "PrintText", Erl)
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)

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
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAYearEndMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintGraphicsByDept()
  Dim DOrigCost#, DBookTotal#, DCDep#, DYDep#, OrigCost#, BookTotal#, CDep#, YDep#
  Dim YrFile As Integer
  Dim FAYear(1) As FAYearEndType
  Dim YearRecNum As Integer
  Dim LastYr$
  Dim ReportFile$
  Dim CurDep#
  Dim ItemCnt&
  Dim RptHandle As Integer
  Dim FAFile As Integer
  Dim FAItemRec As FAItemRecType
  Dim DepFile As Integer
  Dim FADep(1) As FADepFileType
  Dim NumOfDepRecs As Integer
  Dim cnt&
  Dim ItemRecNo As Long
  Dim DeptNumber As Integer
  Dim DCurDep#
  Dim YTDDep#
  Dim NumOfFARecs As Integer
  Dim dlm$, x As Integer
  Dim Employer$
  Dim FASetUpRec As FASetupRecType
  Dim FASHandle As Integer
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxRecNums As Integer
  Dim DeptDesc$, Dpr4Year$
  Dim AccuDpr As Double '9/21/2004
  
  On Error GoTo ERRORSTUFF
  
  OpenDeptIdxFile DIdxHandle
  DIdxRecNums = LOF(DIdxHandle) / Len(DeptIdx)
  If DIdxRecNums > 0 Then
    ReDim DeptIndx(1 To DIdxRecNums) As String
    ReDim DeptNum(1 To DIdxRecNums) As Integer
    For x = 1 To DIdxRecNums
      Get DIdxHandle, x, DeptIdx
      DeptIndx(x) = QPTrim$(DeptIdx.DeptIdxDesc)
      DeptNum(x) = DeptIdx.DeptNumb
    Next x
    Close DIdxHandle
  End If
  
  dlm$ = "~"
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  
  Employer$ = FASetUpRec.TownName
  OpenYearFile YrFile
  YearRecNum = LOF(YrFile) / Len(FAYear(1))
  If YearRecNum = 0 Then
    LastYr$ = "N/A"
  Else
    Get YrFile, 1, FAYear(1)
    LastYr$ = FAYear(1).CurYear
  End If
  Close YrFile
  
  ReportFile$ = "FARPTS\FADEPEDT.RPT"  'Report File Name
  ItemCnt& = 0
  
  RptHandle = FreeFile
  
  Open ReportFile$ For Output As #RptHandle
  
  OpenFAItemFile FAFile
  NumOfFARecs = LOF(FAFile) / Len(FAItemRec)
  
  'Open Deprec Edit File
  OpenDeprEditFile DepFile
  NumOfDepRecs = LOF(DepFile) / Len(FADep(1))
  Get DepFile, 1, FADep(1)
  Dpr4Year$ = QPTrim$(FADep(1).CurrYear)
  
  For cnt& = 1 To NumOfDepRecs
    Get DepFile, cnt&, FADep(1)
    ItemRecNo = FADep(1).AssetRecord
    Get FAFile, ItemRecNo, FAItemRec
    If cnt& = 1 Then
      DeptNumber = FAItemRec.IDEPT
    End If
    
    If DIdxRecNums > 0 Then
      For x = 1 To DIdxRecNums
        If DeptNum(x) = FAItemRec.IDEPT Then
          DeptDesc = QPTrim$(DeptIndx(x))
          Exit For
        End If
      Next x
    End If
    
    If DeptNumber <> FAItemRec.IDEPT Then 'reached the point where
'    'dept totals can be printed
      DeptNumber = FAItemRec.IDEPT
      DOrigCost# = 0
      DCurDep# = 0
      DYDep# = 0
    End If
    'Figure Values
    'Calc Depreciation for This Period
    YTDDep# = FAItemRec.DEP2DATE
    AccuDpr = 0
    AccuDpr = OldRound(FADep(1).CurYrDep + FAItemRec.DEP2DATE)
    '                     0                   1
    Print #RptHandle, Employer; dlm; FAItemRec.ItemTag; dlm;
    '                         2                     3
    Print #RptHandle, FAItemRec.IDESC1; dlm; FAItemRec.IDEPT; dlm;
    '                         4                     5                       6
    Print #RptHandle, FAItemRec.ILIFE; dlm; FAItemRec.ORGCOST; dlm; FAItemRec.DEP2DATE; dlm;
    '                         7                    8                  9                    10
    Print #RptHandle, FADep(1).CurYrDep; dlm; DeptNumber; dlm; FADep(1).CurrYear; dlm; DeptDesc; dlm;
    If FADep(1).PctFlag Then
      '                  11
      Print #RptHandle, "*"; dlm;
    Else
      '                  11
      Print #RptHandle, " "; dlm;
    End If
    '                    12             13
    Print #RptHandle, Dpr4Year$; dlm; AccuDpr
    'SubTotal Here
    ItemCnt& = ItemCnt& + 1
    'Grand Totals Here
    OrigCost# = OrigCost# + FAItemRec.ORGCOST
    CurDep# = CurDep# + FADep(1).CurYrDep
    YDep# = YDep# + YTDDep#
    'Dept Totals Here
    DOrigCost# = DOrigCost# + FAItemRec.ORGCOST
    DCurDep# = DCurDep# + FADep(1).CurYrDep
    DYDep# = DYDep# + YTDDep#
    
    
SkipEm3:
  Next cnt&
  Close         'Close all open files now
  
  arFADprBeforePostRpt.Show
  frmFALoadReport.Show
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAYearEndMenu", "PrintGraphics", Erl)
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

Private Sub PrintTextByFundAsset()
  Dim DOrigCost#, DBookTotal#, DCDep#, DYDep#, OrigCost#, BookTotal#, CDep#, YDep#, TAccuDpr#
  Dim YrFile As Integer
  Dim FAYear(1) As FAYearEndType
  Dim YearRecNum As Integer
  Dim LastYr$
  Dim ReportFile$
  Dim Dash80$
  Dim FF$, CurDep#
  Dim MaxLines As Integer
  Dim LineCnt&, ItemCnt&
  Dim RptHandle As Integer
  Dim FAFile As Integer
  Dim FAItemRec As FAItemRecType
  Dim DepFile As Integer
  Dim FADep(1) As FADepFileType
  Dim NumOfDepRecs As Integer
  Dim cnt&, Page As Integer
  Dim ItemRecNo As Long
  Dim FndAssNumber$ ' As Integer
  Dim DCurDep#
  Dim YTDDep#
  Dim NumOfFARecs As Integer
  Dim x As Integer
  Dim DItemCnt&
  Dim DAccuDpr As Double
  Dim Employer$
  Dim FASetUpRec As FASetupRecType
  Dim FASHandle As Integer
  Dim AccuDpr As Double '9/21/2004
  Dim FundAssetSort$
  Dim Big$
  Dim ThisBig$
  Dim HoldFundAsset$
  Dim HoldRec As Integer
  Dim Nextx As Integer
  Dim ThisCnt As Integer
  Dim ThisFundAsset$
  Dim FHandle As Integer
  Dim FundRec As FAFundCodeType
  Dim NumOfFundRecs As Integer
  Dim CHandle As Integer
  Dim CodeRec As FAAssetCodeRecType
  Dim NumOfCodeRecs As Integer
  Dim ThisFndAssDesc$
  Dim MaxFndAss As Integer
  Dim HoldFndAssDesc$
'  Dim NewRec As Integer
  Dim HoldDesc$
  Dim HoldDpr As FADepFileType
  Dim ThisFADesc As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenFACodeNameFile CHandle
  NumOfCodeRecs = LOF(CHandle) / Len(CodeRec)
  If NumOfCodeRecs = 0 Then
    MsgBox "No asset codes can be found. Report printing aborted."
    Close
    Exit Sub
  End If
  
  ReDim CodeNum(1 To NumOfCodeRecs) As String
  ReDim CodeDesc(1 To NumOfCodeRecs) As String
  
  For x = 1 To NumOfCodeRecs
    Get CHandle, x, CodeRec
    CodeNum(x) = QPTrim$(CodeRec.ASSETCODE)
    CodeDesc(x) = QPTrim$(CodeRec.AssetDesc)
  Next x
  Close CHandle
  
  OpenFAFundCodeFile FHandle
  NumOfFundRecs = LOF(FHandle) / Len(FundRec)
  If NumOfFundRecs = 0 Then
    MsgBox "No fund codes can be found. Report printing aborted."
    Close
    Exit Sub
  End If
  
  ReDim FundNum(1 To NumOfFundRecs) As String
  ReDim FundDesc(1 To NumOfFundRecs) As String
  
  For x = 1 To NumOfFundRecs
    Get FHandle, x, FundRec
    FundNum(x) = CStr(FundRec.FundNum)
    FundDesc(x) = QPTrim$(FundRec.FundDesc)
  Next x
  Close FHandle
  
  OpenYearFile YrFile
  YearRecNum = LOF(YrFile) / Len(FAYear(1))
  If YearRecNum = 0 Then
    LastYr$ = "N/A"
  Else
    Get YrFile, 1, FAYear(1)
    LastYr$ = FAYear(1).CurYear
  End If
  Close YrFile
  
  ReportFile$ = "FADEPEDTFUNDASSET.PRN"  'Report File Name
  Dash80$ = String$(80, "=")
  FF$ = Chr$(12)
  
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  
  Employer$ = FASetUpRec.TownName
  
  MaxLines = 53
  LineCnt& = 0
  ItemCnt& = 0
  DItemCnt& = 0
  
  RptHandle = FreeFile
  
  Open ReportFile$ For Output As #RptHandle
  
  OpenFAItemFile FAFile
  NumOfFARecs = LOF(FAFile) / Len(FAItemRec)
  
  'Open Deprec Edit File
  OpenDeprEditFile DepFile
  NumOfDepRecs = LOF(DepFile) / Len(FADep(1))
  Get DepFile, 1, FADep(1)
  GoSub PrintMasterHeader3
  If NumOfDepRecs = 0 Then
    Close
    MsgBox "No temporary depreciation records have been saved. Use the build depreciation feature to create these records."
    Exit Sub
  End If
  
  ThisFADesc = 1
  
  ReDim FundAsset(1 To NumOfDepRecs) As String 'make arrays that will
  'be used to sort fixed assets by fund and asset code
  ReDim FndAssdesc(1 To NumOfDepRecs) As String 'this array coincides
  'with each fundasset
  
  Nextx = 1
  ReDim SwapDepSort(1 To NumOfDepRecs) As FADepFileType
  For cnt& = 1 To NumOfDepRecs
    Get DepFile, cnt&, FADep(1) 'gather an array that holds only
    'each unique fund/asset number
    SwapDepSort(cnt) = FADep(1)
    ItemRecNo = FADep(1).AssetRecord
    Get FAFile, ItemRecNo, FAItemRec
    ThisFundAsset$ = CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE)
    FundAsset(cnt) = ThisFundAsset
    For x = 1 To NumOfFundRecs
      If CStr(FAItemRec.FundNum) = FundNum(x) Then
        FndAssdesc(Nextx) = FundDesc(x)
        Exit For
      End If
    Next x
    For x = 1 To NumOfCodeRecs
      If QPTrim$(FAItemRec.ASSETCODE) = CodeNum(x) Then
        FndAssdesc(Nextx) = FndAssdesc(cnt) + "/" + CodeDesc(x)
        Exit For
      End If
    Next x
    Nextx = Nextx + 1
  Next cnt
  
  Big = ""
  For x = 1 To NumOfDepRecs
    If FundAsset(x) > Big Then
      Big = FundAsset(x)
    End If
  Next x
  
  Big = Big + "z"
  ThisBig = Big
  Nextx = 1
  Do
    For x = Nextx To NumOfDepRecs
      If FundAsset(x) < Big Then
        Big = FundAsset(x)
        ThisCnt = x
      End If
    Next x
    HoldDesc = FndAssdesc(Nextx)
    FndAssdesc(Nextx) = FndAssdesc(ThisCnt)
    FndAssdesc(ThisCnt) = HoldDesc
    HoldFundAsset = FundAsset(Nextx)
    FundAsset(Nextx) = FundAsset(ThisCnt)
    FundAsset(ThisCnt) = HoldFundAsset
    HoldDpr = SwapDepSort(ThisCnt)
    SwapDepSort(ThisCnt) = SwapDepSort(Nextx)
    SwapDepSort(Nextx) = HoldDpr
    Nextx = Nextx + 1
    If Nextx = NumOfDepRecs + 1 Then Exit Do
    Big = ThisBig
  Loop
  
'  For x = 1 To NumOfDepRecs
'    Debug.Print FndAssdesc(x)
'  Next x
  Nextx = 1
  For cnt& = 1 To NumOfDepRecs
    ItemRecNo = SwapDepSort(cnt).AssetRecord
    
    Get FAFile, ItemRecNo, FAItemRec
    If cnt& = 1 Then
      FndAssNumber = CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE)
    End If
    
    If LineCnt& >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintMasterHeader3
    End If
    If FndAssNumber <> CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE) Then 'data is being read in dept order
      ThisFADesc = ThisFADesc + 1
      'Print Subtotals and Clear
      Print #RptHandle, String$(122, "-")
      Print #RptHandle, "Totals for: "; FndAssNumber; ; "  "; FndAssdesc(cnt - 1); "  "; "#Items:"; DItemCnt;
      Print #RptHandle, Tab(64); Using("###,###,##0.00", DOrigCost#);
      Print #RptHandle, Tab(79); Using("###,###,##0.00", DYDep#);
      Print #RptHandle, Tab(93); Using("###,###,##0.00", DCurDep#);
      Print #RptHandle, Tab(109); Using("###,###,##0.00", DAccuDpr#)
      LineCnt& = LineCnt& + 2
      
      Print #RptHandle, "": LineCnt& = LineCnt& + 1
      Print #RptHandle, "": LineCnt& = LineCnt& + 1
      
      FndAssNumber = CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE) 'FAItemRec.IDEPT
      DOrigCost# = 0
      DCurDep# = 0
      DYDep# = 0
      DItemCnt& = 0
      DAccuDpr = 0
    End If
    
    'Figure Values
    'Calc Depreciation for This Period
    YTDDep# = FAItemRec.DEP2DATE
    AccuDpr = 0
    AccuDpr = OldRound(SwapDepSort(cnt).CurYrDep + YTDDep#)
    Print #RptHandle, FAItemRec.ItemTag; Tab(22); Left$(FAItemRec.IDESC1, 28);
    Print #RptHandle, Tab(51); CStr(FAItemRec.FundNum) + "/" + QPTrim$(FAItemRec.ASSETCODE);
    Print #RptHandle, Tab(58); Using("###", FAItemRec.ILIFE);
    Print #RptHandle, Tab(64); Using("###,###,##0.00", FAItemRec.ORGCOST);
    Print #RptHandle, Tab(79); Using("###,###,##0.00", YTDDep#);
    Print #RptHandle, Tab(93); Using("###,###,##0.00", SwapDepSort(cnt).CurYrDep);
    If FADep(1).PctFlag Then
      Print #RptHandle, "*";
    End If
    Print #RptHandle, Tab(108); Using("###,###,##0.00#", AccuDpr#)
    'SubTotal Here
    LineCnt& = LineCnt& + 1
    ItemCnt& = ItemCnt& + 1
    DItemCnt& = DItemCnt& + 1
    'Grand Totals Here
    OrigCost# = OrigCost# + FAItemRec.ORGCOST
    CurDep# = CurDep# + SwapDepSort(cnt).CurYrDep
    YDep# = YDep# + YTDDep#
    TAccuDpr = TAccuDpr + AccuDpr
    'Fund/Asset Totals Here
    DOrigCost# = DOrigCost# + FAItemRec.ORGCOST
    DCurDep# = DCurDep# + SwapDepSort(cnt).CurYrDep
    DYDep# = DYDep# + YTDDep#
    DAccuDpr# = DAccuDpr# + AccuDpr#
    
SkipEm3:
  Next cnt&
  'First Print Subtotals
  
  Print #RptHandle, String$(122, "-")
  Print #RptHandle, "Totals for: "; FndAssNumber; ; "  "; FndAssdesc(Nextx); "  "; "#Items:"; DItemCnt;
  Print #RptHandle, Tab(64); Using("###,###,##0.00", DOrigCost#);
  Print #RptHandle, Tab(79); Using("###,###,##0.00", DYDep#);
  Print #RptHandle, Tab(93); Using("###,###,##0.00", DCurDep#);
  Print #RptHandle, Tab(109); Using("###,###,##0.00", DAccuDpr#)
  LineCnt& = LineCnt& + 2
  
  Print #RptHandle, "": LineCnt& = LineCnt& + 1
  Print #RptHandle, "": LineCnt& = LineCnt& + 1
  
  GoSub PrintDepRepEnding1
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now
  
  ViewPrint ReportFile$, "Current Depreciation Report", True
  KillFile (ReportFile$)
  
  Exit Sub
  
PrintMasterHeader3:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Master Asset Listing : Depreciation Edit Report For "; FADep(1).CurrYear
  Print #RptHandle, Employer
  Print #RptHandle, "Report Date: "; Date$; Tab(68); "Page #"; Page
  Print #RptHandle, "Asset Number"; Tab(22); "Description"; Tab(51); "Fd/Ast"; Tab(58); "Life"; Tab(65); "Original Cost"; Tab(81); "Dprc To Date"; Tab(94); "Cur Yr Deprec"; Tab(113); "Accum Dprc"
  Print #RptHandle, String$(122, "=")
  LineCnt& = 5
  Return
  
PrintDepRepEnding1:
  Print #RptHandle, String$(122, "-")
  Print #RptHandle, "Grand Totals: "; Tab(15); "# Items: "; Tab(26); Using("######0", ItemCnt);
  Print #RptHandle, Tab(64); Using("###,###,##0.00", OrigCost#);
  Print #RptHandle, Tab(79); Using("###,###,##0.00", YDep#);
  Print #RptHandle, Tab(93); Using("###,###,##0.00", CurDep#);
  Print #RptHandle, Tab(109); Using("###,###,##0.00", TAccuDpr#)
  Print #RptHandle, FF$
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAYearEndMenu", "PrintText", Erl)
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

Private Sub PrintGraphicsByFundAsset()
  Dim DOrigCost#, DBookTotal#, DCDep#, DYDep#, OrigCost#, BookTotal#, CDep#, YDep#
  Dim YrFile As Integer
  Dim FAYear(1) As FAYearEndType
  Dim YearRecNum As Integer
  Dim LastYr$
  Dim ReportFile$
  Dim CurDep#
  Dim ItemCnt&
  Dim RptHandle As Integer
  Dim FAFile As Integer
  Dim FAItemRec As FAItemRecType
  Dim DepFile As Integer
  Dim FADep(1) As FADepFileType
  Dim NumOfDepRecs As Integer
  Dim cnt&
  Dim ItemRecNo As Long
  Dim FndAssNumber$ ' As Integer
  Dim DCurDep#
  Dim YTDDep#
  Dim NumOfFARecs As Integer
  Dim dlm$, x As Integer
  Dim Employer$
  Dim FASetUpRec As FASetupRecType
  Dim FASHandle As Integer
  Dim Dpr4Year$
  Dim AccuDpr As Double '9/21/2004
  Dim FundAssetSort$
  Dim Big$
  Dim ThisBig$
  Dim HoldFundAsset$
  Dim HoldRec As Integer
  Dim Nextx As Integer
  Dim ThisCnt As Integer
  Dim ThisFundAsset$
  Dim FHandle As Integer
  Dim FundRec As FAFundCodeType
  Dim NumOfFundRecs As Integer
  Dim CHandle As Integer
  Dim CodeRec As FAAssetCodeRecType
  Dim NumOfCodeRecs As Integer
'  Dim MaxFndAss As Integer
'  Dim NewRec As Integer
  Dim HoldFndAssDesc$
  Dim HoldDesc$
  Dim HoldDpr As FADepFileType
  
  On Error GoTo ERRORSTUFF
  
  OpenFACodeNameFile CHandle 'gather code data for later use when
  'needing a code description
  NumOfCodeRecs = LOF(CHandle) / Len(CodeRec)
  If NumOfCodeRecs = 0 Then
    MsgBox "No asset codes can be found. Report printing aborted."
    Close
    Exit Sub
  End If
  
  ReDim CodeNum(1 To NumOfCodeRecs) As String
  ReDim CodeDesc(1 To NumOfCodeRecs) As String
  
  For x = 1 To NumOfCodeRecs
    Get CHandle, x, CodeRec
    CodeNum(x) = QPTrim$(CodeRec.ASSETCODE)
    CodeDesc(x) = QPTrim$(CodeRec.AssetDesc)
  Next x
  Close CHandle
  
  'gather fund data for later use when needing a fund description
  OpenFAFundCodeFile FHandle
  NumOfFundRecs = LOF(FHandle) / Len(FundRec)
  If NumOfFundRecs = 0 Then
    MsgBox "No fund codes can be found. Report printing aborted."
    Close
    Exit Sub
  End If
    
  ReDim FundNum(1 To NumOfFundRecs) As String
  ReDim FundDesc(1 To NumOfFundRecs) As String
  
  For x = 1 To NumOfFundRecs
    Get FHandle, x, FundRec
    FundNum(x) = CStr(FundRec.FundNum)
    FundDesc(x) = QPTrim$(FundRec.FundDesc)
  Next x
  Close FHandle
  
  dlm$ = "~"
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  
  Employer$ = FASetUpRec.TownName
  OpenYearFile YrFile
  YearRecNum = LOF(YrFile) / Len(FAYear(1))
  If YearRecNum = 0 Then
    LastYr$ = "N/A"
  Else
    Get YrFile, 1, FAYear(1)
    LastYr$ = FAYear(1).CurYear
  End If
  Close YrFile
  
  ReportFile$ = "FARPTS\FADEPEDTFNDASS.RPT"  'Report File Name
  ItemCnt& = 0
  
  RptHandle = FreeFile
  
  Open ReportFile$ For Output As #RptHandle
  
  OpenFAItemFile FAFile
  NumOfFARecs = LOF(FAFile) / Len(FAItemRec)
  
  'Open Deprec Edit File
  OpenDeprEditFile DepFile
  NumOfDepRecs = LOF(DepFile) / Len(FADep(1))
  Get DepFile, 1, FADep(1)
  Dpr4Year$ = QPTrim$(FADep(1).CurrYear)
  
  ReDim FundAsset(1 To NumOfDepRecs) As String 'make arrays that will
  'be used to sort fixed assets by fund and asset code
  ReDim FndAssdesc(1 To NumOfDepRecs) As String 'this array coincides
  'with each fundasset
  
  Nextx = 1
  ReDim SwapDepSort(1 To NumOfDepRecs) As FADepFileType
  For cnt& = 1 To NumOfDepRecs
    Get DepFile, cnt&, FADep(1) 'gather an array that holds only
    'each unique fund/asset number
    SwapDepSort(cnt) = FADep(1)
    ItemRecNo = FADep(1).AssetRecord
    Get FAFile, ItemRecNo, FAItemRec
    ThisFundAsset$ = CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE)
    FundAsset(cnt) = ThisFundAsset
    For x = 1 To NumOfFundRecs
      If CStr(FAItemRec.FundNum) = FundNum(x) Then
        FndAssdesc(Nextx) = FundDesc(x)
        Exit For
      End If
    Next x
    For x = 1 To NumOfCodeRecs
      If QPTrim$(FAItemRec.ASSETCODE) = CodeNum(x) Then
        FndAssdesc(Nextx) = FndAssdesc(cnt) + "/" + CodeDesc(x)
        Exit For
      End If
    Next x
    Nextx = Nextx + 1
  Next cnt
  
  Big = ""
  For x = 1 To NumOfDepRecs
    If FundAsset(x) > Big Then
      Big = FundAsset(x)
    End If
  Next x
  
  Big = Big + "z"
  ThisBig = Big
  Nextx = 1
  Do
    For x = Nextx To NumOfDepRecs
      If FundAsset(x) < Big Then
        Big = FundAsset(x)
        ThisCnt = x
      End If
    Next x
    HoldDesc = FndAssdesc(Nextx)
    FndAssdesc(Nextx) = FndAssdesc(ThisCnt)
    FndAssdesc(ThisCnt) = HoldDesc
    HoldFundAsset = FundAsset(Nextx)
    FundAsset(Nextx) = FundAsset(ThisCnt)
    FundAsset(ThisCnt) = HoldFundAsset
    HoldDpr = SwapDepSort(ThisCnt)
    SwapDepSort(ThisCnt) = SwapDepSort(Nextx)
    SwapDepSort(Nextx) = HoldDpr
    Nextx = Nextx + 1
    If Nextx = NumOfDepRecs + 1 Then Exit Do
    Big = ThisBig
  Loop
  
  Nextx = 1
  For cnt& = 1 To NumOfDepRecs
    ItemRecNo = SwapDepSort(cnt).AssetRecord
    Get FAFile, ItemRecNo, FAItemRec
    If cnt& = 1 Then
      FndAssNumber = CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE)
    End If
    
    If FndAssNumber <> CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE) Then 'reached the point where
'    'dept totals can be printed
      FndAssNumber = CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE)
      DOrigCost# = 0
      DCurDep# = 0
      DYDep# = 0
      Nextx = Nextx + 1
'      DAccuDpr = 0
    End If
    'Figure Values
    'Calc Depreciation for This Period
    YTDDep# = FAItemRec.DEP2DATE
    AccuDpr = 0
    AccuDpr = OldRound(SwapDepSort(cnt).CurYrDep + FAItemRec.DEP2DATE)
    '                     0                   1
    Print #RptHandle, Employer; dlm; FAItemRec.ItemTag; dlm;
    '                         2                     3
    Print #RptHandle, FAItemRec.IDESC1; dlm; CStr(FAItemRec.FundNum) + QPTrim$(FAItemRec.ASSETCODE); dlm;
    '                         4                     5                       6
    Print #RptHandle, FAItemRec.ILIFE; dlm; FAItemRec.ORGCOST; dlm; FAItemRec.DEP2DATE; dlm;
    '                         7                    8                  9                    10
    Print #RptHandle, SwapDepSort(cnt).CurYrDep; dlm; FndAssNumber; dlm; SwapDepSort(cnt).CurrYear; dlm; FndAssdesc(cnt); dlm;
    If FADep(1).PctFlag Then
      '                  11
      Print #RptHandle, "*"; dlm;
    Else
      '                  11
      Print #RptHandle, " "; dlm;
    End If
    '                    12             13
    Print #RptHandle, Dpr4Year$; dlm; AccuDpr
    'SubTotal Here
    ItemCnt& = ItemCnt& + 1
    'Grand Totals Here
    OrigCost# = OrigCost# + FAItemRec.ORGCOST
    CurDep# = CurDep# + SwapDepSort(cnt).CurYrDep
    YDep# = YDep# + YTDDep#
    'Dept Totals Here
    DOrigCost# = DOrigCost# + FAItemRec.ORGCOST
    DCurDep# = DCurDep# + SwapDepSort(cnt).CurYrDep
    DYDep# = DYDep# + YTDDep#
    
    
SkipEm3:
  Next cnt&
  Close         'Close all open files now
  
  arFADprBeforePostRptFndAss.Show
  frmFALoadReport.Show
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAYearEndMenu", "PrintGraphics", Erl)
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

