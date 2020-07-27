VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmBLCategoryMaintMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Category Maintenance"
   ClientHeight    =   8730
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11565
   Icon            =   "frmCategoryMaintMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   300
      Left            =   5760
      TabIndex        =   1
      Top             =   7152
      Width           =   684
      _Version        =   131072
      _ExtentX        =   1206
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   9405029
      ForeColor       =   8454143
      Text            =   ""
      Shape           =   0
      ShapeRoundWidth =   180
      ShapeRoundHeight=   180
      BorderWidth     =   -1
      BorderColor     =   -2147483630
      ThreeDWidth     =   -1
      ThreeDShadowColor=   -2147483632
      ThreeDHighlightColor=   16777215
      ThreeDText      =   0
      ThreeDTextHighlightColor=   16777215
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignPictureH   =   0
      AlignPictureV   =   0
      PictureStyle    =   0
      WordWrap        =   -1  'True
      ScaleMode       =   1
      ThreeDStyle     =   2
      Position        =   0
      PosBaseX        =   0
      PosBaseY        =   0
      PosOffsetX      =   -100
      PosOffsetY      =   300
      MaxWidth        =   5000
      CloudInset      =   100
      CloudMinWidth   =   600
      TailShape       =   2
      TailType        =   2
      TailBaseOffsetOutside=   300
      TailBaseOffsetInside=   100
      TailBaseAxisOutside=   0
      TailBaseAxisInside=   0
      TailBubbleCount =   3
      AlignTextH      =   1
      AlignTextV      =   1
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      AutoScan        =   0
      ScanAllDescendants=   -1  'True
      Interval        =   500
      IntervalNext    =   200
      AutoSize        =   -1  'True
      UseTagProp      =   -1  'True
      HideOnInactiveApp=   0   'False
      HideOnMouseDown =   2
      HideOnKeyDown   =   2
      HideOnFocus     =   0   'False
      ScanDisabledControls=   -1  'True
      ThreeDAppearance=   0
      FollowFocus     =   0   'False
      TemplateName    =   ""
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAddCatCode 
      Height          =   492
      Left            =   3960
      TabIndex        =   2
      Tag             =   "Click this button to add a brand new category code."
      Top             =   2940
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmCategoryMaintMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditCat 
      Height          =   492
      Left            =   3960
      TabIndex        =   3
      Tag             =   "Click this button to make changes to existing category codes."
      Top             =   3628
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmCategoryMaintMenu.frx":0AB5
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCatListRpt 
      Height          =   492
      Left            =   3960
      TabIndex        =   4
      Tag             =   "Click this button to display and print a report of category listings."
      Top             =   4316
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmCategoryMaintMenu.frx":0CA2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdIndex 
      Height          =   480
      Left            =   3960
      TabIndex        =   5
      Tag             =   $"frmCategoryMaintMenu.frx":0E8D
      Top             =   5010
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmCategoryMaintMenu.frx":0F15
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   492
      Left            =   3960
      TabIndex        =   6
      Top             =   5692
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmCategoryMaintMenu.frx":10FF
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   480
      Left            =   3960
      TabIndex        =   7
      Tag             =   "Click this button to exit this menu and return to the main Citipak menu."
      Top             =   6390
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   847
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmCategoryMaintMenu.frx":12E4
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8666
      X2              =   8666
      Y1              =   2136
      Y2              =   8008
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   155
      Index           =   3
      Left            =   8550
      Top             =   1995
      Width           =   990
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   150
      Index           =   4
      Left            =   1970
      Top             =   2000
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY MAINTENANCE "
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
      Left            =   2775
      TabIndex        =   0
      Top             =   1170
      Width           =   6012
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2086
      X2              =   2086
      Y1              =   2133
      Y2              =   8005
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2086
      X2              =   2795
      Y1              =   8010
      Y2              =   8010
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8666
      X2              =   9359
      Y1              =   8010
      Y2              =   8010
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1092
      Index           =   1
      Left            =   1455
      Top             =   820
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1455
      Top             =   690
      Width           =   8652
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   1966
      Top             =   1890
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2086
      Top             =   2130
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8550
      Top             =   1890
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8655
      Top             =   2130
      Width           =   732
   End
End
Attribute VB_Name = "frmBLCategoryMaintMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "Turn Menu &Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "Turn Menu &Help On"
    btnHelp.AutoScan = fpAutoScanOff
  End If
End Sub

Private Sub cmdAddCatCode_Click()
  frmBLCatEdit.Show
  DoEvents
  Unload frmBLCategoryMaintMenu
End Sub

Private Sub cmdCatListRpt_Click()
  Dim PrintType$
  
  On Error Resume Next
  
  If Not Exist("arcatcodeidx.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "No category index saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  frmBLReportOpt.Show vbModal 'opens small screen from which the
  'user selects the printing method
  PrintType$ = frmBLReportOpt.fptxtPrintType
  Select Case PrintType$
    Case "Graphical"
      Call PrintGraphics
    Case "Text"
      frmBLMessageBoxJr.Label1.Caption = "Pitch 17 is recommended for this report."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Call PrintText
    Case "Exit"
  End Select
  Unload frmBLReportOpt
  
  cmdHelp.Text = "Turn Menu &Help On"
  btnHelp.AutoScan = fpAutoScanOff

End Sub

Private Sub cmdEditCat_Click()
'  If Exist("artmppst.dat") Then
'    frmBLMessageBoxJr.Label1.Caption = "There is a pending business license renewal file that has not been posted. Please post this business license update file before continuing."
'    frmBLMessageBoxJr.Label1.Top = 700
'    frmBLMessageBoxJr.Show vbModal
'    Close
'    Exit Sub
'  End If
  
  frmBLCatCodeLookup.Show
  DoEvents
  Unload frmBLCategoryMaintMenu

End Sub

Private Sub cmdExit_Click()
  frmBLMainMenu.Show
  DoEvents
  Unload frmBLCategoryMaintMenu
End Sub

Private Sub cmdIndex_Click()
  
  If Not Exist("arcode.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Error: There are no category codes saved. Re-indexing aborted."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  Call CreateCatCodeIdx
  cmdHelp.Text = "Turn Menu &Help On"
  btnHelp.AutoScan = fpAutoScanOff
  frmBLMessageBoxJr.Label1.Caption = "Category codes have reindexed successfully."
  frmBLMessageBoxJr.Label1.Top = 900
  frmBLMessageBoxJr.Show vbModal
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  GCatNum = 0
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
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
      Call cmdExit_Click
      SendKeys "%X"
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLCategoryMaintMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

Sub PrintText()
  Dim ReportFile$
  Dim FF$
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim CatCodeRecLen As Integer
  Dim IdxFile As Integer
  Dim CatIdxRec As CatCodeIdxType
  Dim CatIdxNum As Integer
  Dim x As Integer
  Dim ARCatCodeRec As ARNewCatCodeRecType
  Dim TrHandle As Integer
  Dim TRNumRecs As Integer
  Dim RptHandle As Integer
  Dim TotalCodes As Integer
  Dim Page As Integer
  Dim cnt As Integer
  Dim CodeType$
  
  On Error GoTo ERRORSTUFF
  ReportFile$ = "ARCUSTLST.PRN"  'Report File Name
  FF$ = Chr$(12)
  MaxLines = 50
  LineCnt = 0
  OpenCatCodeIdxFile IdxFile
  CatIdxNum = LOF(IdxFile) / Len(CatIdxRec)
  'if CatIdxNum is 0 then it will be trapped in the
  'process command
  ReDim CatIdxArr(1 To CatIdxNum) As Integer
  For x = 1 To CatIdxNum
    Get IdxFile, x, CatIdxRec
      CatIdxArr(x) = CatIdxRec.CatCodeRec
  Next x
  Close IdxFile
    
  OpenCatCodeFile TrHandle
  
  TRNumRecs = LOF(TrHandle) \ Len(ARCatCodeRec)
  
  If TRNumRecs <> CatIdxNum Then
    frmBLMessageBoxJr.Label1.Caption = "Error: The number of code records in the code index and in the code files are not the same. Re-index category codes or call Southern Software at 1-800-842- 8190 before continuing."
    frmBLMessageBoxJr.Label1.Top = 600
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  GoSub PrintRptHeader
  frmBLShowPctComp.Label1 = "Loading Category List"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  DoEvents
  
  For cnt = 1 To TRNumRecs
    Get TrHandle, CatIdxArr(cnt), ARCatCodeRec
    If Left$(ARCatCodeRec.CatCode, 1) <> " " Then
      If Len(QPTrim$(ARCatCodeRec.CodeType)) = 0 Then
        ARCatCodeRec.CodeType = "F"
        Put TrHandle, CatIdxArr(cnt), ARCatCodeRec
      End If
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintRptHeader
      End If
      
      Print #RptHandle, QPTrim$(ARCatCodeRec.CatCode);
      Print #RptHandle, Tab(8); Left$(ARCatCodeRec.CODEDESC, 30);
      Print #RptHandle, Tab(40); GetGLNum(ARCatCodeRec.REVGLNUM);
      Print #RptHandle, Tab(55); GetGLNum(ARCatCodeRec.ARGLACCT);
      Print #RptHandle, Tab(70); GetGLNum(ARCatCodeRec.CASHACCT);
      Print #RptHandle, Tab(85); Using("$#####,#.##", ARCatCodeRec.Fee);
      If ARCatCodeRec.CodeType = "F" Then
        CodeType = "Flat"
      ElseIf ARCatCodeRec.CodeType = "M" Then
        CodeType = "Mult"
      ElseIf CodeType = "S" Then
        CodeType = "Step"
      Else
        CodeType = "*NA*"
      End If
        
      Print #RptHandle, Tab(99); CodeType
      If ARCatCodeRec.CodeType = "S" Then
        Print #RptHandle, "Base Amt"; Tab(15); "Receipts Up To"; Tab(40); "  Plus %"; Tab(55); "On Amount Over"
        Print #RptHandle, Using("#####.##", ARCatCodeRec.BaseAmt1);
        Print #RptHandle, Tab(15); Using("###########,#", ARCatCodeRec.Recpt1);
        Print #RptHandle, Tab(40); Using("##0.000", ARCatCodeRec.Percent1); "%";
        Print #RptHandle, Tab(55); Using("#########,#", ARCatCodeRec.Maximum1)
        Print #RptHandle, Using("#####.##", ARCatCodeRec.BaseAmt2);
        Print #RptHandle, Tab(15); Using("###########,#", ARCatCodeRec.Recpt2);
        Print #RptHandle, Tab(40); Using("##0.000", ARCatCodeRec.Percent2); "%";
        Print #RptHandle, Tab(55); Using("#########,#", ARCatCodeRec.Maximum2)
        Print #RptHandle, Using("#####.##", ARCatCodeRec.BaseAmt3);
        Print #RptHandle, Tab(15); Using("###########,#", ARCatCodeRec.Recpt3);
        Print #RptHandle, Tab(40); Using("##0.000", ARCatCodeRec.Percent3); "%";
        Print #RptHandle, Tab(55); Using("#########,#", ARCatCodeRec.Maximum3)
        Print #RptHandle, Using("#####.##", ARCatCodeRec.BaseAmt4);
        Print #RptHandle, Tab(15); Using("###########,#", ARCatCodeRec.Recpt4);
        Print #RptHandle, Tab(40); Using("##0.000", ARCatCodeRec.Percent4); "%";
        Print #RptHandle, Tab(55); Using("#########,#", ARCatCodeRec.Maximum4)
        Print #RptHandle, Using("#####.##", ARCatCodeRec.BaseAmt5);
        Print #RptHandle, Tab(15); Using("###########,#", ARCatCodeRec.Recpt5);
        Print #RptHandle, Tab(40); Using("##0.000", ARCatCodeRec.Percent5); "%";
        Print #RptHandle, Tab(55); Using("#########,#", ARCatCodeRec.Maximum5)
        Print #RptHandle, Using("#####.##", ARCatCodeRec.BaseAmt6);
        Print #RptHandle, Tab(15); Using("###########,#", ARCatCodeRec.Recpt6);
        Print #RptHandle, Tab(40); Using("##0.000", ARCatCodeRec.Percent6); "%";
        Print #RptHandle, Tab(55); Using("#########,#", ARCatCodeRec.Maximum6)
        Print #RptHandle, String$(79, "-")
        LineCnt = LineCnt + 7
      End If
      TotalCodes = TotalCodes + 1
      LineCnt = LineCnt + 1
    End If
    frmBLShowPctComp.ShowPctComp cnt, TRNumRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next
  GoSub PrintRptEnding
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  
  ViewPrint ReportFile$, "Code Listing", True
  
  Kill ReportFile$
  
  Exit Sub

PrintRptHeader:
  Page = Page + 1
  Print #RptHandle, Tab(27); "A/R System : Category Code Listing "
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle,
  Print #RptHandle, "Code "; Tab(8); "Description"; Tab(40); "Rev GL #"; Tab(55); "A/R GL #"; Tab(70); "Cash GL #"; Tab(87); "FEE AMOUNT  TYPE"
  Print #RptHandle, String$(102, "=")
  LineCnt = 5
Return
  
PrintRptEnding:
  Print #RptHandle, String$(102, "-")
  Print #RptHandle, "Number of Codes .. "; Using("##,##0", TotalCodes)
  Print #RptHandle, FF$
Return
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCategoryMaintMenu", "PrintText", Erl)
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
  

End Sub

Sub PrintGraphics()
  Dim ReportFile$
  Dim CatCodeRecLen As Integer
  Dim IdxFile As Integer
  Dim CatIdxRec As CatCodeIdxType
  Dim CatIdxNum As Integer
  Dim x As Integer
  Dim ARCatCodeRec As ARNewCatCodeRecType
  Dim TrHandle As Integer
  Dim TRNumRecs As Integer
  Dim RptHandle As Integer
  Dim TotalCodes As Integer
  Dim cnt As Integer
  Dim dlm$
  Dim TownName$
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  dlm$ = "~"
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  TownName = QPTrim$(TownRec.TownName)
  
  ReportFile$ = "BLRPTS\MNCODLST.RPT"  'Report File Name
  
  OpenCatCodeIdxFile IdxFile
  CatIdxNum = LOF(IdxFile) / Len(CatIdxRec)
  'if CatIdxNum is 0 then it will be trapped in the
  'process command
  ReDim CatIdxArr(1 To CatIdxNum) As Integer
  For x = 1 To CatIdxNum
    Get IdxFile, x, CatIdxRec
      CatIdxArr(x) = CatIdxRec.CatCodeRec
  Next x
  Close IdxFile
    
  OpenCatCodeFile TrHandle
  TRNumRecs = LOF(TrHandle) \ Len(ARCatCodeRec)
  
  If TRNumRecs <> CatIdxNum Then
    frmBLMessageBoxJr.Label1.Caption = "Error: The number of code records in the code index and in the code files are not the same. Re-index category codes or call Southern Software at 1-800-842-8190 before continuing."
    frmBLMessageBoxJr.Label1.Top = 600
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  frmBLShowPctComp.Label1 = "Loading Category List"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  DoEvents
  
  For cnt = 1 To TRNumRecs
    Get TrHandle, CatIdxArr(cnt), ARCatCodeRec
    If Left$(ARCatCodeRec.CatCode, 1) <> " " Then
      If Len(QPTrim$(ARCatCodeRec.CodeType)) = 0 Then
        ARCatCodeRec.CodeType = "F"
        Put TrHandle, CatIdxArr(cnt), ARCatCodeRec
      End If
      '                     0
      Print #RptHandle, TownName$; dlm;
      '                                  1
      Print #RptHandle, QPTrim$(ARCatCodeRec.CatCode); dlm;
      '                      2
      Print #RptHandle, QPTrim$(ARCatCodeRec.CODEDESC); dlm;
      '                              3
      Print #RptHandle, GetGLNum(ARCatCodeRec.REVGLNUM); dlm;
      '                              4
      Print #RptHandle, GetGLNum(ARCatCodeRec.ARGLACCT); dlm;
      '                              5
      Print #RptHandle, GetGLNum(ARCatCodeRec.CASHACCT); dlm;
      If ARCatCodeRec.CodeType = "M" Then
      '                           6
        Print #RptHandle, ARCatCodeRec.Fee; dlm;
      End If
      If ARCatCodeRec.CodeType = "F" Then
      '                           6
        Print #RptHandle, ARCatCodeRec.Fee; dlm;
      End If
      If ARCatCodeRec.CodeType = "N" Then
      '                           6
        Print #RptHandle, ARCatCodeRec.Fee; dlm;
      End If
      If ARCatCodeRec.CodeType = "S" Then
        '                  6
        Print #RptHandle, ""; dlm;
      '                           7
        Print #RptHandle, ARCatCodeRec.BaseAmt1; dlm;
        '                         8
        Print #RptHandle, ARCatCodeRec.Recpt1; dlm;
        '                         9
        Print #RptHandle, ARCatCodeRec.Percent1; dlm;
        '                         10
        Print #RptHandle, ARCatCodeRec.Maximum1; dlm;
        '                         11
        Print #RptHandle, ARCatCodeRec.BaseAmt2; dlm;
        '                         12
        Print #RptHandle, ARCatCodeRec.Recpt2; dlm;
        '                         13
        Print #RptHandle, ARCatCodeRec.Percent2; dlm;
        '                         14
        Print #RptHandle, ARCatCodeRec.Maximum2; dlm;
        '                         15
        Print #RptHandle, ARCatCodeRec.BaseAmt3; dlm;
        '                         16
        Print #RptHandle, ARCatCodeRec.Recpt3; dlm;
        '                         17
        Print #RptHandle, ARCatCodeRec.Percent3; dlm;
        '                         18
        Print #RptHandle, ARCatCodeRec.Maximum3; dlm;
        '                         19
        Print #RptHandle, ARCatCodeRec.BaseAmt4; dlm;
        '                         20
        Print #RptHandle, ARCatCodeRec.Recpt4; dlm;
        '                         21
        Print #RptHandle, ARCatCodeRec.Percent4; dlm;
        '                         22
        Print #RptHandle, ARCatCodeRec.Maximum4; dlm;
        '                         23
        Print #RptHandle, ARCatCodeRec.BaseAmt5; dlm;
        '                         24
        Print #RptHandle, ARCatCodeRec.Recpt5; dlm;
        '                         25
        Print #RptHandle, ARCatCodeRec.Percent5; dlm;
        '                         26
        Print #RptHandle, ARCatCodeRec.Maximum5; dlm;
        '                         27
        Print #RptHandle, ARCatCodeRec.BaseAmt6; dlm;
        '                         28
        Print #RptHandle, ARCatCodeRec.Recpt6; dlm;
        '                         29
        Print #RptHandle, ARCatCodeRec.Percent6; dlm;
        '                         30
        Print #RptHandle, ARCatCodeRec.Maximum6; dlm;
        GoTo TypeS
      End If
      '                  7
      Print #RptHandle, ""; dlm;
      '                  8
      Print #RptHandle, ""; dlm;
      '                  9
      Print #RptHandle, ""; dlm;
      '                 10
      Print #RptHandle, ""; dlm;
      '                 11
      Print #RptHandle, ""; dlm;
      '                 12
      Print #RptHandle, ""; dlm;
      '                 13
      Print #RptHandle, ""; dlm;
      '                 14
      Print #RptHandle, ""; dlm;
      '                 15
      Print #RptHandle, ""; dlm;
      '                 16
      Print #RptHandle, ""; dlm;
      '                 17
      Print #RptHandle, ""; dlm;
      '                 18
      Print #RptHandle, ""; dlm;
      '                 19
      Print #RptHandle, ""; dlm;
      '                 20
      Print #RptHandle, ""; dlm;
      '                 21
      Print #RptHandle, ""; dlm;
      '                 22
      Print #RptHandle, ""; dlm;
      '                 23
      Print #RptHandle, ""; dlm;
      '                 24
      Print #RptHandle, ""; dlm;
      '                 25
      Print #RptHandle, ""; dlm;
      '                 26
      Print #RptHandle, ""; dlm;
      '                 27
      Print #RptHandle, ""; dlm;
      '                 28
      Print #RptHandle, ""; dlm;
      '                 29
      Print #RptHandle, ""; dlm;
      '                 30
      Print #RptHandle, ""; dlm;
TypeS:
      '                          31
      Print #RptHandle, ARCatCodeRec.CodeType
      TotalCodes = TotalCodes + 1
    End If
    frmBLShowPctComp.ShowPctComp cnt, TRNumRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next
  Close         'Close all open files now
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  
  arBLCodeList.Show
  frmBLLoadReport.Show
  
  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCategoryMaintMenu", "PrintGraphics", Erl)
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
  
  
End Sub

