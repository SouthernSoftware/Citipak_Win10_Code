VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "EDT32X30.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmFAMasterItemListing 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Item Listing"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbOrder 
      Height          =   384
      Left            =   5136
      TabIndex        =   0
      Top             =   2880
      Width           =   3228
      _Version        =   196608
      _ExtentX        =   5694
      _ExtentY        =   677
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
      ColDesigner     =   "frmFAMasterItemListing.frx":0000
   End
   Begin LpLib.fpCombo fpcmbYN 
      Height          =   384
      Left            =   7440
      TabIndex        =   1
      Top             =   3600
      Width           =   780
      _Version        =   196608
      _ExtentX        =   1376
      _ExtentY        =   677
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
      ColDesigner     =   "frmFAMasterItemListing.frx":02BF
   End
   Begin VB.CommandButton cmdDept 
      Caption         =   "F8 &Dept List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   6576
      TabIndex        =   11
      Top             =   4368
      Width           =   1356
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F10 &Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   6924
      TabIndex        =   5
      Top             =   6636
      Width           =   1884
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ESC &Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   2988
      TabIndex        =   4
      Top             =   6636
      Width           =   1884
   End
   Begin EditLib.fpText fptxtDeptNum 
      Height          =   396
      Left            =   4944
      TabIndex        =   2
      ToolTipText     =   $"frmFAMasterItemListing.frx":057E
      Top             =   4368
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 - A L a l"
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
   Begin EditLib.fpDateTime fptxtYear 
      Height          =   372
      Left            =   6192
      TabIndex        =   3
      ToolTipText     =   "Enter the Year to extract W2 information here."
      Top             =   5136
      Width           =   1260
      _Version        =   196608
      _ExtentX        =   2222
      _ExtentY        =   656
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
      ButtonStyle     =   2
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
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   -2147483643
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "2003"
      DateCalcMethod  =   1
      DateTimeFormat  =   5
      UserDefinedFormat=   "yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "19800101"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Current Year:"
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
      TabIndex        =   10
      Top             =   5232
      Width           =   1836
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   1068
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Master Asset Listing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2940
      TabIndex        =   9
      Top             =   1212
      Width           =   6012
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Report Order:"
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
      Left            =   3060
      TabIndex        =   8
      Top             =   2988
      Width           =   1836
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Include Desposed Of Items (Y/N):"
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
      Left            =   3312
      TabIndex        =   7
      Top             =   3696
      Width           =   3852
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dept #"
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
      Left            =   3408
      TabIndex        =   6
      Top             =   4464
      Width           =   1260
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3324
      Left            =   1836
      Top             =   2556
      Width           =   7980
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   1020
      Width           =   8652
   End
End
Attribute VB_Name = "frmFAMasterItemListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdDept_Click()
  frmFADeptList.Show vbModal
End Sub

Private Sub cmdExit_Click()
  frmFAReportMenu.Show
  DoEvents
  Unload frmFAMasterItemListing
End Sub

Private Sub cmdPrint_Click()
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim ItemCnt As Integer
  Dim CYear$, Dash80$, FF$
  Dim RptHandle As Integer
  Dim ReportFile$
  Dim ItemRecNo As Integer
  Dim Cnt As Integer
  Dim Dispose$, Dept$
  Dim DeptNumber As Double
  Dim DItemCnt#
  Dim DOrigCost#
  Dim DDepTot#
  Dim OrigCost#, PValue#
  Dim DepTot#, Disp$
  Dim DevSpec$, Page As Integer
  Dim TagFlag As Boolean
  Dim PCnt As Integer
  Dim PNumOFFARecs As Integer
  
  On Error GoTo ERRORSTUFF
  
  ReDim Arr(1 To 1) As Struct 'Template for the sort Arr
  ReportFile$ = "FAMaster.PRN"  'Report File Name
  Dash80$ = String$(80, "=")
  FF$ = Chr$(12)
  Dispose$ = QPTrim$(fpcmbYN.Text)
  Dept$ = QPTrim$(fptxtDeptNum.Text)
  MaxLines = 50
  LineCnt = 0
  ItemCnt = 0

  CYear = Val(Right$(Date$, 4))
  CYear$ = LTrim$(Str$(CYear))
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  GoSub PrintMasterHeader

  OpenFAItemFile FAHandle
  NumOfFARecs = LOF(FAHandle) / Len(FAItemRec)
  PNumOFFARecs = NumOfFARecs * 4
  
  If NumOfFARecs = 0 Then
    MsgBox "No asset items on record."
    Close FAHandle
    Exit Sub
  End If

  frmFAShowPctComp.Label1 = "Sorting Numbers"
  frmFAShowPctComp.Show
  DoEvents
  
  GoSub GetIndex
  
  For Cnt = 1 To NumOfFARecs
    PCnt = PCnt + 1
    ItemRecNo = Arr(Cnt).RecNum
    Get FAHandle, ItemRecNo, FAItemRec
'    Get FAHandle, Cnt, FAItemRec
    If Cnt = 1 Then DeptNumber = Val(FAItemRec.IDEPT)
'    PValue# = 0
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintMasterHeader
    End If

    'Check For Disposed Of

    If Dispose$ = "N" Then
      If FAItemRec.DISPDATE > 0 Then GoTo SkipEm
    End If

'    If Dept$ = "ALL" Then
'      GoTo PrintIt
'    Else
'      If Val(Dept$) <> Val(FAItemRec.IDEPT) Then GoTo SkipEm
'    End If
    If QPTrim$(fpcmbOrder.Text) = "DEPARTMENT NUMBER" And QPTrim$(fptxtDeptNum.Text) <> "ALL" Then
      If DeptNumber <> Val(FAItemRec.IDEPT) Then
        'Print Subtotals and Clear
        Print #RptHandle, "Subtotal for Dept Number: "; DeptNumber
        Print #RptHandle, "Total Items Printed: "; Using("#####", DItemCnt)
        Print #RptHandle, "Total Original Cost: "; Using("$$########,#.##", DOrigCost#)

        Print #RptHandle, "Total Deprc to Date: "; Using("$$########,#.##", DDepTot#)

        Print #RptHandle, FF$
        GoSub PrintMasterHeader
        DeptNumber = Val(FAItemRec.IDEPT)
        DItemCnt = 0
        DOrigCost# = 0
        DDepTot# = 0
      End If
      'May Need to Skip Them
      DItemCnt = DItemCnt + 1
      DOrigCost# = DOrigCost# + FAItemRec.ORGCOST
      DDepTot# = DDepTot# + FAItemRec.DEP2DATE
    End If
PrintIt:

'    If FAItemRec.DISPDATE > 0 Then Disp$ = "Y" Else Disp$ = "N"
    Print #RptHandle, "Tag #"; FAItemRec.ITEMTAG;
    Print #RptHandle, Tab(40); "Ser #"; FAItemRec.SERIALNO
    Print #RptHandle, "Desc "; FAItemRec.IDESC1;
    Print #RptHandle, Tab(40); "Mfg  "; FAItemRec.ITEMMFG
    Print #RptHandle, Tab(5); FAItemRec.IDESC2;
    Print #RptHandle, Tab(40); "Cont "; FAItemRec.CONTACT
    Print #RptHandle, "Loc  "; QPTrim$(FAItemRec.ITEMLOC);
    Print #RptHandle, Tab(40); "Dept "; FAItemRec.IDEPT
    Print #RptHandle, "Orig Cost: "; Using("$$######,#.##", FAItemRec.ORGCOST);

    Print #RptHandle, Tab(40); "Acquired On "; MakeRegDate(FAItemRec.AQURDATE)
    Print #RptHandle, "Asset Life: "; Using("###", FAItemRec.ILIFE);
    Print #RptHandle, Tab(40); "Depreciation To-Date "; Using("$$######,#.##", FAItemRec.DEP2DATE)

    Print #RptHandle, "Status: "; FAItemRec.ISTATUS;
    Print #RptHandle, Tab(40); "G/L Dep Exp Acct "; FAItemRec.GLACCT
    Print #RptHandle, "Asset Group: "; FAItemRec.ASSETCODE
    Print #RptHandle, String$(79, "-")

    'SubTotal Here
    LineCnt = LineCnt + 8
    ItemCnt = ItemCnt + 1
    OrigCost# = OrigCost# + FAItemRec.ORGCOST
    DepTot# = DepTot# + FAItemRec.DEP2DATE

SkipEm:
    frmFAShowPctComp.ShowPctComp PCnt, PNumOFFARecs
    If frmFAShowPctComp.Out = True Then
      Close
      frmFAShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdExit.Enabled = True
      Me.cmdPrint.Enabled = True
      GoTo ExitRpt
    End If
ExitRpt:
  Next Cnt
  
  Unload frmFAShowPctComp
  GoSub PrintMasterValueEnding
  
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close FAHandle
  Close         'Close all open files now

  
  ViewPrint ReportFile$, "Master Asset Listing", False

  Kill ReportFile$

  Exit Sub

PrintMasterHeader:
  Page = Page + 1
  Print #RptHandle, Tab(29); "Master Asset Listing"
  Print #RptHandle, "Dept # "; Dept$
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle, Dash80$
  LineCnt = 6
Return

PrintMasterValueEnding:
  Print #RptHandle, Dash80$
  Print #RptHandle, "Dept Number: "; Dept$
  Print #RptHandle, "Total Items Printed: "; Using("#####", ItemCnt)
  Print #RptHandle, "Total Original Cost: "; Using("$$########,#.##", OrigCost#)
  Print #RptHandle, "Total Deprc to Date: "; Using("$$########,#.##", DepTot#)
  Print #RptHandle, FF$
Return

GetIndex:
  TagFlag = False
  If QPTrim$(fpcmbOrder.Text) = "TAG NUMBER" Then TagFlag = True
  ReDim Arr(1 To NumOfFARecs) As Struct
  For Cnt = 1 To NumOfFARecs
    PCnt = PCnt + 1
    Get FAHandle, Cnt, FAItemRec
    Arr(Cnt).RecNum = Cnt
    If TagFlag Then
      Arr(Cnt).who = QPTrim$(UCase$(FAItemRec.ITEMTAG))
    Else
      RSet Arr(Cnt).who = QPTrim$(UCase$(FAItemRec.IDEPT)) ' + RTrim$(FAItemRec.ASSETCODE))
    End If
  frmFAShowPctComp.ShowPctComp PCnt, PNumOFFARecs
  Next Cnt

  If TagFlag = True Then
    Call SortTagNums(Arr(), NumOfFARecs, PCnt, PNumOFFARecs)
'    Call SortTagNums(Arr(), NumOfFARecs, 1, 1)
  Else
'    Call SortAssetCodes(Arr(), NumOfFARecs, PCnt, PNumOFFARecs, True)
    Call SortAssetCodes(Arr(), NumOfFARecs, 1, 1, False)
    Unload frmFAShowPctComp
  End If
Return
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAMasterItemListing", "FixedAssets.exe", Erl)
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
    Me.Visible = False
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
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
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
      Call UnloadAllFormsAndOpn
'      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAMasterItemListing.")
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  
  fpcmbOrder.Text = "TAG NUMBER"
  fpcmbOrder.AddItem "TAG NUMBER"
  fpcmbOrder.AddItem "DEPARTMENT NUMBER"
  fpcmbYN.Text = "N"
  fpcmbYN.AddItem "Y"
  fpcmbYN.AddItem "N"
  fptxtDeptNum.Text = "ALL"
  fptxtYear.Enabled = False
  
End Sub

Private Sub fpcmbOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbOrder.ListIndex = -1
  End If
  If fpcmbOrder.ListDown <> True Then
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

Private Sub fpcmbYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYN.ListIndex = -1
  End If
  If fpcmbYN.ListDown <> True Then
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

Private Sub fptxtDeptNum_DblClick(Button As Integer)
  Dim This$
  This$ = Clipboard.GetText
  If This$ = "" Then Exit Sub
  fptxtDeptNum = Clipboard.GetText
  Clipboard.Clear

End Sub

