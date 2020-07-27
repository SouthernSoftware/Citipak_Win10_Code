VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFAPrintDsplList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets Print Disposal List"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAPrintDsplList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4764
      Left            =   1956
      TabIndex        =   2
      Top             =   2064
      Width           =   7740
      _Version        =   196609
      _ExtentX        =   13652
      _ExtentY        =   8403
      _StockProps     =   70
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
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmFAPrintDsplList.frx":08CA
      Begin LpLib.fpCombo fpListDates 
         Height          =   405
         Left            =   3210
         TabIndex        =   0
         Top             =   1875
         Width           =   3240
         _Version        =   196608
         _ExtentX        =   5715
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
         AutoSearchFill  =   -1  'True
         AutoSearchFillDelay=   200
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmFAPrintDsplList.frx":08E6
      End
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3510
         TabIndex        =   1
         ToolTipText     =   "Select Graphical for a more robust but slower processing report. Select Text for a quick report."
         Top             =   2550
         Width           =   2355
         _Version        =   196608
         _ExtentX        =   4154
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
         Object.TabStop         =   -1  'True
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
         MaxEditLen      =   150
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
         AutoSearchFill  =   -1  'True
         AutoSearchFillDelay=   200
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmFAPrintDsplList.frx":0BDD
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   675
         Left            =   1590
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to create the desired report."
         Top             =   3510
         Width           =   1875
         _Version        =   131072
         _ExtentX        =   3307
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
         ButtonDesigner  =   "frmFAPrintDsplList.frx":0ED4
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
         Height          =   675
         Left            =   4410
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to create the report based on the parameters entered above."
         Top             =   3510
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
         ButtonDesigner  =   "frmFAPrintDsplList.frx":10B0
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disposal Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1152
         TabIndex        =   5
         Top             =   1968
         Width           =   1836
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1488
         Top             =   576
         Width           =   4908
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Print Disposal List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   492
         Left            =   1584
         TabIndex        =   4
         Top             =   720
         Width           =   4812
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Print Option:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   1824
         TabIndex        =   3
         Top             =   2628
         Width           =   1500
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   4956
      Left            =   1860
      Top             =   1956
      Width           =   7932
   End
End
Attribute VB_Name = "frmFAPrintDsplList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmFADisposalMenu.Show
  DoEvents
  Unload frmFAPrintDsplList
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
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdPrint_Click
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
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAPrintDsplList.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub cmdPrint_Click()
  Select Case QPTrim$(fpcomboPrintOpt.Text)
    Case "Graphical"
      Call PrintGraphics
    Case "Text"
      MsgBox "Pitch 12 recommended for this report."
      Call PrintText
    Case "Exit"
  End Select
End Sub

Private Sub LoadMe()
  Dim DateRec As TempDisposedOfDate
  Dim GHandle As Integer
  Dim DateCnt As Long
  Dim x As Long
  Dim BigNum As Long
  Dim SmallNum As Long
  Dim StopSpot As Integer
  Dim HoldSpot As Integer
  Dim Y As Long
  Dim Nextx As Long
  Dim ThisDate As Integer
  
  OpenTempDisposedDate GHandle
  DateCnt = LOF(GHandle) / Len(DateRec) 'if there
  'were no valid dates it would have been detected
  'and the user alerted when this form was accessed
  'from the menu
  
  fpListDates.Clear 'start with a clean list
  
  ReDim OrderDate(1 To DateCnt) As Integer
  'we want to display the dates from earliest to latest
  'so sort dates here
  BigNum = 0
  For x = 1 To DateCnt
    Get GHandle, x, DateRec
    If DateRec.DsplDate = 0 Then GoTo DateDeleted
    Y = Y + 1
    OrderDate(x) = DateRec.DsplDate
    If DateRec.DsplDate > BigNum Then
      BigNum = DateRec.DsplDate
    End If
DateDeleted:
  Next x
  Close GHandle
  
  If Y = 0 Then
    KillFile ("FATEMPDISPDATE.DAT")
    GoTo NoMoreDates
  End If
  
  Nextx = 1
  BigNum = BigNum + 1
  SmallNum = BigNum
  Do
    For x = Nextx To DateCnt
      If OrderDate(x) < SmallNum Then
        SmallNum = OrderDate(x)
        StopSpot = x
      End If
    Next x
    HoldSpot = OrderDate(Nextx)
    OrderDate(Nextx) = SmallNum
    OrderDate(StopSpot) = HoldSpot
    If Nextx = DateCnt Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum
  Loop
  
  For x = 1 To DateCnt
    If OrderDate(x) = 0 Then GoTo DateIsZero
    fpListDates.AddItem (MakeRegDate(OrderDate(x)))
DateIsZero:
  Next x
  
  For x = 1 To DateCnt
    If OrderDate(x) = 0 Then GoTo NODate
    fpListDates.Text = MakeRegDate(OrderDate(x))
    Exit For
NODate:
  Next x
  
NoMoreDates:
  fpListDates.Action = ActionSelectAll
  
  fpcomboPrintOpt.AddItem "Graphical"
  fpcomboPrintOpt.AddItem "Text"
  fpcomboPrintOpt.Text = "Graphical"

End Sub
Private Sub PrintText()
  Dim ReportFile$
  Dim DateRec As TempDisposedOfDate
  Dim DHandle As Integer
  Dim x As Long, FF$
  Dim DateCnt As Long
  Dim ThisDate As Integer
  Dim RptHandle As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim FASetUpRec As FASetupRecType
  Dim FASHandle As Integer
  Dim FAHandle As Integer
  Dim FAItemRec As FAItemRecType
  Dim Employer$, Page As Integer
  Dim Nextx As Integer
  Dim TagRec As TagNumbSortIdxType
  Dim TagHandle As Integer
  Dim TagCnt As Long
  
  On Error GoTo ERRORSTUFF
  
  If QPTrim$(fpListDates.Text) = "" Then
    MsgBox "Please make sure a valid date is entered in the Disposal Date field."
    fpListDates.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpListDates.Text) <> "" Then
    ThisDate = Date2Num(fpListDates.Text)
  Else
    MsgBox "No date saved."
    Close
    Exit Sub
  End If
  
  If Exist("FATEMPDISPDATE.DAT") Then
    OpenTempDisposedDate DHandle
    DateCnt = LOF(DHandle) / Len(DateRec)
    If DateCnt = 0 Then
      MsgBox "No disposal dates have been saved."
      fpListDates.SetFocus
      Close DHandle
      Exit Sub
    End If
  End If
  
  For x = 1 To DateCnt
    Get DHandle, x, DateRec
    If DateRec.DsplDate = ThisDate Then 'validate the date entered
      Close DHandle
      Exit For
    End If
  Next x
  
  'x will be greater than DateCnt only if no match was found in the
  'for loop above
  If x > DateCnt Then
    MsgBox "Nothing is saved for this date. Please choose another date or save data for this date."
    fpListDates.SetFocus
    Close DHandle
    Exit Sub
  End If
  
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  FF$ = Chr$(12)
  
  Employer = QPTrim$(FASetUpRec.TownName)
  
  MaxLines = 57
  ReportFile$ = "FATEMPDSPLPRINT.PRT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  GoSub PrintHeader
  
  frmFAShowPctComp.Label1 = "Gathering Disposed Of Item Data"
  frmFAShowPctComp.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdPrint.Enabled = False
  
  OpenTagIdxFile TagHandle
  TagCnt = LOF(TagHandle) / Len(TagRec)
  ReDim TagRecNum(1 To TagCnt)
  For x = 1 To TagCnt
    Get TagHandle, x, TagRec
    TagRecNum(x) = TagRec.DataRecNum 'load array with record pointers arranged by
    'tag numerical order
  Next x
  Close TagHandle
  
  OpenFAItemFile FAHandle
  
  For x = 1 To TagCnt
    Get FAHandle, TagRecNum(x), FAItemRec
    If FAItemRec.DispDate = ThisDate And FAItemRec.DsplFlag = 1 Then
      Print #RptHandle, FAItemRec.ItemTag; Tab(21); FAItemRec.IDESC1; Tab(53); FAItemRec.IDEPT; Tab(58); Using$("$##,###,##0.00", FAItemRec.ORGCOST); Tab(73); Using$("$##,###,##0.00", FAItemRec.CURRVAL)
      Print #RptHandle,
      Print #RptHandle, "Disposal Amount __________________ "; Tab(41); "Method: AUCTION __ SALVAGE __ SOLD__ OTHER __"
      Print #RptHandle, String$(86, "-")
      LineCnt = LineCnt + 4
    End If
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
  
    frmFAShowPctComp.ShowPctComp x, TagCnt
    If frmFAShowPctComp.Out = True Then
      Close
      frmFAShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdExit.Enabled = True
      Me.cmdPrint.Enabled = True
      Unload frmFAShowPctComp
      Exit Sub
    End If
  Next x
  
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdPrint.Enabled = True
  Unload frmFAShowPctComp
  Print #RptHandle, FF$
  Close RptHandle
  ViewPrint ReportFile$, "Master Asset Disposed Of Listing", False
  KillFile (ReportFile$)
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(27); "Fixed Asset List of Items For Disposal"
  Print #RptHandle,
  Print #RptHandle, Employer; Tab(77); "Page "; Tab(83); Page
  Print #RptHandle, "Item Disposal Date: "; Tab(22); fpListDates.Text
  Print #RptHandle,
  Print #RptHandle, Tab(1); "Tag Number"; Tab(25); "Description"; Tab(53); "Dept"; Tab(58); "Purchase Price"; Tab(74); "Current Value"
  Print #RptHandle, String$(86, "=")
  LineCnt = 7
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAPrintDsplList", "PrintText", Erl)
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

Private Sub PrintGraphics()
  Dim ReportFile$
  Dim DateRec As TempDisposedOfDate
  Dim DHandle As Integer
  Dim x As Long
  Dim DateCnt As Long
  Dim ThisDate As Integer
  Dim RptHandle As Integer
  Dim FASetUpRec As FASetupRecType
  Dim FASHandle As Integer
  Dim FAHandle As Integer
  Dim FAItemRec As FAItemRecType
  Dim Employer$, Page As Integer
  Dim Nextx As Integer
  Dim TagRec As TagNumbSortIdxType
  Dim TagHandle As Integer
  Dim TagCnt As Long
  Dim dlm$
  
  On Error GoTo ERRORSTUFF
  
  dlm$ = "~"
  If QPTrim$(fpListDates.Text) <> "" Then
    ThisDate = Date2Num(fpListDates.Text)
  Else
    MsgBox "No date saved."
    Close
    Exit Sub
  End If
  
  If Exist("FATEMPDISPDATE.DAT") Then
    OpenTempDisposedDate DHandle
    DateCnt = LOF(DHandle) / Len(DateRec)
    If DateCnt = 0 Then 'nothing saved
      MsgBox "No disposal dates have been saved."
      fpListDates.SetFocus
      Close DHandle
      Exit Sub
    End If
  End If
  
  If QPTrim$(fpListDates.Text) = "" Then
    MsgBox "Please enter a valid date in the Disposal Date field."
    fpListDates.SetFocus
    Exit Sub
  End If
  
  For x = 1 To DateCnt
    Get DHandle, x, DateRec 'look for the selected disposal date in the records
    If DateRec.DsplDate = ThisDate Then
      Close DHandle
      Exit For
    End If
  Next x
  
  If x > DateCnt Then 'x will be greater than DateCnt if no match
  'was found in the for loop above
    MsgBox "Nothing is saved for this date. Please choose another date or save data for this date."
    fpListDates.SetFocus
    Close DHandle
    Exit Sub
  End If
  
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  
  Employer = QPTrim$(FASetUpRec.TownName)
  
  ReportFile$ = "FARPTS\FATEMPDSPLPRINT.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  frmFAShowPctComp.Label1 = "Gathering Disposed Of Item Data"
  frmFAShowPctComp.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdPrint.Enabled = False
  
  OpenTagIdxFile TagHandle
  TagCnt = LOF(TagHandle) / Len(TagRec)
  ReDim TagRecNum(1 To TagCnt)
  For x = 1 To TagCnt
    Get TagHandle, x, TagRec
    TagRecNum(x) = TagRec.DataRecNum 'load an array with record pointers
    'arranged in numerical order
  Next x
  Close TagHandle
  
  OpenFAItemFile FAHandle
  
  For x = 1 To TagCnt
    Get FAHandle, TagRecNum(x), FAItemRec
    If FAItemRec.DispDate = ThisDate And FAItemRec.DsplFlag = 1 Then
      '                     0                   1                      2                     3
      Print #RptHandle, Employer; dlm; FAItemRec.ItemTag; dlm; FAItemRec.IDESC1; dlm; FAItemRec.IDEPT; dlm;
      '                         4                       5                       6
      Print #RptHandle, FAItemRec.ORGCOST; dlm; FAItemRec.CURRVAL; dlm; fpListDates.Text; dlm;
      '                                    7                                             8
      Print #RptHandle, "Disposal Amount __________________ "; dlm; "Method: AUCTION __ SALVAGE __ SOLD__ OTHER __"
    End If
  
    frmFAShowPctComp.ShowPctComp x, TagCnt
    If frmFAShowPctComp.Out = True Then
      Close
      frmFAShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdExit.Enabled = True
      Me.cmdPrint.Enabled = True
      Unload frmFAShowPctComp
      Exit Sub
    End If
  Next x
  
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdPrint.Enabled = True
  Unload frmFAShowPctComp
  
  Close RptHandle
  
  arFAItemsForDsplList.Show
  frmFALoadReport.Show
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAPrintDsplList", "PrintGraphics", Erl)
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

Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcomboPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboPrintOpt.ListIndex = -1
  End If
  If fpcomboPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdExit.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpListDates_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpListDates.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpListDates.ListIndex = -1
  End If
  If fpListDates.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcomboPrintOpt.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

