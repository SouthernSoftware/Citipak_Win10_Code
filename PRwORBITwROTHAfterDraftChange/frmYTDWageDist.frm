VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmYTDWageDist 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "YTD Wage Distribution Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmYTDWageDist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   4332
      Left            =   2112
      TabIndex        =   2
      Top             =   2256
      Width           =   7404
      _Version        =   196609
      _ExtentX        =   13060
      _ExtentY        =   7641
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
      FrameThreeDShadowColor=   -2147483633
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmYTDWageDist.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3405
         TabIndex        =   1
         Top             =   2445
         Width           =   2370
         _Version        =   196608
         _ExtentX        =   4180
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
         ColDesigner     =   "frmYTDWageDist.frx":08E6
      End
      Begin EditLib.fpDateTime fptxtYear 
         Height          =   396
         Left            =   4608
         TabIndex        =   0
         Top             =   1728
         Width           =   1116
         _Version        =   196608
         _ExtentX        =   1968
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
         ButtonStyle     =   2
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   1
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
         Text            =   "2002"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "YYYY"
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4368
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to generate the Year To Date Wage Distribution report."
         Top             =   3168
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
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
         ButtonDesigner  =   "frmYTDWageDist.frx":0BDD
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1296
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   3168
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
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
         ButtonDesigner  =   "frmYTDWageDist.frx":0DBC
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
         Left            =   1584
         TabIndex        =   5
         Top             =   2544
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "YTD Wage Distribution Report"
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
         Height          =   444
         Left            =   1584
         TabIndex        =   4
         Top             =   720
         Width           =   4332
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Summary Using Year:"
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
         Left            =   1584
         TabIndex        =   3
         Top             =   1824
         Width           =   2604
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   732
         Left            =   1392
         Top             =   528
         Width           =   4716
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   4632
      Left            =   1968
      Top             =   2100
      Width           =   7728
   End
End
Attribute VB_Name = "frmYTDWageDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdEscape_Click()
   frmReportsProcessing.Show
   DoEvents
   Unload frmYTDWageDist
End Sub

Private Sub cmdProcess_Click()
  If fpcomboPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
    Exit Sub
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    Call PrintText
  Else
    Exit Sub
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdEscape_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Dim Today As String * 11
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
'  Date$ = FormatDateTime(Date, vbShortDate)
  Today = Date '$
  fptxtYear.Text = Mid(Today, 7, 4)
  Me.HelpContextID = hlpYearToDateWage
  fpcomboPrintOpt.AddItem "Graphical"
  fpcomboPrintOpt.AddItem "Text"
  fpcomboPrintOpt.Text = "Graphical"
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub PrintGraphics()
  Dim ExitFlag As Boolean
  Dim Year$
  Dim LowDate As Long, HiDate As Long
  Dim RptName$, RptTitle$, SysFileName As Integer
  Dim UnitFileName As Integer, FundNumLen As Long
  Dim FundYTDLen As Long, FundCnt As Long
  Dim FirstFund As Boolean, Image1$
  Dim MaxLines As Integer, Image2$, Image3$
  Dim TRecSize&, NumOfRecs&, RHandle As Integer
  Dim THandle As Integer, DHandle As Integer
  Dim TCol As Integer, PctRow As Integer
  Dim cnt As Long, TAcct$, Cnt2 As Integer
  Dim DashPos As Integer, TFundNum&
  Dim TotalsFlag As Boolean, RegTot#, OTTot#
  Dim RegWageTot#, OTWageTot#, MonthCnt As Integer
  Dim LineCnt As Integer, FF$, FundPointer&
  Dim NewFund As Boolean, Fcnt As Long
  Dim Page As Integer, UTemp$, NumPrinted As Integer
  Dim Dash As String * 80, MonthNum As Integer
  Dim Temp As YTDFundRptType
  Dim bigNo As Long, y As Integer
  Dim smallNo As Long, idx As Integer
  Dim Number As String, NextCnt As Integer
  Dim NumberT As String, Month$, Trip As Integer
  Dim dlm$, FundCntT As Long, Trip2 As Integer
  Dim RptEndName$, EndHandle As Integer
  Dim ThisCnt As Integer
  
  dlm$ = "~"
  If fptxtYear.Text = "" Then
     MsgBox "Please enter a Year"
     fptxtYear.SetFocus
     Exit Sub
  End If

  If Val(fptxtYear.Text) < 1920 Or Val(fptxtYear.Text) > 2099 Then
     MsgBox "Please enter a valid Year (####)"
     fptxtYear.SetFocus
     Exit Sub
  End If
  
  ExitFlag = False
  Year$ = QPTrim$(fptxtYear.Text)
  LowDate = Date2Num("01-01-" + Year$)
  HiDate = Date2Num("12-31-" + Year$)

  RptName$ = "PRRPTS\YTDWAGEG.RPT"
  RptTitle$ = "YTD Wage Distribution Report."

  RptEndName$ = "PRRPTS\YTDWAGETOTAL.RPT"
  
  ReDim THRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  
  ReDim FundYTD(1 To 1) As YTDFundRptType
  ReDim ToDisk(1) As String * 78
  ReDim SysRec(1) As RegDSysFileRecType
  
  OpenSysFile SysFileName
  Get SysFileName, 1, SysRec(1)
  Close SysFileName
  OpenUnitFile UnitFileName
  Get UnitFileName, 1, Unit(1)
  Close UnitFileName

  FundNumLen = SysRec(1).AcctCnt
  FundYTDLen = Len(FundYTD(1))
  FundCnt = 0
  FirstFund = True
  Image1$ = "####0.00"
  Image2$ = "#####0.00"
  Image3$ = "######0.00"

  TRecSize = Len(THRec(1))

  RHandle = FreeFile
  On Error GoTo ErrorHandler
  Open RptName$ For Output As RHandle
  
  EndHandle = FreeFile
  Open RptEndName$ For Output As EndHandle
  
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TRecSize
  
  NumOfRecs& = LOF(THandle) \ Len(THRec(1))
  If NumOfRecs& = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  ReDim FundList(1 To NumOfRecs&) As Long
  
  OpenEmpData2File DHandle

  TCol = 40 - (Len(RptTitle$) \ 2) + 5
  PctRow = 11
  FrmShowPctComp.Label1 = "YTD Wage Distribution Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False

  For cnt = 1 To NumOfRecs&
    Get THandle, CLng(cnt), THRec(1)
    If (THRec(1).CheckDate >= LowDate) And (THRec(1).CheckDate <= HiDate) Then
      For Cnt2 = 1 To 8  'eight possiable distributions
        TAcct$ = THRec(1).TDist(Cnt2).DAcct
        Do
          DashPos = InStr(TAcct$, "-")
          If DashPos Then
            TAcct$ = Left$(TAcct$, DashPos - 1) + Mid$(TAcct$, DashPos + 1)
          End If
        Loop While DashPos
        If Len(QPTrim$(TAcct$)) > 0 Then
          TFundNum& = Val(QPTrim$(Left$(TAcct$, FundNumLen)))
          GoSub Parse2Fund
        End If
      Next
    End If

    FrmShowPctComp.ShowPctComp cnt, NumOfRecs&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next

  Close THandle
  'sort by fund number
  bigNo = 0
  For cnt = 1 To FundCnt
     If FundYTD(cnt).FundNum > bigNo Then
        bigNo = FundYTD(cnt).FundNum
     End If
  Next cnt
  smallNo = bigNo + 1
  y = 0
  y = 1
  If FundCnt = 0 Then
    Me.cmdEscape.Enabled = True
    Me.cmdProcess.Enabled = True
    EnableCloseButton Me.hwnd, True
    GoTo NoFundsSkip
  End If
  Do
    For cnt = y To FundCnt
    If FundYTD(cnt).FundNum < smallNo Then
       smallNo = FundYTD(cnt).FundNum
       idx = cnt
    End If
    Next cnt
    If y = FundCnt Then Exit Do
    Temp = FundYTD(y)
    FundYTD(y) = FundYTD(idx)
    FundYTD(idx) = Temp
    y = y + 1
    smallNo = bigNo + 1
  Loop
NoFundsSkip:

  For cnt = 1 To FundCnt
    GoSub PrintFundInfo
  Next
  TotalsFlag = True
  For cnt = 1 To FundCntT
    RegTot# = 0
    OTTot# = 0
    RegWageTot# = 0
    OTWageTot# = 0
    For MonthCnt = 1 To 12
      RegTot# = OldRound(RegTot# + FundYTD(cnt).Mths(MonthCnt).RegHrs)
      RegWageTot# = OldRound(RegWageTot# + FundYTD(cnt).Mths(MonthCnt).RegWage)
      OTTot# = OldRound(OTTot# + FundYTD(cnt).Mths(MonthCnt).OTHrs)
      OTWageTot# = OldRound(OTWageTot# + FundYTD(cnt).Mths(MonthCnt).OTWage)
    Next
    NumberT = Val(FundYTD(cnt).FundNum)
    ThisCnt = ThisCnt + 1
    '                   0                     1                          2
    Print #EndHandle, NumberT; dlm; Using(Image2$, RegTot#); dlm; Using(Image2$, OTTot#); dlm;
    '                           3                                  4
    Print #EndHandle, Using(Image3$, RegWageTot#); dlm; Using(Image3$, OTWageTot#)

  Next
  Close RHandle
  Close DHandle
  Close EndHandle
  
  If ThisCnt = 0 Then
    MsgBox "There are no records for the parameters entered."
    Exit Sub
  End If

  arYTDWageRpt.Show
  frmLoadingRpt.Show
  MainLog ("YTD Wage Distribution report processed.")
Exit Sub

Parse2Fund:

  FundPointer = 0
  NewFund = True     'assume this is a new fund

  If FirstFund Then       'if this is the first fund processed
    FirstFund = False     'skip search for fund part
  Else
    For Fcnt = 1 To FundCnt              'look through the fund list and see
      If FundList(Fcnt) = TFundNum& Then 'if this fund is already in the list
        FundPointer = Fcnt               'it is point to this fund in array
        NewFund = False                  'Not a new fund
        Exit For                         '
      End If
    Next
  End If

  If NewFund Then                    'if this fund wasn't found in the list
    FundCnt = FundCnt + 1
    FundCntT = FundCnt 'total funds count
    FundCnt4Rpt = FundCnt 'global
    FundPointer = FundCnt            'set the array pointer to the new entry
    ReDim Preserve FundYTD(1 To FundCnt) As YTDFundRptType
                                     'resize the YTD Fund array
    FundList(FundCnt) = TFundNum&    'add this fund to the fund list
    FundYTD(FundCnt).FundNum = TFundNum& 'set the new fund into the array
  End If

  MonthNum = Left$(MakeRegDate(THRec(1).CheckDate), 2)
  'get the month of this transaction

  FundYTD(FundPointer).Mths(MonthNum).RegHrs = OldRound(FundYTD(FundPointer).Mths(MonthNum).RegHrs + THRec(1).TDist(Cnt2).DRHrs)
  FundYTD(FundPointer).Mths(MonthNum).OTHrs = OldRound(FundYTD(FundPointer).Mths(MonthNum).OTHrs + THRec(1).TDist(Cnt2).DOHrs)
  FundYTD(FundPointer).Mths(MonthNum).RegWage = OldRound(FundYTD(FundPointer).Mths(MonthNum).RegWage + THRec(1).TDist(Cnt2).DRWage)
  FundYTD(FundPointer).Mths(MonthNum).OTWage = OldRound(FundYTD(FundPointer).Mths(MonthNum).OTWage + THRec(1).TDist(Cnt2).DOWage)
  'set trans data into the YTD Fund Array for the month

Return


PrintFundInfo:
    Number = Val(FundYTD(cnt).FundNum)

    NextCnt = 1
    Trip2 = 1
    For MonthCnt = 1 To 12
      Month = Mid("JanFebMarAprMayJunJulAugSepOctNovDec", NextCnt, 3)
      '                  0           1               2                    3                      4            5           6
      Print #RHandle, Month; dlm; FundCnt; dlm; TotalsFlag; dlm; QPTrim$(Unit(1).UFEMPR); dlm; Year$; dlm; Date$; dlm; Number; dlm;
      '                                 7                                                     8
      Print #RHandle, Using(Image1$, FundYTD(cnt).Mths(MonthCnt).RegHrs); dlm; Using(Image1$, FundYTD(cnt).Mths(MonthCnt).OTHrs); dlm;
      '                             9                                                         10
      Print #RHandle, Using(Image2$, FundYTD(cnt).Mths(MonthCnt).RegWage); dlm; Using(Image2$, FundYTD(cnt).Mths(MonthCnt).OTWage); dlm;
      '                  11                    12                             13
      Print #RHandle, NumberT; dlm; Using(Image2$, RegTot#); dlm; Using(Image2$, OTTot#); dlm;
      '                           14                                  15                    16              17
      Print #RHandle, Using(Image3$, RegWageTot#); dlm; Using(Image3$, OTWageTot#); dlm; "Fund No."; dlm; "Month"; dlm;
      '                18
      Print #RHandle, Trip; dlm; Trip2
      
      NextCnt = NextCnt + 3
    Next
    Trip = Trip + 1
    Me.cmdEscape.Enabled = True
    Me.cmdProcess.Enabled = True
    EnableCloseButton Me.hwnd, True

Return

ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      Call Terminate
      MainLog ("Payroll.exe terminated via menu bar on frmYTDWageDist.")
      End
    End If
  End If
End Sub

Private Sub PrintText()
  Dim ExitFlag As Boolean
  Dim Year$
  Dim LowDate As Long, HiDate As Long
  Dim RptName$, RptTitle$, SysFileName As Integer
  Dim UnitFileName As Integer, FundNumLen As Long
  Dim FundYTDLen As Long, FundCnt As Long
  Dim FirstFund As Boolean, Image1$
  Dim MaxLines As Integer, Image2$, Image3$
  Dim TRecSize&, NumOfRecs&, RHandle As Integer
  Dim THandle As Integer, DHandle As Integer
  Dim TCol As Integer, PctRow As Integer
  Dim cnt As Long, TAcct$, Cnt2 As Integer
  Dim DashPos As Integer, TFundNum&
  Dim TotalsFlag As Boolean, RegTot#, OTTot#
  Dim RegWageTot#, OTWageTot#, MonthCnt As Integer
  Dim LineCnt As Integer, FF$, FundPointer&
  Dim NewFund As Boolean, Fcnt As Long
  Dim Page As Integer, UTemp$, NumPrinted As Integer
  Dim Dash As String * 80, MonthNum As Integer
  Dim Temp As YTDFundRptType
  Dim bigNo As Long, y As Integer
  Dim smallNo As Long, idx As Integer
  Dim Number As String
  Dim ThisCnt As Integer
  
  If fptxtYear.Text = "" Then
     MsgBox "Please enter a Year"
     fptxtYear.SetFocus
     Exit Sub
  End If

  If Val(fptxtYear.Text) < 1920 Or Val(fptxtYear.Text) > 2099 Then
     MsgBox "Please enter a valid Year (####)"
     fptxtYear.SetFocus
     Exit Sub
  End If
  
  ExitFlag = False
  FF$ = Chr$(12)
  Year$ = QPTrim$(fptxtYear.Text)
  LowDate = Date2Num("01-01-" + Year$)
  HiDate = Date2Num("12-31-" + Year$)

  RptName$ = "PRRPTS\YTDWAGE.RPT"
  RptTitle$ = "YTD Wage Distribution Report."

  ReDim THRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  
  ReDim FundYTD(1 To 1) As YTDFundRptType
  ReDim ToDisk(1) As String * 78
  ReDim SysRec(1) As RegDSysFileRecType
  
  OpenSysFile SysFileName
  Get SysFileName, 1, SysRec(1)
  Close SysFileName
  OpenUnitFile UnitFileName
  Get UnitFileName, 1, Unit(1)
  Close UnitFileName

  FundNumLen = SysRec(1).AcctCnt
  FundYTDLen = Len(FundYTD(1))
  FundCnt = 0
  FirstFund = True
  MaxLines = 59

  Dash = String$(80, "-")
  Image1$ = "####0.00"
  Image2$ = "#####0.00"
  Image3$ = "######0.00"

  TRecSize = Len(THRec(1))

  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  RPTSetupPRN 9, RHandle
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TRecSize
  
  NumOfRecs& = LOF(THandle) \ Len(THRec(1))
  If NumOfRecs& = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  ReDim FundList(1 To NumOfRecs&) As Long
  
  OpenEmpData2File DHandle

  TCol = 40 - (Len(RptTitle$) \ 2) + 5
  PctRow = 11
  FrmShowPctComp.Label1 = "YTD Wage Distribution Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False

  For cnt = 1 To NumOfRecs&
    Get THandle, CLng(cnt), THRec(1)
    If (THRec(1).CheckDate >= LowDate) And (THRec(1).CheckDate <= HiDate) Then
      For Cnt2 = 1 To 8  'eight possiable distributions
        TAcct$ = THRec(1).TDist(Cnt2).DAcct
        Do
          DashPos = InStr(TAcct$, "-")
          If DashPos Then
            TAcct$ = Left$(TAcct$, DashPos - 1) + Mid$(TAcct$, DashPos + 1)
          End If
        Loop While DashPos
        If Len(QPTrim$(TAcct$)) > 0 Then
          TFundNum& = Val(QPTrim$(Left$(TAcct$, FundNumLen)))
          GoSub Parse2Fund
        End If
      Next
    End If

    FrmShowPctComp.ShowPctComp cnt, NumOfRecs&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next

  Close THandle
  'sort by fund number
  bigNo = 0
  For cnt = 1 To FundCnt
     If FundYTD(cnt).FundNum > bigNo Then
        bigNo = FundYTD(cnt).FundNum
     End If
  Next cnt
  smallNo = bigNo + 1
  y = 0
  y = 1
  If FundCnt = 0 Then
    Me.cmdEscape.Enabled = True
    Me.cmdProcess.Enabled = True
    EnableCloseButton Me.hwnd, True
    GoTo NoFundsSkip
  End If
  Do
    For cnt = y To FundCnt
    If FundYTD(cnt).FundNum < smallNo Then
       smallNo = FundYTD(cnt).FundNum
       idx = cnt
    End If
    Next cnt
    If y = FundCnt Then Exit Do
    Temp = FundYTD(y)
    FundYTD(y) = FundYTD(idx)
    FundYTD(idx) = Temp
    y = y + 1
    smallNo = bigNo + 1
  Loop
NoFundsSkip:
  GoSub PrintRptHeader

  For cnt = 1 To FundCnt
    GoSub PrintFundInfo
  Next
  Print #RHandle, FF$
  TotalsFlag = True
  GoSub PrintRptHeader
  For cnt = 1 To FundCnt
    RegTot# = 0
    OTTot# = 0
    RegWageTot# = 0
    OTWageTot# = 0
    For MonthCnt = 1 To 12
      RegTot# = OldRound(RegTot# + FundYTD(cnt).Mths(MonthCnt).RegHrs)
      RegWageTot# = OldRound(RegWageTot# + FundYTD(cnt).Mths(MonthCnt).RegWage)
      OTTot# = OldRound(OTTot# + FundYTD(cnt).Mths(MonthCnt).OTHrs)
      OTWageTot# = OldRound(OTWageTot# + FundYTD(cnt).Mths(MonthCnt).OTWage)
    Next
    Number = Val(FundYTD(cnt).FundNum)
    LSet ToDisk(1) = "  " + Number  'FundYTD(cnt).FundNum
    Mid$(ToDisk(1), 19) = Using(Image2$, RegTot#)
    Mid$(ToDisk(1), 34) = Using(Image2$, OTTot#)
    Mid$(ToDisk(1), 48) = Using(Image3$, RegWageTot#)
    Mid$(ToDisk(1), 63) = Using(Image3$, OTWageTot#)
    Print #RHandle, ToDisk(1)
    LineCnt = LineCnt + 1
    If LineCnt > MaxLines Then
      Print #RHandle, FF$
      GoSub PrintRptHeader
    End If
  Next
  Print #RHandle, FF$
  RPTSetupPRN 123, RHandle '7/24
  Close RHandle
  Close DHandle
  
  If ThisCnt = 0 Then
    MsgBox "There are no records for the parameters entered."
    Exit Sub
  End If
  
  ViewPrint RptName$, RptTitle$
  MainLog ("YTD Wage Distribution report processed.")
Exit Sub

Parse2Fund:

  FundPointer = 0
  NewFund = True     'assume this is a new fund

  If FirstFund Then       'if this is the first fund processed
    FirstFund = False     'skip search for fund part
  Else
    For Fcnt = 1 To FundCnt              'look through the fund list and see
      If FundList(Fcnt) = TFundNum& Then 'if this fund is already in the list
        FundPointer = Fcnt               'it is point to this fund in array
        NewFund = False                  'Not a new fund
        Exit For                         '
      End If
    Next
  End If

  If NewFund Then                    'if this fund wasn't found in the list
    FundCnt = FundCnt + 1            'total funds count
    FundPointer = FundCnt            'set the array pointer to the new entry
    ReDim Preserve FundYTD(1 To FundCnt) As YTDFundRptType
                                     'resize the YTD Fund array
    FundList(FundCnt) = TFundNum&    'add this fund to the fund list
    FundYTD(FundCnt).FundNum = TFundNum& 'set the new fund into the array
  End If

  MonthNum = Left$(MakeRegDate(THRec(1).CheckDate), 2)
  'get the month of this transaction

  FundYTD(FundPointer).Mths(MonthNum).RegHrs = OldRound(FundYTD(FundPointer).Mths(MonthNum).RegHrs + THRec(1).TDist(Cnt2).DRHrs)
  FundYTD(FundPointer).Mths(MonthNum).OTHrs = OldRound(FundYTD(FundPointer).Mths(MonthNum).OTHrs + THRec(1).TDist(Cnt2).DOHrs)
  FundYTD(FundPointer).Mths(MonthNum).RegWage = OldRound(FundYTD(FundPointer).Mths(MonthNum).RegWage + THRec(1).TDist(Cnt2).DRWage)
  FundYTD(FundPointer).Mths(MonthNum).OTWage = OldRound(FundYTD(FundPointer).Mths(MonthNum).OTWage + THRec(1).TDist(Cnt2).DOWage)
  'set trans data into the YTD Fund Array for the month

Return

PrintRptHeader:
  Page = Page + 1
  RSet Pg(1) = Str$(Page)
  UTemp$ = Space$(71)
  LSet UTemp$ = QPTrim$(Unit(1).UFEMPR)
  Mid$(UTemp$, 62) = "Page:" + Pg(1)
  Print #RHandle, UTemp$
  Print #RHandle, "YTD Wage Distribution Report for Year: " + Year$ ' + CrLf$
  Print #RHandle, "Report Date: " + Date$
  If TotalsFlag Then
    Print #RHandle, "Totals"
    Print #RHandle, "Fund No.            Reg Hrs         OT Hrs      Reg Wages       OT Wages" '  + CrLf$
    LineCnt = 6
  Else
    Print #RHandle, "Fund No."
    Print #RHandle, "Month               Reg Hrs         OT Hrs      Reg Wages       OT Wages" ' + CrLf$
  End If
  Print #RHandle, Dash
Return

PrintFundInfo:
    NumPrinted = NumPrinted + 1
    If NumPrinted = 4 Then
      NumPrinted = 1
      Print #RHandle, FF$
      GoSub PrintRptHeader
    End If
    Number = Val(FundYTD(cnt).FundNum)
    ThisCnt = ThisCnt + 1
    Print #RHandle, "Fund: " + Number 'FundYTD(cnt).FundNum ' + CrLf$
    For MonthCnt = 1 To 12

      LSet ToDisk(1) = "  " + Mid$("JanFebMarAprMayJunJulAugSepOctNovDec", (((MonthCnt - 1) * 3)) + 1, 3)
      Mid$(ToDisk(1), 22) = Using(Image1$, FundYTD(cnt).Mths(MonthCnt).RegHrs)
      Mid$(ToDisk(1), 37) = Using(Image1$, FundYTD(cnt).Mths(MonthCnt).OTHrs)
      Mid$(ToDisk(1), 51) = Using(Image2$, FundYTD(cnt).Mths(MonthCnt).RegWage)
      Mid$(ToDisk(1), 66) = Using(Image2$, FundYTD(cnt).Mths(MonthCnt).OTWage)

      Print #RHandle, QPTrim$(ToDisk(1))
    Next
    Print #RHandle,

    Me.cmdEscape.Enabled = True
    Me.cmdProcess.Enabled = True
    EnableCloseButton Me.hwnd, True

Return

End Sub

Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboPrintOpt.ListIndex = -1
  End If
  If fpcomboPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdEscape.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

