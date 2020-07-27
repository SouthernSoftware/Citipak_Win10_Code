VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFAEditDisposedOf 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Disposed Of List"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAEditDisposedOf.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpList fplistMethods 
      Height          =   1680
      Left            =   6210
      TabIndex        =   6
      ToolTipText     =   $"frmFAEditDisposedOf.frx":08CA
      Top             =   2130
      Width           =   2745
      _Version        =   196608
      _ExtentX        =   4842
      _ExtentY        =   2963
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      Columns         =   0
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
      ColDesigner     =   "frmFAEditDisposedOf.frx":095B
   End
   Begin LpLib.fpList fpListDates 
      Height          =   1680
      Left            =   2610
      TabIndex        =   3
      ToolTipText     =   "Click on a date to bring up the fixed assets designated for disposal on that date."
      Top             =   2160
      Width           =   2745
      _Version        =   196608
      _ExtentX        =   4842
      _ExtentY        =   2963
      TextAlias       =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      Columns         =   0
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
      ColDesigner     =   "frmFAEditDisposedOf.frx":0BE7
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4305
      Left            =   285
      TabIndex        =   7
      Top             =   4200
      Width           =   11250
      _Version        =   196613
      _ExtentX        =   19844
      _ExtentY        =   7594
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      ShadowColor     =   13684944
      SpreadDesigner  =   "frmFAEditDisposedOf.frx":0E73
   End
   Begin EditLib.fpDateTime fpDateDisp 
      Height          =   396
      Left            =   5616
      TabIndex        =   0
      ToolTipText     =   "Items listed below are slated for disposal on this date."
      Top             =   1152
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
      Text            =   "2/28/2003"
      DateCalcMethod  =   0
      DateTimeFormat  =   0
      UserDefinedFormat=   ""
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
      PopUpType       =   0
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   690
      Left            =   9442
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1650
      _Version        =   131072
      _ExtentX        =   2910
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
      ButtonDesigner  =   "frmFAEditDisposedOf.frx":2824
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   690
      Left            =   562
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Click on this button to save all data displayed on the spreadsheet  for the active date."
      Top             =   3360
      Width           =   1650
      _Version        =   131072
      _ExtentX        =   2910
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
      ButtonDesigner  =   "frmFAEditDisposedOf.frx":2A00
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2310
      Left            =   2317
      Top             =   1725
      Width           =   7050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   5722
      X2              =   5722
      Y1              =   1720
      Y2              =   4000
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Assign Method To Entire List:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   6074
      TabIndex        =   5
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pending Disposal Dates:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   2624
      TabIndex        =   4
      Top             =   1830
      Width           =   2700
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Active Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   4128
      TabIndex        =   2
      Top             =   1248
      Width           =   1404
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   192
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed Asset Disposal Edit"
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
      Top             =   336
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   144
      Width           =   8652
   End
End
Attribute VB_Name = "frmFAEditDisposedOf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim DateOnScreen As Integer
  Dim DateChangeFlag As Boolean
  Dim ListCnt As Long
  Dim ThisRecNum() As Long
  Dim ActiveX As Long
  
Private Sub cmdExit_Click()
  frmFADisposalMenu.Show
  Close
  DoEvents
  Unload frmFAEditDisposedOf
  
End Sub

Private Sub cmdSave_Click()
  Dim x As Integer
  Dim Nextx As Long
  Dim PHandle As Integer
  Dim DsplRec As PrePostDsplType
  
  On Error GoTo ERRORSTUFF
  If ActiveX = 0 Then
    MsgBox "There are no fixed assets to be saved for this date"
    Exit Sub
  End If
  
  'this for loop moves through the spreadsheet making sure that
  'the editable values have valid values
  For x = 1 To ActiveX
    vaSpread1.Col = 5
    vaSpread1.Row = x
    If QPTrim$(vaSpread1.Text) = "$0.00" Then
      vaSpread1.Col = 2
      vaSpread1.Row = x
      If MsgBox("A disposal value of $0.00 is assigned to " + vaSpread1.Text + " on row " + CStr(x) + ". Do you wish to return and edit before saving.?", vbYesNo) = vbYes Then
        Close 'FAHandle
        vaSpread1.SetFocus
        vaSpread1.SetActiveCell 5, x
        Exit Sub
      End If
    End If
    vaSpread1.Col = 6
    vaSpread1.Row = x
    If QPTrim$(vaSpread1.Text) = "" Then
      vaSpread1.Col = 2
      vaSpread1.Row = x
      If MsgBox("No disposal method has been assigned to " + vaSpread1.Text + " on row " + CStr(x) + ". Do you wish to return and edit before saving.?", vbYesNo) = vbYes Then
        Close
        vaSpread1.SetFocus
        vaSpread1.SetActiveCell 6, x
        Exit Sub
      End If
    End If
  Next x
  'data has been examined and has passed with no problems or
  'the user has been alerted to problems and he has OK'd the
  'data
  'destroy existing file because it may contain any number of
  'unneeded deletd files that tend to cause problems with populating
  'a spreadsheet because you always have to filter out any deleted items
  If Exist(PrepostDsplName + CStr(Date2Num(fpDateDisp)) + ".DAT") Then
    KillFile PrepostDsplName + CStr(Date2Num(fpDateDisp)) + ".DAT"
  End If
  'start a brand new edit file for this date
  OpenPrePostDsplData PHandle, Date2Num(fpDateDisp)
  
  For x = 1 To ActiveX
      DsplRec.ThisRec = ThisRecNum(x)
      vaSpread1.Col = 5
      vaSpread1.Row = x
      DsplRec.DisposAmt = CDbl(vaSpread1.Text)
      vaSpread1.Col = 6
      vaSpread1.Row = x
      DsplRec.DsplMethod = QPTrim$(vaSpread1.Text)
    Put PHandle, x, DsplRec
  Next x
  Close
  MsgBox "Your data has been saved."
  Call LoadMe
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditDisposedOf", "cmdSave_Click", Erl)
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
  Call FixSpread
  fpDateDisp.Text = Date 'Loadme is called thru fpDateDisp Change
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
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAEditDisposedOf.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim x As Double
  Dim Nextx As Integer
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim TagIdxCnt As Long
  Dim ThisDate As Integer
  Dim DateRec As TempDisposedOfDate
  Dim GHandle As Integer
  Dim DateCnt As Integer
  Dim Y As Integer
  Dim BigNum As Integer
  Dim SmallNum
  Dim StopSpot As Integer
  Dim HoldSpot As Integer
  Dim Method$
  Dim StrDate$
  Dim PHandle As Integer
  Dim DsplRec As PrePostDsplType
  Dim DsplCnt As Long
  
  StrDate = CStr(Date2Num(fpDateDisp))
  
  On Error GoTo ERRORSTUFF
  
  ActiveX = 0
  
  OpenTagIdxFile TagIdxHandle
  TagIdxCnt = LOF(TagIdxHandle) / Len(TagIdx)
  
  GoSub LoadDates
  
  ReDim TagRecNum(1 To TagIdxCnt)
  For x = 1 To TagIdxCnt
    Get TagIdxHandle, x, TagIdx
    TagRecNum(x) = TagIdx.DataRecNum
  Next x
  
  OpenFAItemFile FAHandle
  Nextx = 1
  ReDim ThisRecNum(1 To 1) As Long
  If DateChangeFlag = True Then
    For x = 1 To TagIdxCnt 'clear the spreadsheet in preparation
    'for loading new disposal date data ...could have used a spreadsheet
    'clear function but would have had to go back and unlock the last
    'two columns anyway so went with clearing one line at a time
      vaSpread1.Col = 1
      vaSpread1.Row = x
      vaSpread1.Text = ""
      vaSpread1.Col = 2
      vaSpread1.Row = x
      vaSpread1.Text = ""
      vaSpread1.Col = 3
      vaSpread1.Row = x
      vaSpread1.Text = ""
      vaSpread1.Col = 4
      vaSpread1.Row = x
      vaSpread1.Text = ""
      vaSpread1.Col = 5
      vaSpread1.Row = x
      vaSpread1.Lock = False
      vaSpread1.Text = ""
      vaSpread1.Col = 6
      vaSpread1.Row = x
      vaSpread1.Lock = False
      vaSpread1.Text = ""
    Next x
  End If
  
  If Exist(PrepostDsplName + StrDate + ".DAT") Then 'edit has already taken place
    Nextx = 1
    OpenPrePostDsplData PHandle, Date2Num(fpDateDisp)
    DsplCnt = LOF(PHandle) / Len(DsplRec)
    For x = 1 To DsplCnt
      Get PHandle, x, DsplRec
      If DsplRec.Deleted = True Then GoTo SkipIt 'deleted is set in build disposal list
      Get FAHandle, DsplRec.ThisRec, FAItemRec
        ActiveX = ActiveX + 1
        ReDim Preserve ThisRecNum(1 To ActiveX) As Long
        ThisRecNum(ActiveX) = DsplRec.ThisRec
        vaSpread1.Col = 1
        vaSpread1.Row = Nextx
        vaSpread1.Text = QPTrim$(FAItemRec.ItemTag)
        vaSpread1.Col = 2
        vaSpread1.Row = Nextx
        vaSpread1.Text = QPTrim$(FAItemRec.IDESC1)
        vaSpread1.Col = 3
        vaSpread1.Row = Nextx
        vaSpread1.Text = FAItemRec.IDEPT
        vaSpread1.TypeHAlign = TypeHAlignCenter
        vaSpread1.Col = 4
        vaSpread1.Row = Nextx
        vaSpread1.Text = FAItemRec.ORGCOST
        vaSpread1.Col = 5
        vaSpread1.Row = Nextx
        vaSpread1.Text = DsplRec.DisposAmt
        vaSpread1.Col = 6
        vaSpread1.Row = Nextx
        vaSpread1.Text = QPTrim$(DsplRec.DsplMethod)
        vaSpread1.TypeHAlign = TypeHAlignCenter
        Nextx = Nextx + 1
SkipIt:
    Next x
  Else
    For x = 1 To TagIdxCnt
      Get FAHandle, TagRecNum(x), FAItemRec
      If FAItemRec.DispDate = 0 Then GoTo DateIsZero
      If FAItemRec.DispDate = DateOnScreen And FAItemRec.DsplFlag = 1 Then
        ActiveX = ActiveX + 1
        ReDim Preserve ThisRecNum(1 To ActiveX) As Long
        ThisRecNum(ActiveX) = TagRecNum(x)
        vaSpread1.Col = 1
        vaSpread1.Row = Nextx
        vaSpread1.Text = QPTrim$(FAItemRec.ItemTag)
        vaSpread1.Col = 2
        vaSpread1.Row = Nextx
        vaSpread1.Text = QPTrim$(FAItemRec.IDESC1)
        vaSpread1.Col = 3
        vaSpread1.Row = Nextx
        vaSpread1.Text = FAItemRec.IDEPT
        vaSpread1.TypeHAlign = TypeHAlignCenter
        vaSpread1.Col = 4
        vaSpread1.Row = Nextx
        vaSpread1.Text = FAItemRec.ORGCOST
        vaSpread1.Col = 5
        vaSpread1.Row = Nextx
        vaSpread1.Text = FAItemRec.DisposAmt
        vaSpread1.Col = 6
        vaSpread1.Row = Nextx
        vaSpread1.Text = QPTrim$(FAItemRec.DsplMethod)
        vaSpread1.TypeHAlign = TypeHAlignCenter
        Nextx = Nextx + 1
      End If
DateIsZero:
    Next x
  End If
  Close FAHandle
  Close TagIdxHandle
  Close PHandle
  
  fplistMethods.Action = ActionClear
  
  fplistMethods.AddItem ("Clear")
  fplistMethods.AddItem ("Auction")
  fplistMethods.AddItem ("Salvage")
  fplistMethods.AddItem ("Sold")
  fplistMethods.AddItem ("Other")
  
  ListCnt = ActiveX
  
  Exit Sub
   
LoadDates:
  OpenTempDisposedDate GHandle
  DateCnt = LOF(GHandle) / Len(DateRec)
  fpListDates.Clear
  If DateCnt = 0 Then
    fpListDates.AddItem ("NONE")
    Close GHandle
    GoTo NoMoreDates
  End If
  'we've got active dates so sort from earliest to latest
  ReDim OrderDate(1 To DateCnt) As Integer
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
  
  If Y = 0 Then 'FATEMPDISPDATE exists but it's full
  'of zero date records...no more valid dates
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
  
  For x = 1 To DateCnt 'dates exist but some may be zeros
    If OrderDate(x) = 0 Then GoTo NoDsplDate
    fpListDates.AddItem (MakeRegDate(OrderDate(x)))
NoDsplDate:
  Next x
    
  If DateOnScreen = 0 Then
    ThisDate = OrderDate(x - 1) 'if DateOnScreen is empty assign last valid date to ThisDate
    fpDateDisp.Text = MakeRegDate(ThisDate)
    DateOnScreen = ThisDate
  Else
    ThisDate = DateOnScreen
    fpDateDisp.Text = MakeRegDate(DateOnScreen)
  End If
  
NoMoreDates:
  Return
    
ERRORSTUFF:
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAEditDisposedOf", "LoadMe", Erl)
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

Private Sub fpDateDisp_Change()
  If QPTrim$(fpDateDisp.Text) = "" Then
    fpDateDisp.Text = MakeRegDate(DateOnScreen)
    Exit Sub
  End If
  
  'DateOnScreen maintains a constant anchor to a valid date
  DateOnScreen = Date2Num(fpDateDisp.Text)
  Close
  DateChangeFlag = True
  Call LoadMe
  
End Sub

Private Sub fpListDates_Click()
  fpDateDisp.Text = fpListDates.Text
  
End Sub

Private Sub fplistMethods_Click()
  Dim Method$
  Dim x As Long
  'here we assigning all listed fixed assets with one method...
  'saves time if there are lots of fixed assets and they are all
  'being disposed of at an event such as an auction, etc.
  Method = QPTrim$(fplistMethods.Text)
  
  For x = 1 To ListCnt
    vaSpread1.Col = 6
    vaSpread1.Row = x
    vaSpread1.Text = Method
    vaSpread1.TypeHAlign = TypeHAlignCenter
  Next x
  
End Sub

Private Sub FixSpread()
  Dim COne As Integer
  Dim CTwo As Integer
  Dim CThree As Integer
  Dim CFour As Integer
  Dim CFive As Integer
  Dim CSix As Integer
  Dim cnt As Integer
  '-1 means all rows or all columns....0 means headers
'    GoTo SkipAdjust
    Select Case ScreenW
      Case 1280
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 5
        coladj = 10
        vaSpread1.FontSize = 18
        vaSpread1.RowHeight(-1) = 22
        vaSpread1.RowHeight(0) = 22
      Else
        COne = 13
        coladj = 4.9
        vaSpread1.RowHeight(-1) = 18
        vaSpread1.RowHeight(0) = 18
      End If
      Case 1152
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 14
        coladj = 7
        vaSpread1.FontSize = 14
        vaSpread1.RowHeight(0) = 18.5
        vaSpread1.RowHeight(-1) = 18.5
      Else
        COne = 6.65
        coladj = 2.5
        vaSpread1.RowHeight(0) = 16
        vaSpread1.RowHeight(-1) = 17
      End If
      Case 1024
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 13.49
        coladj = 4.75
        vaSpread1.RowHeight(0) = 14
        vaSpread1.RowHeight(-1) = 14
      Else
        COne = 1.2
        coladj = 0
      End If
      Case 800
        COne = 0
        coladj = -0.5
        vaSpread1.Font.Size = 10
        vaSpread1.RowHeight(-1) = 14
      Case Else
    End Select
SkipAdjust:
    vaSpread1.ColWidth(1) = vaSpread1.ColWidth(1)
    vaSpread1.ColWidth(2) = vaSpread1.ColWidth(2) + coladj
    vaSpread1.ColWidth(3) = vaSpread1.ColWidth(3) + coladj
    vaSpread1.ColWidth(4) = vaSpread1.ColWidth(4) + coladj
    vaSpread1.ColWidth(5) = vaSpread1.ColWidth(5) + coladj
    vaSpread1.ColWidth(6) = vaSpread1.ColWidth(6) + coladj

End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
  
  If Col = 5 Then
    vaSpread1.Row = Row
    vaSpread1.Col = 1
    If vaSpread1.Text = "" Then
      vaSpread1.Col = Col
      vaSpread1.Text = ""
    End If
  End If

  If Col = 6 Then
    vaSpread1.Row = Row
    vaSpread1.Col = 1
    If vaSpread1.Text = "" Then
      vaSpread1.Col = Col
      vaSpread1.Text = ""
    End If
  End If
  

End Sub

