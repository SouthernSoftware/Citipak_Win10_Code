VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmFAYearEndPost 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Year End Posting"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11580
   Icon            =   "frmFAYearEndPost.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   2316
      Left            =   1992
      TabIndex        =   1
      Top             =   2758
      Width           =   7692
      _Version        =   196609
      _ExtentX        =   13568
      _ExtentY        =   4085
      _StockProps     =   70
      Caption         =   $"frmFAYearEndPost.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   192
      AlignTextH      =   1
      AlignTextV      =   1
      Caption         =   $"frmFAYearEndPost.frx":09E9
      ForeColor       =   8454143
      Picture         =   "frmFAYearEndPost.frx":0B08
   End
   Begin EditLib.fpDateTime fptxtDispYear 
      Height          =   372
      Left            =   5136
      TabIndex        =   2
      ToolTipText     =   "Read only year for which this post will affect."
      Top             =   5590
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
      ControlType     =   1
      Text            =   "2018"
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
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   675
      Left            =   3600
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6855
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
      ButtonDesigner  =   "frmFAYearEndPost.frx":0B24
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   675
      Left            =   6090
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to commit the depreciation for the year displayed above to memory."
      Top             =   6840
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
      ButtonDesigner  =   "frmFAYearEndPost.frx":0D00
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Depreciation Period to be Posted:"
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
      Left            =   3840
      TabIndex        =   3
      Top             =   5158
      Width           =   3852
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3660
      Left            =   1632
      Top             =   2506
      Width           =   8412
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1464
      Top             =   1383
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "YEAR END FIXED ASSET POSTING"
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
      Left            =   2256
      TabIndex        =   0
      Top             =   1530
      Width           =   7068
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1464
      Top             =   1335
      Width           =   8652
   End
End
Attribute VB_Name = "frmFAYearEndPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  If Exist("fromItemMaintMenu.dat") Then
    KillFile "fromItemMaintMenu.dat"
    frmFAItemMaintMenu.Show
  Else
    frmFAYearEndMenu.Show
  End If
  
  Close
  DoEvents
  Unload frmFAYearEndPost
End Sub

Private Sub Form_Load()
  Dim DepFile As Integer
  Dim x As Integer
  Dim FADep(1) As FADepFileType
  Dim NumOfDprRecs As Long
  Dim DprHistRec As DprHistType
  
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
'  lblDone.Visible = False
  
  DepFile = FreeFile
  Open "FADPREDT.DAT" For Random Access Read Write Shared As #DepFile Len = Len(FADep(1))
  NumOfDprRecs = LOF(DepFile) / Len(FADep(1))
  For x = 1 To NumOfDprRecs
    Get DepFile, x, FADep(1) 'find the year for this depreciation
    'by examining the temporary depreciation file created for this year
    If QPTrim$(FADep(1).CurrYear) <> "" Then
      fptxtDispYear = QPTrim$(FADep(1).CurrYear)
      Exit For
    End If
  Next x
  Close
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
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAYearEndPost.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub cmdProcess_Click()
  Dim CurYear$
  Dim DepFile As Integer
  Dim FADep(1) As FADepFileType
  Dim NumOfDepRecs As Integer
  Dim YearHandle As Integer
  Dim FAYear As FAYearEndType
  Dim YearSize As Integer
  Dim FAFile As Integer
  Dim FAItemRec As FAItemRecType
  Dim NumOfFARecs As Integer
  Dim cnt&
  Dim ItemRecNo As Integer
  Dim DHHandle As Integer
  Dim DprHistCnt As Long
  Dim DprRec As DprHistType
  Dim BYear$, x As Integer
  Dim TotLife As Integer
  Dim LifeLeft As Double
  Dim NmlDprAmt As Double
  Dim CurrValue As Double
  Dim Dif As Double
  Dim Remain As String * 2
  Dim GetThisRec As Long
  Dim RepeatDprFlag As Boolean
  Dim OldRec As Long
  Dim ThisCnt As Integer
  Dim DprHistRec As DprHistType
  Dim HHandle As Integer
  Dim HistCnt As Long
  
  On Error GoTo ERRORSTUFF
  
  frmFAWarnPostDepr.Show vbModal
  If frmFAWarnPostDepr.fptxtAnswer = "Exit" Then
    Unload frmFAWarnPostDepr
    Exit Sub
  End If
  
  RepeatDprFlag = False
  DepFile = FreeFile
  OldRec = 0
  Open "FADPREDT.DAT" For Random Access Read Write Shared As #DepFile Len = Len(FADep(1))
  NumOfDepRecs = LOF(DepFile) / Len(FADep(1))
  'if no temporary depreciation records have been created it will be
  'trapped when this form is accessed
  
  For x = 1 To NumOfDepRecs
    Get DepFile, x, FADep(1) 'get year for this depreciation
    If QPTrim$(FADep(1).CurrYear) <> "" Then
      CurYear = Mid(FADep(1).CurrYear, 1, 4)
      Exit For
    End If
  Next x
  
  If QPTrim$(CurYear) = "" Then 'this should never happen
    MsgBox "The current depreciation year was not saved while building the depreciation file. Please build the depreciation file again."
    Close
    Exit Sub
  End If
  'Build Deprecition File
  'Open Deprec Edit File
  OpenFAItemFile FAFile
  NumOfFARecs = LOF(FAFile) / Len(FAItemRec)
  
  If NumOfFARecs = 0 Then
    Close
    MsgBox "There are no fixed assets saved."
    Exit Sub
  End If
  
  OpenDprHistFile DHHandle
  DprHistCnt = (LOF(DHHandle) / Len(DprRec)) + 1
  
  For x = 1 To DprHistCnt
    Get DHHandle, x, DprRec
    If Not QPTrim$(DprRec.DprYear) = CurYear$ Then GoTo NotThisYear
      If DprRec.SoSoftFlag = True Then 'look to see if this
      'depreciation is a follow up to a depreciation reversal
        RepeatDprFlag = True
        Exit For
      End If
NotThisYear:
  Next x
'  ReDim BadCnt(1 To 1) As String
  For cnt& = 1 To NumOfDepRecs 'temp recs for this depreciation only
    Get DepFile, cnt&, FADep(1) 'get temp record
    ItemRecNo = FADep(1).AssetRecord
    Get FAFile, ItemRecNo, FAItemRec
'    If QPTrim$(FAItemRec.ItemTag) = "440-0385" Then Stop
    If RepeatDprFlag = True Then 'if this is a depreciation event after
'    'a sosoft clear then we want to overwrite the old recs with the new data
'    'instead of creating new records
      GoSub GetOldRec
      If OldRec = 0 Then GoTo OldRecIsZero
      Get DHHandle, OldRec, DprRec
    End If
    'start updating item and depreciation history records
    FAItemRec.DEP2DATE = FAItemRec.DEP2DATE + FADep(1).CurYrDep
    DprRec.DprAmt = FADep(1).CurYrDep
    FAItemRec.CDEPDATE = FADep(1).DprDay
    DprRec.ThisDept = FAItemRec.IDEPT
    DprRec.OrigCost = FAItemRec.ORGCOST
    If QPTrim$(FAItemRec.IDESC1) <> "" Then
      DprRec.ThisDesc1 = QPTrim$(FAItemRec.IDESC1)
    Else
      DprRec.ThisDesc1 = QPTrim$(FAItemRec.IDESC2)
    End If
    DprRec.DprToDate = FAItemRec.DEP2DATE
'    If FAItemRec.ILIFE = 0 Then
'      ThisCnt = ThisCnt + 1
'      ReDim Preserve BadCnt(1 To ThisCnt) As String
'      BadCnt(ThisCnt) = QPTrim$(FAItemRec.ItemTag)
'      FAItemRec.ILIFE = 1
'    End If
    DprRec.Life = FAItemRec.ILIFE
    DprRec.DprYear = QPTrim$(CurYear)
    DprRec.ItemTag = QPTrim$(FAItemRec.ItemTag)
    BYear = MakeRegDate(FAItemRec.AQURDATE)
    DprRec.PurchYear = Mid(BYear, 7, 4)
    If RepeatDprFlag = False Then
      If FAItemRec.LastDprRec >= 0 Then
        DprRec.PrevDprRec = FAItemRec.LastDprRec
      Else
        DprRec.PrevDprRec = 0
      End If
    End If
    DprRec.BookTotal = OldRound(FAItemRec.ORGCOST - FAItemRec.DEP2DATE)
    FAItemRec.CURRVAL = OldRound(FAItemRec.ORGCOST - FAItemRec.DEP2DATE)
    If RepeatDprFlag = False Then
      FAItemRec.LastDprRec = DprHistCnt
    End If
    GoSub GetLifeLeft
    FAItemRec.LifeLeft = LifeLeft
    DprRec.LifeLeft = LifeLeft
    DprRec.SoSoftFlag = False
    Put FAFile, ItemRecNo, FAItemRec
    If RepeatDprFlag = True Then
      Put DHHandle, OldRec, DprRec
    Else
      Put DHHandle, DprHistCnt, DprRec
      DprHistCnt = DprHistCnt + 1
    End If
    GoTo Reversal
OldRecIsZero:
    'added 12/23/04 to fix an issue where if a post was reversed
    'the program would only post those items that were included
    'in the reversed post file
    FAItemRec.DEP2DATE = FAItemRec.DEP2DATE + FADep(1).CurYrDep
    DprRec.DprAmt = FADep(1).CurYrDep
    FAItemRec.CDEPDATE = FADep(1).DprDay
    DprRec.ThisDept = FAItemRec.IDEPT
    DprRec.OrigCost = FAItemRec.ORGCOST
    If QPTrim$(FAItemRec.IDESC1) <> "" Then
      DprRec.ThisDesc1 = QPTrim$(FAItemRec.IDESC1)
    Else
      DprRec.ThisDesc1 = QPTrim$(FAItemRec.IDESC2)
    End If
    DprRec.DprToDate = FAItemRec.DEP2DATE
    DprRec.Life = FAItemRec.ILIFE
    DprRec.DprYear = QPTrim$(CurYear)
    DprRec.ItemTag = QPTrim$(FAItemRec.ItemTag)
    BYear = MakeRegDate(FAItemRec.AQURDATE)
    DprRec.PurchYear = Mid(BYear, 7, 4)
    If FAItemRec.LastDprRec >= 0 Then
      DprRec.PrevDprRec = FAItemRec.LastDprRec
    Else
      DprRec.PrevDprRec = 0
    End If
    DprRec.BookTotal = OldRound(FAItemRec.ORGCOST - FAItemRec.DEP2DATE)
    FAItemRec.CURRVAL = OldRound(FAItemRec.ORGCOST - FAItemRec.DEP2DATE)
    FAItemRec.LastDprRec = DprHistCnt
    GoSub GetLifeLeft
    FAItemRec.LifeLeft = LifeLeft
    DprRec.LifeLeft = LifeLeft
    DprRec.SoSoftFlag = False
    Put FAFile, ItemRecNo, FAItemRec
    Put DHHandle, DprHistCnt, DprRec
    DprHistCnt = DprHistCnt + 1
Reversal:
  Next cnt&
  Close
  
  OpenYearFile YearHandle
  FAYear.CurYear = CurYear$
  FAYear.LastYear = CurYear$
  
  Put YearHandle, 1, FAYear
  Close YearHandle
  'Now Clear Edit File
  '---------added 1/4/05-------------
  OpenDprHistFile HHandle
  HistCnt = LOF(HHandle) / Len(DprHistRec)
  If HistCnt = 0 Then
    GoTo NoCount
  End If
  
  For x = 1 To HistCnt
  Get HHandle, x, DprHistRec
    If DprHistRec.SoSoftFlag = True Then
      DprHistRec.SoSoftFlag = False
      Put HHandle, x, DprHistRec
    End If
  Next x
  
NoCount:
  Close HHandle
  '----------added 1/4/05^^^^^^^^^^^^^^^...added because in some cases
  'if a reversal has taken place and a user edits an asset that was in thge reversal so that it will
  'now be skipped during the temporary depreciation build then the sosoft flags
  'will not be reset to false...this causes flags to go off whenever a new depreciation
  'takes place or if another reversal is required. This universal sosoft flag
  'check automatically sets all flags to false at the end of every depreciation post.

'FromDprDeletion:
  KillFile ("FADPREDT.DAT") 'destroy temp file
  
  frmFAYrEndPostOK.Show vbModal
  If RepeatDprFlag = False Then
    MainLog ("Depreciation for year " + CurYear$ + " posted in frmFAYearEndPost.")
  Else
    MainLog ("Repeat depreciation for year " + CurYear$ + " posted in frmFAYearEndPost.")
  End If
  Call cmdExit_Click
  
  Exit Sub
  
  
GetLifeLeft: 'this procedure is commented in frmFAUnPostSnglDspl
  If DprRec.Life = 0 Then
    LifeLeft = 0
    Return
  End If
  CurrValue = OldRound(DprRec.OrigCost) - OldRound(DprRec.DprToDate) 'current value
  NmlDprAmt = OldRound(DprRec.OrigCost / DprRec.Life) 'normal yearly depreciation
  If NmlDprAmt = 0 Then
    LifeLeft = FAItemRec.LifeLeft 'added 12/22/04 because users are
    'entering .01 for the purchase price causing a division by zero error...
    'in this case the lifeleft will always remain the same
    Return
  End If
  
  If CurrValue = 0 Then
    LifeLeft = 0
    GoTo LifeIsZero
  ElseIf OldRound(DprRec.DprAmt) > OldRound(DprRec.BookTotal) Then
    LifeLeft = 1
    GoTo LifeIsZero
  ElseIf OldRound(DprRec.DprAmt) < OldRound(NmlDprAmt) Then 'first year of depreciation
    'for this item
    LifeLeft = DprRec.Life
  Else
    LifeLeft = DprRec.BookTotal / NmlDprAmt
    LifeLeft = OldRound(LifeLeft)
    LifeLeft = OldRound(LifeLeft * 100)
    Remain = Right(CStr(LifeLeft), 2)
    If Val(Remain) = 0 Or Val(Remain) = 98 Or Val(Remain) = 99 Then
      LifeLeft = DprRec.BookTotal / NmlDprAmt
      LifeLeft = CInt(LifeLeft)
    ElseIf Val(Remain) = 50 Then
      LifeLeft = OldRound(DprRec.BookTotal / NmlDprAmt)
      LifeLeft = OldRound(LifeLeft + 0.5)
    Else
      LifeLeft = OldRound(LifeLeft / 100)
      Remain = Right(CStr(LifeLeft), 2)
      If InStr(1, Remain, ".") Then Remain = Remain * 100
      If Val(Remain) < 50 Then
        LifeLeft = DprRec.BookTotal / NmlDprAmt
        LifeLeft = CInt(LifeLeft) + 1
      Else
        LifeLeft = DprRec.BookTotal / NmlDprAmt
        LifeLeft = CInt(LifeLeft)
      End If
    End If
  End If
LifeIsZero:
  Return
  
GetOldRec:
  'this posting is a post reversal posting and therefore we
  'don't want to create new records...we want to overwrite
  'the files that were reversed...if we don't then unneeded
  'data will be saved and reported in any depreciation history
  'related report
  
  OldRec = FAItemRec.LastDprRec
  Do
    If OldRec = 0 Then Exit Do
    Get DHHandle, OldRec, DprRec
    If QPTrim$(DprRec.DprYear) = CurYear$ Then
      Exit Do
    ElseIf DprRec.PrevDprRec >= 0 Then
      OldRec = DprRec.PrevDprRec
    End If
  Loop

  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAYearEndPost", "cmdProcess_Click", Erl)
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
  
  
End Sub

