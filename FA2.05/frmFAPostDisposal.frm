VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFAPostDisposal 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Post Disposed Of Items"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAPostDisposal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   3756
      Left            =   2448
      TabIndex        =   1
      Top             =   2880
      Width           =   6684
      _Version        =   196609
      _ExtentX        =   11790
      _ExtentY        =   6625
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
      BackColor       =   8421504
      Caption         =   ""
      Picture         =   "frmFAPostDisposal.frx":08CA
      Begin LpLib.fpList fpListDates 
         Height          =   915
         Left            =   1965
         TabIndex        =   3
         ToolTipText     =   "Select the date to post and either press F10 or double click the highlighted choice to begin posting procedure."
         Top             =   1245
         Width           =   2745
         _Version        =   196608
         _ExtentX        =   4842
         _ExtentY        =   1614
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
         ColDesigner     =   "frmFAPostDisposal.frx":08E6
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   825
         Left            =   1005
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to create the desired report."
         Top             =   2685
         Width           =   1845
         _Version        =   131072
         _ExtentX        =   3254
         _ExtentY        =   1455
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
         ButtonDesigner  =   "frmFAPostDisposal.frx":0B72
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdPost 
         Height          =   825
         Left            =   4125
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to create the desired report."
         Top             =   2685
         Width           =   1845
         _Version        =   131072
         _ExtentX        =   3254
         _ExtentY        =   1455
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
         ButtonDesigner  =   "frmFAPostDisposal.frx":0D4E
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H0080FFFF&
         BorderWidth     =   3
         Height          =   2028
         Left            =   1632
         Top             =   384
         Width           =   3420
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Active Dates:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   348
         Left            =   2352
         TabIndex        =   2
         Top             =   624
         Width           =   1836
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3996
      Left            =   2352
      Top             =   2796
      Width           =   6876
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed Asset Disposal Post Procedure"
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
      TabIndex        =   0
      Top             =   1152
      Width           =   6012
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   1008
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   960
      Width           =   8652
   End
End
Attribute VB_Name = "frmFAPostDisposal"
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
  Close
  DoEvents
  Unload frmFAPostDisposal
End Sub

Private Sub cmdPost_Click()
  
  If QPTrim$(fpListDates.Text) = "" Then
    MsgBox "No dates to post."
    Exit Sub
  End If
  
  If Not Exist(PrepostDsplName + CStr(Date2Num(fpListDates.Text)) + ".DAT") Then
'    If MsgBox("You are attempting to post fixed assets for disposal but no disposal prices or disposal methods have been saved for this date. Do you wish to jump to the edit screen where these values can be saved?", vbYesNo) = vbYes Then
    frmFAPostDsplMess.Label1.Caption = "You are attempting to post fixed assets for disposal but no disposal prices or disposal methods have been saved for this date. Do you wish to jump to the edit screen where these values can be saved or do you wish to return to the disposal menu?"
    DoEvents
    frmFAPostDsplMess.Label2.Caption = "Any fixed assets that are disposed of require a method of disposal and disposal price before posting takes place."
    frmFAPostDsplMess.Show vbModal
    If frmFAPostDsplMess.fptxtChoice.Text = "abort" Then
      Close
      Unload frmFAPostDsplMess
      frmFADisposalMenu.Show
      DoEvents
      Unload frmFAPostDisposal
      Exit Sub
    Else
      Close
      Unload frmFAPostDsplMess
      frmFAEditDisposedOf.Show
      DoEvents
      Unload frmFAPostDisposal
      Exit Sub
    End If
  End If
 
  frmFAWarnPostDspl.Show vbModal
  If frmFAWarnPostDspl.fptxtAnswer = "Exit" Then
    Unload frmFAWarnPostDspl
    Exit Sub
  Else
    Unload frmFAWarnPostDspl
    Call PostThisData
  End If
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
      Call cmdPost_Click
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
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAPostDisposal.")
      Call Terminate
      End
    End If
  End If
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
    
NoMoreDates:
  fpListDates.Action = ActionSelectAll

End Sub

Private Sub PostThisData()
  Dim DateRec As TempDisposedOfDate
  Dim DateCnt As Integer
  Dim DSPLHandle As Integer
  Dim DTHandle As Integer
  Dim DsplCnt As Long
  Dim FAItemRec As FAItemRecType
  Dim FAHandle As Integer
  Dim FACnt As Long
  Dim x As Long, Nextx As Long
  Dim ThisCnt As Long
  Dim ThisTag$
  Dim ThisDate As Integer
  Dim LogDate$
  Dim PHandle As Integer
  Dim DsplRec As PrePostDsplType
  Dim StrDate$
  
  On Error GoTo ERRORSTUFF
  
  If QPTrim$(fpListDates.Text) = "" Then
    MsgBox "Please highlight the date you wish to post."
    fpListDates.Col = 1
    fpListDates.Row = 1
    fpListDates.SetFocus
    Exit Sub
  End If
  
  ThisDate = Date2Num(fpListDates.Text) 'grab and store date selected
  StrDate = CStr(ThisDate)
  LogDate = MakeRegDate(ThisDate)
  
  OpenPrePostDsplData PHandle, ThisDate
  DsplCnt = LOF(PHandle) / Len(DsplRec)
  
  OpenFAItemFile FAHandle
  FACnt = LOF(FAHandle) / Len(FAItemRec)
  
  For x = 1 To DsplCnt
    Get PHandle, x, DsplRec
    Get FAHandle, DsplRec.ThisRec, FAItemRec 'look for a depreciation date
    'that is after this disposal date...
    If FAItemRec.CDEPDATE > ThisDate Then
      Exit For
    End If
  Next x
  
  If x <= DsplCnt Then 'if the user elects to continue after
  'getting this warning then the town will have disposed of
  'an item(s) that have been depreciated already...they will
  'have gotten the benefit of depreciation on items they did not own
    DoEvents
    frmFADsplMess.Label1.Caption = "A depreciation processing date (" + MakeRegDate(FAItemRec.CDEPDATE) + ") comes after the disposal date entered (" + MakeRegDate(ThisDate) + "):"
    DoEvents
    frmFADsplMess.Label2.Top = 1500
    frmFADsplMess.Label2.Caption = "1. Fixed assets should not be depreciated after they are disposed of."
    frmFADsplMess.Label3.Top = 3000
    frmFADsplMess.Label3.Caption = "2. Continuing at this point would record assets that are no longer owned as being depreciated."
    frmFADsplMess.Show vbModal
    If frmFADsplMess.fptxtChoice.Text = "abort" Then
      Unload frmFADsplMess
      Close
      Exit Sub
    Else
      Unload frmFADsplMess
      MainLog ("User warned that the disposal date (" + LogDate + ") is before a later depreciation date (" + MakeRegDate(FAItemRec.CDEPDATE) + ") and posted anyway in frmFAPostDisposal.")
    End If
  End If
  
  For x = 1 To DsplCnt
    Get PHandle, x, DsplRec
    If DsplRec.Deleted = True Then GoTo NoMatch
    Get FAHandle, DsplRec.ThisRec, FAItemRec 'do the post procedure but only
    'after the asset has been pre processed for disposal (dsplflag = 1)
      FAItemRec.DEPYN = "N" 'further depreciation stopped
      FAItemRec.ISTATUS = "I" 'further depreciation stopped
      FAItemRec.LifeLeft = 0 'no longer relevant
      FAItemRec.DsplFlag = 2 'this seals it
      FAItemRec.CURRVAL = 0 'no longer relevant
      FAItemRec.DisposAmt = DsplRec.DisposAmt
      FAItemRec.DsplMethod = DsplRec.DsplMethod
      Put FAHandle, DsplRec.ThisRec, FAItemRec
NoMatch:
  Next x
  Close
  
  KillFile PrepostDsplName + StrDate + ".DAT"
  
  OpenTempDisposedDate DTHandle
  DateCnt = LOF(DTHandle) / Len(DateRec)
  For x = 1 To DateCnt
    Get DTHandle, x, DateRec
    If DateRec.DsplDate <> 0 Then
      If DateRec.DsplDate = ThisDate Then
        DateRec.DsplDate = 0 'zero out this disposal date
        Put DTHandle, x, DateRec
      End If
    ThisCnt = ThisCnt + 1
    End If
  Next x
  
  Close DTHandle
  
  'with no valid dates left the file that collects
  'dates can safely be deleted
  If ThisCnt = 0 Then KillFile (TempDispDateName)
  
  MsgBox ("The disposal data for " + MakeRegDate(ThisDate) + " has been posted.")
  MainLog ("Disposal data for " + LogDate + " was posted in frmFAPostDisposal.")
  frmFADisposalMenu.Show
  DoEvents
  Unload frmFAPostDisposal
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAPostDisposal", "PostThisData", Erl)
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

Private Sub fpListDates_DblClick()
  Call cmdPost_Click
End Sub
