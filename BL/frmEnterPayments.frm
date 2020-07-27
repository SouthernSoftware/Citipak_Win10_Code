VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmBLEnterPayments 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Enter Payments Menu"
   ClientHeight    =   8730
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11565
   Icon            =   "frmEnterPayments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   300
      Left            =   5376
      TabIndex        =   1
      Top             =   7008
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
   Begin fpBtnAtlLibCtl.fpBtn cmdEnterPayTrans 
      Height          =   492
      Left            =   3960
      TabIndex        =   2
      Tag             =   "Press to begin the transaction entry process."
      Top             =   2685
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
      ButtonDesigner  =   "frmEnterPayments.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEditPayTrans 
      Height          =   492
      Left            =   3960
      TabIndex        =   3
      Tag             =   "Press to bring up an interactive list of all transactions waiting to be posted."
      Top             =   3405
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
      ButtonDesigner  =   "frmEnterPayments.frx":0AB7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintEditList 
      Height          =   492
      Left            =   3960
      TabIndex        =   4
      Tag             =   "Press to bring up an option to print a graphical or text report listing all outstanding unposted transactions."
      Top             =   4125
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
      ButtonDesigner  =   "frmEnterPayments.frx":0CA3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   492
      Left            =   3960
      TabIndex        =   5
      Tag             =   "Press to begin the posting process for all outstanding transactions."
      Top             =   4845
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
      ButtonDesigner  =   "frmEnterPayments.frx":0E86
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   5565
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEnterPayments.frx":1067
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Tag             =   "Press to exit this screen and return to the main Business License menu."
      Top             =   6288
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmEnterPayments.frx":124C
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   150
      Index           =   3
      Left            =   1970
      Top             =   2000
      Width           =   990
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   155
      Index           =   4
      Left            =   8550
      Top             =   1995
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PAYMENT ENTRY"
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
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8666
      X2              =   8666
      Y1              =   2136
      Y2              =   8005
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2086
      X2              =   2086
      Y1              =   2136
      Y2              =   8008
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2086
      X2              =   2795
      Y1              =   8025
      Y2              =   8025
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8666
      X2              =   9369
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
Attribute VB_Name = "frmBLEnterPayments"
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

Private Sub cmdEditPayTrans_Click()
  frmBLEditTransList2.Show
  DoEvents
  Unload frmBLEnterPayments
End Sub

Private Sub cmdEnterPayTrans_Click()
  EditFlag = False
  frmBLTransEntry.Show
  DoEvents
  Unload frmBLEnterPayments
End Sub

Private Sub cmdExit_Click()
  PayDate$ = ""
  frmBLMainMenu.Show
  DoEvents
  Unload frmBLEnterPayments
End Sub

Private Sub cmdPost_Click()
  Dim PayHandle As Integer
  Dim PayRec As AREditPaymentRecType
  Dim NumOfPayRecs As Integer
  
  OpenPayFile PayHandle, OPERNUM
  NumOfPayRecs = LOF(PayHandle) / Len(PayRec)
  If NumOfPayRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no transaction records saved for current operator."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close PayHandle
    Exit Sub
  End If
  
  frmBLPostTrans.Show
  DoEvents
  Unload frmBLEnterPayments
End Sub

Private Sub cmdPrintEditList_Click()
  Dim PrintType$
  frmBLReportOpt.Show vbModal 'opens small screen from which the
  'user selects the printing method
  PrintType$ = frmBLReportOpt.fptxtPrintType
  Select Case PrintType$
    Case "Graphical"
      Call PrintGraphics
    Case "Text"
      frmBLMessageBoxJr.Label1.Caption = "Pitch 10 is recommended for this report."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Call PrintText
    Case "Exit"
  End Select
  Unload frmBLReportOpt
  cmdHelp.Text = "Turn Menu &Help On"
  btnHelp.AutoScan = fpAutoScanOff

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
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
      SendKeys "%M"
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLIssueAppsLics.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintText()
  Dim Oper$
  Dim PayHandle As Integer
  Dim PayRec As AREditPaymentRecType
  Dim NumOfPayRecs As Integer
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim FF$
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim CRec As Integer
  Dim Header$
  Dim CustNo&
  Dim CHANGE#
  Dim TChange#
  Dim TDue#
  Dim TAmt#
  Dim TotalCust As Integer
  Dim Page As Integer
  Dim cnt As Integer
  Dim TotCash As Double
  Dim TotCheck As Double
  Dim TotCredit As Double
  Dim NetRev As Double
  Dim Method$
  
  On Error GoTo ERRORSTUFF
  Oper$ = QPTrim$(Str$(OPERNUM))
  
  ReportFile$ = "AREDPY" + Oper$ + ".PRN" 'Report File Name
  FF$ = Chr$(12)
  MaxLines = 53
  LineCnt = 0
  Header$ = "Business License Payment Journal"
  OpenCustFile CustHandle
  NumOfCustRecs = LOF(CustHandle) / Len(CustRec)
  
  OpenPayFile PayHandle, OPERNUM
  NumOfPayRecs = LOF(PayHandle) \ Len(PayRec)
  
  If NumOfPayRecs = 0 Then
    Close
    frmBLMessageBoxJr.Label1.Caption = "There are no transactions saved for current operator."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  GoSub PrintRptHeader
  frmBLShowPctComp.Label1 = "Loading Detailed Customer List"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  
  For cnt = 1 To NumOfPayRecs ' NumOfARRecs
    Get PayHandle, cnt, PayRec
    CRec = Val(PayRec.CustNumber)
    If CRec <= 0 Then
      GoTo SkipDeleted
    End If
    Get CustHandle, Val(PayRec.CustNumber), CustRec
  
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintRptHeader
    End If
    If PayRec.Amount <> 0 Then
      CustNo& = Val(QPTrim$(CustRec.CustNumb))
      CHANGE# = OldRound(PayRec.AMTPAID - OldRound(PayRec.LICPAID + PayRec.PENPAID + PayRec.ISSPAID))
      TChange# = OldRound(TChange# + CHANGE#)
      TDue# = OldRound(TDue# + PayRec.TOTDUE)
      TAmt# = OldRound(TAmt# + PayRec.AMTPAID)
      TotCash = OldRound(TotCash + PayRec.CASHAMT)
      TotCheck = OldRound(TotCheck + PayRec.CHKAMT)
      TotCredit = OldRound(TotCredit + PayRec.CREDITAM)
      If QPTrim$(PayRec.CASHCHK) = "Both" Then
        Method = "Cash/Check"
      Else
        Method = QPTrim$(PayRec.CASHCHK)
      End If
      Print #RptHandle, Using("####0", CustNo&);
      Print #RptHandle, Tab(8); Left$(CustRec.BillName, 25); Tab(37); Method$; Tab(50); Using("$##,##0.00", PayRec.TOTDUE); Tab(64); Using("$##,##0.00", PayRec.AMTPAID); Tab(76); Using("$##,##0.00", CHANGE#)
      Print #RptHandle, QPTrim$(PayRec.DESC)
      TotalCust = TotalCust + 1
      LineCnt = LineCnt + 2
    End If
SkipDeleted:
    frmBLShowPctComp.ShowPctComp cnt, NumOfPayRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  
  GoSub PrintRptEnding
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now
  
  ViewPrint ReportFile, Header, True
  Kill ReportFile$
  
  Exit Sub
  
  
PrintRptHeader:
  Page = Page + 1
  Print #RptHandle, Tab(24); "Business License : Payment Journal"
  Print #RptHandle, Tab(27); "      Report Date: "; Date$; Tab(68); "Page #"; Page
  Print #RptHandle, ""
  Print #RptHandle, "Cust#"; Tab(8); "Billing Name"; Tab(37); "Method"; Tab(53); "Amt Due"; Tab(66); "Amt Paid"; Tab(80); "Change"
  Print #RptHandle, "Transaction Description"
  Print #RptHandle, String$(85, "=")
  LineCnt = 6
  Return
  
PrintRptEnding:
  Print #RptHandle, FF$
  Page = Page + 1
  Print #RptHandle, Tab(24); "Business License : Payment Journal"
  Print #RptHandle, Tab(27); "      Report Date: "; Date$; Tab(68); "Page #"; Page
  Print #RptHandle, Tab(42); "Summary"
  Print #RptHandle, String$(85, "-")
  Print #RptHandle, "Number of Entries: "; Using("###0", TotalCust);
  Print #RptHandle, Tab(50); Using("$##,##0.00", TDue#); Tab(64); Using("$##,##0.00", TAmt#); Tab(76); Using("$##,##0.00", TChange#)
  Print #RptHandle,
  Print #RptHandle, Tab(8); "Total Cash Paid"; Tab(40); Using("$###,##0.00", TotCash)
  Print #RptHandle, Tab(8); "Total Checks Paid"; Tab(40); Using("$###,##0.00", TotCheck)
  Print #RptHandle, Tab(8); "Total Charges Paid"; Tab(40); Using("$###,##0.00", TotCredit)
  NetRev = TotCash + TotCheck + TotCredit
  Print #RptHandle,
  Print #RptHandle, Tab(8); "Total Collected"; Tab(40); Using("$###,##0.00", NetRev)
  Print #RptHandle,
  Print #RptHandle, Tab(8); "Less Change"; Tab(40); Using("$###,##0.00", -TChange#)
  NetRev = NetRev - TChange#
  Print #RptHandle,
  Print #RptHandle, Tab(8); "Net Revenue Collected"; Tab(40); Using("$###,##0.00", NetRev)
  
  
  Print #RptHandle, FF$
  
  Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLEnterPayments", "PrintText", Erl)
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

Private Sub PrintGraphics()
  Dim Oper$
  Dim PayHandle As Integer
  Dim PayRec As AREditPaymentRecType
  Dim NumOfPayRecs As Integer
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim CRec As Integer
  Dim Header$
  Dim CustNo&
  Dim CHANGE#
  Dim cnt As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim ThisTown$
  Dim dlm$
  Dim TotCash As Double
  Dim TotCheck As Double
  Dim TotCredit As Double
  Dim TotCollected As Double
  Dim NetRev As Double
  Dim Method$
  
  On Error GoTo ERRORSTUFF
  
  dlm$ = "~"
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  ThisTown$ = TownRec.TownName
  
  Oper$ = QPTrim$(Str$(OPERNUM))
  
  ReportFile$ = "BLRPTS\AREDPY" + Oper$ + ".RPT" 'Report File Name
  OpenCustFile CustHandle
  NumOfCustRecs = LOF(CustHandle) / Len(CustRec)
  If NumOfCustRecs = 0 Then
    Close
    frmBLMessageBoxJr.Label1.Caption = "There are no customer files saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  OpenPayFile PayHandle, OPERNUM
  
  NumOfPayRecs = LOF(PayHandle) \ Len(PayRec)
  If NumOfPayRecs = 0 Then
    Close
    frmBLMessageBoxJr.Label1.Caption = "There are no transactions saved for current operator."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  frmBLShowPctComp.Label1 = "Loading Detailed Customer List"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  
  For cnt = 1 To NumOfPayRecs ' NumOfARRecs
    Get PayHandle, cnt, PayRec
    CRec = Val(PayRec.CustNumber)
    If CRec <= 0 Then
      GoTo SkipDeleted
    End If
    Get CustHandle, Val(PayRec.CustNumber), CustRec
    If PayRec.Amount <> 0 Then
      If QPTrim$(PayRec.CASHCHK) = "Both" Then
        Method = "Cash/Check"
      Else
        Method = QPTrim$(PayRec.CASHCHK)
      End If
      TotCash = TotCash + PayRec.CASHAMT
      TotCheck = TotCheck + PayRec.CHKAMT
      TotCredit = TotCredit + PayRec.CREDITAM
      TotCollected = TotCash + TotCheck + TotCredit
      CustNo& = Val(QPTrim$(CustRec.CustNumb))
      CHANGE# = OldRound(PayRec.AMTPAID - OldRound(PayRec.LICPAID + PayRec.PENPAID + PayRec.ISSPAID))
      NetRev = TotCollected - CHANGE#
      '                    0              1                 2
      Print #RptHandle, ThisTown; dlm; CustNo; dlm; CustRec.BillName; dlm;
      '                        3                    4                  5                 6                7
      Print #RptHandle, PayRec.AMTPAID; dlm; PayRec.TOTDUE; dlm; Method; dlm; CHANGE#; dlm; PayRec.DESC; dlm;
      '                    8              9              10
      Print #RptHandle, TotCash; dlm; TotCheck; dlm; TotCredit; dlm; TotCollected; dlm; NetRev; dlm; -CHANGE#
    End If
SkipDeleted:
    frmBLShowPctComp.ShowPctComp cnt, NumOfPayRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close         'Close all open files now
  Call arBLPayTransList.Show
  frmBLLoadReport.Show
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLEnterPayments", "PrintGraphics", Erl)
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

