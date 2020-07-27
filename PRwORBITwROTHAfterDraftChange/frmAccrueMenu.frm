VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmAccrueMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accrual Menu"
   ClientHeight    =   8880
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11655
   Icon            =   "frmAccrueMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdAccrue 
      Height          =   495
      Left            =   4005
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3360
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmAccrueMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrint 
      Height          =   491
      Left            =   4005
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4152
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   866
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
      ButtonDesigner  =   "frmAccrueMenu.frx":0AAD
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   480
      Left            =   4005
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4950
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmAccrueMenu.frx":0C95
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
      Height          =   491
      Left            =   4005
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5760
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   866
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
      ButtonDesigner  =   "frmAccrueMenu.frx":0E7A
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   2101
      Top             =   2103
      Width           =   972
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   8593
      Top             =   2103
      Width           =   971
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   1097
      Left            =   1500
      Top             =   897
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Benefits Accrual Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2820
      TabIndex        =   1
      Top             =   1250
      Width           =   6012
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1214
      Left            =   1500
      Top             =   770
      Width           =   8652
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   9412.576
      Y1              =   7884.973
      Y2              =   7884.973
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   8710.757
      Y1              =   2151.243
      Y2              =   7892.757
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2220.428
      X2              =   2220.428
      Y1              =   2151.243
      Y2              =   7880.108
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   2205.432
      X2              =   2919.248
      Y1              =   7881.081
      Y2              =   7881.081
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2101
      Top             =   1971
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   0
      Left            =   2220
      Top             =   2201
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8593
      Top             =   1971
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5910
      Index           =   1
      Left            =   8712
      Top             =   2201
      Width           =   732
   End
End
Attribute VB_Name = "frmAccrueMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdAccrue_Click()
  InFileNames(1) = "PRDATA\PRLEAVE.DAT"
  InFileNames(2) = "PRDATA\PRUNIT.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 2) = False Then
    Close
    Exit Sub
  End If
  
  frmAccrueLv.Show
  DoEvents
  Unload frmAccrueMenu
End Sub

Private Sub cmdEscape_Click()
  frmPayrollProcessingMenu.Show
  DoEvents
  Unload frmAccrueMenu
End Sub

Private Sub cmdPost_Click()
  Dim EmpIdxNNameHandle As Integer
  Dim INumOfRecs As Long
  Dim TempAccrualHandle As Integer
  Dim DHandle As Integer
  Dim EmpRec2(1) As EmpData2Type
  Dim TempAccRec As TempAccrualType
  Dim NumOfRecs As Long
  Dim RecNo As Long, x As Long
  Dim DoWhatFlag As PRTAccrue '8/7
  Dim Today$
  
  If Not Exist(TempAccrualName) Then
    MsgBox "No accrual data to post. Please run Accrue Benefits first."
    Exit Sub
  End If
  
  InFileNames(1) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 1) = False Then
    Close
    Exit Sub
  End If
  
'  Date$ = FormatDateTime(Date, vbShortDate)
  Today = Date '$
  
  'this select statement stops the process and asks
  'the user if he/she is absolutely sure they want to continue
  '...the reason is because it is easy to inadvertantly hit
  'the process key and once done it is not reversible easily
  DoWhatFlag = PromptPRTAccrue(Me) '8/7
  Select Case DoWhatFlag '8/7
  Case PRTAccrue.prtaEscape '8/7
    Exit Sub '8/7
  Case PRTAccrue.prtaContinue '8/7
  End Select '8/7
  
  
  OpenEmpIdxNNameFile EmpIdxNNameHandle
  INumOfRecs = LOF(EmpIdxNNameHandle) \ 2

  ReDim IdxBuf(1 To INumOfRecs) As Integer
  For x = 1 To INumOfRecs
     Get EmpIdxNNameHandle, x, IdxBuf(x)
  Next x
  Close EmpIdxNNameHandle
  
  If INumOfRecs = 0 Then
    MsgBox "No records on file."
    Close
    Exit Sub
  End If
  
  OpenTempAccrualFile TempAccrualHandle '12/11/02
  
  OpenEmpData2File DHandle
  NumOfRecs = LOF(DHandle) \ Len(EmpRec2(1))
  FrmShowPctComp.Label1 = "Employee Leave Accrual Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdAccrue.Enabled = False '12/11/02
  
  For RecNo = 1 To NumOfRecs
    Get TempAccrualHandle, IdxBuf(RecNo), TempAccRec
    Get DHandle, IdxBuf(RecNo), EmpRec2(1)
'    If QPTrim$(EmpRec2(1).EmpLName) = "WHITAKER" Then Stop
    If EmpRec2(1).EMPHDATE = 0 Or EmpRec2(1).EMPHDATE < -11000 Then GoTo BadHireDate
    If EmpRec2(1).EMPTDATE = 0 And EmpRec2(1).EMPBCODE > 0 And Not EmpRec2(1).Deleted Then
      EmpRec2(1).EMPSLBAL = TempAccRec.EMPSLBAL
      EmpRec2(1).EMPSLE = TempAccRec.EMPSLE
      EmpRec2(1).EMPVBAL = TempAccRec.EMPVBAL
      EmpRec2(1).EMPVACE = TempAccRec.EMPVACE
      EmpRec2(1).HOLBAL = TempAccRec.EMPHBAL
      EmpRec2(1).HOLERN = TempAccRec.EMPHOLE
      EmpRec2(1).PERBAL = TempAccRec.EMPPBAL
      EmpRec2(1).PERERN = TempAccRec.EMPPERE
      Put DHandle, IdxBuf(RecNo), EmpRec2(1)
    End If
    
BadHireDate:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdAccrue.Enabled = True '12/11/02
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  Me.cmdEscape.Enabled = True
  Me.cmdAccrue.Enabled = True '12/11/02
  EnableCloseButton Me.hwnd, True
  Call ProcessDate
  Close DHandle
  Close TempAccrualHandle '12/11/02
  KillFile TempAccrualName 'TempAccrualName is only used
  'for holding one accrual period at a time...when accrual is
  'posted then this file is destroyed
  If Exist("PRData\firstAccrual.dat") Then 'firstAccrual.dat is only used
  'when a new user has not started the beginning date for the accruing
  'process...from this point forward the starting date is automatically
  'assigned
    KillFile "PRData\firstAccrual.dat"
  End If
  MsgBox "Accrual posting has been completed."
  MainLog ("Employee accrual data posted.")
End Sub

Private Sub cmdPrint_Click()
  Dim AccrualRec As AccrualDates
  Dim AccrueHandle As Integer
  
  If Not Exist(TempAccrualName) Then
    MsgBox "No accrual data to print. Please run Accrue Benefits first."
    Exit Sub
  End If
  
  InFileNames(1) = "PRDATA\PRLEAVE.DAT"
  InFileNames(2) = "PRDATA\PRUNIT.DAT"
  InFileNames(3) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 3) = False Then
    Close
    Exit Sub
  End If
  
  If Exist("PRDATA\PRACCRUE.DAT") Then
    OpenAccrualDatesFile AccrueHandle
    Get AccrueHandle, 1, AccrualRec
    AccrualDate = AccrualRec.PreviousDate
    Close AccrueHandle
  End If
  
  frmReportOpt.Show vbModal
  If RptOpt = 2 Then
    Call ProcessAccrualT(AccrualDate)
    Exit Sub
  ElseIf RptOpt = 1 Then
    Call ProcessAccrualG(AccrualDate)
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
      SendKeys "%X"
      Call cmdEscape_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpAccrueLeave
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmAccrual Menu.")
      Call Terminate
      End
    End If
  End If
End Sub

Sub ProcessAccrualG(AccrualDate)

  Dim VAmt#, VADJFlag As Boolean, StableEntry As Long
  Dim SAmt#, EmpName$, TotalSick#, TotalVac#, TotalHol#, TotalPer#
  Dim SADJFlag As Boolean
  Dim LRecLen As Long
  Dim NumLeaveRec As Long
  Dim LeaveHandle As Integer
  Dim UnitHandle As Integer, x As Long
  Dim Image$, TImage$, Image1$, Image2$
  Dim TblPos As Integer, YrsPos As Integer, BenPos As Integer
  Dim VacPos As Integer, SickPos As Integer, LineCnt As Integer
  Dim MaxLines As Integer, EmpRecSize As Long
  Dim NumOfRecs As Long, IdxRecLen As Integer
  Dim IdxFileSize&, INumOfRecs As Long
  Dim DHandle As Integer, RHandle As Integer
  Dim RptTitle$, RptName$, RptFile As Integer
  Dim THandle As Integer, AccrualRptFile$
  Dim RecNo As Long
  Dim EmpTotal As Long
  Dim HireDate As Long
  Dim WhatLeaveTbl As Integer
  Dim AccrualDays As Long
  Dim YearsOfService As Integer
  Dim cnt As Long
  Dim VTableEntry As Long
  Dim HTableEntry As Long, HAmt#, HADJFlag As Boolean, HolPos As Integer
  Dim PTableEntry As Long, PAmt#, PADJFlag As Boolean, PerPos As Integer
  Dim EmpIdxNNameHandle As Integer
  Dim dlm$, starS$, yrsEmplyd$, starV$, starH$, StarP$
  Dim TempAccrualHandle As Integer '12/11/02
  Dim TempAccRec As TempAccrualType '12/11/02
  Dim AccrueHandle As Integer
  Dim AccrueRec As AccrualDates
  Dim LastDate As Integer
  
  If Exist("PRDATA\PRACCRUE.DAT") Then
    OpenAccrualDatesFile AccrueHandle
    Get AccrueHandle, 1, AccrueRec 'LastDate
    AccrualDate = AccrueRec.CurrentDate 'LastDate
    Close AccrueHandle
  End If
  
  If AccrualDate <= 0 Then
    MsgBox "Please run Accrue Benefits first."
    Exit Sub
  End If
  
  dlm$ = "~"
  starS$ = ""
  starV$ = ""
  ReDim Unit(1) As UnitFileRecType
  ReDim EmpRec2(1) As EmpData2Type
  ReDim TwoPrint(1) As String * 79
  ReDim Tot(1) As String * 8
  ReDim LeaveRec(1) As LeaveRecType
  
  LRecLen = Len(LeaveRec(1))
  OpenLeaveFileName LeaveHandle
  NumLeaveRec = LOF(LeaveHandle) \ Len(LeaveRec(1))
  If NumLeaveRec = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  ReDim LeaveRec(1 To NumLeaveRec) As LeaveRecType
  
  For x = 1 To NumLeaveRec
    Get LeaveHandle, x, LeaveRec(x)
  Next x
  Close LeaveHandle
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  Image$ = "##0.00"
  TImage$ = "####0.00"
  Image1$ = "##"
  Image2$ = "###"

  TblPos = 41
  YrsPos = 46
  BenPos = 53
  VacPos = 63
  SickPos = 72

  EmpRecSize = Len(EmpRec2(1))

  OpenEmpIdxNNameFile EmpIdxNNameHandle
  INumOfRecs = LOF(EmpIdxNNameHandle) \ 2

  ReDim IdxBuf(1 To INumOfRecs) As Integer
  For x = 1 To INumOfRecs
     Get EmpIdxNNameHandle, x, IdxBuf(x)
  Next x
  Close EmpIdxNNameHandle
  
  If INumOfRecs = 0 Then
    MsgBox "No records on file."
    Close
    Exit Sub
  End If
  
  
  OpenEmpData2File DHandle
  NumOfRecs = LOF(DHandle) \ Len(EmpRec2(1))
  RptName$ = "PRRPTS\ACCRUALG.RPT"
  RHandle = FreeFile
  On Error GoTo ErrorHandler
  Open RptName$ For Output As RHandle
  FrmShowPctComp.Label1 = "Employee Leave Accrual Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdAccrue.Enabled = False '12/11/02
  
  For RecNo = 1 To NumOfRecs
    Get DHandle, IdxBuf(RecNo), EmpRec2(1)
    If EmpRec2(1).EMPTDATE = 0 And EmpRec2(1).EMPBCODE > 0 And Not EmpRec2(1).Deleted Then
    'if employee not terminated AND they get benefits.
      EmpTotal = EmpTotal + 1
      HireDate = EmpRec2(1).EMPHDATE
      If HireDate <= -11000 Or HireDate = 0 Then 'roughly 1950
        GoSub BadHireDate
        GoTo BadDateSkip
      End If
      WhatLeaveTbl = EmpRec2(1).LeaveTbl 'get data from
      'leave table assigned to this employee
      If WhatLeaveTbl < 1 Then
        GoTo BadDateSkip
'        WhatLeaveTbl = 1
      End If
      AccrualDays = AccrualDate - HireDate
      If AccrualDays > 365 Then
        YearsOfService = Int(AccrualDays / 365)
      Else
        YearsOfService = 0
      End If
      For cnt = 1 To 20
        If YearsOfService <= LeaveRec(WhatLeaveTbl).VEntry(cnt).YEARS Then
          Exit For
        End If
      Next
      If cnt > 20 Then cnt = 20
      If YearsOfService = LeaveRec(WhatLeaveTbl).VEntry(cnt).YEARS Then
        VTableEntry = cnt
      Else
        VTableEntry = cnt - 1
      End If
      If VTableEntry = 0 Then VTableEntry = 1
      VAmt# = OldRound#(LeaveRec(WhatLeaveTbl).VEntry(VTableEntry).EARN * (EmpRec2(1).EMPBCODE * 0.01))
      If VAmt# > 0 Then           ' if there is amount to add
        If EmpRec2(1).EMPVBAL + VAmt# > LeaveRec(WhatLeaveTbl).VacMax Then     'if > max amt
          VAmt# = LeaveRec(WhatLeaveTbl).VacMax - EmpRec2(1).EMPVBAL   'set amt to max
          VADJFlag = True
        End If                                             '
      End If
      
      For cnt = 1 To 20
        If YearsOfService <= LeaveRec(WhatLeaveTbl).SEntry(cnt).YEARS Then
          Exit For
        End If
      Next
      If cnt > 20 Then cnt = 20
      If YearsOfService = LeaveRec(WhatLeaveTbl).SEntry(cnt).YEARS Then
        StableEntry = cnt
      Else
        StableEntry = cnt - 1
      End If
      If StableEntry = 0 Then StableEntry = 1
      SAmt# = OldRound#(LeaveRec(WhatLeaveTbl).SEntry(StableEntry).EARN * (EmpRec2(1).EMPBCODE * 0.01)) '8/5
      If SAmt# > 0 Then           ' if there is amount to add
        If EmpRec2(1).EMPSLBAL + SAmt# > LeaveRec(WhatLeaveTbl).SICKMAX Then   'if > max amt
          SADJFlag = True
          SAmt# = LeaveRec(WhatLeaveTbl).SICKMAX - EmpRec2(1).EMPSLBAL
        End If
      End If
      
      For cnt = 1 To 20
        If YearsOfService <= LeaveRec(WhatLeaveTbl).HEntry(cnt).YEARS Then
          Exit For
        End If
      Next
      If cnt > 20 Then cnt = 20
      If YearsOfService = LeaveRec(WhatLeaveTbl).HEntry(cnt).YEARS Then
        HTableEntry = cnt
      Else
        HTableEntry = cnt - 1
      End If
      If HTableEntry = 0 Then HTableEntry = 1
      HAmt# = OldRound#(LeaveRec(WhatLeaveTbl).HEntry(HTableEntry).EARN * (EmpRec2(1).EMPBCODE * 0.01))
      If HAmt# > 0 Then           ' if there is amount to add
        If EmpRec2(1).HOLBAL + HAmt# > LeaveRec(WhatLeaveTbl).HolMax Then     'if > max amt
          HAmt# = LeaveRec(WhatLeaveTbl).HolMax - EmpRec2(1).HOLBAL   'set amt to max
          HADJFlag = True
        End If                                             '
      End If
      
      For cnt = 1 To 20
        If YearsOfService <= LeaveRec(WhatLeaveTbl).PEntry(cnt).YEARS Then
          Exit For
        End If
      Next
      If cnt > 20 Then cnt = 20
      If YearsOfService = LeaveRec(WhatLeaveTbl).PEntry(cnt).YEARS Then
        PTableEntry = cnt
      Else
        PTableEntry = cnt - 1
      End If
      If PTableEntry = 0 Then PTableEntry = 1
      PAmt# = OldRound#(LeaveRec(WhatLeaveTbl).PEntry(PTableEntry).EARN * (EmpRec2(1).EMPBCODE * 0.01))
      If PAmt# > 0 Then           ' if there is amount to add
        If EmpRec2(1).PERBAL + PAmt# > LeaveRec(WhatLeaveTbl).PerMax Then     'if > max amt
          PAmt# = LeaveRec(WhatLeaveTbl).PerMax - EmpRec2(1).PERBAL   'set amt to max
          PADJFlag = True
        End If                                             '
      End If
      GoSub UpDateReport
    End If

BadDateSkip:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdAccrue.Enabled = True '12/11/02
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  Me.cmdEscape.Enabled = True
  Me.cmdAccrue.Enabled = True '12/11/02
  EnableCloseButton Me.hwnd, True

  Close DHandle
  Close RHandle
  Close

  arLvBnfts.Show
  frmLoadingRpt.Show
  Exit Sub

BadHireDate:

  EmpName$ = Space$(28)
  LSet EmpName$ = QPTrim$(EmpRec2(1).EmpLName) + ", " + QPTrim$(EmpRec2(1).EmpFName)
  LSet TwoPrint(1) = LTrim$(EmpRec2(1).EmpNo)    'set number
  Mid$(TwoPrint(1), 13, 28) = EmpName$  'set name
  Mid$(TwoPrint(1), YrsPos) = "Invalid hire date."

  '                      0                       1                            2
  Print #RHandle, Unit(1).UFEMPR; dlm; MakeRegDate(AccrualDate); dlm; EmpRec2(1).EmpNo; dlm;
  '                  3            4       5
  Print #RHandle, EmpName$; dlm; -1; dlm; ""; dlm;
  '               6        7        8          9          10         11       12         13          14
  Print #RHandle, ""; dlm; ""; dlm; ""; dlm; starS; dlm; starV; dlm; ""; dlm; ""; dlm; starH; dlm; StarP

  Return


UpDateReport:
  TotalSick# = OldRound#(TotalSick# + SAmt#)
  TotalVac# = OldRound#(TotalVac# + VAmt#)
  TotalHol# = OldRound#(TotalHol# + HAmt#)
  TotalPer# = OldRound#(TotalPer# + PAmt#)
  EmpName$ = Space$(28)
  LSet EmpName$ = QPTrim$(EmpRec2(1).EmpLName) + ", " + QPTrim$(EmpRec2(1).EmpFName)

'  LSet TwoPrint(1) = LTrim$(EmpRec2(1).EmpNo)    'set number
'  Mid$(TwoPrint(1), 13, 28) = EmpName$  'set name

'  Mid$(TwoPrint(1), TblPos, 3) = Using(Image1$, Str$(WhatLeaveTbl)) 'set benefit
  If (AccrualDays \ 365) >= 1 Then
'    Mid$(TwoPrint(1), YrsPos, 6) = Using(Image1$, Str$(AccrualDays \ 365)) 'set benefit
    yrsEmplyd = Using(Image1$, Str$(AccrualDays \ 365))
  Else
'    Mid$(TwoPrint(1), YrsPos, 6) = " 0" 'set benefit
    yrsEmplyd = " 0"
  End If
'  Mid$(TwoPrint(1), BenPos, 6) = Using(Image$, Str$(EmpRec2(1).EMPBCODE)) 'set benefit
  
'  Mid$(TwoPrint(1), VacPos, 6) = Using(Image$, Str$(VAmt#))               'set vac
  starV = ""
  If VADJFlag Then
    VADJFlag = False
'    Mid$(TwoPrint(1), VacPos + 6) = "*"
    starV = "*"
  End If
  
  starS = ""
'  Mid$(TwoPrint(1), SickPos, 6) = Using(Image$, Str$(SAmt#))             'set sick
  If SADJFlag Then
    SADJFlag = False
'    Mid$(TwoPrint(1), SickPos + 6) = "*"
    starS = "*"
  End If

  starH = ""
'  Mid$(TwoPrint(1), HolPos, 6) = Using(Image$, Str$(HAmt#))             'set sick
  If HADJFlag Then
    HADJFlag = False
'    Mid$(TwoPrint(1), HolPos + 6) = "*"
    starH = "*"
  End If

  StarP = ""
'  Mid$(TwoPrint(1), PerPos, 6) = Using(Image$, Str$(PAmt#))             'set sick
  If PADJFlag Then
    PADJFlag = False
'    Mid$(TwoPrint(1), PerPos + 6) = "*"
    StarP = "*"
  End If

  '                      0                       1                            2
  Print #RHandle, Unit(1).UFEMPR; dlm; MakeRegDate(AccrualDate); dlm; EmpRec2(1).EmpNo; dlm;
  '                  3                4                   5
  Print #RHandle, EmpName$; dlm; WhatLeaveTbl; dlm; yrsEmplyd; dlm;
  '                        6                  7          8            9           10         11          12          13          14
  Print #RHandle, EmpRec2(1).EMPBCODE; dlm; VAmt#; dlm; SAmt#; dlm; starS; dlm; starV; dlm; HAmt#; dlm; PAmt#; dlm; starH; dlm; StarP

Return
'
ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdAccrue.Enabled = True '12/11/02
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."
End Sub

Private Sub ProcessDate()
   Dim AccrueHandle As Integer
   Dim AccrualDate As Integer
   Dim AccrualDateString$
   Dim AccrualRec As AccrualDates
   Dim LastDate As Integer
   
   OpenAccrualDatesFile AccrueHandle
   Get AccrueHandle, 1, AccrualRec
   AccrualRec.PreviousDate = AccrualRec.CurrentDate
   Put AccrueHandle, 1, AccrualRec
   Close AccrueHandle
   
   MainLog ("Accrual Leave Benefits report processed.")
End Sub

Sub ProcessAccrualT(AccrualDate)

  Dim VAmt#, VADJFlag As Boolean, StableEntry As Long
  Dim SAmt#, EmpName$, TotalSick#, TotalVac#, TotalHol#, TotalPer#
  Dim SADJFlag As Boolean
  Dim LRecLen As Long
  Dim NumLeaveRec As Long
  Dim LeaveHandle As Integer
  Dim UnitHandle As Integer, x As Long
  Dim Image$, TImage$, Image1$, Image2$
  Dim TblPos As Integer, YrsPos As Integer, BenPos As Integer
  Dim VacPos As Integer, SickPos As Integer, LineCnt As Integer
  Dim MaxLines As Integer, EmpRecSize As Long
  Dim NumOfRecs As Long, IdxRecLen As Integer
  Dim IdxFileSize&, INumOfRecs As Long
  Dim DHandle As Integer, RHandle As Integer
  Dim RptTitle$, RptName$, RptFile As Integer
  Dim THandle As Integer, AccrualRptFile$
  Dim RecNo As Long
  Dim EmpTotal As Long
  Dim HireDate As Long
  Dim WhatLeaveTbl As Integer
  Dim AccrualDays As Long
  Dim YearsOfService As Integer
  Dim cnt As Long
  Dim VTableEntry As Long
  Dim HTableEntry As Long, HAmt#, HADJFlag As Boolean, HolPos As Integer
  Dim PTableEntry As Long, PAmt#, PADJFlag As Boolean, PerPos As Integer
  Dim EmpIdxNNameHandle As Integer
  Dim TempAccrualHandle As Integer '12/11/02
  Dim TempAccRec As TempAccrualType '12/11/02
  Dim AccrueHandle As Integer
  Dim AccrueRec As AccrualDates
  Dim LastDate As Integer
        
  ReDim Unit(1) As UnitFileRecType
  ReDim EmpRec2(1) As EmpData2Type
  ReDim TwoPrint(1) As String * 100
  ReDim Tot(1) As String * 8
  ReDim LeaveRec(1) As LeaveRecType
  
  If Exist("PRDATA\PRACCRUE.DAT") Then
    OpenAccrualDatesFile AccrueHandle
    Get AccrueHandle, 1, AccrueRec 'LastDate
    AccrualDate = AccrueRec.CurrentDate 'LastDate
    Close AccrueHandle
  End If
  
  LRecLen = Len(LeaveRec(1))
  OpenLeaveFileName LeaveHandle
  NumLeaveRec = LOF(LeaveHandle) \ Len(LeaveRec(1))
  If NumLeaveRec = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  ReDim LeaveRec(1 To NumLeaveRec) As LeaveRecType
  
  For x = 1 To NumLeaveRec
    Get LeaveHandle, x, LeaveRec(x)
  Next x
  Close LeaveHandle
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  Image$ = "##0.00"
  TImage$ = "####0.00"
  Image1$ = "##"
  Image2$ = "###"

  TblPos = 41
  YrsPos = 46
  BenPos = 53
  VacPos = 63
  SickPos = 72
  HolPos = 82
  PerPos = 92
  LineCnt = 0
  MaxLines = 50
  
  EmpRecSize = Len(EmpRec2(1))

  OpenEmpIdxNNameFile EmpIdxNNameHandle
  INumOfRecs = LOF(EmpIdxNNameHandle) \ 2

  ReDim IdxBuf(1 To INumOfRecs) As Integer
  For x = 1 To INumOfRecs
     Get EmpIdxNNameHandle, x, IdxBuf(x)
  Next x
  Close EmpIdxNNameHandle
  
  If INumOfRecs = 0 Then
    MsgBox "No records on file."
    Close
    Exit Sub
  End If
  
'  OpenTempAccrualFile TempAccrualHandle '12/11/02

  OpenEmpData2File DHandle
  NumOfRecs = LOF(DHandle) \ Len(EmpRec2(1))
  RptName$ = "PRRPTS\ACCRUAL.RPT"
  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  RPTSetupPRN 8, RHandle '8 is the position on the
  'Printer setup screen for this report
  FrmShowPctComp.Label1 = "Employee Leave Accrual Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdAccrue.Enabled = False '12/11/02
  
  GoSub PrintLeaveHeader
  For RecNo = 1 To NumOfRecs
    Get DHandle, IdxBuf(RecNo), EmpRec2(1)
    If EmpRec2(1).EMPTDATE = 0 And EmpRec2(1).EMPBCODE > 0 And Not EmpRec2(1).Deleted Then
    'if employee not terminated AND they get benefits.
      EmpTotal = EmpTotal + 1
      HireDate = EmpRec2(1).EMPHDATE
      If HireDate <= -11000 Or HireDate = 0 Then 'roughly 1950
        GoSub BadHireDate
        GoTo BadDateSkip
        
      End If
      WhatLeaveTbl = EmpRec2(1).LeaveTbl 'get data from
      'leave table assigned to this employee
      If WhatLeaveTbl < 1 Then
        GoTo BadDateSkip
'        WhatLeaveTbl = 1
      End If
      AccrualDays = AccrualDate - HireDate
      If AccrualDays > 365 Then
        YearsOfService = Int(AccrualDays / 365)
      Else
        YearsOfService = 0
      End If
      
      For cnt = 1 To 20
        If YearsOfService <= LeaveRec(WhatLeaveTbl).VEntry(cnt).YEARS Then
          Exit For
        End If
      Next
      If cnt > 20 Then cnt = 20
      If YearsOfService = LeaveRec(WhatLeaveTbl).VEntry(cnt).YEARS Then
        VTableEntry = cnt
      Else
        VTableEntry = cnt - 1
      End If
      If VTableEntry = 0 Then VTableEntry = 1
      VAmt# = OldRound#(LeaveRec(WhatLeaveTbl).VEntry(VTableEntry).EARN * (EmpRec2(1).EMPBCODE * 0.01))
      If VAmt# > 0 Then           ' if there is amount to add
        If EmpRec2(1).EMPVBAL + VAmt# > LeaveRec(WhatLeaveTbl).VacMax Then     'if > max amt
          VAmt# = LeaveRec(WhatLeaveTbl).VacMax - EmpRec2(1).EMPVBAL   'set amt to max
          VADJFlag = True
        End If                                             '
      End If
      
      For cnt = 1 To 20
        If YearsOfService <= LeaveRec(WhatLeaveTbl).SEntry(cnt).YEARS Then
          Exit For
        End If
      Next
      If cnt > 20 Then cnt = 20
      If YearsOfService = LeaveRec(WhatLeaveTbl).SEntry(cnt).YEARS Then
        StableEntry = cnt
      Else
        StableEntry = cnt - 1
      End If
      If StableEntry = 0 Then StableEntry = 1
      SAmt# = OldRound#(LeaveRec(WhatLeaveTbl).SEntry(StableEntry).EARN * (EmpRec2(1).EMPBCODE * 0.01)) '8/5
      If SAmt# > 0 Then           ' if there is amount to add
        If EmpRec2(1).EMPSLBAL + SAmt# > LeaveRec(WhatLeaveTbl).SICKMAX Then   'if > max amt
          SADJFlag = True
          SAmt# = LeaveRec(WhatLeaveTbl).SICKMAX - EmpRec2(1).EMPSLBAL
        End If
      End If
      
      For cnt = 1 To 20
        If YearsOfService <= LeaveRec(WhatLeaveTbl).HEntry(cnt).YEARS Then
          Exit For
        End If
      Next
      If cnt > 20 Then cnt = 20
      If YearsOfService = LeaveRec(WhatLeaveTbl).HEntry(cnt).YEARS Then
        HTableEntry = cnt
      Else
        HTableEntry = cnt - 1
      End If
      If HTableEntry = 0 Then HTableEntry = 1
      HAmt# = OldRound#(LeaveRec(WhatLeaveTbl).HEntry(HTableEntry).EARN * (EmpRec2(1).EMPBCODE * 0.01))
      If HAmt# > 0 Then           ' if there is amount to add
        If EmpRec2(1).HOLBAL + HAmt# > LeaveRec(WhatLeaveTbl).HolMax Then     'if > max amt
          HAmt# = LeaveRec(WhatLeaveTbl).HolMax - EmpRec2(1).HOLBAL   'set amt to max
          HADJFlag = True
        End If                                             '
      End If
      
      For cnt = 1 To 20
        If YearsOfService <= LeaveRec(WhatLeaveTbl).PEntry(cnt).YEARS Then
          Exit For
        End If
      Next
      If cnt > 20 Then cnt = 20
      If YearsOfService = LeaveRec(WhatLeaveTbl).PEntry(cnt).YEARS Then
        PTableEntry = cnt
      Else
        PTableEntry = cnt - 1
      End If
      If PTableEntry = 0 Then PTableEntry = 1
      PAmt# = OldRound#(LeaveRec(WhatLeaveTbl).PEntry(PTableEntry).EARN * (EmpRec2(1).EMPBCODE * 0.01))
      If PAmt# > 0 Then           ' if there is amount to add
        If EmpRec2(1).PERBAL + PAmt# > LeaveRec(WhatLeaveTbl).PerMax Then     'if > max amt
          PAmt# = LeaveRec(WhatLeaveTbl).PerMax - EmpRec2(1).PERBAL   'set amt to max
          PADJFlag = True
        End If                                             '
      End If
      
      GoSub UpDateReport
    End If

BadDateSkip:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdAccrue.Enabled = True '12/11/02
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  Me.cmdEscape.Enabled = True
  Me.cmdAccrue.Enabled = True '12/11/02
  EnableCloseButton Me.hwnd, True

  GoSub PrintLeaveSummary
  Close DHandle
  RPTSetupPRN 123, RHandle '8/15
  Close RHandle
'  Close TempAccrualHandle '12/11/02
  Close

  RptTitle$ = "Employee Leave Accrual Report"
  ViewPrint RptName$, RptTitle$, True
  Exit Sub

BadHireDate:

  EmpName$ = Space$(28)
  LSet EmpName$ = QPTrim$(EmpRec2(1).EmpLName) + ", " + QPTrim$(EmpRec2(1).EmpFName)
  LSet TwoPrint(1) = LTrim$(EmpRec2(1).EmpNo)    'set number
  Mid$(TwoPrint(1), 13, 28) = EmpName$  'set name
  Mid$(TwoPrint(1), YrsPos) = "Invalid hire date."

  Print #RHandle, TwoPrint(1)
  LineCnt = LineCnt + 1
  GoSub Check4NewPage
  Return

PrintLeaveHeader:
  Print #RHandle, Unit(1).UFEMPR
  Print #RHandle, "Leave Benefits Earned"
  Print #RHandle, "Accrual Date: " + MakeRegDate(AccrualDate)
  Print #RHandle,
  Print #RHandle, "Number      Name                       Tbl   Yrs   Benefit% Vacation     Sick   Holiday  Personal" '  + CrLf$
  Print #RHandle, "-------------------------------------------------------------------------------------------------" '  + CrLf$
  LineCnt = LineCnt + 6
  GoSub Check4NewPage
Return

UpDateReport:
  TotalSick# = OldRound#(TotalSick# + SAmt#)
  TotalVac# = OldRound#(TotalVac# + VAmt#)
  TotalHol# = OldRound#(TotalHol# + HAmt#)
  TotalPer# = OldRound#(TotalPer# + PAmt#)
  EmpName$ = Space$(28)
  LSet EmpName$ = QPTrim$(EmpRec2(1).EmpLName) + ", " + QPTrim$(EmpRec2(1).EmpFName)
  LSet TwoPrint(1) = LTrim$(EmpRec2(1).EmpNo)    'set number
  Mid$(TwoPrint(1), 13, 28) = EmpName$  'set name

  Mid$(TwoPrint(1), TblPos, 3) = Using(Image1$, Str$(WhatLeaveTbl)) 'set benefit
  If (AccrualDays \ 365) >= 1 Then
    Mid$(TwoPrint(1), YrsPos, 6) = Using(Image1$, Str$(AccrualDays \ 365)) 'set benefit
  Else
    Mid$(TwoPrint(1), YrsPos, 6) = " 0" 'set benefit
  End If
  Mid$(TwoPrint(1), BenPos, 6) = Using(Image$, Str$(EmpRec2(1).EMPBCODE)) 'set benefit
  Mid$(TwoPrint(1), VacPos, 6) = Using(Image$, Str$(VAmt#))               'set vac
  If VADJFlag Then
    VADJFlag = False
    Mid$(TwoPrint(1), VacPos + 6) = "*"
  End If

  Mid$(TwoPrint(1), SickPos, 6) = Using(Image$, Str$(SAmt#))             'set sick
  If SADJFlag Then
    SADJFlag = False
    Mid$(TwoPrint(1), SickPos + 6) = "*"
  End If

  Mid$(TwoPrint(1), HolPos, 6) = Using(Image$, Str$(HAmt#))             'set sick
  If HADJFlag Then
    HADJFlag = False
    Mid$(TwoPrint(1), HolPos + 6) = "*"
  End If

  Mid$(TwoPrint(1), PerPos, 6) = Using(Image$, Str$(PAmt#))             'set sick
  If PADJFlag Then
    PADJFlag = False
    Mid$(TwoPrint(1), PerPos + 6) = "*"
  End If

  Print #RHandle, TwoPrint(1)
  LineCnt = LineCnt + 1
  GoSub Check4NewPage

Return

PrintLeaveSummary:
  LSet TwoPrint(1) = "Totals          Employees"
  Mid$(TwoPrint(1), 13, 3) = Using(Image2$, Str$(EmpTotal))
  Mid$(TwoPrint(1), VacPos - 2, 8) = Using(TImage$, Str$(TotalVac#))             'set vac
  Mid$(TwoPrint(1), SickPos - 2, 8) = Using(TImage$, Str$(TotalSick#))           'set sick
  Mid$(TwoPrint(1), HolPos - 2, 8) = Using(TImage$, Str$(TotalHol#))           'set hol
  Mid$(TwoPrint(1), PerPos - 2, 8) = Using(TImage$, Str$(TotalPer#))           'set per

  Print #RHandle, "-------------------------------------------------------------------------------------------------" '  + CrLf$
  Print #RHandle, TwoPrint(1)
  Print #RHandle,
  LSet TwoPrint(1) = "NOTE: " + Chr$(34) + "*" + Chr$(34) + " Indicates maxium balance reached."
  Print #RHandle, TwoPrint(1)
  Print #RHandle, Chr$(12)
Return

Check4NewPage:
  If LineCnt > MaxLines Then
    LineCnt = 0
    Print #RHandle, Chr$(12)
    GoSub PrintLeaveHeader
  End If
Return

End Sub

