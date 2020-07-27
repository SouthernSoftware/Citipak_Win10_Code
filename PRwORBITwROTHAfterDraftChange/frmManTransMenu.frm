VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmManTransMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manual Transactions Menu"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   11655
   Icon            =   "frmManTransMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintManual 
      Height          =   495
      Left            =   4005
      TabIndex        =   1
      Top             =   4155
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
      ButtonDesigner  =   "frmManTransMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn EmpFileMaintCmmd 
      Height          =   492
      Left            =   4005
      TabIndex        =   0
      Top             =   3348
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
      ButtonDesigner  =   "frmManTransMenu.frx":0AAC
   End
   Begin fpBtnAtlLibCtl.fpBtn PostManCmmd 
      Height          =   495
      Left            =   4005
      TabIndex        =   2
      Top             =   4945
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
      ButtonDesigner  =   "frmManTransMenu.frx":0C99
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   4005
      TabIndex        =   3
      Top             =   5760
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
      ButtonDesigner  =   "frmManTransMenu.frx":0E85
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1097
      Index           =   1
      Left            =   1500
      Top             =   897
      Width           =   8655
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   2101
      Top             =   2103
      Width           =   971
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      Height          =   126
      Left            =   8593
      Top             =   2103
      Width           =   971
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   9412.576
      Y1              =   7894.417
      Y2              =   7894.417
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8710.757
      X2              =   8710.757
      Y1              =   2153.909
      Y2              =   7888.569
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2205.432
      X2              =   2919.248
      Y1              =   7894.417
      Y2              =   7894.417
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2220.428
      X2              =   2220.428
      Y1              =   2150.985
      Y2              =   7888.569
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MANUAL TRANSACTION MENU"
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
      TabIndex        =   4
      Top             =   1250
      Width           =   6012
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5900
      Index           =   1
      Left            =   8712
      Top             =   2201
      Width           =   732
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000B&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5900
      Index           =   0
      Left            =   2220
      Top             =   2201
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1214
      Left            =   1500
      Top             =   770
      Width           =   8652
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
   Begin VB.Shape Shape2 
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   8592
      Top             =   1971
      Width           =   972
   End
End
Attribute VB_Name = "frmManTransMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
'  Dim EmpRec As EmpData2Type
'  Dim EmpHandle As Integer
'  Dim NumOfEmpRecs As Integer
'  Dim THandle As Integer
'  Dim TRec As TransRecType
'  Dim NumOfTransRecs As Integer
'  Dim PHandle As Integer
'  Dim PPDRec As PeriodDefaultRecType
'  Dim X As Integer
'  Dim EmpIdxNNameRec As NumbSortIdxType
'  Dim NHandle As Integer
'  Dim NumOfIdxRecs As Integer
  
'  OpenEmpIdxNNameFile NHandle
'  NumOfIdxRecs = LOF(NHandle) / 2
'  ReDim IdxBuff(1 To NumOfIdxRecs) As Integer
'
'  For X = 1 To NumOfIdxRecs
'    Get NHandle, X, IdxBuff(X)
'  Next X
'  Close NHandle
'
'  OpenTransWorkFile THandle
'  NumOfTransRecs = LOF(THandle) / Len(TRec)
'
'  OpenEmpData2File EmpHandle
'  NumOfEmpRecs = LOF(EmpHandle) / Len(EmpRec)
'  For X = 1 To NumOfIdxRecs
'    Get EmpHandle, IdxBuff(X), EmpRec
'    If Not EmpRec.Deleted And EmpRec.EMPTDATE = 0 Then
'      Get THandle, IdxBuff(X), TRec
'        If TRec.TActive = -1 Then
'          Exit For
'        End If
'    End If
'  Next X
'  Close EmpHandle
'  Close THandle
'
'  If X > NumOfEmpRecs Then
'    PPDRec.MACTIVE = 0
'  Else
'    PPDRec.MACTIVE = -1
'  End If
'
'  PPDRec.PACTIVE = 0  '
  
'  OpenPPDefaultFile PHandle
'  Put PHandle, 1, PPDRec
'  Close PHandle
  
  frmPayrollProcessingMenu.Show
  DoEvents
  Unload frmManTransMenu
End Sub

Private Sub EmpFileMaintCmmd_Click()
  InFileNames(1) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(2) = "PRDATA\PRERNCOD.DAT"
  InFileNames(3) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 3) = False Then
    Close
    Exit Sub
  End If
  frmTransEntryEdit.Show
  DoEvents
  Unload frmManTransMenu
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
      SendKeys "%T"
      Call cmdExit_Click
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
  Me.HelpContextID = hlpManualTransaction
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub cmdPrintManual_Click()
  frmReportOpt.Show vbModal
  If RptOpt = 2 Then
    Call PCPrintManRegisterT
    Exit Sub
  ElseIf RptOpt = 1 Then
    Call PCPrintManRegisterG
  Else
    Exit Sub
  End If
End Sub
Sub PCPrintManRegisterG()
  Dim RptTitle$, TotHrsPaid#
  Dim Cnt2 As Integer
  Dim AddLine As Integer
  Dim RegHrs#, VACHRS#, SICKHRS#, HOLHRS#, COMPHRS#, PerHRS#
  Dim TotalHrs#, TotEIC#, TotHrs#, TOTPaid#, TOTComp#
  Dim TRegWage#, TOTWage#
  Dim GPay#, SSTax#, MTax#, FTax#, STax#, RETTOT#
  Dim TNetPay#, GFedGross#, GStaGross#, GSocGross#, GMedGross#, GRetGross#
  Dim TotalHrsPaid#
  Dim SumDed$(1 To 5), SumErn$
  Dim NumOfDeds As Integer, dlm$
  ReDim TransRec(1) As TransRecType
  ReDim EmpRec1(1) As EmpData1Type
  ReDim Unit(1) As UnitFileRecType

  ReDim DedCodes(1 To 50) As DedCodeRecType
  Dim DedRec As DedCodeRecType
  Dim DedDesc(1 To 50) As String * 8
  ReDim TotDeds(1 To 50) As Double

  ReDim EDRHrs(1) As String * 11
  ReDim EDOHrs(1) As String * 11
  ReDim EDRPay(1) As String * 11
  ReDim EDOPay(1) As String * 11
  ReDim EDEarn(1) As String * 11
  ReDim EDGroP(1) As String * 11

  ReDim EDSAmt(1) As String * 11
  ReDim EDMAmt(1) As String * 11
  ReDim EDRAmt(1) As String * 11

  ReDim ENumb(1) As String * 13
  ReDim EName(1) As String * 32

  ReDim MTDates(1 To 3) As String * 22
  ReDim ChkNo(1) As String * 10
  ReDim DAcct(1 To 4) As String * 22
  ReDim DAmts(1 To 4) As String * 22

  ReDim FedGro(1) As String * 11
  ReDim StaGro(1) As String * 11
  ReDim SocGro(1) As String * 11
  ReDim MedGro(1) As String * 11
  ReDim RetGro(1) As String * 11

  ReDim BRat(1) As String * 11
  ReDim ORat(1) As String * 11
  ReDim Fill11(1) As String * 11

  ReDim SCnt(1) As String * 11
  ReDim HCnt(1) As String * 11

  ReDim RHrs(1) As String * 11
  ReDim VHrs(1) As String * 11
  ReDim SHrs(1) As String * 11
  ReDim HHrs(1) As String * 11
  ReDim CHrs(1) As String * 11
  ReDim PHrs(1) As String * 11
  ReDim THrs(1) As String * 11
  ReDim OTHrs(1) As String * 11
  ReDim OTPaid(1) As String * 11
  ReDim OTComp(1) As String * 11

  ReDim RErnP(1) As String * 11
  ReDim OErnP(1) As String * 11

  ReDim GPayP(1) As String * 11
  ReDim SSTaxP(1) As String * 11
  ReDim MTaxP(1) As String * 11
  ReDim FTaxP(1) As String * 11
  ReDim STaxP(1) As String * 11
  ReDim RetirP(1) As String * 11
  ReDim NetPayP(1) As String * 11
  ReDim Ded(1) As String * 11

  ReDim EEicP(1) As String * 11
  ReDim Ern(1) As String * 11
  ReDim Pg(1) As String * 5
  ReDim EMPLine(1) As String * 132

  Dim Dash(1) As String * 132
  Dim FileHandle As Integer
  Dim DedHandle As Integer
  Dim x As Integer
  Dim FF$
  Dim TransRecLen As Integer, Emp1RecLen As Integer
  Dim IdxRecLen As Integer
  Dim NumOfRecs As Integer
  Dim IdxFileSize&, SalCnt As Integer, HrlCnt As Integer
  Dim LineCnt As Integer, MaxLines As Integer, Page As Integer
  Dim IdxNHandle As Integer
  Dim Image0$, Image$, Image3$, Image5$
  Dim DTitle$(1 To 5), TDed$, LastDed As Integer
  Dim Nextx As Integer, tripCnt As Integer
  Dim ETitle$, SumHeader2$, RHandle As Integer
  Dim ManRegisterRptName$, NHandle As Integer
  Dim THandle As Integer, cnt As Integer
  Dim ManRegDedDescName$
  Dim EmpCnt As Integer
  
  dlm$ = "~"
  InFileNames(1) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(2) = "PRDATA\PRERNCOD.DAT"
  InFileNames(3) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 3) = False Then
    Close
    Exit Sub
  End If
'  FF$ = Chr$(12)
  RptTitle$ = "Manual Register Report"
  OpenUnitFile FileHandle
  Get FileHandle, 1, Unit(1)
  Close FileHandle
  
  OpenDedCodeFile DedHandle
  NumOfDeds = LOF(DedHandle) / Len(DedRec)
  For x = 1 To 50 'NumOfDeds '50
    Get DedHandle, x, DedRec
    If Len(QPTrim$(DedRec.DCDESC1)) > 0 Then
      DedCodes(x) = DedRec
'      NumOfDeds = NumOfDeds + 1
    End If
  Next x
  Close DedHandle

  
  FrmShowPctComp.Label1 = "Manual Transaction Report"
  FrmShowPctComp.Show
  Image0$ = "###0"
  Image$ = "##,##0.00"
  Image3$ = "###,##0.00"
  Image5$ = "####,##0.00"
  GFedGross# = 0
  GStaGross# = 0
  GMedGross# = 0
  GSocGross# = 0
  GRetGross# = 0
  
'  LSet Dash(1) = String$(132, "-")
'  LSet Fill11(1) = ""

  TransRecLen = Len(TransRec(1))
  Emp1RecLen = Len(EmpRec1(1))

  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxNName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  SalCnt = 0
  HrlCnt = 0

'  LineCnt = 0
'  MaxLines = 45
'  Page = 1

  ReDim IdxBuff(1 To NumOfRecs) As Integer
  OpenEmpIdxNNameFile IdxNHandle
  
  For x = 1 To NumOfRecs
    Get IdxNHandle, x, IdxBuff(x)
  Next x
  Close IdxNHandle
  
  ManRegDedDescName$ = "PRRPTS\MANDEDDESC.RPT"
  DedHandle = FreeFile
  Open ManRegDedDescName$ For Output As DedHandle
  For x = 1 To 50
    DedDesc(x) = DedCodes(x).DCDESC1
    Print #DedHandle, QPTrim$(DedDesc(x)); dlm;
  Next x
  Print #DedHandle, ""
  Close DedHandle
'---------------------------------------------
  ETitle$ = "   Reg Earn   O/T Earn                        Gross Pay        EIC    Soc Sec   Medicare        FWT        SWT     Retire    Net Pay"
  SumHeader2$ = ETitle$
'------------------------------------------------------------------
  ManRegisterRptName$ = "PRRPTS\MANREGISG.RPT"
  RHandle = FreeFile
  Open ManRegisterRptName$ For Output As RHandle
  
  OpenEmpData1File NHandle
  OpenTransWorkFile THandle
  For cnt = 1 To NumOfRecs
    If IdxBuff(cnt) <> 0 Then
      Get THandle, IdxBuff(cnt), TransRec(1)
      If TransRec(1).TActive = True Then
        Get NHandle, IdxBuff(cnt), EmpRec1(1)
        GoSub SumAndPrintTime
      End If
      FrmShowPctComp.ShowPctComp cnt, NumOfRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Unload FrmShowPctComp
        Exit Sub
      End If
    End If
  Next
  
  Close THandle
  Close NHandle
  Close RHandle

'----------------------------------------------------------------------------
  If EmpCnt = 0 Then '5/26/04
    MsgBox "No manual payroll records have been saved."
  Else
    arManTranEntry.Show
  End If
  
  MainLog ("Manual Transaction Register was processed.")
Exit Sub
  
SumAndPrintTime:
 RegHrs# = OldRound#(RegHrs# + TransRec(1).RegHrsWork)
 VACHRS# = OldRound#(VACHRS# + TransRec(1).VacUsed)
 SICKHRS# = OldRound#(SICKHRS# + TransRec(1).SickUsed)
 HOLHRS# = OldRound#(HOLHRS# + TransRec(1).HOLHOURS)
 COMPHRS# = OldRound#(COMPHRS# + TransRec(1).CompUsed)
 PerHRS# = OldRound#(PerHRS# + TransRec(1).PerHours)
 
 TotalHrs# = OldRound(TotalHrs# + TransRec(1).RegHrsWork + TransRec(1).VacUsed + TransRec(1).SickUsed + TransRec(1).HOLHOURS + TransRec(1).CompUsed + TransRec(1).PerHours)

 TotEIC# = OldRound#(TotEIC# + TransRec(1).EICAmt)

'-=-=-=-=-=-=-=
 TotHrs# = OldRound#(TotHrs# + TransRec(1).OTHrsPaid)

 If TransRec(1).TotOTWage <> 0 Then
   TOTPaid# = OldRound#(TOTPaid# + TransRec(1).TotOTWage)
 End If

 TOTComp# = OldRound#(TOTComp# + TransRec(1).OT2Comp)

 TRegWage# = OldRound#(TRegWage# + TransRec(1).TotRegWage)

 If TransRec(1).TotOTWage > 0 Then
   TOTWage# = OldRound#(TOTWage# + TransRec(1).TotOTWage)
 End If
 GPay# = OldRound#(GPay# + TransRec(1).GrossPay)
 SSTax# = OldRound#(SSTax# + TransRec(1).SocTaxAmt)
 MTax# = OldRound#(MTax# + TransRec(1).MedTaxAmt)
 FTax# = OldRound#(FTax# + TransRec(1).FedTaxAmt)
 STax# = OldRound#(STax# + TransRec(1).StaTaxAmt)
 If TransRec(1).RetireAmt <> 0 Then
   RETTOT# = OldRound(RETTOT# + TransRec(1).RetireAmt)
 End If

 TNetPay# = OldRound#(TNetPay# + TransRec(1).NetPay)

 GFedGross# = OldRound#(GFedGross# + TransRec(1).FedGrossPay)
 GStaGross# = OldRound#(GStaGross# + TransRec(1).StaGrossPay)
 GSocGross# = OldRound#(GSocGross# + TransRec(1).SocGrossPay)
 GMedGross# = OldRound#(GMedGross# + TransRec(1).MedGrossPay)
 GRetGross# = OldRound#(GRetGross# + TransRec(1).RetGrossPay)

' RSet FedGro(1) = TransRec(1).FedGrossPay
' RSet StaGro(1) = TransRec(1).StaGrossPay
' RSet MedGro(1) = TransRec(1).MedGrossPay
' RSet SocGro(1) = TransRec(1).SocGrossPay
' RSet RetGro(1) = TransRec(1).RetGrossPay

' LSet ENumb(1) = LTrim$(EmpRec1(1).EmpNo)
 LSet EName(1) = QPTrim$(EmpRec1(1).EmpLName) + ", " + QPTrim$(EmpRec1(1).EmpFName)

' RSet BRat(1) = TransRec(1).BaseRate
' RSet ORat(1) = TransRec(1).OTRate
'look here

' RSet RHrs(1) = TransRec(1).RegHrsWork

' RSet VHrs(1) = TransRec(1).VacUsed
' RSet SHrs(1) = TransRec(1).SickUsed
' RSet HHrs(1) = TransRec(1).HOLHOURS
' RSet CHrs(1) = TransRec(1).CompUsed
' RSet PHrs(1) = TransRec(1).PerHours
 TotHrsPaid# = TransRec(1).RegHrsPaid + TransRec(1).VacUsed + TransRec(1).SickUsed + TransRec(1).HOLHOURS + TransRec(1).CompUsed

' RSet THrs(1) = TotHrsPaid#

' RSet OTHrs(1) = TransRec(1).OTHours
' RSet OTPaid(1) = TransRec(1).TotOTWage
' RSet OTComp(1) = TransRec(1).OT2Comp

' RSet EEicP(1) = TransRec(1).EICAmt

 Select Case TransRec(1).PayType
   Case "S"
'     RSet RHrs(1) = "Salaried"
     SalCnt = SalCnt + 1
   Case Else
     HrlCnt = HrlCnt + 1
 End Select

'=======Using(Image3$,
'  RSet RErnP(1) = TransRec(1).TotRegWage
'  RSet OErnP(1) = TransRec(1).TotOTWage

'  RSet GPayP(1) = TransRec(1).GrossPay
'  RSet SSTaxP(1) = TransRec(1).SocTaxAmt
'  RSet MTaxP(1) = TransRec(1).MedTaxAmt
'  RSet FTaxP(1) = TransRec(1).FedTaxAmt
'  RSet STaxP(1) = TransRec(1).StaTaxAmt

'  RSet RetirP(1) = TransRec(1).RetireAmt

'  RSet NetPayP(1) = TransRec(1).NetPay
 
  For Cnt2 = 1 To 50 'NumOfDeds
    TotDeds(Cnt2) = OldRound#(TotDeds(Cnt2) + TransRec(1).DAmt(Cnt2))
'    RSet Ded(1) = TransRec(1).DAmt(Cnt2)
  Next

'  LSet MTDates(1) = MakeRegDate(TransRec(1).PayPdStart)
'  LSet MTDates(2) = MakeRegDate(TransRec(1).PayPdEnd)
'  LSet MTDates(3) = MakeRegDate(TransRec(1).CheckDate)

'  RSet ChkNo(1) = TransRec(1).CheckNum

  For Cnt2 = 1 To 4
    RSet DAcct(Cnt2) = QPTrim$(TransRec(1).TDist(Cnt2).DAcct)
    RSet DAmts(Cnt2) = TransRec(1).TDist(Cnt2).DRWage
  Next
 '                       0               1               2                   3                         4                                       5                                    6                                        7
  Print #RHandle, Unit(1).UFEMPR; dlm; Date$; dlm; EmpRec1(1).EmpNo; dlm; EName(1); dlm; MakeRegDate(TransRec(1).PayPdStart); dlm; MakeRegDate(TransRec(1).PayPdEnd); dlm; MakeRegDate(TransRec(1).CheckDate); dlm; TransRec(1).CheckNum; dlm;
  '                          8                          9                         10                          11                       12                        13                         14                         15                     16                    17                         18                     19
  Print #RHandle, TransRec(1).BaseRate; dlm; TransRec(1).OTRate; dlm; TransRec(1).RegHrsWork; dlm; TransRec(1).VacUsed; dlm; TransRec(1).SickUsed; dlm; TransRec(1).HOLHOURS; dlm; TransRec(1).CompUsed; dlm; TransRec(1).PerHours; dlm; TotHrsPaid#; dlm; TransRec(1).OTHours; dlm; TransRec(1).TotOTWage; dlm; TransRec(1).OT2Comp; dlm;
  '                        20                           21                           22                       23                          24                          25                         26                         27                         28                           29
  Print #RHandle, TransRec(1).TotRegWage; dlm; TransRec(1).TotOTWage; dlm; TransRec(1).GrossPay; dlm; TransRec(1).EICAmt; dlm; TransRec(1).SocTaxAmt; dlm; TransRec(1).MedTaxAmt; dlm; TransRec(1).FedTaxAmt; dlm; TransRec(1).StaTaxAmt; dlm; TransRec(1).RetireAmt; dlm; TransRec(1).NetPay; dlm;
 
  For x = 1 To 50
  '                         30  to  79
    Print #RHandle, TransRec(1).DAmt(x); dlm;
  Next x
  '                  80             81             82             83                     84                         85                              86                              87
  Print #RHandle, DAcct(1); dlm; DAcct(2); dlm; DAcct(3); dlm; DAcct(4); dlm; TransRec(1).FedGrossPay; dlm; TransRec(1).StaGrossPay; dlm; TransRec(1).SocGrossPay; dlm; TransRec(1).MedGrossPay; dlm;
  '                  88             89             90             91                     92
  Print #RHandle, DAmts(1); dlm; DAmts(2); dlm; DAmts(3); dlm; DAmts(4); dlm; TransRec(1).RetGrossPay; dlm;
  
'  RSet SCnt(1) = Using(Image0$, SalCnt)
'  RSet HCnt(1) = Using(Image0$, HrlCnt)

'  RSet THrs(1) = Using(Image3$, TotalHrs#)
'  RSet RHrs(1) = Using(Image3$, RegHrs#)
'  RSet VHrs(1) = Using(Image3$, VACHRS#)
'  RSet SHrs(1) = Using(Image3$, SICKHRS#)
'  RSet HHrs(1) = Using(Image3$, HOLHRS#)
'  RSet CHrs(1) = Using(Image3$, COMPHRS#)
'  RSet PHrs(1) = Using(Image3$, PerHRS#)
'  RSet OTHrs(1) = Using(Image3$, TOTHrs#)
'  RSet OTPaid(1) = Using(Image3$, TOTPaid#)
'  RSet OTComp(1) = Using(Image3$, TOTComp#)

'  RSet RErnP(1) = Using(Image3$, TRegWage#)
'  RSet OErnP(1) = Using(Image3$, TOTWage#)

'  RSet GPayP(1) = Using(Image3$, GPay#)
'  RSet SSTaxP(1) = Using(Image3$, SSTax#)
'  RSet MTaxP(1) = Using(Image3$, MTax#)
'  RSet FTaxP(1) = Using(Image3$, FTax#)
'  RSet STaxP(1) = Using(Image3$, STax#)

'  RSet RetirP(1) = Using(Image3$, RETTOT#)
'  RSet NetPayP(1) = Using(Image3$, TNetPay#)
  
'  RSet EEicP(1) = Using(Image5$, TotEIC#)

'  RSet FTaxP(1) = Using(Image5$, GFedGross#)
'  RSet STaxP(1) = Using(Image5$, GStaGross#)
'  RSet MTaxP(1) = Using(Image5$, GMedGross#)
'  RSet SSTaxP(1) = Using(Image5$, GSocGross#)
 
  '                  93           94          95            96            97            98             99            100             101           102            103             104
  Print #RHandle, SalCnt; dlm; HrlCnt; dlm; RegHrs#; dlm; VACHRS#; dlm; SICKHRS#; dlm; HOLHRS#; dlm; COMPHRS#; dlm; PerHRS#; dlm; TotalHrs#; dlm; TotHrs#; dlm; TOTPaid#; dlm; TOTComp#; dlm;
  '                 105             106           107          109          110          111         112        112
  Print #RHandle, TRegWage#; dlm; TOTWage#; dlm; GPay#; dlm; TotEIC#; dlm; SSTax#; dlm; MTax#; dlm; FTax#; dlm; STax#; dlm;
  '                 113           114
  Print #RHandle, RETTOT#; dlm; TNetPay#; dlm;
  '
  For x = 1 To 50
  '                       115 to 164
    Print #RHandle, Using(Image3$, TotDeds(x)); dlm;
  Next x
  '                 165               166              167              168              169
  Print #RHandle, GFedGross#; dlm; GStaGross#; dlm; GMedGross#; dlm; GSocGross#; dlm; GRetGross#; dlm;
  '               170
  Print #RHandle, ""
  
  EmpCnt = EmpCnt + 1
Return
  
  
End Sub

Private Sub PostManCmmd_Click()

  Dim PPDHandle As Integer
  Dim PDR(1) As PeriodDefaultRecType
  
  InFileNames(1) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(2) = "PRDATA\PRERNCOD.DAT"
  InFileNames(3) = "PRDATA\PREMP2.DAT"
  InFileNames(4) = "PRDATA\PRSYS.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 4) = False Then
    Close
    Exit Sub
  End If
  OpenPPDefaultFile PPDHandle
  Get PPDHandle, 1, PDR(1)
  Close PPDHandle
  
  If PDR(1).MACTIVE = False And PDR(1).PACTIVE = False Then
    frmWarnNoActiveTrans.Show vbModal, Me
    Close
    Exit Sub
  End If
   
  frmWarningPostPayroll.Show
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmManTransMenu.")
      Call Terminate
      End
    End If
  End If
End Sub

Sub PCPrintManRegisterT()
  Dim RptTitle$, TotHrsPaid#
  Dim Cnt2 As Integer
  Dim AddLine As Integer
  Dim RegHrs#, VACHRS#, SICKHRS#, HOLHRS#, COMPHRS#, PerHRS#
  Dim TotalHrs#, TotEIC#, TotHrs#, TOTPaid#, TOTComp#
  Dim TRegWage#, TOTWage#
  Dim GPay#, SSTax#, MTax#, FTax#, STax#, RETTOT#
  Dim TNetPay#, GFedGross#, GStaGross#, GSocGross#, GMedGross#, GRetGross#
  Dim TotalHrsPaid#
  Dim SumDed$(1 To 5), SumErn$
  Dim NumOfDeds As Integer
  Dim EmpCnt As Integer
  ReDim TransRec(1) As TransRecType
  ReDim EmpRec1(1) As EmpData1Type
  ReDim Unit(1) As UnitFileRecType

  ReDim DedCodes(1 To 50) As DedCodeRecType
  Dim DedRec As DedCodeRecType

  ReDim TotDeds(1 To 50) As Double

  ReDim EDRHrs(1) As String * 11
  ReDim EDOHrs(1) As String * 11
  ReDim EDRPay(1) As String * 11
  ReDim EDOPay(1) As String * 11
  ReDim EDEarn(1) As String * 11
  ReDim EDGroP(1) As String * 11

  ReDim EDSAmt(1) As String * 11
  ReDim EDMAmt(1) As String * 11
  ReDim EDRAmt(1) As String * 11

  ReDim ENumb(1) As String * 13
  ReDim EName(1) As String * 32

  ReDim MTDates(1 To 3) As String * 22
  ReDim ChkNo(1) As String * 10
  ReDim DAcct(1 To 4) As String * 22
  ReDim DAmts(1 To 4) As String * 22

  ReDim FedGro(1) As String * 11
  ReDim StaGro(1) As String * 11
  ReDim SocGro(1) As String * 11
  ReDim MedGro(1) As String * 11
  ReDim RetGro(1) As String * 11

  ReDim BRat(1) As String * 11
  ReDim ORat(1) As String * 11
  ReDim Fill11(1) As String * 11

  ReDim SCnt(1) As String * 11
  ReDim HCnt(1) As String * 11

  ReDim RHrs(1) As String * 11
  ReDim VHrs(1) As String * 11
  ReDim SHrs(1) As String * 11
  ReDim HHrs(1) As String * 11
  ReDim CHrs(1) As String * 11
  ReDim PHrs(1) As String * 11
  ReDim THrs(1) As String * 11
  ReDim OTHrs(1) As String * 11
  ReDim OTPaid(1) As String * 11
  ReDim OTComp(1) As String * 11

  ReDim RErnP(1) As String * 11
  ReDim OErnP(1) As String * 11

  ReDim GPayP(1) As String * 11
  ReDim SSTaxP(1) As String * 11
  ReDim MTaxP(1) As String * 11
  ReDim FTaxP(1) As String * 11
  ReDim STaxP(1) As String * 11
  ReDim RetirP(1) As String * 11
  ReDim NetPayP(1) As String * 11
  ReDim Ded(1) As String * 11

  ReDim EEicP(1) As String * 11
  ReDim Ern(1) As String * 11
  ReDim Pg(1) As String * 5
  ReDim EMPLine(1) As String * 132

  Dim Dash(1) As String * 132
  Dim FileHandle As Integer
  Dim DedHandle As Integer
  Dim x As Integer
  Dim FF$
  Dim TransRecLen As Integer, Emp1RecLen As Integer
  Dim IdxRecLen As Integer
  Dim NumOfRecs As Integer
  Dim IdxFileSize&, SalCnt As Integer, HrlCnt As Integer
  Dim LineCnt As Integer, MaxLines As Integer, Page As Integer
  Dim IdxNHandle As Integer
  Dim Image0$, Image$, Image3$, Image5$
  Dim DTitle$(1 To 5), TDed$, LastDed As Integer
  Dim Nextx As Integer, tripCnt As Integer
  Dim ETitle$, SumHeader2$, RHandle As Integer
  Dim ManRegisterRptName$, NHandle As Integer
  Dim THandle As Integer, cnt As Integer
  
  InFileNames(1) = "PRDATA\PRDEDCOD.DAT"
  InFileNames(2) = "PRDATA\PRERNCOD.DAT"
  InFileNames(3) = "PRDATA\PREMP2.DAT"
  If FilesROK(Me, InFileNames(), OutFileNames(), 3) = False Then
    Close
    Exit Sub
  End If
  FF$ = Chr$(12)
  RptTitle$ = "Manual Register Report"
  OpenUnitFile FileHandle
  Get FileHandle, 1, Unit(1)
  Close FileHandle
  
  OpenDedCodeFile DedHandle
  NumOfDeds = LOF(DedHandle) / Len(DedRec)
  For x = 1 To NumOfDeds
    Get DedHandle, x, DedRec
    If Len(QPTrim$(DedRec.DCDESC1)) > 0 Then
      DedCodes(x) = DedRec
    End If
  Next x
  Close DedHandle

  
  FrmShowPctComp.Label1 = "Manual Transaction Report"
  FrmShowPctComp.Show
  Image0$ = "###0"
  Image$ = "##,##0.00"
  Image3$ = "###,##0.00"
  Image5$ = "####,##0.00"
  GFedGross# = 0
  GStaGross# = 0
  GMedGross# = 0
  GSocGross# = 0
  GRetGross# = 0
  
  LSet Dash(1) = String$(132, "-")
  LSet Fill11(1) = ""

  TransRecLen = Len(TransRec(1))
  Emp1RecLen = Len(EmpRec1(1))

  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxNName)
  NumOfRecs = IdxFileSize& \ IdxRecLen

  SalCnt = 0
  HrlCnt = 0

  LineCnt = 0
  MaxLines = 45
  Page = 1

  ReDim IdxBuff(1 To NumOfRecs) As Integer
  OpenEmpIdxNNameFile IdxNHandle
  
  For x = 1 To NumOfRecs
    Get IdxNHandle, x, IdxBuff(x)
  Next x
  Close IdxNHandle
  
  For x = 1 To 5
    DTitle$(x) = ""
  Next x
  
  tripCnt = 1
  Nextx = 1
  For cnt = 1 To NumOfDeds
    If tripCnt = 13 Then
      tripCnt = 1
      Nextx = Nextx + 1
    End If
    TDed$ = QPTrim$(DedCodes(cnt).DCDESC1)
    If Len(TDed$) > 0 Then
      LastDed = LastDed + 1
      RSet Ded(1) = TDed$
      DTitle$(Nextx) = DTitle$(Nextx) + Ded(1)
    End If
    tripCnt = tripCnt + 1
  Next
'---------------------------------------------
  ETitle$ = "   Reg Earn   O/T Earn                        Gross Pay        EIC    Soc Sec   Medicare        FWT        SWT     Retire    Net Pay"
  SumHeader2$ = ETitle$
'------------------------------------------------------------------
  ManRegisterRptName$ = "PRRPTS\MANREGIS.RPT"
  RHandle = FreeFile
  Open ManRegisterRptName$ For Output As RHandle
  OpenEmpData1File NHandle
  OpenTransWorkFile THandle
  GoSub PrintManualHeader
  For cnt = 1 To NumOfRecs
    If IdxBuff(cnt) <> 0 Then
      Get THandle, IdxBuff(cnt), TransRec(1)
    
      If TransRec(1).TActive = True Then
        Get NHandle, IdxBuff(cnt), EmpRec1(1)
        GoSub SumAndPrintTime
        LineCnt = LineCnt + 7 + (AddLine - 1)
        If LineCnt >= MaxLines And cnt < NumOfRecs Then
          LineCnt = 0
          Print #RHandle, FF$
          GoSub PrintManualHeader
        End If
      End If
      FrmShowPctComp.ShowPctComp cnt, NumOfRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Unload FrmShowPctComp
        Exit Sub
      End If
    End If
  Next
  
  GoSub PrintSumTotal
  Close THandle
  Close NHandle
  Close RHandle

'----------------------------------------------------------------------------
  If EmpCnt = 0 Then '5/26/04
    MsgBox "No manual payroll records have been saved."
  Else
    ViewPrint ManRegisterRptName, RptTitle$, True
  End If
  
  MainLog ("Manual Transaction Register was processed.")
Exit Sub
  
PrintManualHeader:
  RSet Pg(1) = Str$(Page)
  Print #RHandle, Unit(1).UFEMPR + Space$(87) + "Page:" + Pg(1)
  Print #RHandle, "Manual Transaction Register"
  Print #RHandle, "Report Date: " + Date$
  Print #RHandle,
  Print #RHandle, "Employee No   Name                             Beg Date              End Date              Chk Date                         Check No"
  Print #RHandle, "  Base Rate   O/T Rate    Reg Hrs      Vacat       Sick        Hol       Comp   Personal      Total    O/T Hrs   O/T Paid   O/T Comp"
  Print #RHandle, ETitle$
  For x = 1 To 5
    If Len(QPTrim$(DTitle(x))) > 0 Then
      AddLine = AddLine + 1
      Print #RHandle, DTitle$(x)
    End If
  Next x
  
  Print #RHandle, "                Dist 1                Dist 2                Dist 3                Dist 4  Fed Gross  Sta Gross  Med Gross  Soc Gross"
  Print #RHandle, "                   Amt                   Amt                   Amt                   Amt                                   Ret Gross"
  Print #RHandle, Dash(1)
  LineCnt = LineCnt + 11 + (AddLine - 1)
  Page = Page + 1
Return

SumAndPrintTime:
 RegHrs# = OldRound#(RegHrs# + TransRec(1).RegHrsWork)
 VACHRS# = OldRound#(VACHRS# + TransRec(1).VacUsed)
 SICKHRS# = OldRound#(SICKHRS# + TransRec(1).SickUsed)
 HOLHRS# = OldRound#(HOLHRS# + TransRec(1).HOLHOURS)
 COMPHRS# = OldRound#(COMPHRS# + TransRec(1).CompUsed)
 PerHRS# = OldRound#(PerHRS# + TransRec(1).PerHours)
 
 TotalHrs# = OldRound(TotalHrs# + TransRec(1).RegHrsWork + TransRec(1).VacUsed + TransRec(1).SickUsed + TransRec(1).HOLHOURS + TransRec(1).CompUsed + TransRec(1).PerHours)

 TotEIC# = OldRound#(TotEIC# + TransRec(1).EICAmt)

'-=-=-=-=-=-=-=
 TotHrs# = OldRound#(TotHrs# + TransRec(1).OTHrsPaid)

 If TransRec(1).TotOTWage <> 0 Then
   TOTPaid# = OldRound#(TOTPaid# + TransRec(1).TotOTWage)
 End If

 TOTComp# = OldRound#(TOTComp# + TransRec(1).OT2Comp)

 TRegWage# = OldRound#(TRegWage# + TransRec(1).TotRegWage)

 If TransRec(1).TotOTWage > 0 Then
   TOTWage# = OldRound#(TOTWage# + TransRec(1).TotOTWage)
 End If
 GPay# = OldRound#(GPay# + TransRec(1).GrossPay)
 SSTax# = OldRound#(SSTax# + TransRec(1).SocTaxAmt)
 MTax# = OldRound#(MTax# + TransRec(1).MedTaxAmt)
 FTax# = OldRound#(FTax# + TransRec(1).FedTaxAmt)
 STax# = OldRound#(STax# + TransRec(1).StaTaxAmt)

 If TransRec(1).RetireAmt <> 0 Then
   RETTOT# = OldRound(RETTOT# + TransRec(1).RetireAmt)
 End If

 TNetPay# = OldRound#(TNetPay# + TransRec(1).NetPay)

 GFedGross# = OldRound#(GFedGross# + TransRec(1).FedGrossPay)
 GStaGross# = OldRound#(GStaGross# + TransRec(1).StaGrossPay)
 GSocGross# = OldRound#(GSocGross# + TransRec(1).SocGrossPay)
 GMedGross# = OldRound#(GMedGross# + TransRec(1).MedGrossPay)
 GRetGross# = OldRound#(GRetGross# + TransRec(1).RetGrossPay)

 RSet FedGro(1) = Using(Image3$, TransRec(1).FedGrossPay)
 RSet StaGro(1) = Using(Image3$, TransRec(1).StaGrossPay)
 RSet MedGro(1) = Using(Image3$, TransRec(1).MedGrossPay)
 RSet SocGro(1) = Using(Image3$, TransRec(1).SocGrossPay)
 RSet RetGro(1) = Using(Image3$, TransRec(1).RetGrossPay)

 LSet ENumb(1) = LTrim$(EmpRec1(1).EmpNo)
 LSet EName(1) = QPTrim$(EmpRec1(1).EmpLName) + ", " + QPTrim$(EmpRec1(1).EmpFName)

 RSet BRat(1) = Using(Image3$, TransRec(1).BaseRate)
 RSet ORat(1) = Using(Image3$, TransRec(1).OTRate)
'look here

 RSet RHrs(1) = Using(Image$, TransRec(1).RegHrsWork)

 RSet VHrs(1) = Using(Image$, TransRec(1).VacUsed)
 RSet SHrs(1) = Using(Image$, TransRec(1).SickUsed)
 RSet HHrs(1) = Using(Image$, TransRec(1).HOLHOURS)
 RSet CHrs(1) = Using(Image$, TransRec(1).CompUsed)
 RSet PHrs(1) = Using(Image$, TransRec(1).PerHours)
 TotHrsPaid# = TransRec(1).RegHrsPaid + TransRec(1).VacUsed + TransRec(1).SickUsed + TransRec(1).HOLHOURS + TransRec(1).CompUsed

 RSet THrs(1) = Using(Image$, TotHrsPaid#)

 RSet OTHrs(1) = Using(Image$, TransRec(1).OTHours)
 RSet OTPaid(1) = Using(Image$, TransRec(1).TotOTWage)
 RSet OTComp(1) = Using(Image$, TransRec(1).OT2Comp)

 RSet EEicP(1) = Using(Image3$, TransRec(1).EICAmt)

 Select Case TransRec(1).PayType
   Case "S"
     RSet RHrs(1) = "Salaried"
     SalCnt = SalCnt + 1
   Case Else
     HrlCnt = HrlCnt + 1
 End Select

'=======
 RSet RErnP(1) = Using(Image3$, TransRec(1).TotRegWage)
 RSet OErnP(1) = Using(Image3$, TransRec(1).TotOTWage)

 RSet GPayP(1) = Using(Image3$, TransRec(1).GrossPay)
 RSet SSTaxP(1) = Using(Image3$, TransRec(1).SocTaxAmt)
 RSet MTaxP(1) = Using(Image3$, TransRec(1).MedTaxAmt)
 RSet FTaxP(1) = Using(Image3$, TransRec(1).FedTaxAmt)
 RSet STaxP(1) = Using(Image3$, TransRec(1).StaTaxAmt)

 RSet RetirP(1) = Using(Image3$, TransRec(1).RetireAmt)

 RSet NetPayP(1) = Using(Image3$, TransRec(1).NetPay)
 
 For x = 1 To 5
   SumDed$(x) = ""
 Next x
 
 tripCnt = 1
 Nextx = 1
 AddLine = 0
 For Cnt2 = 1 To NumOfDeds
   If tripCnt = 13 Then
     tripCnt = 1
     Nextx = Nextx + 1
     AddLine = AddLine + 1
   End If
     TotDeds(Cnt2) = OldRound#(TotDeds(Cnt2) + TransRec(1).DAmt(Cnt2))
     RSet Ded(1) = Using(Image3$, TransRec(1).DAmt(Cnt2))
   SumDed$(Nextx) = SumDed$(Nextx) + Ded(1)
   tripCnt = tripCnt + 1
 Next

'----------------------------------------------
  SumErn$ = Space$(22)

  LSet MTDates(1) = MakeRegDate(TransRec(1).PayPdStart)
  LSet MTDates(2) = MakeRegDate(TransRec(1).PayPdEnd)
  LSet MTDates(3) = MakeRegDate(TransRec(1).CheckDate)

  RSet ChkNo(1) = TransRec(1).CheckNum

 For Cnt2 = 1 To 4
   RSet DAcct(Cnt2) = QPTrim$(TransRec(1).TDist(Cnt2).DAcct)
   RSet DAmts(Cnt2) = Using(Image3$, TransRec(1).TDist(Cnt2).DRWage)
 Next
'---------------------------------------Fill11(1)----------------
 Print #RHandle, ENumb(1) + EName(1) + MTDates(1) + MTDates(2) + MTDates(3) + Fill11(1) + ChkNo(1)
 Print #RHandle, BRat(1) + ORat(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1) + CHrs(1) + PHrs(1) + THrs(1) + OTHrs(1) + OTPaid(1) + OTComp(1)
 Print #RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + EEicP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1) + STaxP(1) + RetirP(1) + NetPayP(1)

 For x = 1 To 5
   If Len(QPTrim$(SumDed$(x))) > 0 Then
     Print #RHandle, SumDed$(x)
   End If
 Next x
 Print #RHandle, DAcct(1) + DAcct(2) + DAcct(3) + DAcct(4) + FedGro(1) + StaGro(1) + SocGro(1) + MedGro(1)
 Print #RHandle, DAmts(1) + DAmts(2) + DAmts(3) + DAmts(4) + Fill11(1) + Fill11(1) + Fill11(1) + RetGro(1)
 Print #RHandle,
 EmpCnt = EmpCnt + 1
Return

PrintSumTotal:
  RSet SCnt(1) = Using(Image0$, SalCnt)
  RSet HCnt(1) = Using(Image0$, HrlCnt)

  RSet THrs(1) = Using(Image3$, TotalHrs#)
  RSet RHrs(1) = Using(Image3$, RegHrs#)
  RSet VHrs(1) = Using(Image3$, VACHRS#)
  RSet SHrs(1) = Using(Image3$, SICKHRS#)
  RSet HHrs(1) = Using(Image3$, HOLHRS#)
  RSet CHrs(1) = Using(Image3$, COMPHRS#)
  RSet PHrs(1) = Using(Image3$, PerHRS#)
  RSet OTHrs(1) = Using(Image3$, TotHrs#)
  RSet OTPaid(1) = Using(Image3$, TOTPaid#)
  RSet OTComp(1) = Using(Image3$, TOTComp#)

  RSet RErnP(1) = Using(Image3$, TRegWage#)
  RSet OErnP(1) = Using(Image3$, TOTWage#)

  RSet GPayP(1) = Using(Image3$, GPay#)
  RSet SSTaxP(1) = Using(Image3$, SSTax#)
  RSet MTaxP(1) = Using(Image3$, MTax#)
  RSet FTaxP(1) = Using(Image3$, FTax#)
  RSet STaxP(1) = Using(Image3$, STax#)

  RSet RetirP(1) = Using(Image3$, RETTOT#)
  RSet NetPayP(1) = Using(Image3$, TNetPay#)
  
  For x = 1 To 5
    SumDed$(x) = ""
  Next x
  
  Nextx = 1
  tripCnt = 1
  For Cnt2 = 1 To NumOfDeds
    If tripCnt = 13 Then
      tripCnt = 1
      Nextx = Nextx + 1
    End If
    RSet Ded(1) = Using(Image3$, TotDeds(Cnt2))
    SumDed$(Nextx) = SumDed$(Nextx) + Ded(1)
    tripCnt = tripCnt + 1
  Next
  SumErn$ = Space$(22)
  RSet EEicP(1) = Using(Image5$, TotEIC#)
  RSet Pg(1) = Str$(Page)
  Print #RHandle, FF$
  Print #RHandle, Unit(1).UFEMPR + Space$(87) + "Page:" + Pg(1)
  Print #RHandle, "Manual Transaction Summary"
  Print #RHandle, "Report Date: " + Date$
  Print #RHandle,
  Print #RHandle, Dash(1)
  Print #RHandle,
  Print #RHandle, "   Salaried     Hourly    Reg Hrs      Vacat       Sick        Hol       Comp   Personal      Total    O/T Hrs   O/T Paid   O/T Comp"
  Print #RHandle, SCnt(1) + HCnt(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1) + CHrs(1) + PHrs(1) + THrs(1) + OTHrs(1) + OTPaid(1) + OTComp(1)
  Print #RHandle,
  Print #RHandle, SumHeader2$
  Print #RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + EEicP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1) + STaxP(1) + RetirP(1) + NetPayP(1)
  Print #RHandle,
  For x = 1 To 5
    If Len(QPTrim$(DTitle(x))) > 0 Then
      Print #RHandle, DTitle$(x)
      Print #RHandle, SumDed$(x)
    End If
  Next x
  Print #RHandle,
  Print #RHandle, "  Fed Gross  Sta Gross  Med Gross  Soc Gross  Ret Gross"

  RSet FTaxP(1) = Using(Image5$, GFedGross#)
  RSet STaxP(1) = Using(Image5$, GStaGross#)
  RSet MTaxP(1) = Using(Image5$, GMedGross#)
  RSet SSTaxP(1) = Using(Image5$, GSocGross#)
  RSet RetGro(1) = Using(Image5$, GRetGross#)

  Print #RHandle, FTaxP(1) + STaxP(1) + MTaxP(1) + SSTaxP(1) + RetGro(1)

  Print #RHandle,
  Print #RHandle, Dash(1)
  Print #RHandle, FF$

Return

End Sub

