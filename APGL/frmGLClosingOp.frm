VERSION 5.00
Begin VB.Form frmGLClosingOpMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Ledger Closing Operations"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   ClipControls    =   0   'False
   FillColor       =   &H00C0C0C0&
   Icon            =   "frmGLClosingOp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSelectFunds 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Select Funds to Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   3612
   End
   Begin VB.CommandButton cmdPreClosing 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Pre-Closing Operations"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   3612
   End
   Begin VB.CommandButton cmdCloseYear 
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Close Fiscal Year"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   3612
   End
   Begin VB.CommandButton cmdExitCloseOpMenu 
      BackColor       =   &H00D0D0D0&
      Caption         =   "E&xit Closing Operations"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   3612
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   132
      Left            =   8880
      Top             =   2280
      Width           =   972
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9720
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9000
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   132
      Left            =   2400
      Top             =   2280
      Width           =   972
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Left            =   2400
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2520
      X2              =   3240
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2520
      Top             =   2400
      Width           =   732
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GENERAL LEDGER CLOSING OPERATIONS"
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
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   1440
      Width           =   6372
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   9000
      Top             =   2400
      Width           =   732
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1800
      Top             =   960
      Width           =   8652
   End
End
Attribute VB_Name = "frmGLClosingOpMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim GLSetup As GLSetupRecType
Dim Acct    As GLAcctRecType
Dim FundIdx As GLFundIndexType
Dim AcctIdx As GLAcctIndexType
Dim Trans   As GLTransRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim FY1BegDate As Integer, FY1EndDate As Integer, FY2BegDate As Integer, FY2EndDate As Integer
Dim FirstFund As String, LastFund As String
Dim ActiveYear As Integer, FBAcct As String, UnPosted As Integer
Dim BadTrans As Integer
Public Sub KillFileCO(FileName$)
Dim xxonce As Integer
xxonce = 0
  On Local Error GoTo ErrorCatch
tryagain:
  If ExistCO(FileName$) Then
    Kill FileName$
  End If
  Exit Sub
  'In wrightsville they were the error below when adding glaccts added retry and do not to terminate.
ErrorCatch:
  Select Case Err
    Case Is <> 53
      xxonce = xxonce + 1
      MainLog ("Killfile error code is " + Str$(Err) + " .")
      If FileName$ <> "GLAcct.IDX" Then
       MsgBox ("File deletion permission denied " + Str$(Err) + " . PLEASE CONTACT SOUTHERN SOFTWARE @ 1-800-842-8190."), vbOKOnly
       GLTerminate
      Else
        If xxonce < 3 Then
          Resume tryagain
        Else
          MsgBox ("File deletion permission denied " + Str$(Err) + " . PLEASE CONTACT SOUTHERN SOFTWARE @ 1-800-842-8190."), vbOKOnly
        End If
      End If
    Case 53
      Resume ExitFillFile
  End Select
    
ExitFillFile:
  
End Sub
Public Function ExistCO(FileName$)
  Dim FileHandle As Integer
  Dim FileSize As Long
 On Error GoTo LOGTHIS
  FileHandle = FreeFile
  Open FileName$ For Binary Shared As FileHandle
  FileSize = LOF(FileHandle)
  Close FileHandle
  If FileSize > 0 Then
    ExistCO = True
  Else
    ExistCO = False
    Kill FileName$
    MainLog ("ExistCO NOT-File " + FileName$ + "##@ Does not exist.")
  End If
  Exit Function
  
LOGTHIS:
  Call MainLog("Problem with CO " & FileName$)
  Resume Next
End Function
Private Sub cmdSelectFunds_Click()
  KillFileCO ("CLOSETB.PRN")
  frmFundSelClose.Show
  Unload frmGLClosingOpMenu
End Sub

Private Sub cmdPreClosing_Click()
  If Exist("FCLOSE.LST") Then
    
    frmReportOpt.Show 1
    If rptopt = 1 Then
      CheckClosingEntry
    ElseIf rptopt = 2 Then
      CheckClosingEntry2
    End If
    
  Else
    MsgBox "Select the funds to close first!", vbOKOnly, "GL Closing"
  End If
End Sub
Private Sub cmdCloseYear_Click()
 Dim PreGo As String, oy As String
  oy$ = Right$(Format(DateAdd("d", (FY1EndDate), "12-31-1979"), "mm/dd/yyyy"), 2)
  PreGo = Dir("PreClose" + oy$, vbDirectory)
  'SHOULD NEVER CLOSE YEAR PAST CURRENT DATE*&#%^%#^%q@^@
  If FY1EndDate < DateDiff("d", "12/31/1979", Date) Then
    If Exist("CLOSETB.PRN") And Exist("FCLOSE.LST") Then 'Exist("FBADJ.DAT") And
      frmWarning.Label1.Caption = "YOU ARE ABOUT TO CLOSE"
      frmWarning.Label6.Caption = "FISCAL YEAR ENDING"
      frmWarning.Label5.Caption = (Format(DateAdd("d", (FY1EndDate), "12-31-1979")))
      frmWarning.Label2.Caption = "ARE YOU SURE?"
      frmWarning.Show 1
      If frmWarning.nogo = False Then
        If PreGo <> "" Then
          MsgBox "A PreClosing Directory Already Exists, Please Contact Software Support Before Trying To Close Year.", vbOKOnly, "Closing Halted"
          Call MainLog("Preclose Exists msg, close aborted ")
          Exit Sub
        Else
          frmReportOpt.Show 1
          DeActivateControls frmGLClosingOpMenu
          PostAdj (UnPosted)
          If rptopt = 1 Then
            MakeOpenEntries
          ElseIf rptopt = 2 Then
            MakeOpenEntries2
          End If
          SplitTransFile
          UpDateNewYear
          RelinkBgtTrans frmGLClosingOpMenu, True
          RepostNewYearTrans
          If Exist("potrans.dat") Then
            PurgePOs
          End If
          ResetYears
          MsgBox "Closing Operations Complete", vbOKOnly, "GL Closing"
          Call MainLog("Closing Complete ")
          If UnPosted > 0 Then
            MsgBox "Errors During Closing", vbOKOnly, "GL Closing"
            Call MainLog("Close Errors ")
          End If
          ActivateControls frmGLClosingOpMenu
          cmdExitCloseOpMenu_Click
        End If
      End If
    Else
      MsgBox "You Must Run PreClosing Option First", vbOKOnly, "Close Canceled"
    End If
  Else
    frmWarning.Label1.Caption = "YOU ARE TRYING TO CLOSE"
    frmWarning.Label6.Caption = "FISCAL YEAR ENDING"
    frmWarning.Label5.Caption = (Format(DateAdd("d", (FY1EndDate), "12-31-1979")))
    frmWarning.Label4.Caption = "PLEASE CONTACT SOFTWARE SUPPORT"
    frmWarning.Label2.Caption = "YOU MAY NOT CONTINUE."
    frmWarning.Label3.Caption = "THIS OPERATION HAS BEEN TERMINATED!"
    frmWarning.cmdContinue.Enabled = False
    
    frmWarning.Show 1
    'If frmWarning.nogo = False Then
  End If
End Sub


Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  Me.HelpContextID = hlpGLClosing
  KillFileCO ("CLOSETB.PRN")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitCloseOpMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        KillFileCO ("Fclose.opn")
        KillFileCO ("FBADJ.DAT")
        KillFileCO ("CLOSETB.PRN")
        KillFileCO ("FCLOSE.LST")

        Call MainLog("Close via GLClose ")
        ClearInUse PWcnt
      End If
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub cmdExitCloseOpMenu_Click()
  frmGLSetupMenu.Show
  Call MainLog("Exit Close Menu ")
  KillFileCO ("Fclose.opn")
  KillFileCO ("FBADJ.DAT")
  KillFileCO ("CLOSETB.PRN")
  KillFileCO ("FCLOSE.LST")
  Unload frmGLClosingOpMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitCloseOpMenu_Click
      KeyCode = 0
      DoEvents
    Case Else:
  End Select
End Sub
  '--call close trial bal to calculate ending account balances
  'also copies ending acct balances to Prior Year Actual field
Private Sub CheckClosingEntry()
  Dim GLFBAdj As GLFBAdjRecType
  Dim fmt As String, NumFundstoClose As Integer, FundBalAdjFileName As String
  Dim FundIdxFileNum As Integer, NumFunds As Integer, EndDate As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, RecNo As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long, PageNum As Integer
  Dim FBARecLen As Integer, FundBalAdjFile As Integer, cnt As Integer
  Dim PRNFile As Integer, ReportFile As String, FundCode As String
  Dim ClosingThisFund As Boolean, F As Integer, DC As String, B As String
  Dim AdjAmt As Double, FBAcct As String, FundBalAcct As String
  Dim GoodAcct As Integer, MsgFlag As Boolean, RecCnt As Integer
  Dim ToPrint As String
  CloseTrialBal
  DeActivateControls frmGLClosingOpMenu
  fmt$ = "#,###,###,###.##"
 
  GetFBAcct FBAcct$
  ReDim Accts(1) As GLAcctRecType
  ReDim AcctIndex(1) As GLAcctIndexType

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFileNum, NumFunds

  ReDim FundList(1) As String                   'List of all active Funds
  GetFundList FundList$(), NumFunds
  ReDim FundCloseList$(1)                'List of Funds to close
  GetFundCloseList FundCloseList$(), NumFundstoClose

  FundBalAdjFileName$ = "FBADJ.DAT"
  If Exist(FundBalAdjFileName$) Then
     KillFileCO FundBalAdjFileName$
  End If

  
  FBARecLen = Len(GLFBAdj)
  FundBalAdjFile = FreeFile
  Open FundBalAdjFileName$ For Random As FundBalAdjFile Len = FBARecLen

  ReDim FundTotRev#(1 To NumFunds)       'List of total revenues by fund
  ReDim FundTotExp#(1 To NumFunds)       'list of tot exp by fund
  ReDim FundBalAdj#(1 To NumFunds)

  PRNFile = FreeFile
  ReportFile$ = "GLCLOSE.PRN"
  Open ReportFile$ For Output As #PRNFile
  If NumGLAccts <> 0 Then
    FrmShowPctComp.Label1 = "Checking Closing Entries"
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show , Me
    DeActivateControls frmGLClosingOpMenu
    DoEvents
  End If
  For cnt = 1 To NumGLAccts   'NumGLAccts
    FrmShowPctComp.ShowPctComp cnt, NumGLAccts

    Get AcctIdxFileNum, cnt, AcctIndex(1)
    Get AcctFileNum, AcctIndex(1).RecNum, Accts(1)

    FundCode$ = Left$(Accts(1).Num, GLFundLen)

    ClosingThisFund = False
    For F = 1 To NumFundstoClose
      If FundCode$ = QPTrim$(FundCloseList$(F)) Then
        ClosingThisFund = True
        Exit For
      End If
    Next

    If ClosingThisFund Then

      '--find position in list
      For F = 1 To NumFunds
        If FundCode$ = QPTrim$(FundList$(F)) Then
          Exit For
        End If
      Next
      If F = 0 Then
        Cls
        'QPrintRC "Fatal Error Code: TIHSHO", 1, 1, 12
        Close: End
      End If

      '--Apply balance to proper fund
      Select Case Accts(1).Typ
        'CASE "A"
        '  IF Accts(1).Bal >= 0 THEN
        '    'Debit
        '    DebitAmt# = Accts(1).Bal
        '    CreditAmt# = 0
        '  ELSE
        '    DebitAmt# = 0
        '    CreditAmt# = ABS(Accts(1).Bal)
        '  END IF
        '  GOSUB WriteToInterfaceFile
        'CASE "L"
        '  IF Accts(1).Bal >= 0 THEN
        '    DebitAmt# = 0
        '    CreditAmt# = Accts(1).Bal
        '  ELSE
        '    DebitAmt# = ABS(Accts(1).Bal)
        '    CreditAmt# = 0
        '  END IF
        '  GOSUB WriteToInterfaceFile
        Case "R"
          FundTotRev#(F) = FundTotRev#(F) + Accts(1).Bal
        Case "E"
          FundTotExp#(F) = FundTotExp#(F) + Accts(1).Bal
       End Select

    End If
  Next

  For F = 1 To NumFunds
    'IF FundTotRev#(F) <> 0 AND FundTotExp#(F) <> 0 THEN
    If FundTotRev#(F) <> 0 Or FundTotExp#(F) <> 0 Then
      FundBalAdj#(F) = FundTotRev#(F) - FundTotExp#(F)
      If FundBalAdj#(F) < 0 Then
        DC$ = " Debit"
      Else
        DC$ = " Credit"
      End If
      AdjAmt# = Abs(FundBalAdj#(F))
      FundBalAcct$ = QPTrim$(FundList$(F)) + FBAcct$
      GoodAcct = 0
      GoodAcct = AcctFind(FundBalAcct$)
      If GoodAcct = 0 Then
        B$ = " *"
        MsgFlag = True
      Else
        B$ = "  "
      End If
      ToPrint$ = FundBalAcct$ + B$ + "~" + Using$(fmt$, Str$(AdjAmt#))
      ToPrint$ = ToPrint$ + "~" + DC$
      Print #PRNFile, ToPrint$
      '--Write out adjustments
      RecCnt = RecCnt + 1
      GLFBAdj.AcctNum = FundBalAcct$
      GLFBAdj.AdjAmt = FundBalAdj#(F)
      Put FundBalAdjFile, RecCnt, GLFBAdj
    End If
  Next
  Close
  ActivateControls frmGLClosingOpMenu
  Load frmLoadingRpt
  If MsgFlag Then
    ARptCloseEntries.Label4.Caption = "* Account Does not Exist!   Setup account before closing!"
    ARptCloseEntries.Label4.Visible = True
    Call MainLog("Error GL Close -Setup acct before closing.")
  End If
  ARptCloseEntries.txtRptDate = "Fiscal Year Ending " + Format(DateAdd("d", (FY1EndDate), "12-31-1979"), "mm/dd/yyyy")
  ARptCloseEntries.txtDate = Now
  ARptCloseEntries.txtTown = GLUserName$
  ARptCloseEntries.GetName ReportFile$
  ARptCloseEntries.startrpt

Exit Sub

End Sub
Private Sub CheckClosingEntry2()
  Dim GLFBAdj As GLFBAdjRecType
  Dim fmt As String, NumFundstoClose As Integer, FundBalAdjFileName As String
  Dim FundIdxFileNum As Integer, NumFunds As Integer, EndDate As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, RecNo As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long, PageNum As Integer
  Dim FBARecLen As Integer, FundBalAdjFile As Integer, cnt As Integer
  Dim PRNFile As Integer, ReportFile As String, FundCode As String
  Dim ClosingThisFund As Boolean, F As Integer, DC As String, B As String
  Dim AdjAmt As Double, FBAcct As String, FundBalAcct As String
  Dim GoodAcct As Integer, MsgFlag As Boolean, RecCnt As Integer
  CloseTrialBal2
  DeActivateControls frmGLClosingOpMenu
  fmt$ = "#,###,###,###.##"
 
  GetFBAcct FBAcct$
  ReDim Accts(1) As GLAcctRecType
  ReDim AcctIndex(1) As GLAcctIndexType

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFileNum, NumFunds

  ReDim FundList(1) As String                   'List of all active Funds
  GetFundList FundList$(), NumFunds
  ReDim FundCloseList$(1)                'List of Funds to close
  GetFundCloseList FundCloseList$(), NumFundstoClose

  FundBalAdjFileName$ = "FBADJ.DAT"
  If Exist(FundBalAdjFileName$) Then
     Kill FundBalAdjFileName$
  End If

  
  FBARecLen = Len(GLFBAdj)
  FundBalAdjFile = FreeFile
  Open FundBalAdjFileName$ For Random As FundBalAdjFile Len = FBARecLen

  ReDim FundTotRev#(1 To NumFunds)       'List of total revenues by fund
  ReDim FundTotExp#(1 To NumFunds)       'list of tot exp by fund
  ReDim FundBalAdj#(1 To NumFunds)

  PRNFile = FreeFile
  ReportFile$ = "GLCLOSE.PRN"
  Open ReportFile$ For Output As #PRNFile
  If NumGLAccts <> 0 Then
    FrmShowPctComp.Label1 = "Checking Closing Entries"
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show , Me
    DeActivateControls frmGLClosingOpMenu
    DoEvents
  End If
  For cnt = 1 To NumGLAccts   'NumGLAccts
    FrmShowPctComp.ShowPctComp cnt, NumGLAccts

    Get AcctIdxFileNum, cnt, AcctIndex(1)
    Get AcctFileNum, AcctIndex(1).RecNum, Accts(1)

    FundCode$ = Left$(Accts(1).Num, GLFundLen)

    ClosingThisFund = False
    For F = 1 To NumFundstoClose
      If FundCode$ = QPTrim$(FundCloseList$(F)) Then
        ClosingThisFund = True
        Exit For
      End If
    Next

    If ClosingThisFund Then

      '--find position in list
      For F = 1 To NumFunds
        If FundCode$ = QPTrim$(FundList$(F)) Then
          Exit For
        End If
      Next
      If F = 0 Then
        Cls
        'QPrintRC "Fatal Error Code: TIHSHO", 1, 1, 12
        Close: End
      End If

      '--Apply balance to proper fund
      Select Case Accts(1).Typ
        'CASE "A"
        '  IF Accts(1).Bal >= 0 THEN
        '    'Debit
        '    DebitAmt# = Accts(1).Bal
        '    CreditAmt# = 0
        '  ELSE
        '    DebitAmt# = 0
        '    CreditAmt# = ABS(Accts(1).Bal)
        '  END IF
        '  GOSUB WriteToInterfaceFile
        'CASE "L"
        '  IF Accts(1).Bal >= 0 THEN
        '    DebitAmt# = 0
        '    CreditAmt# = Accts(1).Bal
        '  ELSE
        '    DebitAmt# = ABS(Accts(1).Bal)
        '    CreditAmt# = 0
        '  END IF
        '  GOSUB WriteToInterfaceFile
        Case "R"
          FundTotRev#(F) = FundTotRev#(F) + Accts(1).Bal
        Case "E"
          FundTotExp#(F) = FundTotExp#(F) + Accts(1).Bal
       End Select

    End If
  Next

  GoSub PrnHeader
  For F = 1 To NumFunds
    'IF FundTotRev#(F) <> 0 AND FundTotExp#(F) <> 0 THEN
    If FundTotRev#(F) <> 0 Or FundTotExp#(F) <> 0 Then
      FundBalAdj#(F) = FundTotRev#(F) - FundTotExp#(F)
      If FundBalAdj#(F) < 0 Then
        DC$ = " Debit"
      Else
        DC$ = " Credit"
      End If
      AdjAmt# = Abs(FundBalAdj#(F))
      FundBalAcct$ = QPTrim$(FundList$(F)) + FBAcct$
      GoodAcct = 0
      GoodAcct = AcctFind(FundBalAcct$)
      If GoodAcct = 0 Then
        B$ = " *"
        MsgFlag = True
      Else
        B$ = "  "
      End If
      Print #PRNFile, FundBalAcct$ + B$;
      Print #PRNFile, Tab(15); Using$(fmt$, Str$(AdjAmt#))
      Print #PRNFile, DC$
      '--Write out adjustments
      RecCnt = RecCnt + 1
      GLFBAdj.AcctNum = FundBalAcct$
      GLFBAdj.AdjAmt = FundBalAdj#(F)
      Put FundBalAdjFile, RecCnt, GLFBAdj
    End If
  Next
  If MsgFlag Then
    Print #PRNFile,
    Print #PRNFile, "* Account Does not Exist!"
    Print #PRNFile, "Setup account before closing!"
    Call MainLog("Error GL Close -Setup acct before closing.")
  End If
  Print #PRNFile, Chr$(12)

  Close
  ActivateControls frmGLClosingOpMenu
  ViewPrint ReportFile$, "G/L Closing Entries"
Exit Sub
PrnHeader:
  Print #PRNFile, GLUserName$
  Print #PRNFile, "Fund Balance Closing Entries"
  Print #PRNFile, "Fiscal Year Ending "; Format(DateAdd("d", (FY1EndDate), "12-31-1979"), "mm/dd/yyyy")
  Print #PRNFile,
  Print #PRNFile, "Fund                Amount  Entry"
  Print #PRNFile, "---------------------------------"
Return
End Sub

'*********************************************************************
'calculates account balances for each fund in the fund close out list
'and sets the p/y actual field to the ending balance
'prints closing trial balance report to disk file
'***********************************************************************
  'QPrintRC "Generating Closing Trial Balance Report.  Please wait.", 5, 10, 15
Private Sub CloseTrialBal()
  Dim MaxLines As Integer, LookFor As String, CrLF As String
  Dim Linecnt As Integer, PRNFile As Integer, FundCnt As Integer
  Dim ReportFile As String, ToPrint As String, SumLine As String
  Dim FF As String, Header As String, StartFund As String, EndFund As String
  Dim PRNFileNum As Integer, cnt As Integer, Howmany As Integer
  Dim FundCode As String, DivLine As String, DivLine2 As String
  Dim CommaFmt As String, TotalFmt As String, FundNumber As String
  Dim TotDr As Double, TotCr As Double, TranCashTot As Double, CalcBal As Double
  Dim NumFundstoClose As Integer, ClosingThisFund As Boolean, F As Integer
  Dim FundDr As Double, FundCr As Double, FundRecNum As Integer
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer
  Dim FundIdxFileNum As Integer, NumFunds As Integer, EndDate As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, RecNo As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long, PageNum As Integer
  Dim Debit As String, Credit As String, Diff As Double, PYFundBal As Double
  ReDim FundList(1) As String
  DeActivateControls frmGLClosingOpMenu
  GetFundList FundList(), NumFunds
  'Define vars used for printing
      'make sure funds are in ascending order

  GetFundCodes FirstFund$, LastFund$
  ReportFile$ = "CLOSETB.PRN"                'Report File Name
  Header$ = "Closing Trial Balance"
  ReDim Desc$(1)
  Desc$(1) = "Acct Number     Title                                      Debit          Credit"
  CommaFmt$ = "###,###,###.##"  'format takes 13 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 14 chars
  SumLine$ = String$(14, "-")   'column summary line
  DivLine$ = String$(80, "-")   'dashed line
  DivLine2$ = String$(80, "=")  'Double Line

  CrLF$ = Chr$(13) + Chr$(10)
  FF$ = Chr$(12)
  MaxLines = 55
  TotDr# = 0
  TotCr# = 0

  ReDim FundCloseList$(1)
  GetFundCloseList FundCloseList$(), NumFundstoClose

  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile

'  QPrintRC Space$(80), 25, 1, -1
'  QPrintRC "Processing:", 25, 2, -1

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFileNum, NumFunds
  If NumGLAccts <> 0 Then
    FrmShowPctComp.Label1 = "Calculating Closing Trial Balance"
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show , Me
    DoEvents
  End If
  For cnt = 1 To NumGLAccts   'NumGLAccts
    FrmShowPctComp.ShowPctComp cnt, NumGLAccts
    Get AcctIdxFileNum, cnt, AcctIdx
    Get AcctFileNum, AcctIdx.RecNum, Acct

    FundCode$ = Left$(Acct.Num, GLFundLen)
    ClosingThisFund = False
    For F = 1 To NumFundstoClose
      If FundCode$ = QPTrim$(FundCloseList$(F)) Then
        ClosingThisFund = True
        Exit For
      End If
    Next

    'QPrintRC Acct.Num, 25, 14, -1

    If ClosingThisFund Then

      CalcBal# = Round#(Acct.BegBal)            'get the beginning balance

      NextTr& = Acct.FrstTran   'get the first trans for this acct

      Do Until NextTr& = 0      'keep going 'til we run out of trans
        Get TransFileNum, NextTr&, Trans

          If Trans.TRDATE <= FY1EndDate Then

            Select Case Acct.Typ
              Case "A", "E"
                CalcBal# = Round#(CalcBal# + Trans.DrAmt - Trans.CrAmt)
              Case "L", "R"
                CalcBal# = Round#(CalcBal# + Trans.CrAmt - Trans.DrAmt)
            End Select

          End If

          NextTr& = Trans.NextTran                'Get the next transaction

      Loop

      Acct.Bal = CalcBal#
      Acct.Work = CalcBal# 'hold this
      Put AcctFileNum, AcctIdx.RecNum, Acct

    End If   'test for account in fund range
  Next       'next account

  'PrintHelp "Please wait..."

  TotDr# = 0: TotCr# = 0
  FrmShowPctComp.Label1 = "Calculating Closing Trial Balance Fund Totals"
  FrmShowPctComp.Show , Me
  DoEvents

  For cnt = 1 To NumFunds
    FrmShowPctComp.ShowPctComp cnt, NumFunds
    FundDr# = 0: FundCr# = 0
    Get FundIdxFileNum, cnt, FundIdx

    FundNumber$ = QPTrim$(FundIdx.FundNum)
    ClosingThisFund = False
    For F = 1 To NumFundstoClose
      If FundNumber$ = QPTrim$(FundCloseList$(F)) Then
        ClosingThisFund = True
        Exit For
      End If
    Next

    If ClosingThisFund Then

      FundRecNum = FindFund(FundNumber$)
      FundName$ = QPTrim$(GetFundTitle$(FundRecNum))

     ' GoSub PrintPageHeader

      For RecNo = 1 To NumGLAccts             'Active Accts

        Get AcctIdxFileNum, RecNo, AcctIdx

        FundCode$ = Left$(AcctIdx.AcctNum, GLFundLen)
        If FundCode$ = FundNumber$ Then
          Get AcctFileNum, AcctIdx.RecNum, Acct


          ToPrint$ = ""
          ToPrint$ = FundCode$ + "~" + FundName$ + "~" + Acct.Num
          ToPrint$ = ToPrint$ + "~" + Acct.Title
          Select Case Acct.Typ
          Case "A", "E"
            If Acct.Bal >= 0 Then
              Debit$ = Using$(CommaFmt$, Str$(Acct.Bal))
              TotDr# = TotDr# + Acct.Bal
              FundDr# = FundDr# + Acct.Bal
              Credit$ = "0"
            Else
              Credit$ = Using$(CommaFmt$, Str$(Abs(Acct.Bal)))
              TotCr# = TotCr# + Abs(Acct.Bal)
              FundCr# = FundCr# + Abs(Acct.Bal)
              Debit$ = "0"
            End If

          Case "L", "R"
            If Acct.Bal >= 0 Then
              Credit$ = Using$(CommaFmt$, Str$(Acct.Bal))
              TotCr# = TotCr# + Acct.Bal
              FundCr# = FundCr# + Acct.Bal
              Debit$ = "0"
            Else
              Debit$ = Using$(CommaFmt$, Str$(Abs(Acct.Bal)))
              TotDr# = TotDr# + Abs(Acct.Bal)
              FundDr# = FundDr# + Abs(Acct.Bal)
              Credit$ = "0"
            End If

          End Select

          ToPrint$ = ToPrint$ + "~" + Debit$ + "~" + Credit$
          Print #PRNFile, ToPrint$
        End If
      Next
      Diff# = Round#(FundDr# - FundCr#)
'      If Diff# <> 0 Then
'
'        ARptTrialBal.Label9.Caption = "Fund is out of balance :"
'        ARptTrialBal.txtDiff = Using$(CommaFmt$, Str$(Diff#))
'
'      End If
'
    Else
      'were not closing this fund
      FundRecNum = FindFund(FundNumber$)
      FundName$ = QPTrim$(GetFundTitle$(FundRecNum))
      ToPrint$ = ""
      ToPrint$ = FundNumber$ + "~" + FundName$ + "~" + "FUND: " + FundNumber$ + "~" + "Is Not Being Closed." + "~0~0"
      Print #PRNFile, ToPrint$
      ToPrint$ = ""
      'ARptTrialBal.Label9.Caption = "Fund " + FundNumber$ + " " + FundName$ + " is not being closed."
    End If    'if fund is in range test
  Next        'next fund

  Close
  Load frmLoadingRpt
  ActivateControls frmGLClosingOpMenu
  'End Report Processing
  ARptCloseTrialBal.txtRptDate = "Period Ending: " + Format(DateAdd("d", (FY1EndDate), "12-31-1979"), "mm/dd/yyyy")
  'ARptCloseTrialBal.gTotDebits = Using$(TotalFmt$, Str$(TotDr#))
  'ARptCloseTrialBal.gTotCredits = Using$(TotalFmt$, Str$(TotCr#))
  
  ARptCloseTrialBal.txtDate = Now
  ARptCloseTrialBal.txtTown = GLUserName$
  ARptCloseTrialBal.GetName ReportFile$
  ARptCloseTrialBal.startrpt

 ' ViewPrint ReportFile$, Header$
  'KILL ReportFile$
Exit Sub

End Sub
Private Sub CloseTrialBal2()
  Dim MaxLines As Integer, LookFor As String, CrLF As String
  Dim Linecnt As Integer, PRNFile As Integer, FundCnt As Integer
  Dim ReportFile As String, ToPrint As String, SumLine As String
  Dim FF As String, Header As String, StartFund As String, EndFund As String
  Dim PRNFileNum As Integer, cnt As Integer, Howmany As Integer
  Dim FundCode As String, DivLine As String, DivLine2 As String
  Dim CommaFmt As String, TotalFmt As String, FundNumber As String
  Dim TotDr As Double, TotCr As Double, TranCashTot As Double, CalcBal As Double
  Dim NumFundstoClose As Integer, ClosingThisFund As Boolean, F As Integer
  Dim FundDr As Double, FundCr As Double, FundRecNum As Integer
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer
  Dim FundIdxFileNum As Integer, NumFunds As Integer, EndDate As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, RecNo As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long, PageNum As Integer
  Dim Debit As String, Credit As String, Diff As Double, PYFundBal As Double
  ReDim FundList(1) As String
  DeActivateControls frmGLClosingOpMenu
  GetFundList FundList(), NumFunds
  'Define vars used for printing
      'make sure funds are in ascending order

  GetFundCodes FirstFund$, LastFund$
  ReportFile$ = "CLOSETB.PRN"                'Report File Name
  Header$ = "Closing Trial Balance"
  ReDim Desc$(1)
  Desc$(1) = "Acct Number     Title                                      Debit          Credit"
  CommaFmt$ = "###,###,###.##"  'format takes 13 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 14 chars
  SumLine$ = String$(14, "-")   'column summary line
  DivLine$ = String$(80, "-")   'dashed line
  DivLine2$ = String$(80, "=")  'Double Line

  CrLF$ = Chr$(13) + Chr$(10)
  FF$ = Chr$(12)
  MaxLines = 55
  TotDr# = 0
  TotCr# = 0

  ReDim FundCloseList$(1)
  GetFundCloseList FundCloseList$(), NumFundstoClose

  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile

'  QPrintRC Space$(80), 25, 1, -1
'  QPrintRC "Processing:", 25, 2, -1

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFileNum, NumFunds
  If NumGLAccts <> 0 Then
    FrmShowPctComp.Label1 = "Calculating Closing Trial Balance"
    FrmShowPctComp.cmdCancel.Enabled = False
    FrmShowPctComp.Show , Me
    DoEvents
  End If
  For cnt = 1 To NumGLAccts   'NumGLAccts
    FrmShowPctComp.ShowPctComp cnt, NumGLAccts
    Get AcctIdxFileNum, cnt, AcctIdx
    Get AcctFileNum, AcctIdx.RecNum, Acct

    FundCode$ = Left$(Acct.Num, GLFundLen)
    ClosingThisFund = False
    For F = 1 To NumFundstoClose
      If FundCode$ = QPTrim$(FundCloseList$(F)) Then
        ClosingThisFund = True
        Exit For
      End If
    Next

    'QPrintRC Acct.Num, 25, 14, -1

    If ClosingThisFund Then

      CalcBal# = Round#(Acct.BegBal)            'get the beginning balance

      NextTr& = Acct.FrstTran   'get the first trans for this acct

      Do Until NextTr& = 0      'keep going 'til we run out of trans
        Get TransFileNum, NextTr&, Trans

          If Trans.TRDATE <= FY1EndDate Then

            Select Case Acct.Typ
              Case "A", "E"
                CalcBal# = Round#(CalcBal# + Trans.DrAmt - Trans.CrAmt)
              Case "L", "R"
                CalcBal# = Round#(CalcBal# + Trans.CrAmt - Trans.DrAmt)
            End Select

          End If

          NextTr& = Trans.NextTran                'Get the next transaction

      Loop

      Acct.Bal = CalcBal#
      Acct.Work = CalcBal# 'hold this
      Put AcctFileNum, AcctIdx.RecNum, Acct

    End If   'test for account in fund range
  Next       'next account

  'PrintHelp "Please wait..."

  TotDr# = 0: TotCr# = 0
  FrmShowPctComp.Label1 = "Calculating Closing Trial Balance Fund Totals"
  FrmShowPctComp.Show , Me
  DoEvents

  For cnt = 1 To NumFunds
    FrmShowPctComp.ShowPctComp cnt, NumFunds
    FundDr# = 0: FundCr# = 0
    Get FundIdxFileNum, cnt, FundIdx

    FundNumber$ = QPTrim$(FundIdx.FundNum)
    ClosingThisFund = False
    For F = 1 To NumFundstoClose
      If FundNumber$ = QPTrim$(FundCloseList$(F)) Then
        ClosingThisFund = True
        Exit For
      End If
    Next

    If ClosingThisFund Then

      FundRecNum = FindFund(FundNumber$)
      FundName$ = QPTrim$(GetFundTitle$(FundRecNum))

      GoSub PrintPageHeader

      For RecNo = 1 To NumGLAccts             'Active Accts

        Get AcctIdxFileNum, RecNo, AcctIdx

        FundCode$ = Left$(AcctIdx.AcctNum, GLFundLen)
        If FundCode$ = FundNumber$ Then
          Get AcctFileNum, AcctIdx.RecNum, Acct

          Linecnt = Linecnt + 1
          If Linecnt >= MaxLines Then
             Print #PRNFile, FF$
             GoSub PrintPageHeader
          End If

          ToPrint$ = Space$(80)
          LSet ToPrint$ = Acct.Num
          Mid$(ToPrint$, 17) = Acct.Title
          Select Case Acct.Typ
          Case "A", "E"
            If Acct.Bal >= 0 Then
              Debit$ = Using$(CommaFmt$, Str$(Acct.Bal))
              TotDr# = TotDr# + Acct.Bal
              FundDr# = FundDr# + Acct.Bal
              Credit$ = ""
            Else
              Credit$ = Using$(CommaFmt$, Str$(Abs(Acct.Bal)))
              TotCr# = TotCr# + Abs(Acct.Bal)
              FundCr# = FundCr# + Abs(Acct.Bal)
              Debit$ = ""
            End If

          Case "L", "R"
            If Acct.Bal >= 0 Then
              Credit$ = Using$(CommaFmt$, Str$(Acct.Bal))
              TotCr# = TotCr# + Acct.Bal
              FundCr# = FundCr# + Acct.Bal
              Debit$ = ""
            Else
              Debit$ = Using$(CommaFmt$, Str$(Abs(Acct.Bal)))
              TotDr# = TotDr# + Abs(Acct.Bal)
              FundDr# = FundDr# + Abs(Acct.Bal)
              Credit$ = ""
            End If

          End Select

          Mid$(ToPrint$, 50) = Debit$
          Mid$(ToPrint$, 67) = Credit$
          Print #PRNFile, ToPrint$
        End If

      Next

      ToPrint$ = Space$(80)
      Mid$(ToPrint$, 50) = SumLine$
      Mid$(ToPrint$, 67) = SumLine$
      Print #PRNFile, ToPrint$

      ToPrint$ = Space$(80)
      LSet ToPrint$ = FundName$ + " " + "Totals"
      Mid$(ToPrint$, 48) = Using$(TotalFmt$, Str$(FundDr#))
      Mid$(ToPrint$, 65) = Using$(TotalFmt$, Str$(FundCr#))
      Print #PRNFile, ToPrint$

      Diff# = Round#(FundDr# - FundCr#)
      If Diff# <> 0 Then
        ToPrint$ = Space$(80)
        LSet ToPrint$ = "Fund is out of balance :"
        Mid$(ToPrint$, 25) = Using$(CommaFmt$, Str$(Diff#))
        Print #PRNFile, ToPrint$
      End If

      Print #PRNFile,
      Print #PRNFile, FF$

    Else
      'were not closing this fund
      FundRecNum = FindFund(FundNumber$)
      FundName$ = QPTrim$(GetFundTitle$(FundRecNum))
      GoSub PrintPageHeader
      Print #PRNFile, "Fund " + FundNumber$ + " " + FundName$ + " is not being CLOSED."
      Print #PRNFile, FF$
    End If    'if fund is in range test
  Next        'next fund

  ToPrint$ = Space$(80)

  '--Print a grand total if more than one fund
  If StartFund$ <> EndFund$ Then
    Print #PRNFile, "Combined totals - All Funds"
    Print #PRNFile, "Total Debits  : " + Using$(TotalFmt$, Str$(TotDr#))
    Print #PRNFile, "Total Credits : " + Using$(TotalFmt$, Str$(TotCr#))
    Print #PRNFile, Chr$(12)
  End If

  Close
  Load frmLoadingRpt
  'End Report Processing
  ActivateControls frmGLClosingOpMenu
  ViewPrint ReportFile$, Header$
  'KILL ReportFile$
Exit Sub

PrintPageHeader:
   Print #PRNFile, GLUserName$
   Print #PRNFile, FundName$ + " " + Header$
   Print #PRNFile, "Period Ending: " + Format(DateAdd("d", (FY1EndDate), "12-31-1979"), "mm/dd/yyyy")
   Print #PRNFile,
   Print #PRNFile, Desc$(1)
   Print #PRNFile, DivLine$
   Linecnt = 6
Return


End Sub


Private Sub GetFundCodes(FirstFund$, LastFund$)
  Dim FundIdxFileNum As Integer, NumFIdxRecs As Integer
   OpenFundIdx FundIdxFileNum, NumFIdxRecs

   If NumFIdxRecs = 0 Then
      FirstFund = 0
      LastFund = 0
      Exit Sub
   End If

   Get FundIdxFileNum, 1, FundIdx
   FirstFund$ = QPTrim$(FundIdx.FundNum)

   Get FundIdxFileNum, NumFIdxRecs, FundIdx
   LastFund$ = QPTrim$(FundIdx.FundNum)

   Close FundIdxFileNum

End Sub

Private Sub GetFundCloseList(FundCloseList$(), NumFundstoClose)
  Dim FundsToClose As GLFundCloseRecType
  Dim CloseListFileName As String, FundCloseListFile As Integer
  Dim RecLen As Integer, r As Integer
  CloseListFileName$ = "FCLOSE.LST"
  RecLen = Len(FundsToClose)
  FundCloseListFile = FreeFile
  Open CloseListFileName$ For Random As FundCloseListFile Len = RecLen
  NumFundstoClose = LOF(FundCloseListFile) \ RecLen

  If NumFundstoClose = 0 Then Exit Sub

  ReDim FundCloseList$(1 To NumFundstoClose)

  For r = 1 To NumFundstoClose
    Get FundCloseListFile, r, FundsToClose
    FundCloseList$(r) = FundsToClose.FundNum
  Next

  Close FundCloseListFile

End Sub

'****************************************************************************
'Retrieves the fund title from the fund data file.
'****************************************************************************
'
Private Function GetFundTitle$(FundRecNum)

   Dim FundRec As GLFundRecType
   Dim FundFileNum As Integer, NumFunds As Integer
   OpenFundFile FundFileNum, NumFunds
   Get FundFileNum, FundRecNum, FundRec
   GetFundTitle$ = FundRec.Title
   Close FundFileNum

End Function
'********************
'  ON LOCAL ERROR RESUME NEXT
  'Posts the fund balance adjustments to the accts file.
  'Must run prior to creating the opening entries for new year
  'backup files first
'************
Private Sub PostAdj(UnPosted As Integer)
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer
  Dim FundBalAdjFileName As String, FBARecLen As Integer
  Dim FundBalAdjFile As Integer, NumFBARecs As Integer
  Dim cnt As Integer, AcctRec As Integer, PreGo As String, oy As String
  oy$ = Right$(Format(DateAdd("d", (FY1EndDate), "12-31-1979"), "mm/dd/yyyy"), 2)
  FrmShowPctComp.Label1 = "Copying Accounting Data Files"
  FrmShowPctComp.Show , Me
  DoEvents

  'On Error Resume Next
    MkDir "PreClose" + oy$
  Call MainLog("Dir Made - PreClose" + oy$)
   'On Error GoTo 0
  'QPrintRC "Copying Files...", 5, 10, 15
  FrmShowPctComp.ShowPctComp 0, 5
  If Exist("Gltrans.dat") Then
    SH_CopyFile "gl*.dat", "PreClose" + oy$, , True
  End If
  FrmShowPctComp.ShowPctComp 1, 5
  If Exist("Glacct.idx") Then
    SH_CopyFile "gl*.idx", "PreClose" + oy$, , True
  End If
  FrmShowPctComp.ShowPctComp 2, 5
  If Exist("bgttrans.dat") Then
    SH_CopyFile "bgt*.dat", "PreClose" + oy$, , True
  End If
  FrmShowPctComp.ShowPctComp 3, 5
  If Exist("apvendor.dat") Then
    SH_CopyFile "ap*.dat", "PreClose" + oy$, , True
  End If
  FrmShowPctComp.ShowPctComp 4, 5
  If Exist("POTRANS.DAT") Then
    SH_CopyFile "po*.dat", "PreClose" + oy$, , True
  End If
  FrmShowPctComp.ShowPctComp 5, 5
 ' QPrintRC "Posting Adjustments...", 5, 10, 15
  'PrintHelp "Wait"
  Call MainLog("Copied GL,BG,AP,PO data")
  ReDim Accts(1) As GLAcctRecType
  OpenAcctFile AcctFileNum, NumGLAcctRecs

  FundBalAdjFileName$ = "FBADJ.DAT"
  Dim GLFBAdj As GLFBAdjRecType
  FBARecLen = Len(GLFBAdj)
  FundBalAdjFile = FreeFile
  Open FundBalAdjFileName$ For Random As FundBalAdjFile Len = FBARecLen
  NumFBARecs = LOF(FundBalAdjFile) \ FBARecLen
  FrmShowPctComp.Label1 = "Posting Adjustments....."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DoEvents


  For cnt = 1 To NumFBARecs
    FrmShowPctComp.ShowPctComp cnt, NumFBARecs

    Get FundBalAdjFile, cnt, GLFBAdj
      AcctRec = AcctFind(QPTrim$(GLFBAdj.AcctNum))
      If AcctRec > 0 Then
        Get AcctFileNum, AcctRec, Accts(1)
        Accts(1).Bal = Round#(Accts(1).Bal + GLFBAdj.AdjAmt)
        Put AcctFileNum, AcctRec, Accts(1)
      Else
        UnPosted = UnPosted + 1
      End If
  Next
  Call MainLog("Adjustments applied to Acct Balances")
End Sub
Private Sub MakeOpenEntries()
  Dim fmt As String, FF As String, TransRecLen As Integer
  Dim OEFileNum As Integer, AcctIdxFileNum As Integer, NumGLAccts As Integer
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, FundCode As String
  Dim NumFundstoClose As Integer, PRNFile As Integer, ReportFile As String
  Dim ToPrint As String, FirstTime As Boolean, cnt As Integer
  Dim ClosingThisFund As Boolean, F As Integer, RecCnt As Integer
  Dim TotDr As Double, TotCr As Double, OutofBal As Double
  'ToPrint$ = "Creating Opening Entries"
  'QPrintRC ToPrint$, 5, 10, 15

  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  fmt$ = "#,###,###,###.##"
  FF$ = Chr$(12)
  KillFileCO "OPNENTRY.DAT"
  ReDim Accts(1) As GLAcctRecType
  ReDim AcctIndex(1) As GLAcctIndexType
  ReDim OpenEntry(1) As GLTransRecType

  TransRecLen = Len(OpenEntry(1))
  OEFileNum = FreeFile
  Open "OPNENTRY.DAT" For Random As OEFileNum Len = TransRecLen
  'NumTrans& = LOF(TransFileNum) \ TransRecLen

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  'REDIM FundList$(1)                     'List of all active Funds
  'GetFundList FundList$(), NumFunds

  ReDim FundCloseList$(1)                'List of Funds to close
  GetFundCloseList FundCloseList$(), NumFundstoClose

  'FundBalAdjFileName$ = "FBADJ.DAT"
  'IF Exist(FundBalAdjFileName$) THEN
  '   KILL FundBalAdjFileName$
  'END IF

  'DIM GLFBAdj AS GLFBAdjRecType
  'FBARecLen = LEN(GLFBAdj)
  'FundBalAdjFile = FREEFILE
  'OPEN FundBalAdjFileName$ FOR RANDOM AS FundBalAdjFile LEN = FBARecLen

  'REDIM FundTotRev#(1 TO NumFunds)       'List of total revenues by fund
  'REDIM FundTotExp#(1 TO NumFunds)       'list of tot exp by fund
  'REDIM FundBalAdj#(1 TO NumFunds)
  KillFileCO "OPNENTRY.PRN"
  PRNFile = FreeFile
  ReportFile$ = "OPNENTRY.PRN"
  Open ReportFile$ For Output As #PRNFile
  ToPrint$ = ""

  'GoSub PrnOEHeader

  OpenEntry(1).TRDATE = FY2BegDate
  OpenEntry(1).Desc = "Opening Entry"
  OpenEntry(1).Ref = "Sys"
  OpenEntry(1).Src = "GJ" + Left$(Date$, 2) + Mid$(Date$, 4, 2) + Right$(Date$, 2)
  OpenEntry(1).NextTran = 0

  FirstTime = True
  FrmShowPctComp.Label1 = "Creating Opening Entries"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DoEvents


  For cnt = 1 To NumGLAccts   'NumGLAccts
    FrmShowPctComp.ShowPctComp cnt, NumGLAccts

    Get AcctIdxFileNum, cnt, AcctIndex(1)
    Get AcctFileNum, AcctIndex(1).RecNum, Accts(1)

    FundCode$ = Left$(Accts(1).Num, GLFundLen)
    ClosingThisFund = False
    For F = 1 To NumFundstoClose
      If FundCode$ = QPTrim$(FundCloseList$(F)) Then
        ClosingThisFund = True
        Exit For
      End If
    Next

    If ClosingThisFund Then
      If FirstTime Then
        FirstTime = False
        LastFund$ = FundCode$
      End If

      If LastFund$ <> FundCode$ Then
        GoSub SubTotalFund
        'GoSub PrnOEHeader
      End If

      If Round#(Accts(1).Bal) <> 0 Then
        Select Case Accts(1).Typ
          Case "A"
            If Accts(1).Bal > 0 Then
              OpenEntry(1).DrAmt = Round#(Accts(1).Bal)
              OpenEntry(1).CrAmt = 0
            Else
              OpenEntry(1).DrAmt = 0
              OpenEntry(1).CrAmt = Round#(Abs(Accts(1).Bal))
            End If
            RecCnt = RecCnt + 1
            OpenEntry(1).AcctNum = Accts(1).Num
            GoSub WriteToInterfaceFile
            GoSub PrintEntry
            LastFund$ = Left$(Accts(1).Num, GLFundLen)
          Case "L"
            If Accts(1).Bal > 0 Then
              OpenEntry(1).DrAmt = 0
              OpenEntry(1).CrAmt = Round#(Accts(1).Bal)
            Else
              OpenEntry(1).DrAmt = Round#(Abs(Accts(1).Bal))
              OpenEntry(1).CrAmt = 0
            End If
            RecCnt = RecCnt + 1
            OpenEntry(1).AcctNum = Accts(1).Num
            GoSub WriteToInterfaceFile
            GoSub PrintEntry
            LastFund$ = Left$(Accts(1).Num, GLFundLen)
          Case Else
        End Select
      End If    'Account has balance
    End If  'Fund being closed
  Next

  GoSub SubTotalFund
  Call MainLog("Opening Entries Created")
  Close
  'ViewPrint ReportFile$, "G/L Opening Entries"
  Load frmLoadingRpt
  ActivateControls frmGLClosingOpMenu
  ARptCloseTrialBal.txtRptDate = "Fiscal Year Beginning: " + Format(DateAdd("d", (FY2BegDate), "12-31-1979"), "mm/dd/yyyy")
  ARptCloseTrialBal.Title.Caption = "Opening Entries"
  ARptCloseTrialBal.txtDate = Now
  ARptCloseTrialBal.txtTown = GLUserName$
  ARptCloseTrialBal.GetName ReportFile$
  ARptCloseTrialBal.startrpt


Exit Sub

PrintEntry:
  ToPrint$ = FundCode$ + "~~" + Accts(1).Num
  ToPrint$ = ToPrint$ + "~" + Accts(1).Title
  ToPrint$ = ToPrint$ + "~" + Using$(fmt$, Str$(OpenEntry(1).DrAmt))
  ToPrint$ = ToPrint$ + "~" + Using$(fmt$, Str$(OpenEntry(1).CrAmt))
  Print #PRNFile, ToPrint$
  
  TotDr# = Round#(TotDr# + OpenEntry(1).DrAmt)
  TotCr# = Round#(TotCr# + OpenEntry(1).CrAmt)
Return

SubTotalFund:

  OutofBal# = Round(TotDr# - TotCr#)
  If OutofBal# <> 0 Then
    'rint #PRNFile, ToPrint$
  End If

 'Print #PRNFile, FF$

 ' LSet ToPrint$ = ""
  FirstTime = True
  TotDr# = 0
  TotCr# = 0

Return

WriteToInterfaceFile:
  Put OEFileNum, RecCnt, OpenEntry(1)
Return


End Sub

Private Sub MakeOpenEntries2()
  Dim fmt As String, FF As String, TransRecLen As Integer
  Dim OEFileNum As Integer, AcctIdxFileNum As Integer, NumGLAccts As Integer
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, FundCode As String
  Dim NumFundstoClose As Integer, PRNFile As Integer, ReportFile As String
  Dim ToPrint As String, FirstTime As Boolean, cnt As Integer
  Dim ClosingThisFund As Boolean, F As Integer, RecCnt As Integer
  Dim TotDr As Double, TotCr As Double, OutofBal As Double
  'ToPrint$ = "Creating Opening Entries"
  'QPrintRC ToPrint$, 5, 10, 15

  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  fmt$ = "#,###,###,###.##"
  FF$ = Chr$(12)
  KillFileCO "OPNENTRY.DAT"
  ReDim Accts(1) As GLAcctRecType
  ReDim AcctIndex(1) As GLAcctIndexType
  ReDim OpenEntry(1) As GLTransRecType

  TransRecLen = Len(OpenEntry(1))
  OEFileNum = FreeFile
  Open "OPNENTRY.DAT" For Random As OEFileNum Len = TransRecLen
  'NumTrans& = LOF(TransFileNum) \ TransRecLen

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  'REDIM FundList$(1)                     'List of all active Funds
  'GetFundList FundList$(), NumFunds

  ReDim FundCloseList$(1)                'List of Funds to close
  GetFundCloseList FundCloseList$(), NumFundstoClose

  'FundBalAdjFileName$ = "FBADJ.DAT"
  'IF Exist(FundBalAdjFileName$) THEN
  '   KILL FundBalAdjFileName$
  'END IF

  'DIM GLFBAdj AS GLFBAdjRecType
  'FBARecLen = LEN(GLFBAdj)
  'FundBalAdjFile = FREEFILE
  'OPEN FundBalAdjFileName$ FOR RANDOM AS FundBalAdjFile LEN = FBARecLen

  'REDIM FundTotRev#(1 TO NumFunds)       'List of total revenues by fund
  'REDIM FundTotExp#(1 TO NumFunds)       'list of tot exp by fund
  'REDIM FundBalAdj#(1 TO NumFunds)
  KillFileCO "OPNENTRY.PRN"
  PRNFile = FreeFile
  ReportFile$ = "OPNENTRY.PRN"
  Open ReportFile$ For Output As #PRNFile
  ToPrint$ = Space$(80)

  GoSub PrnOEHeader

  OpenEntry(1).TRDATE = FY2BegDate
  OpenEntry(1).Desc = "Opening Entry"
  OpenEntry(1).Ref = "Sys"
  OpenEntry(1).Src = "GJ" + Left$(Date$, 2) + Mid$(Date$, 4, 2) + Right$(Date$, 2)
  OpenEntry(1).NextTran = 0

  FirstTime = True
  FrmShowPctComp.Label1 = "Creating Opening Entries"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DoEvents


  For cnt = 1 To NumGLAccts   'NumGLAccts
    FrmShowPctComp.ShowPctComp cnt, NumGLAccts

    Get AcctIdxFileNum, cnt, AcctIndex(1)
    Get AcctFileNum, AcctIndex(1).RecNum, Accts(1)

    FundCode$ = Left$(Accts(1).Num, GLFundLen)
    ClosingThisFund = False
    For F = 1 To NumFundstoClose
      If FundCode$ = QPTrim$(FundCloseList$(F)) Then
        ClosingThisFund = True
        Exit For
      End If
    Next

    If ClosingThisFund Then
      If FirstTime Then
        FirstTime = False
        LastFund$ = FundCode$
      End If

      If LastFund$ <> FundCode$ Then
        GoSub SubTotalFund
        GoSub PrnOEHeader
      End If

      If Round#(Accts(1).Bal) <> 0 Then
        Select Case Accts(1).Typ
          Case "A"
            If Accts(1).Bal > 0 Then
              OpenEntry(1).DrAmt = Round#(Accts(1).Bal)
              OpenEntry(1).CrAmt = 0
            Else
              OpenEntry(1).DrAmt = 0
              OpenEntry(1).CrAmt = Round#(Abs(Accts(1).Bal))
            End If
            RecCnt = RecCnt + 1
            OpenEntry(1).AcctNum = Accts(1).Num
            GoSub WriteToInterfaceFile
            GoSub PrintEntry
            LastFund$ = Left$(Accts(1).Num, GLFundLen)
          Case "L"
            If Accts(1).Bal > 0 Then
              OpenEntry(1).DrAmt = 0
              OpenEntry(1).CrAmt = Round#(Accts(1).Bal)
            Else
              OpenEntry(1).DrAmt = Round#(Abs(Accts(1).Bal))
              OpenEntry(1).CrAmt = 0
            End If
            RecCnt = RecCnt + 1
            OpenEntry(1).AcctNum = Accts(1).Num
            GoSub WriteToInterfaceFile
            GoSub PrintEntry
            LastFund$ = Left$(Accts(1).Num, GLFundLen)
          Case Else
        End Select
      End If    'Account has balance
    End If  'Fund being closed
  Next

  GoSub SubTotalFund
  Call MainLog("Opening Entries Created")
  Close
  ViewPrint ReportFile$, "G/L Opening Entries"

Exit Sub

PrnOEHeader:
  Print #PRNFile, GLUserName$
  Print #PRNFile, "Opening Entries"
  Print #PRNFile, "Fiscal Year Beginning "; Format(DateAdd("d", (FY2BegDate), "12-31-1979"), "mm/dd/yyyy")

  Print #PRNFile,
  Print #PRNFile, "Acct"
  Print #PRNFile, "-----------------------------------------------------------"
Return
PrintEntry:
  LSet ToPrint$ = Accts(1).Num
  Mid$(ToPrint$, 17) = Accts(1).Title
  Mid$(ToPrint$, 49) = Using$(fmt$, Str$(OpenEntry(1).DrAmt))
  Mid$(ToPrint$, 65) = Using$(fmt$, Str$(OpenEntry(1).CrAmt))
  Print #PRNFile, ToPrint$
  LSet ToPrint$ = ""
  TotDr# = Round#(TotDr# + OpenEntry(1).DrAmt)
  TotCr# = Round#(TotCr# + OpenEntry(1).CrAmt)
Return

SubTotalFund:

  LSet ToPrint$ = ""
  Print #PRNFile, ToPrint$
  LSet ToPrint$ = "Fund Totals"
  Mid$(ToPrint$, 49) = Using$(fmt$, Str$(TotDr#))
  Mid$(ToPrint$, 65) = Using$(fmt$, Str$(TotCr#))
  Print #PRNFile, ToPrint$

  LSet ToPrint$ = ""
  Print #PRNFile, ToPrint$
  OutofBal# = Round(TotDr# - TotCr#)
  If OutofBal# <> 0 Then
    Print #PRNFile, ToPrint$
    LSet ToPrint$ = "Fund is Out of Balance :" + Using$(fmt$, Str$(OutofBal#))
    Print #PRNFile, ToPrint$
  End If

  Print #PRNFile, FF$

  LSet ToPrint$ = ""
  FirstTime = True
  TotDr# = 0
  TotCr# = 0

Return

WriteToInterfaceFile:
  Put OEFileNum, RecCnt, OpenEntry(1)
Return


End Sub
Private Sub SplitTransFile()
  Dim BgtTrans As GLTransRecType
  Dim Tr As GLTransRecType
  Dim Y1Trans As GLTransRecType
  Dim Y2Trans As GLTransRecType
  Dim Bgtr As GLTransRecType
  Dim Y1BGt As GLTransRecType
  Dim Y2Bgt As GLTransRecType
  Dim BgRec As Long, Bgtreclen As Integer, BgtTransFile As Integer
  Dim TotLen As Integer, Yr As String, NumFundstoClose As Integer
  Dim TransRecLen As Integer, GLTransFile As Integer, NumTrans As Long
  Dim Y1TransFile As Integer, Y2TransFile As Integer, TRRec As Long
  Dim FundCode As String, ClosingThisFund As Boolean, F As Integer
  Dim Y1Recs As Long, Y2Recs As Long, HistDir As String, tstdir As String
  Dim NumBgTrans As Long, Y1BGFile As Integer, Y2BGFile As Integer
  Dim Y1BgRecs As Long, Y2BgRecs As Long
 'QPrintRC "Updating Transaction Files.", 5, 10, 15
  'PrintHelp "Parsing Histories..."
  Dim Part As Double
  Part = Timer
   

  TotLen% = GLFundLen + GLAcctLen + GLDetLen
  Yr$ = Right$(Format(DateAdd("d", (FY1EndDate), "12-31-1979"), "mm/dd/yyyy"), 4)
  '--List of Funds to close
  'Clean up old closing files
  KillFileCO "GLTRANS.Y1"
  KillFileCO "GLTRANS.Y2"
  KillFileCO "gltrans.d97"
  ReDim FundCloseList$(1)
  GetFundCloseList FundCloseList$(), NumFundstoClose
  TransRecLen = Len(Trans)
  OpenTransFile GLTransFile, NumTrans&
  Y1TransFile = FreeFile
  Open "GLTRANS.Y1" For Random As Y1TransFile Len = TransRecLen
  Y2TransFile = FreeFile
  Open "GLTRANS.Y2" For Random As Y2TransFile Len = TransRecLen
  FrmShowPctComp.Label1 = "Updating Transaction Files."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DoEvents
  For TRRec& = 1 To NumTrans&
    FrmShowPctComp.ShowPctComp TRRec&, NumTrans&
    Get GLTransFile, TRRec&, Tr
    'Complete# = (TRRec& / NumTrans&) * 100
    'LOCATE 6, 10
    'Color 15
    'Print Using; "Processing files.  ###% Complete.        "; Complete#

    FundCode$ = Left$(Tr.AcctNum, GLFundLen)
    ClosingThisFund = False
    For F = 1 To NumFundstoClose
      If FundCode$ = QPTrim$(FundCloseList$(F)) Then
        ClosingThisFund = True
        Exit For
      End If
    Next
'Year 1 is year closing, Year 2 is if have trans in 2nd year that will be unclosed
    If Tr.TRDATE <= FY1EndDate And ClosingThisFund Then
      '--copy Year 1 trans
      Y1Recs& = Y1Recs& + 1
      Y1Trans.AcctNum = Tr.AcctNum
      Y1Trans.TRDATE = Tr.TRDATE
      Y1Trans.Desc = Tr.Desc
      Y1Trans.LDesc = Tr.LDesc
      Y1Trans.CrAmt = Tr.CrAmt
      Y1Trans.DrAmt = Tr.DrAmt
      Y1Trans.Ref = Tr.Ref
      Y1Trans.Src = Tr.Src
      Y1Trans.NextTran = 0 'Tr.NextTran
      Put Y1TransFile, Y1Recs&, Y1Trans
    Else
      '--copy Year 2 trans
      Y2Recs& = Y2Recs& + 1
      Y2Trans.AcctNum = Tr.AcctNum
      Y2Trans.TRDATE = Tr.TRDATE
      Y2Trans.Desc = Tr.Desc
      Y2Trans.LDesc = Tr.LDesc
      Y2Trans.CrAmt = Tr.CrAmt
      Y2Trans.DrAmt = Tr.DrAmt
      Y2Trans.Ref = Tr.Ref
      Y2Trans.Src = Tr.Src
      Y2Trans.NextTran = 0  'Will be reposted
      Put Y2TransFile, Y2Recs&, Y2Trans
    End If

  Next

  Close

  '--save the original transaction file

  Name "gltrans.dat" As "gltrans.d97"

  '--rename the gltrans.y1 file to .dat the relink y1
  Name "gltrans.y1" As "gltrans.dat"

  'Year 1 data now has .dat file extensions. Call relink to tie
  'everything together so we can run reports accurately in the future.
  ReLinkTrans frmGLClosingOpMenu, True
  
  
  
  FrmShowPctComp.Label1 = "Creating History Directory"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DoEvents

  Call MainLog("years split")
  'QPrintRC Space$(80), 1, 1, 112
  HistDir$ = "GL" + Yr$
  tstdir = Dir(HistDir$, vbDirectory)
  If tstdir = "" Then
    MkDir HistDir$
  Else
    'MsgBox "Prior Year/End Directory Will Be Renamed.", vbOKOnly, "Prior Y/E Dir"
    'Had to rem out line above so wouldn't stop processing...
    HistDir$ = "GL" + QPTrim(Str(CLng(Part)))
    'HistDir$ = "GLNw2002"
    MkDir HistDir$
  End If
  FrmShowPctComp.ShowPctComp 1, 4
  '--after linking copy the *.dat to the archive directory
  If Exist("Gltrans.dat") Then
    SH_CopyFile "gl*.dat", HistDir$, , True
  End If
  FrmShowPctComp.ShowPctComp 2, 4
  If Exist("glacct.idx") Then
    SH_CopyFile "gl*.idx", HistDir$, , True
  End If
  FrmShowPctComp.ShowPctComp 3, 4
  If Exist("bgttrans.dat") Then
    SH_CopyFile "bgt*.dat", HistDir$
  End If
  If Exist("POTRANS.DAT") Then
    SH_CopyFile "po*.dat", HistDir$
  End If
  FrmShowPctComp.ShowPctComp 4, 4
  Call MainLog(HistDir$ + " created,data copied - gl,bg,po")
  
  '~~~~~~~~~~~~~~~~~~~~~
'BUDGET STUFF
  KillFileCO "bgtTRANS.Y1"
  KillFileCO "bgtTRANS.Y2"
  KillFileCO "bgttrans.d97"
  ReDim FundCloseList$(1)
  GetFundCloseList FundCloseList$(), NumFundstoClose
  Bgtreclen = Len(BgtTrans)
  OpenBgtTransFile BgtTransFile, NumBgTrans&
  Y1BGFile = FreeFile
  Open "BGtTRANS.Y1" For Random As Y1BGFile Len = Bgtreclen
  Y2BGFile = FreeFile
  Open "BGtTRANS.Y2" For Random As Y2BGFile Len = Bgtreclen
  FrmShowPctComp.Label1 = "Updating Budget Files."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DoEvents
  For BgRec& = 1 To NumBgTrans&
    FrmShowPctComp.ShowPctComp BgRec&, NumBgTrans&
    Get BgtTransFile, BgRec&, Bgtr
    'Complete# = (TRRec& / NumTrans&) * 100
    'LOCATE 6, 10
    'Color 15
    'Print Using; "Processing files.  ###% Complete.        "; Complete#

    FundCode$ = Left$(Bgtr.AcctNum, GLFundLen)
    ClosingThisFund = False
    For F = 1 To NumFundstoClose
      If FundCode$ = QPTrim$(FundCloseList$(F)) Then
        ClosingThisFund = True
        Exit For
      End If
    Next
'Year 1 is year closing, Year 2 is if have trans in 2nd year that will be unclosed
    If Bgtr.TRDATE <= FY1EndDate And ClosingThisFund Then
      '--copy Year 1 trans
      Y1BgRecs& = Y1BgRecs& + 1
      Y1BGt.AcctNum = Bgtr.AcctNum
      Y1BGt.TRDATE = Bgtr.TRDATE
      Y1BGt.Desc = Bgtr.Desc
      Y1BGt.CrAmt = Bgtr.CrAmt
      Y1BGt.DrAmt = Bgtr.DrAmt
      Y1BGt.Ref = Bgtr.Ref
      Y1BGt.Src = Bgtr.Src
      Y1BGt.NextTran = 0 'Tr.NextTran
      Put Y1BGFile, Y1BgRecs&, Y1BGt
    Else
      '--copy Year 2 trans
      Y2BgRecs& = Y2BgRecs& + 1
      Y2Bgt.AcctNum = Bgtr.AcctNum
      Y2Bgt.TRDATE = Bgtr.TRDATE
      Y2Bgt.Desc = Bgtr.Desc
      Y2Bgt.CrAmt = Bgtr.CrAmt
      Y2Bgt.DrAmt = Bgtr.DrAmt
      Y2Bgt.Ref = Bgtr.Ref
      Y2Bgt.Src = Bgtr.Src
      Y2Bgt.NextTran = 0  'Will be reposted
      Put Y2BGFile, Y2BgRecs&, Y2Bgt
    End If

  Next

  Close

  '--save the original transaction file
  
  Name "bgttrans.dat" As "bgttrans.d97"

  '--rename the gltrans.y1 file to .dat the relink y1
  Name "bgttrans.y1" As "bgttrans.dat"

'~~~~~~~~~~~~~~~~~~~~~~~~~

  KillFileCO "BGtTRANS.DAT"
  KillFileCO "GLTRANS.DAT"

  Name "BGTtRANS.Y2" As "BGTtRANS.DAT"
  RelinkBgtTrans frmGLClosingOpMenu, True

  'NewBgtFileName$ = "BGTTRANS.D" + Yr$
  'IF Exist("BGTTRANS.DAT") THEN
  '  NAME "BGTTRANS.DAT" AS NewBgtFileName$
  'END IF

  '--rename y2 to current & new year transactions
  'NAME "gltrans.y2" AS "gltrans.dat"
  'ReLinkTrans

  '--opening balances and y2 transactions will be posted
  'this is a copy of y1 trans file so kill it.  New trans will
  'be created when opening entries and gltrans.y2 files are
  'posted

  'IF Exist("BGTTRANS.DAT") THEN
  '  KILL "BGTTRANS.DAT"
  'END IF

End Sub

'*******************************************************************
  'operations to update the system for the new year.
  'last operation to be called.  Year one files are archived.

  'before this operation glacct is a copy of the file which is
  'linked to last years transactions.
'*******************************************************************
Private Sub UpDateNewYear()
  Dim NumFundstoClose As Integer, AcctFileNum As Integer, NumGLAcctRecs As Integer
  Dim BgtTransFile As Integer, BgtTransRecLen As Integer, r As Integer
  Dim FundCode As String, ClosingThisFund As Boolean, F As Integer
  Dim BRec As Long
  '--List of Funds to close
  ReDim FundCloseList$(1)
  GetFundCloseList FundCloseList$(), NumFundstoClose

  FrmShowPctComp.Label1 = "Updating System For New Year."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DoEvents

  'QPrintRC "Updating files for new year", 5, 10, 15

  ReDim Accts(1) As GLAcctRecType
  OpenAcctFile AcctFileNum, NumGLAcctRecs

  ReDim BgtTrn(1) As GLTransRecType

  BgtTransFile = FreeFile
  BgtTransRecLen = Len(BgtTrn(1))
  Open "BgtTrans.DAT" For Random As BgtTransFile Len = BgtTransRecLen
  BRec = LOF(BgtTransFile) \ BgtTransRecLen
  'Set the vars for the budget transaction history
  BgtTrn(1).TRDATE = FY2BegDate
  BgtTrn(1).Desc = "Appropriation"
  BgtTrn(1).Ref = "Sys"
  BgtTrn(1).Src = "GJ" + Left$(Date$, 2) + Mid$(Date$, 4, 2) + Right$(Date$, 2)
  BgtTrn(1).NextTran = 0
  BgtTrn(1).Marked = 0

  For r = 1 To NumGLAcctRecs
    FrmShowPctComp.ShowPctComp r, NumGLAcctRecs

    Get AcctFileNum, r, Accts(1)
    'IF LEFT$(Accts(1).Num, 9) = "70-151-01" THEN STOP
    FundCode$ = Left$(Accts(1).Num, GLFundLen)
    ClosingThisFund = False
    For F = 1 To NumFundstoClose
      If FundCode$ = QPTrim$(FundCloseList$(F)) Then
        ClosingThisFund = True
        Exit For
      End If
    Next

   If ClosingThisFund Then
    'm$ = "Updating Account: " + Accts(1).Num
    'QPrintRC m$, 6, 10, -1

    '--create budget history file using the acct n/y approved budget field
    If Accts(1).Typ = "R" Or Accts(1).Typ = "E" Then
     If Not Accts(1).Deleted Then
      If Round#(Accts(1).NYApp) > -99999999.99 And Round#(Accts(1).NYApp) < 99999999.99 Then
       If Round#(Accts(1).NYApp) <> 0 Then
        BRec = BRec + 1
        BgtTrn(1).AcctNum = Accts(1).Num
        Select Case Accts(1).Typ
          Case "R"
            If Accts(1).NYApp >= 0 Then
              BgtTrn(1).CrAmt = Accts(1).NYApp
              BgtTrn(1).DrAmt = 0
            Else
              BgtTrn(1).CrAmt = 0
              BgtTrn(1).DrAmt = Abs(Accts(1).NYApp)
            End If
          Case "E"
            If Accts(1).NYApp >= 0 Then
              BgtTrn(1).DrAmt = Accts(1).NYApp
              BgtTrn(1).CrAmt = 0
            Else
              BgtTrn(1).DrAmt = 0
              BgtTrn(1).CrAmt = Abs(Accts(1).NYApp)
           End If
          Case Else
        End Select
        Put BgtTransFile, BRec, BgtTrn(1)
       End If
      End If
     End If
    End If

    Accts(1).FrstTran = 0
    Accts(1).LastTran = 0
    Accts(1).PYAct = Accts(1).Work
    Accts(1).BegBal = 0
    Accts(1).Bgt = 0
    Accts(1).Bal = 0
    Accts(1).Encumb = 0
    Accts(1).MTD = 0
    Accts(1).YTD = 0
    Accts(1).NYEst = 0
    Accts(1).NYReq = 0
    Accts(1).NYRec = 0
    Accts(1).NYApp = 0
    Accts(1).FrstBTran = 0
    Accts(1).LastBTran = 0
    Accts(1).FrstPTran = 0
    Accts(1).LastPTran = 0
    'Accts(1).Work  'copy to p/y then 0
    Accts(1).Marked = 0
    Put AcctFileNum, r, Accts(1)

    'reset work field
    Accts(1).Work = 0
    Put AcctFileNum, r, Accts(1)
   End If
  Next
  Close
  Call MainLog("Budget updated, Acct approved updated")
'************************
'acct field names
'Deleted
'Num
'Title
'Typ
'FrstTran
'LastTran
'PYAct
'BegBal
'Bgt
'Bal
'Encumb
'MTD
'YTD
'NYEst
'NYReq
'NYRec
'NYApp
'FrstBTran
'LastBTran
'FrstPTran
'LastPTran
'Work
'Res
'Marked

End Sub

Private Sub RepostNewYearTrans()

     ' QPrintRC "Posting Opening Entries...", 5, 10, 15
      Post2GL "OPNENTRY.DAT", BadTrans, frmGLClosingOpMenu, True
      'QPrintRC "Updating Transactions...  ", 5, 10, 15
      Post2GL "gltrans.y2", BadTrans, frmGLClosingOpMenu, True
      Call MainLog("Opening Entries and glnewyr Posted")
End Sub

Private Sub ResetYears()
  Dim GLSetup As GLSetupRecType
  Dim SetUpRecLen As Integer, SetupFile As Integer
  Dim FY1B As String, FY1BYr As String, NewYear As Integer, NF1YB As String
  Dim FY1E As String, FY1EYr As String, NF1YE As String
  Dim FY2B As String, FY2BYr As String, NF2YB As String
  Dim FY2E As String, FY2EYr As String, NF2YE As String
   SetUpRecLen = Len(GLSetup)

   SetupFile = FreeFile
   Open "GLSETUP.DAT" For Random As SetupFile Len = SetUpRecLen

   Get SetupFile, 1, GLSetup

   FY1BegDate = GLSetup.FYBeg
   FY1EndDate = GLSetup.FYEnd
   FY2BegDate = GLSetup.NYBeg
   FY2EndDate = GLSetup.NYEnd

   FY1B$ = Format(DateAdd("d", (FY1BegDate), "12-31-1979"), "mm/dd/yyyy")
   FY1BYr$ = Right$(FY1B$, 4)
   NewYear = Val(FY1BYr$) + 1
   NF1YB$ = Left$(FY1B$, 6) + QPTrim$(Str$(NewYear))

   FY1E$ = Format(DateAdd("d", (FY1EndDate), "12-31-1979"), "mm/dd/yyyy")
   FY1EYr$ = Right$(FY1E$, 4)
   NewYear = Val(FY1EYr$) + 1
   NF1YE$ = Left$(FY1E$, 6) + QPTrim$(Str$(NewYear))

   FY2B$ = Format(DateAdd("d", (FY2BegDate), "12-31-1979"), "mm/dd/yyyy")
   FY2BYr$ = Right$(FY2B$, 4)
   NewYear = Val(FY2BYr$) + 1
   NF2YB$ = Left$(FY2B$, 6) + QPTrim$(Str$(NewYear))

   FY2E$ = Format(DateAdd("d", (FY2EndDate), "12-31-1979"), "mm/dd/yyyy")
   FY2EYr$ = Right$(FY2E$, 4)
   NewYear = Val(FY2EYr$) + 1
   NF2YE$ = Left$(FY2E$, 6) + QPTrim$(Str$(NewYear))

   GLSetup.FYBeg = DateDiff("d", "12/31/1979", NF1YB$)
   GLSetup.FYEnd = DateDiff("d", "12/31/1979", NF1YE$)
   GLSetup.NYBeg = DateDiff("d", "12/31/1979", NF2YB$)
   GLSetup.NYEnd = DateDiff("d", "12/31/1979", NF2YE$)

   Put SetupFile, 1, GLSetup

   Close SetupFile
   Call MainLog("Old Fiscar Yr SET - " + FY1B$ + "," + FY1E$ + "," + FY2B$ + "," + FY2E$)
   Call MainLog("Fiscal Years Updated To - " + NF1YB$ + "," + NF1YE$ + "," + NF2YB$ + "," + NF2YE$)
End Sub

Private Sub PurgePOs()
  Dim PO As GLTransRecType
  Dim PY As GLTransRecType
  Dim TotLen As Integer, Yr As String
  Dim TransRecLen As Integer, POTransFile As Integer, NumTrans As Long
  Dim PYTransFile As Integer, PORec As Long
  Dim FundCode As String, ClosingThisFund As Boolean, F As Integer
  Dim PYRec As Long, HistDir As String, tstdir As String
  'QPrintRC "Updating Transaction Files.", 5, 10, 15
  'PrintHelp "Parsing Histories..."
  TotLen% = GLFundLen + GLAcctLen + GLDetLen
  Yr$ = Right$(Format(DateAdd("d", (FY1BegDate), "12-31-1979"), "mm/dd/yyyy"), 4)
  '--List of Funds to close
  'Clean up old closing files
  KillFileCO "POTrans.PY"
  KillFileCO "potrans.oyr"
  TransRecLen = Len(PO)
  POTransFile = FreeFile
  Open "POTRANS.DAT" For Random Access Read Write Shared As POTransFile Len = TransRecLen
  NumTrans& = LOF(POTransFile) \ TransRecLen

  PYTransFile = FreeFile
  Open "POTRANS.PY" For Random As PYTransFile Len = TransRecLen
  
  FrmShowPctComp.Label1 = "Updating PO Transaction Files."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DoEvents
  For PORec& = 1 To NumTrans&
    FrmShowPctComp.ShowPctComp PORec&, NumTrans&
    Get POTransFile, PORec&, PO
    If PO.TRDATE >= FY1BegDate Then
'only copy records for current years
      PYRec& = PYRec& + 1
'      PY.AcctNum = PO.AcctNum
'      PY.AcctRec = PO.AcctRec
'      PY.CrAmt = PO.CrAmt
'      PY.DESC = PO.DESC
'      PY.DrAmt = PO.DrAmt
'      PY.Marked = PO.Marked
'      PY.NextTran = 0
'      PY.Ref = PO.Ref
'      PY.Res = PO.Res
'      PY.Src = PO.Src
'      PY.TRDATE = PO.TRDATE
      LSet PY = PO
      Put PYTransFile, PYRec&, PY
    End If

  Next

  Close

  '--save the original transaction file

  Name "potrans.dat" As "potrans.oyr"

  '--rename the potrans.py file to .dat then relink
  Name "potrans.py" As "potrans.dat"
  Call MainLog("POs Purged to " + Yr$)
  'Year 1 data now has .dat file extensions. Call relink to tie
  'everything together so we can run reports accurately in the future.
  'ReLinkTrans frmGLClosingOpMenu
  ReLinkPOTrans frmGLClosingOpMenu, True
    

  'QPrintRC Space$(80), 1, 1, 112
'  HistDir$ = "GL" + Yr$
'  tstdir = Dir(HistDir$, vbDirectory)
'  If tstdir = "" Then
'    MkDir HistDir$
'  Else
'    MsgBox "Prior Year/End Directory Will Be Renamed.", vbOKOnly, "Prior Y/E Dir"
'    HistDir$ = "GLNewYr" + Yr$
'    MkDir HistDir$
'  End If
'  FrmShowPctComp.ShowPctComp 1, 4
  '--after linking copy the *.dat to the archive directory
'  SH_CopyFile "gl*.dat", HistDir$
'  FrmShowPctComp.ShowPctComp 2, 4
'  SH_CopyFile "gl*.idx", HistDir$
'  FrmShowPctComp.ShowPctComp 3, 4
'  SH_CopyFile "po*.dat", HistDir$
'  FrmShowPctComp.ShowPctComp 4, 4
'


End Sub

