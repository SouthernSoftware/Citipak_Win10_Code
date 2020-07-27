VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmGenJournalMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Journal "
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   ClipControls    =   0   'False
   Icon            =   "frmGenJournal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdEnterGenJournal 
      Height          =   480
      Left            =   4305
      TabIndex        =   0
      Top             =   3390
      Width           =   3615
      _Version        =   131072
      _ExtentX        =   6376
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmGenJournal.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintGenJournal 
      Height          =   492
      Left            =   4302
      TabIndex        =   1
      Top             =   4200
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
      ButtonDesigner  =   "frmGenJournal.frx":0AC2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPostGenJournal 
      Height          =   492
      Left            =   4302
      TabIndex        =   2
      Top             =   5016
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
      ButtonDesigner  =   "frmGenJournal.frx":0CB8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitGenJournalMenu 
      Height          =   492
      Left            =   4302
      TabIndex        =   3
      Top             =   5832
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
      ButtonDesigner  =   "frmGenJournal.frx":0EAC
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GENERAL JOURNAL"
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
      Left            =   3840
      TabIndex        =   4
      Top             =   1440
      Width           =   4692
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   132
      Left            =   2400
      Top             =   2280
      Width           =   972
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2520
      X2              =   3240
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   132
      Left            =   8880
      Top             =   2280
      Width           =   972
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9000
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9720
      Y1              =   8280
      Y2              =   8280
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
End
Attribute VB_Name = "frmGenJournalMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GJEdit As TrEditRecType
Dim GLSetup As GLSetupRecType
Dim GLFundIdx As GLFundIndexType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim GJEditFNum As Integer, NumEdTrans As Integer
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdEnterGenJournal_Click()
  Dim FileHandle As Integer, WhosOnFirst As String
  If Exist("GJEdit.opn") Then
    FileHandle = FreeFile
    Open "GJEdit.opn" For Input As FileHandle
    Line Input #FileHandle, WhosOnFirst$
    Close FileHandle
    MsgBox "The General Journal File Is In Use By: " + WhosOnFirst$, vbOKOnly, "File Not Accessible"
  Else
    FileHandle = FreeFile
    Open "GJEdit.opn" For Output As FileHandle
    Print #FileHandle, ComputerName$
    Close FileHandle
    frmGenJournalEntry.Show
    Unload frmGenJournalMenu
  End If
End Sub

Private Sub cmdPostGenJournal_Click()
  Dim FileHandle As Integer, WhosOnFirst As String
  If Exist("GJEdit.opn") Then
    FileHandle = FreeFile
    Open "GJEdit.opn" For Input As FileHandle
    Line Input #FileHandle, WhosOnFirst$
    Close FileHandle
    MsgBox "The General Journal File Is In Use By: " + WhosOnFirst$, vbOKOnly, "File Not Accessible"
  Else
    FileHandle = FreeFile
    Open "GJEdit.opn" For Binary As FileHandle
    Put FileHandle, , FileHandle
    Close FileHandle
    frmPostGJ.Show
  End If
End Sub

Private Sub cmdPrintGenJournal_Click()
  frmReportOpt.Show 1
  If rptopt = 1 Then
    Call PrintEditList
  ElseIf rptopt = 2 Then
    Call PrintEditList2
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
    End If
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Me.HelpContextID = hlpGJ
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub cmdExitGenJournalMenu_Click()
  frmGLMainMenu.Show
  Unload frmGenJournalMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitGenJournalMenu_Click
      KeyCode = 0
      DoEvents
    Case Else:
  End Select
End Sub
Private Sub PrintEditList()
  Dim GJEditFile As Integer, GJEditFileNum As Integer, NumEdTrans As Integer
  Dim LookFor As String, PRNFileNum2 As Integer
  Dim ReportFile2 As String, ToPrintF As String
  Dim ReportFile As String, ToPrint As String
  Dim Header As String, Newrp As String
  Dim PRNFileNum As Integer, cnt As Integer, Howmany As Integer
  Dim FundCode As String, BalMsg As String
  Dim CommaFmt As String, FundNum As String
  Dim TotDr As Double, TotCr As Double, TranCashTot As Double
  ReDim FundList(1) As String
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer, NumFunds As Integer
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  GetFundList FundList(), NumFunds
  ReDim FundDr(1 To NumFunds) As Double
  ReDim FundCr(1 To NumFunds) As Double
  OpenGJEditFile GJEditFileNum, NumEdTrans
  If GJEditFileNum = -1 Then
    Exit Sub
  End If
  PRNFileNum = FreeFile
  Newrp = "GJREG"
  GetRPTName Newrp
  ReportFile$ = Newrp
  Open ReportFile$ For Output As #PRNFileNum
  PRNFileNum2 = FreeFile
  ReportFile2$ = "GJRegF.prn"
  Open ReportFile2$ For Output As #PRNFileNum2

  'Define vars used for printing
  'MaxLines = 55
  'FF$ = Chr$(12)
  Header$ = "General Journal Register"
  CommaFmt$ = "###,###,###.##"
  For cnt = 1 To NumEdTrans
    Get GJEditFileNum, cnt, GJEdit
    If GJEdit.Deleted = 0 Then
      Howmany = Howmany + 1
      ToPrint$ = ""
      ToPrint$ = Format(DateAdd("d", (GJEdit.TRDATE), "12-31-1979"), "mm/dd/yyyy")
      ToPrint$ = ToPrint$ + "~" + QPTrim(GJEdit.Desc) + " " + QPTrim(GJEdit.LDesc) + "~" + GJEdit.Ref
      ToPrint$ = ToPrint$ + "~" + GJEdit.AcctNum + "~" + GJEdit.AcctName
      ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, GJEdit.DrAmt)
      ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, GJEdit.CrAmt)
      Print #PRNFileNum, ToPrint$
     
      ' Sum total debits and credits
      TotDr# = Round#(TotDr# + GJEdit.DrAmt)
      TotCr# = Round#(TotCr# + GJEdit.CrAmt)
      ' Sum into proper fund
      Found = False
      For Fund = 1 To NumFunds
        FundNum$ = Left$(GJEdit.AcctNum, GLFundLen)
        If FundNum$ = FundList$(Fund) Then
          Found = True
          FundDr#(Fund) = Round#(FundDr#(Fund) + GJEdit.DrAmt)
          FundCr#(Fund) = Round#(FundCr#(Fund) + GJEdit.CrAmt)
          Exit For
        End If
      Next
    End If
  Next
'  Print #PRNFileNum, String$(80, "-")
'  LSet ToPrint$ = "File Totals"
'  Mid$(ToPrint$, 53) =
'  Mid$(ToPrint$, 67) =
'  Print #PRNFileNum, ToPrint$
'  Linecnt = Linecnt + 1
'  If Linecnt > MaxLines Then
'    Print #PRNFileNum, FF$
'    GoSub PrintGJHeader
'  End If
   'Print Summary by Fund
  TranCashTot# = 0
  FundOutofBal = False
  For Fund = 1 To NumFunds
    If FundDr#(Fund) <> 0 Or FundCr#(Fund) <> 0 Then
      If FundDr#(Fund) <> FundCr#(Fund) Then FundOutofBal = True
        ToPrintF$ = ""
        ToPrintF$ = FundList$(Fund)
        ToPrintF$ = ToPrintF$ + "~" + Using$(CommaFmt$, FundDr#(Fund))
        ToPrintF$ = ToPrintF$ + "~" + Using$(CommaFmt$, FundCr#(Fund))
        Print #PRNFileNum2, ToPrintF$
      End If
   Next
   
   Close
   Load frmLoadingRpt
   If FundOutofBal Then
    ' Tell User they're screwing up
      BalMsg$ = "Entries are not in balance!"
      ARptEditList.Label9.Visible = True
      ARptEditList.Label9.Caption = BalMsg$
   End If
   ARptEditList.totDebits = Using$(CommaFmt$, TotDr#)
   ARptEditList.totCredits = Using$(CommaFmt$, TotCr#)
   ARptEditList.txtDate = Now
   ARptEditList.txtTown = GLUserName$
   ARptEditList.Title = "General Journal Report"
   ARptEditList.GetName ReportFile$, ReportFile2$
   ARptEditList.startrpt

'don't open file because viewprint does it there
'   ViewPrint ReportFile$, "General Journal Report"
'   Kill ReportFile$
Exit Sub
'PrintGJHeader:
'  PageNum = PageNum + 1
'  Print #PRNFileNum, GLUserName$
'  Print #PRNFileNum, "General Journal Register"
'  Print #PRNFileNum, Tab(70); "Page: "; PageNum
'  Print #PRNFileNum, "Date        Description           Reference"
'  Print #PRNFileNum, "            G/L Account                                     Debit         Credit"
'  Print #PRNFileNum, "--------------------------------------------------------------------------------"
'  Linecnt = 5
'Return
End Sub
Private Sub PrintEditList2()
  Dim GJEditFile As Integer, GJEditFileNum As Integer, NumEdTrans As Integer
  Dim MaxLines As Integer, LookFor As String
  Dim Linecnt As Integer, PageNum As Integer
  Dim ReportFile As String, ToPrint As String
  Dim FF As String, Header As String, Newrp As String
  Dim PRNFileNum As Integer, cnt As Integer, Howmany As Integer
  Dim FundCode As String
  Dim CommaFmt As String, FundNum As String
  Dim TotDr As Double, TotCr As Double, TranCashTot As Double
  ReDim FundList(1) As String
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer, NumFunds As Integer
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  GetFundList FundList(), NumFunds
  ReDim FundDr(1 To NumFunds) As Double
  ReDim FundCr(1 To NumFunds) As Double
  OpenGJEditFile GJEditFileNum, NumEdTrans
  If GJEditFileNum = -1 Then
    Exit Sub
  End If
  PRNFileNum = FreeFile
  Newrp = "GJREG"
  GetRPTName Newrp
  ReportFile$ = Newrp
  Open ReportFile$ For Output As #PRNFileNum
  'Define vars used for printing
  MaxLines = 55
  FF$ = Chr$(12)
  Header$ = "General Journal Register"
  CommaFmt$ = "###,###,###.##"
  GoSub PrintGJHeader
  For cnt = 1 To NumEdTrans
    Get GJEditFileNum, cnt, GJEdit
    If GJEdit.Deleted = 0 Then
      Howmany = Howmany + 1
      ToPrint$ = Space$(80)
      LSet ToPrint$ = Format(DateAdd("d", (GJEdit.TRDATE), "12-31-1979"), "mm/dd/yyyy")
      Mid$(ToPrint$, 13) = QPTrim(GJEdit.Desc)
      Mid$(ToPrint$, 35) = GJEdit.Ref
      Print #PRNFileNum, ToPrint$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #PRNFileNum, ToPrint$
        GoSub PrintGJHeader
      End If
      If Len(QPTrim$(GJEdit.LDesc)) > 0 Then
        Print #PRNFileNum, Tab(13); QPTrim$(GJEdit.LDesc)
        Linecnt = Linecnt + 1
      End If
        ' 2nd Line
      ToPrint$ = Space$(80)
      Mid$(ToPrint$, 13) = GJEdit.AcctNum
      Mid$(ToPrint$, 27) = GJEdit.AcctName
      Mid$(ToPrint$, 53) = Using$(CommaFmt$, GJEdit.DrAmt)
      Mid$(ToPrint$, 67) = Using$(CommaFmt$, GJEdit.CrAmt)
      Print #PRNFileNum, ToPrint$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #PRNFileNum, FF$
        GoSub PrintGJHeader
      End If
      ' 3rd Line
      Print #PRNFileNum,
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #PRNFileNum, FF$
        GoSub PrintGJHeader
      End If
      ' Sum total debits and credits
      TotDr# = Round#(TotDr# + GJEdit.DrAmt)
      TotCr# = Round#(TotCr# + GJEdit.CrAmt)
      ' Sum into proper fund
      Found = False
      For Fund = 1 To NumFunds
        FundNum$ = Left$(GJEdit.AcctNum, GLFundLen)
        If FundNum$ = FundList$(Fund) Then
          Found = True
          FundDr#(Fund) = Round#(FundDr#(Fund) + GJEdit.DrAmt)
          FundCr#(Fund) = Round#(FundCr#(Fund) + GJEdit.CrAmt)
          Exit For
        End If
      Next
    End If
  Next
  Print #PRNFileNum, String$(80, "-")
  Linecnt = Linecnt + 1
  If Linecnt > MaxLines Then
    Print #PRNFileNum, FF$
    GoSub PrintGJHeader
  End If
  ToPrint$ = Space$(80)
  LSet ToPrint$ = "File Totals"
  Mid$(ToPrint$, 53) = Using$(CommaFmt$, TotDr#)
  Mid$(ToPrint$, 67) = Using$(CommaFmt$, TotCr#)
  Print #PRNFileNum, ToPrint$
  Linecnt = Linecnt + 1
  If Linecnt > MaxLines Then
    Print #PRNFileNum, FF$
    GoSub PrintGJHeader
  End If
   'Print Summary by Fund
  TranCashTot# = 0
  FundOutofBal = False
  For Fund = 1 To NumFunds
    If FundDr#(Fund) <> 0 Or FundCr#(Fund) <> 0 Then
      If FundDr#(Fund) <> FundCr#(Fund) Then FundOutofBal = True
        ToPrint$ = Space$(80)
        Mid$(ToPrint$, 4) = "Fund# " + FundList$(Fund)
        Mid$(ToPrint$, 53) = Using$(CommaFmt$, FundDr#(Fund))
        Mid$(ToPrint$, 67) = Using$(CommaFmt$, FundCr#(Fund))
        Print #PRNFileNum, ToPrint$
        Linecnt = Linecnt + 1
        If Linecnt > MaxLines Then
          Print #PRNFileNum, FF$
          GoSub PrintGJHeader
        End If
      End If
   Next
   If FundOutofBal Then
   ' skip a line
      ToPrint$ = Space$(80)
      Print #PRNFileNum, ToPrint$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #PRNFileNum, FF$
        GoSub PrintGJHeader
      End If
    ' Tell User they're screwing up
      ToPrint$ = Space$(80)
      Mid$(ToPrint$, 4) = "Entries are not in balance!"
      Print #PRNFileNum, ToPrint$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #PRNFileNum, FF$
        GoSub PrintGJHeader
      End If
    End If
   Print #PRNFileNum, FF$
   Close
'don't open file because viewprint does it there
   ViewPrint ReportFile$, "General Journal Report"
   Kill ReportFile$
Exit Sub
PrintGJHeader:
  PageNum = PageNum + 1
  Print #PRNFileNum, GLUserName$
  Print #PRNFileNum, "General Journal Register"
  Print #PRNFileNum, Tab(70); "Page: "; PageNum
  Print #PRNFileNum, "Date        Description           Reference"
  Print #PRNFileNum, "            G/L Account                                     Debit         Credit"
  Print #PRNFileNum, "--------------------------------------------------------------------------------"
  Linecnt = 5
Return
End Sub

  
  
