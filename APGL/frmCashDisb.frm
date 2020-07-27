VERSION 5.00
Begin VB.Form frmCashDisbMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Disbursements Journal"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   ClipControls    =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmCashDisb.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdEnterCashDisb 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Enter/Edit Cash Disbursements"
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
      Left            =   4344
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   3612
   End
   Begin VB.CommandButton cmdPrintCashDisb 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Print Cash Disbursements Journal"
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
      Left            =   4344
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4344
      Width           =   3612
   End
   Begin VB.CommandButton cmdPostCashDisb 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Post Cash &Disbursements"
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
      Left            =   4344
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   3612
   End
   Begin VB.CommandButton cmdExitCashDisbMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "E&xit Cash Disbursements Menu"
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
      Left            =   4344
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   3612
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CASH DISBURSEMENTS JOURNAL"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   1440
      Width           =   5772
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   132
      Left            =   2400
      Top             =   2280
      Width           =   972
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2520
      X2              =   3240
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   132
      Left            =   8880
      Top             =   2280
      Width           =   972
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9720
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9000
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
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
   Begin VB.Shape Shape6 
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
Attribute VB_Name = "frmCashDisbMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLCDEd(1) As CJEditRecType
Dim GLSetup As GLSetupRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim CJType As Integer
Private Temp_Class As Resize_Class

Private Sub cmdEnterCashDisb_Click()
  Dim CDBusy As Boolean
  If Exist("GLCDEd.DAT") Then CDBusy = GetAttr("GLCDEd.DAT") And vbReadOnly
  If Not CDBusy Then
    frmCashDisbEntry.Show
    Unload frmCashDisbMenu
    frmCashDisbEntry.FirstOpenCD
  Else
    MsgBox "Posting Is In Progress, Editing Not Allowed At This Time.", vbOKOnly, "Request Canceled"
  End If
End Sub

Private Sub cmdExitCashDisbMenu_Click()
  frmGLMainMenu.Show
  Unload frmCashDisbMenu
End Sub

Private Sub cmdPostCashDisb_Click()
'  Dim FileHandle As Integer, WhosOnFirst As String
'  If Exist("GLCDED.opn") Then
'    FileHandle = FreeFile
'    Open "GLCDED.opn" For Input As FileHandle
'    Line Input #FileHandle, WhosOnFirst$
'    Close FileHandle
'    MsgBox "The Cash Disbursements File Is Being Edited By: " + WhosOnFirst$, vbOKOnly, "File Not Accessible"
'  Else
'    FileHandle = FreeFile
'    Open "GLCDED.opn" For Binary As FileHandle
'    Put FileHandle, , FileHandle
'    Close FileHandle
    frmPostCD.Show
'  End If
End Sub

Private Sub cmdPrintCashDisb_Click()
  frmReportOpt.Show 1
  If rptopt = 1 Then
    Call PrintEditList
  ElseIf rptopt = 2 Then
    PrintEditList2
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitCashDisbMenu_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Me.HelpContextID = hlpCashDisbursements
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub PrintEditList()
  Dim CJEditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim LookFor As String, Newrp As String
  Dim ToPrintF As String, ToPrintD As String
  Dim ReportFile As String, ToPrint As String
  Dim Header As String, PRNFileNum2 As Integer
  Dim PRNFileNum As Integer, cnt As Integer, Howmany As Integer, CntD As Integer
  Dim FundCode As String, ReportFile2 As String
  Dim CommaFmt As String, FundNum As String
  Dim totdist As Double, GRANDTOT As Double, AllFundTot As Double
  ReDim FundList(1) As String
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer, NumFunds As Integer
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  GetFundList FundList(), NumFunds
  ReDim FundTot(1 To NumFunds) As Double
  CJType = 2
  OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
  PRNFileNum = FreeFile
  Newrp = "CDREG"
  GetRPTName Newrp
  ReportFile$ = Newrp
  Open ReportFile$ For Output As #PRNFileNum
  PRNFileNum2 = FreeFile
  ReportFile2$ = "CDFund.prn"
  Open ReportFile2$ For Output As #PRNFileNum2
  'Define vars used for printing
  Header$ = "Cash Disbursement Register"
  CommaFmt$ = "###,###,###,###.##"
  For cnt = 1 To NumEdTrans
    Get CJEditFileNum, cnt, GLCDEd(1)
    If GLCDEd(1).DelFlag = 0 Then
      Howmany = Howmany + 1
      ToPrint$ = ""
      ToPrint$ = Str(cnt) + "~" + Format(DateAdd("d", (GLCDEd(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
      ToPrint$ = ToPrint$ + "~" + Trim(GLCDEd(1).Desc) + " " + Trim(GLCDEd(1).LDesc)
      ToPrint$ = ToPrint$ + "~" + GLCDEd(1).DOCREF
      ToPrint$ = ToPrint$ + "~" + Str$(GLCDEd(1).RECCODE)
      ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, GLCDEd(1).Amt)
      GRANDTOT = Round#(GRANDTOT + GLCDEd(1).Amt)
        ' Distribution Heading
     
      totdist = 0
      For CntD = 1 To 36
        If Val(GLCDEd(1).Dist(CntD).DACREC) > 0 Then
          ToPrintD$ = ""
          ToPrintD$ = QPTrim(GLCDEd(1).Dist(CntD).DACN)
          ToPrintD$ = ToPrintD$ + " " + QPTrim(GLCDEd(1).Dist(CntD).DACNM)
          ToPrintD$ = ToPrintD$ + "~" + Using$(CommaFmt$, GLCDEd(1).Dist(CntD).DAMT)
          Print #PRNFileNum, ToPrint$ + "~" + ToPrintD$
         
          totdist = Round#(totdist + GLCDEd(1).Dist(CntD).DAMT)
        
        ' Sum into proper fund
          Found = False
          For Fund = 1 To NumFunds
            FundNum$ = Left$(GLCDEd(1).Dist(CntD).DACN, GLFundLen)
            If FundNum$ = FundList$(Fund) Then
              Found = True
              FundTot#(Fund) = Round#(FundTot#(Fund) + GLCDEd(1).Dist(CntD).DAMT)
              Exit For
            End If
           Next
        Else
          Exit For
        End If
        
      Next
      
    End If
  Next
   'Print Summary by Fund
  AllFundTot = 0
  FundOutofBal = False
  For Fund = 1 To NumFunds
    If FundTot#(Fund) <> 0 Then
      AllFundTot = Round#(AllFundTot + FundTot#(Fund))
      ToPrintF$ = ""
      ToPrintF$ = FundList$(Fund) + "~" + Using$(CommaFmt$, FundTot#(Fund))
      Print #PRNFileNum2, ToPrintF$
    End If
  Next
  Close
   Load frmLoadingRpt
   ARptCRCDEdit.totAmount = Using$(CommaFmt$, GRANDTOT)
   ARptCRCDEdit.totTrans = Howmany
   ARptCRCDEdit.txtDate = Now
   ARptCRCDEdit.txtTown = GLUserName$
   ARptCRCDEdit.Title = Header$
   ARptCRCDEdit.GetName ReportFile$, ReportFile2$
   ARptCRCDEdit.startrpt

Exit Sub
 
End Sub
Private Sub PrintEditList2()
  Dim CJEditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim MaxLines As Integer, LookFor As String, Newrp As String
  Dim Linecnt As Integer, PageNum As Integer
  Dim ReportFile As String, ToPrint As String
  Dim FF As String, Header As String
  Dim PRNFileNum As Integer, cnt As Integer, Howmany As Integer, CntD As Integer
  Dim FundCode As String
  Dim CommaFmt As String, FundNum As String
  Dim totdist As Double, GRANDTOT As Double, AllFundTot As Double
  ReDim FundList(1) As String
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer, NumFunds As Integer
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  GetFundList FundList(), NumFunds
  ReDim FundTot(1 To NumFunds) As Double
  CJType = 2
  OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
  PRNFileNum = FreeFile
  Newrp = "CDREG"
  GetRPTName Newrp
  ReportFile$ = Newrp
  Open ReportFile$ For Output As #PRNFileNum
  'Define vars used for printing
  MaxLines = 55
  FF$ = Chr$(12)
  Header$ = "Cash Disbursement Register"
  CommaFmt$ = "###,###,###,###.##"
  GoSub PrintCDHeader
  For cnt = 1 To NumEdTrans
    Get CJEditFileNum, cnt, GLCDEd(1)
    If GLCDEd(1).DelFlag = 0 Then
      Howmany = Howmany + 1
      ToPrint$ = Space$(80)
      LSet ToPrint$ = Format(DateAdd("d", (GLCDEd(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
      Mid$(ToPrint$, 13) = GLCDEd(1).Desc
      Mid$(ToPrint$, 35) = GLCDEd(1).DOCREF
      Mid$(ToPrint$, 46) = GLCDEd(1).RECCODE
      Mid$(ToPrint$, 63) = Using$(CommaFmt$, GLCDEd(1).Amt)
      Print #PRNFileNum, ToPrint$
      If Len(QPTrim$(GLCDEd(1).LDesc)) > 0 Then
        Print #PRNFileNum, Tab(13); QPTrim$(GLCDEd(1).LDesc)
        Linecnt = Linecnt + 1
      End If
      GRANDTOT = Round#(GRANDTOT + GLCDEd(1).Amt)
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #PRNFileNum, FF$
        GoSub PrintCDHeader
      End If
      'skip a line
      ToPrint$ = Space$(80)
      Print #PRNFileNum, ToPrint$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #PRNFileNum, FF$
        GoSub PrintCDHeader
      End If
        ' Distribution Heading
      ToPrint$ = Space$(80)
      LSet ToPrint$ = "Accounting Distribution:"
      Print #PRNFileNum, ToPrint$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #PRNFileNum, FF$
        GoSub PrintCDHeader
      End If
      ToPrint$ = Space$(80)
      Mid$(ToPrint$, 2) = "Account Number"
      Mid$(ToPrint$, 20) = "Name"
      Mid$(ToPrint$, 60) = "Amount"
      Print #PRNFileNum, ToPrint$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #PRNFileNum, FF$
        GoSub PrintCDHeader
      End If
      totdist = 0
      For CntD = 1 To 36
        If Val(GLCDEd(1).Dist(CntD).DACREC) > 0 Then
          ToPrint$ = Space$(80)
          Mid$(ToPrint$, 2) = GLCDEd(1).Dist(CntD).DACN
          Mid$(ToPrint$, 20) = GLCDEd(1).Dist(CntD).DACNM
          Mid$(ToPrint$, 48) = Using$(CommaFmt$, GLCDEd(1).Dist(CntD).DAMT)
          Print #PRNFileNum, ToPrint$
          Linecnt = Linecnt + 1
          totdist = Round#(totdist + GLCDEd(1).Dist(CntD).DAMT)
          If Linecnt > MaxLines Then
            Print #PRNFileNum, FF$
            GoSub PrintCDHeader
          End If
        
        ' Sum into proper fund
          Found = False
          For Fund = 1 To NumFunds
            FundNum$ = Left$(GLCDEd(1).Dist(CntD).DACN, GLFundLen)
            If FundNum$ = FundList$(Fund) Then
              Found = True
              FundTot#(Fund) = Round#(FundTot#(Fund) + GLCDEd(1).Dist(CntD).DAMT)
              Exit For
            End If
           Next
        Else
          Exit For
        End If
        
      Next
      ToPrint$ = Space$(80)
      Mid$(ToPrint$, 49) = "-----------------"
      Print #PRNFileNum, ToPrint$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #PRNFileNum, FF$
        GoSub PrintCDHeader
      End If
      ToPrint$ = Space$(80)
      Mid$(ToPrint$, 2) = "Total Distributed"
      Mid$(ToPrint$, 48) = Using$(CommaFmt$, totdist)
      Print #PRNFileNum, ToPrint$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #PRNFileNum, FF$
        GoSub PrintCDHeader
      End If
      Print #PRNFileNum, "================================================================================"
      Linecnt = Linecnt + 1
      
    End If
  Next
  'Print #PRNFileNum, String$(80, "-")
  'LineCnt = LineCnt + 1
  If Linecnt > MaxLines Then
    Print #PRNFileNum, FF$
    GoSub PrintCDHeader
  End If
  ToPrint$ = Space$(80)
  LSet ToPrint$ = "File Totals:"
 'Mid$(ToPrint$, 53) = Using$(CommaFmt$, TotDr#)
 ' Mid$(ToPrint$, 67) = Using$(CommaFmt$, TotCr#)
  Print #PRNFileNum, ToPrint$
  Linecnt = Linecnt + 1
  If Linecnt > MaxLines Then
    Print #PRNFileNum, FF$
    GoSub PrintCDHeader
  End If
  ToPrint$ = Space$(80)
  LSet ToPrint$ = "Number of Transactions"
  Mid$(ToPrint$, 38) = Howmany
  Print #PRNFileNum, ToPrint$
  Linecnt = Linecnt + 1
  If Linecnt > MaxLines Then
    Print #PRNFileNum, FF$
    GoSub PrintCDHeader
  End If
  ToPrint$ = Space$(80)
  LSet ToPrint$ = "Grand Totals"
  Mid$(ToPrint$, 35) = Using$(CommaFmt$, GRANDTOT)
  Print #PRNFileNum, ToPrint$
  Linecnt = Linecnt + 1
  If Linecnt > MaxLines Then
    Print #PRNFileNum, FF$
    GoSub PrintCDHeader
  End If
  ToPrint$ = Space$(80)
  Print #PRNFileNum, ToPrint$
  Linecnt = Linecnt + 1
  If Linecnt > MaxLines Then
    Print #PRNFileNum, FF$
    GoSub PrintCDHeader
  End If
   'Print Summary by Fund
  AllFundTot = 0
  FundOutofBal = False
  For Fund = 1 To NumFunds
    If FundTot#(Fund) <> 0 Then
      AllFundTot = Round#(AllFundTot + FundTot#(Fund))
      ToPrint$ = Space$(80)
      LSet ToPrint$ = "Fund# " + FundList$(Fund)
      Mid$(ToPrint$, 35) = Using$(CommaFmt$, FundTot#(Fund))
      'Mid$(ToPrint$, 67) = Using$(CommaFmt$, FundCr#(Fund))
      Print #PRNFileNum, ToPrint$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #PRNFileNum, FF$
        GoSub PrintCDHeader
      End If
    End If
  Next
  ToPrint$ = Space$(80)
  LSet ToPrint$ = "Total All Funds"
  Mid$(ToPrint$, 35) = Using$(CommaFmt$, AllFundTot)
  Print #PRNFileNum, ToPrint$
  Linecnt = Linecnt + 1
  If Linecnt > MaxLines Then
    Print #PRNFileNum, FF$
    GoSub PrintCDHeader
  End If
  If AllFundTot <> GRANDTOT Then
    FundOutofBal = True
    If FundOutofBal Then
' skip a line
      ToPrint$ = Space$(80)
      Print #PRNFileNum, ToPrint$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #PRNFileNum, FF$
        GoSub PrintCDHeader
      End If
 ' Tell User they're screwing up
      ToPrint$ = Space$(80)
      Mid$(ToPrint$, 4) = "Entries are not in balance!"
      Print #PRNFileNum, ToPrint$
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        Print #PRNFileNum, FF$
        GoSub PrintCDHeader
      End If
    End If
    Print #PRNFileNum, FF$
  End If
  Print #PRNFileNum, FF$
  Close
'don't open file because viewprint does it there
   ViewPrint ReportFile$, "Cash Disbursement Report"
   Kill ReportFile$
Exit Sub
 
PrintCDHeader:
  PageNum = PageNum + 1
  Print #PRNFileNum, GLUserName$
  Print #PRNFileNum, "Cash Disbursement Register"
  Print #PRNFileNum, Tab(70); "Page: "; PageNum
  Print #PRNFileNum, "Date      Description           Reference    Bank                         Amount"
  Print #PRNFileNum, "================================================================================"
  Linecnt = 5
Return
End Sub

Public Sub GetFundList(FundList$(), NumFunds)
  Dim FundIndex As GLFundIndexType
  Dim FundIdxFile As Integer, cnt As Integer
  OpenFundIdx FundIdxFile, NumFunds
  If NumFunds = 0 Then
    MsgBox "No Funds", vbOKOnly, "No Funds"
    Close
    Exit Sub
  End If
  ReDim FundList$(1 To NumFunds)
  For cnt = 1 To NumFunds
    Get FundIdxFile, cnt, FundIndex
    FundList$(cnt) = Trim$(FundIndex.FundNum)
  Next
  Close FundIdxFile
End Sub

