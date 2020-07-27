VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Begin VB.Form frmGetDistMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Distributions"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   ClipControls    =   0   'False
   Icon            =   "frmGetDist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdGrabTrans 
      Height          =   492
      Left            =   4296
      TabIndex        =   0
      Top             =   2712
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
      ButtonDesigner  =   "frmGetDist.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintJournalReg 
      Height          =   492
      Left            =   4296
      TabIndex        =   1
      Top             =   3408
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
      ButtonDesigner  =   "frmGetDist.frx":0AB3
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPostJournEntries 
      Height          =   492
      Left            =   4296
      TabIndex        =   3
      Top             =   4800
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
      ButtonDesigner  =   "frmGetDist.frx":0CA1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTransfertoGJ 
      Height          =   492
      Left            =   4296
      TabIndex        =   4
      Top             =   5496
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
      ButtonDesigner  =   "frmGetDist.frx":0E8D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdInitTransIF 
      Height          =   492
      Left            =   4296
      TabIndex        =   5
      Top             =   6192
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
      ButtonDesigner  =   "frmGetDist.frx":1080
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitGetDistMenu 
      Height          =   480
      Left            =   4290
      TabIndex        =   7
      Top             =   7590
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
      ButtonDesigner  =   "frmGetDist.frx":127D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSummarize 
      Height          =   480
      Left            =   4290
      TabIndex        =   2
      Top             =   4110
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
      ButtonDesigner  =   "frmGetDist.frx":1470
   End
   Begin fpBtnAtlLibCtl.fpBtn fpcmdUtilDetail 
      Height          =   492
      Left            =   4296
      TabIndex        =   6
      Top             =   6888
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
      ButtonDesigner  =   "frmGetDist.frx":1663
   End
   Begin VB.Shape Shape3 
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
   Begin VB.Shape Shape2 
      BackColor       =   &H00D0D0D0&
      BorderColor     =   &H00000000&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   9000
      X2              =   9000
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   9000
      X2              =   9720
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Shape Shape5 
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
      Index           =   0
      Left            =   2400
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   2520
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2520
      X2              =   3240
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Caption         =   "GET DISTRIBUTIONS MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Index           =   1
      Left            =   2520
      TabIndex        =   8
      Top             =   1440
      Width           =   7092
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
End
Attribute VB_Name = "frmGetDistMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Private Temp_Class As Resize_Class
Private Sub cmdGrabTrans_Click()
  frmGrabTrans.Show
  Unload frmGetDistMenu
End Sub

Private Sub cmdInitTransIF_Click()
  frmIFInitialize.Show
  Unload frmGetDistMenu
End Sub

Private Sub cmdPostJournEntries_Click()
  If Exist("GLUBTran.dat") And GLUBKill <> 1 Then
    MsgBox "Utility detail report MUST be printed before posting.", vbOKOnly, "Option Canceled."
    Exit Sub
  End If
 frmPostIF.Show
  
End Sub

Private Sub cmdPrintJournalReg_Click()
  frmReportOpt.Show 1
  If rptopt = 1 Then
    PrnEditList 1
  ElseIf rptopt = 2 Then
    PrnEditList2 1
  End If
End Sub

Private Sub cmdSummarize_Click()
  If MsgBox("This will summarize transactions, Continue???", vbYesNo, "Trans Summary") = vbYes Then
    DeActivateControls frmGetDistMenu
    SummaryEditList2
    ActivateControls frmGetDistMenu
    frmReportOpt.Show 1
    If rptopt = 1 Then
      PrnEditList 1
    ElseIf rptopt = 2 Then
      PrnEditList2 1
    End If
  End If
End Sub

Private Sub cmdTransfertoGJ_Click()
  Dim FileHandle As Integer, WhosOnFirst As String
  If Exist("GLUBTran.dat") And GLUBKill <> 1 Then
    MsgBox "Utility detail report MUST be printed before transferring.", vbOKOnly, "Option Canceled."
    Exit Sub
  End If

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
    DeActivateControls frmGetDistMenu
    Trxfer2GJ
    ActivateControls frmGetDistMenu
    KillFileD "GJEdit.opn"
  End If
End Sub
Private Sub fpcmdUtilDetail_Click()
  frmRptTransJournal.Show
  Unload Me
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Me.HelpContextID = hlpGetDistributions
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub cmdExitGetDistMenu_Click()
  frmGLMainMenu.Show
  Unload frmGetDistMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitGetDistMenu_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitGetDistMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        ClearInUse PWcnt
      End If
    End If
  End If
End Sub
'Rptflag = 1 for report called from menu
'          2 for report from grab operation
Public Sub PrnEditList(Rptflag As Integer)
  Dim GJReclen As Integer, IFEditFileNum As Integer
  Dim LookFor As String, NumIfTrans As Long, PRNFileNum2 As Integer
  Dim User As String, PRNfileName2 As String, ToPrintF As String
  Dim PRNfileName As String, ToPrint As String, RptTitle As String
  Dim Header As String, Newrp As String, i As Long, ToPrintE As String
  Dim PRNFileNum As Integer, cnt As Long, Howmany As Integer
  Dim FundCode As String, NumFunds As Integer, BadAcctFlag As Integer
  Dim CommaFmt As String, FundNum As String, CrLF As String
  Dim TotDr As Double, TotCr As Double, TranCashTot As Double
  ReDim FundList(1) As String
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer
  Dim IFEdit As TrEditRecType
  Dim GJRec(1) As TrEditRecType
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  GetFundList FundList(), NumFunds
  ReDim FundDr(1 To NumFunds) As Double
  ReDim FundCr(1 To NumFunds) As Double
  'Set Bad Acct Flag to No Here
  BadAcctFlag = 0
  ToPrintE$ = ""
  '--Get a list of active funds
  GJReclen = Len(GJRec(1))
  
  ReDim FundList$(1)
  GetFundList FundList$(), NumFunds
  ReDim FundDr#(1 To NumFunds)
  ReDim FundCr#(1 To NumFunds)

  OpenIFEditFile IFEditFileNum, NumIfTrans
  If NumIfTrans > 0 Then
    FrmShowPctComp.Label1 = "Creating Journal Report"
    FrmShowPctComp.Show , Me
    DoEvents
    DeActivateControls frmGetDistMenu
  End If
  PRNFileNum = FreeFile
  Newrp = "XIF"
  GetRPTName Newrp
  PRNfileName$ = Newrp
  Open PRNfileName$ For Output As #PRNFileNum
  PRNFileNum2 = FreeFile
  PRNfileName2$ = "Fundtot.prn"
  Open PRNfileName2$ For Output As #PRNFileNum2
  '--Report Variables
  User$ = QPTrim$(GLUserName$)
  If Rptflag = 1 Then
    RptTitle$ = "General Ledger Interface Report"
    Header$ = "General Ledger Interface Report"
  Else
    RptTitle$ = "Interface Error Report"
    Header$ = "Interface Transaction List w/Errors"
  End If

  CommaFmt$ = "###,###,###.##"

  '--Start of printing loop
  For i = 1 To NumIfTrans
    FrmShowPctComp.ShowPctComp i, NumIfTrans
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmGetDistMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    Get IFEditFileNum, i, IFEdit
    If Not IFEdit.Deleted Then
      If UCase$(Left$(IFEdit.AcctName, 9)) = "UNDEFINED" Then BadAcctFlag = 1
      '--First Line
      ToPrint$ = ""
      ToPrint$ = Format(DateAdd("d", (IFEdit.TRDATE), "12-31-1979"), "mm/dd/yyyy")
      ToPrint$ = ToPrint$ + "~" + QPTrim(IFEdit.Desc) + " " + QPTrim(IFEdit.LDesc)
      ToPrint$ = ToPrint$ + "~" + QPTrim(IFEdit.Ref)
    

    If UCase$(Left$(IFEdit.AcctName, 9)) = "UNDEFINED" Then
      ToPrint$ = ToPrint$ + "~" + IFEdit.AcctNum + "???"
    Else
      ToPrint$ = ToPrint$ + "~" + IFEdit.AcctNum
    End If
    ToPrint$ = ToPrint$ + "~" + IFEdit.AcctName
    ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(IFEdit.DrAmt))
    ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(IFEdit.CrAmt))
    Print #PRNFileNum, ToPrint$
    


    '--Sum total debits and credits
    TotDr# = Round#(TotDr# + IFEdit.DrAmt)
    TotCr# = Round#(TotCr# + IFEdit.CrAmt)

    '--Sum into proper fund
    Found = False
    For Fund = 1 To NumFunds
      FundNum$ = Left$(IFEdit.AcctNum, GLFundLen)
      If FundNum$ = FundList$(Fund) Then
        Found = True
        FundDr#(Fund) = Round#(FundDr#(Fund) + IFEdit.DrAmt)
        FundCr#(Fund) = Round#(FundCr#(Fund) + IFEdit.CrAmt)
        Exit For
      End If
    Next
  End If
Next

'Mid$(ToPrint$, 53) = Using$(CommaFmt$, Str$(TotDr#))
'Mid$(ToPrint$, 67) = Using$(CommaFmt$, Str$(TotCr#))
'Print #PRNFileNum, ToPrint$

'--Print Summary by fund
TranCashTot# = 0
FundOutofBal = False
For Fund = 1 To NumFunds
  If FundDr#(Fund) <> 0 Or FundCr#(Fund) <> 0 Then
    If FundDr#(Fund) <> FundCr#(Fund) Then FundOutofBal = True
    ToPrintF$ = ""
    ToPrintF$ = FundList$(Fund)
    ToPrintF$ = ToPrintF$ + "~" + Using$(CommaFmt$, Str$(FundDr#(Fund)))
    ToPrintF$ = ToPrintF$ + "~" + Using$(CommaFmt$, Str$(FundCr#(Fund)))
    Print #PRNFileNum2, ToPrintF$
  End If

Next
Close
If FundOutofBal And Rptflag = 1 Then
  '--Tell user they're screwing up
  ToPrintE$ = "WARNING: Entries are not in balance!"
  ToPrintE$ = ToPrintE$ + "File WILL NOT POST Due to Bad Account Number"
  ToPrintE$ = ToPrintE$ + "Please Transfer to the General Journal and Correct"
  ARptEditList.Label10.Visible = True
  ARptEditList.Label10.Caption = ToPrintE$
End If
If BadAcctFlag = 1 And Rptflag = 1 Then
  
  ToPrintE$ = "WARNING: "
  ToPrintE$ = ToPrintE$ + "File WILL NOT POST Due to Bad Account Number('s)"
  ToPrintE$ = ToPrintE$ + "Please Transfer to the General Journal and Correct"
  ToPrintE$ = ToPrintE$ + "Look for ??? or Invalid Acct"
  ARptEditList.Label10.Visible = True
  ARptEditList.Label10.Caption = ToPrintE$
End If
If Rptflag = 2 Then
 
  ToPrintE$ = "WARNING: "
  ToPrintE$ = ToPrintE$ + "File WAS NOT CREATED Due to Bad Account Number('s)"
  ToPrintE$ = ToPrintE$ + "Please Correct and Try Grab Tranactions Again"
  ToPrintE$ = ToPrintE$ + "Look for ??? or Invalid Acct"
  ARptEditList.Label10.Visible = True
  ARptEditList.Label10.Caption = ToPrintE$
End If


ActivateControls frmGetDistMenu
Load frmLoadingRpt
'ViewPrint PRNfileName$, RptTitle$
'Kill PRNfileName$
   ARptEditList.totDebits = Using$(CommaFmt$, TotDr#)
   ARptEditList.totCredits = Using$(CommaFmt$, TotCr#)
   ARptEditList.txtDate = Now
   ARptEditList.txtTown = GLUserName$
   ARptEditList.Title = RptTitle$
   ARptEditList.GetName PRNfileName$, PRNfileName2$
   ARptEditList.startrpt

Exit Sub

CancelExit:
  Exit Sub
End Sub
'Rptflag = 1 for report called from menu
'          2 for report from grab operation
Public Sub PrnEditList2(Rptflag As Integer)
  Dim GJReclen As Integer, IFEditFileNum As Integer
  Dim MaxLines As Integer, LookFor As String, NumIfTrans As Integer
  Dim Linecnt As Integer, Page As Integer, User As String
  Dim PRNfileName As String, ToPrint As String, RptTitle As String
  Dim FF As String, Header As String, Newrp As String, i As Integer
  Dim PRNFileNum As Integer, cnt As Integer, Howmany As Integer
  Dim FundCode As String, NumFunds As Integer, BadAcctFlag As Integer
  Dim CommaFmt As String, FundNum As String, CrLF As String
  Dim TotDr As Double, TotCr As Double, TranCashTot As Double
  ReDim FundList(1) As String
  Dim Found As Boolean, FundOutofBal As Boolean, Fund As Integer
  Dim IFEdit As TrEditRecType
  Dim GJRec(1) As TrEditRecType
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  GetFundList FundList(), NumFunds
  ReDim FundDr(1 To NumFunds) As Double
  ReDim FundCr(1 To NumFunds) As Double
  'Set Bad Acct Flag to No Here
  BadAcctFlag = 0

  '--Get a list of active funds
  GJReclen = Len(GJRec(1))

  ReDim FundList$(1)
  GetFundList FundList$(), NumFunds
  ReDim FundDr#(1 To NumFunds)
  ReDim FundCr#(1 To NumFunds)

  OpenIFEditFile IFEditFileNum, NumIfTrans
  If NumIfTrans > 0 Then
    FrmShowPctComp.Label1 = "Creating Journal Report"
    FrmShowPctComp.Show , Me
    DoEvents
    DeActivateControls frmGetDistMenu
  End If
  PRNFileNum = FreeFile
  Newrp = "XIF"
  GetRPTName Newrp
  PRNfileName$ = Newrp
  Open PRNfileName$ For Output As #PRNFileNum

  '--Report Variables
  MaxLines = 55
  Page = 0
  User$ = QPTrim$(GLUserName$)
  If Rptflag = 1 Then
    RptTitle$ = "General Ledger Interface Report"
    Header$ = "General Ledger Interface Report"
  Else
    RptTitle$ = "Interface Error Report"
    Header$ = "Interface Transaction List w/Errors"
  End If
  CrLF$ = Chr$(13) + Chr$(10)
  FF$ = Chr$(12)

  GoSub PrintGJHeader
  CommaFmt$ = "###,###,###.##"

  '--Start of printing loop
  For i = 1 To NumIfTrans
    FrmShowPctComp.ShowPctComp i, NumIfTrans
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmGetDistMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

    Get IFEditFileNum, i, IFEdit
    If Not IFEdit.Deleted Then
      If UCase$(Left$(IFEdit.AcctName, 9)) = "UNDEFINED" Then BadAcctFlag = 1
      '--First Line
      ToPrint$ = Space$(80)
      LSet ToPrint$ = Format(DateAdd("d", (IFEdit.TRDATE), "12-31-1979"), "mm/dd/yyyy")
      Mid$(ToPrint$, 13) = IFEdit.Desc
      Mid$(ToPrint$, 35) = IFEdit.Ref
    Print #PRNFileNum, ToPrint$
    Linecnt = Linecnt + 1
      If Len(QPTrim$(IFEdit.LDesc)) > 0 Then
        Print #PRNFileNum, Tab(13); QPTrim$(IFEdit.LDesc)
        Linecnt = Linecnt + 1
      End If

    '--2nd Line

    ToPrint$ = Space$(80)
    If UCase$(Left$(IFEdit.AcctName, 9)) = "UNDEFINED" Then
      Mid$(ToPrint$, 4) = "???"
    End If
    Mid$(ToPrint$, 13) = IFEdit.AcctNum
    Mid$(ToPrint$, 27) = IFEdit.AcctName
    Mid$(ToPrint$, 53) = Using$(CommaFmt$, Str$(IFEdit.DrAmt))
    Mid$(ToPrint$, 67) = Using$(CommaFmt$, Str$(IFEdit.CrAmt))
    Print #PRNFileNum, ToPrint$
    Linecnt = Linecnt + 1

    '--3rd line (Blank)
    Print #PRNFileNum,
    Linecnt = Linecnt + 1
    If Linecnt > MaxLines Then
      Print #PRNFileNum, FF$
      GoSub PrintGJHeader
    End If

    '--Sum total debits and credits
    TotDr# = Round#(TotDr# + IFEdit.DrAmt)
    TotCr# = Round#(TotCr# + IFEdit.CrAmt)

    '--Sum into proper fund
    Found = False
    For Fund = 1 To NumFunds
      FundNum$ = Left$(IFEdit.AcctNum, GLFundLen)
      If FundNum$ = FundList$(Fund) Then
        Found = True
        FundDr#(Fund) = Round#(FundDr#(Fund) + IFEdit.DrAmt)
        FundCr#(Fund) = Round#(FundCr#(Fund) + IFEdit.CrAmt)
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
Mid$(ToPrint$, 53) = Using$(CommaFmt$, Str$(TotDr#))
Mid$(ToPrint$, 67) = Using$(CommaFmt$, Str$(TotCr#))
Print #PRNFileNum, ToPrint$
Linecnt = Linecnt + 1
If Linecnt > MaxLines Then
  Print #PRNFileNum, FF$
  GoSub PrintGJHeader
End If

'--Print Summary by fund
TranCashTot# = 0
FundOutofBal = False
For Fund = 1 To NumFunds
  If FundDr#(Fund) <> 0 Or FundCr#(Fund) <> 0 Then
    If FundDr#(Fund) <> FundCr#(Fund) Then FundOutofBal = True
    ToPrint$ = Space$(80)
    Mid$(ToPrint$, 4) = "Fund# " + FundList$(Fund)
    Mid$(ToPrint$, 53) = Using$(CommaFmt$, Str$(FundDr#(Fund)))
    Mid$(ToPrint$, 67) = Using$(CommaFmt$, Str$(FundCr#(Fund)))
    Print #PRNFileNum, ToPrint$
    Linecnt = Linecnt + 1
    If Linecnt > MaxLines Then
      Print #PRNFileNum, FF$
      GoSub PrintGJHeader
    End If
  End If

Next

If FundOutofBal And Rptflag = 1 Then
  '--skip a line
  ToPrint$ = Space$(80)
  Print #PRNFileNum, ToPrint$
  Linecnt = Linecnt + 1
  '--Tell user they're screwing up
  Print #PRNFileNum, "WARNING:"
  Print #PRNFileNum, "Entries are not in balance!"
  Print #PRNFileNum, "File WILL NOT POST Due to Bad Account Number"
  Print #PRNFileNum, "Please Transfer to the General Journal and Correct"
  Linecnt = Linecnt + 4
End If
If BadAcctFlag = 1 And Rptflag = 1 Then
  Print #PRNFileNum, ""
  Print #PRNFileNum, "WARNING:"
  Print #PRNFileNum, "File WILL NOT POST Due to Bad Account Number('s)"
  Print #PRNFileNum, "Please Transfer to the General Journal and Correct"
  Print #PRNFileNum, "Look for ??? or Invalid Acct"
  Linecnt = Linecnt + 4
End If
If Rptflag = 2 Then
  Print #PRNFileNum, ""
  Print #PRNFileNum, "WARNING:"
  Print #PRNFileNum, "File WAS NOT CREATED Due to Bad Account Number('s)"
  Print #PRNFileNum, "Please Correct and Try Grab Tranactions Again"
  Print #PRNFileNum, "Look for ??? or Invalid Acct"
  Linecnt = Linecnt + 4
End If
Print #PRNFileNum, FF$

Close
ViewPrint PRNfileName$, RptTitle$
Kill PRNfileName$
ActivateControls frmGetDistMenu
Exit Sub

PrintGJHeader:
Page = Page + 1
Print #PRNFileNum, Tab(40 - (Int(Len(User$) / 2))); User$
Print #PRNFileNum, Tab(40 - (Int(Len(Header$) / 2))); Header$
Print #PRNFileNum,
Print #PRNFileNum, "Report Date: "; Date$; Tab(67); "Page #"; Page
Print #PRNFileNum, "Date        Description           Reference"
Print #PRNFileNum, "            G/L Account                                     Debit         Credit"
Print #PRNFileNum, "--------------------------------------------------------------------------------"
Linecnt = 5
Return

CancelExit:
  Exit Sub
End Sub

Private Sub Trxfer2GJ()
  Dim GJReclen As Integer, GJFile As Integer, NumEdTrans As Long
  Dim cnt As Integer, NumRecs As Long
  Dim GJRec(1) As TrEditRecType
  NumRecs = 0
  GJReclen = Len(GJRec(1))
  GJFile = FreeFile
  Open "GJEDIT.DAT" For Random Access Read Write Shared As GJFile Len = GJReclen
  NumEdTrans = LOF(GJFile) \ GJReclen
  For cnt = 1 To NumEdTrans
    Get GJFile, cnt, GJRec(1)
    If GJRec(1).Deleted = 0 Then
      NumRecs = NumRecs + 1
    End If
  Next
  Close

  If NumRecs > 0 Then
    MsgBox "The General Journal Edit File Must Be Empty Before You Can Transfer The Interface Transactions.", vbOKOnly, "Transfer Cancelled"
    Exit Sub
  End If
  NumRecs = 0
  GJReclen = Len(GJRec(1))
  GJFile = FreeFile
  Open "GLTRXED.DAT" For Random Access Read Write Shared As GJFile Len = GJReclen
  NumEdTrans = LOF(GJFile) \ GJReclen
  For cnt = 1 To NumEdTrans
    Get GJFile, cnt, GJRec(1)
    If GJRec(1).Deleted = 0 Then
      NumRecs = NumRecs + 1
    End If
  Next
  Close

  If NumEdTrans = 0 Then
    MsgBox "No Transactions To Transfer.", vbOKOnly, "No Trans"
    If ExistD("GLUBTran.dat") And GLUBKill = 1 Then
      Kill "GLUBTran.dat"
      GLUBKill = 0
    End If
    Exit Sub
  End If

  'CLEAR

  'OK to Rename
  SH_CopyFile "GLTRXED.DAT", "GJEDIT.DAT"
  KillFileD "GLTRXED.DAT"
  If ExistD("GLUBTran.dat") And GLUBKill = 1 Then
    Kill "GLUBTran.dat"
    GLUBKill = 0
  End If

  MsgBox "File Transfer Is Complete.", vbOKOnly, "File Transfered"


End Sub

Public Sub SummaryEditList2()
  Dim IFEditFileNum As Integer, AcctNum As String, EntryT As String
  Dim NumIfTrans As Long, NumAcctTrans As Long, SumEditNum As Integer
  Dim cnt As Integer, i As Long, DDay As Integer
  Dim IFEdit(1) As TrEditRecType
  Dim GLSumEd As TrEditRecType
  KillFileD "GLSumEd.DAT"
  OpenIFEditFile IFEditFileNum, NumIfTrans
  If NumIfTrans > 0 Then
    FrmShowPctComp.Label1 = "Creating Journal Summary"
    FrmShowPctComp.Show , Me
    DoEvents
  End If
  
  ReDim Trsort(1 To 1) As TrEditRecType
  AcctNum$ = "0"
  EntryT$ = " "
  DDay = 0
  NumAcctTrans = 0
  '--Start of summary loop
  For i = 1 To NumIfTrans
    FrmShowPctComp.ShowPctComp i, NumIfTrans
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
    Get IFEditFileNum, i, IFEdit(1)
    
      If IFEdit(1).Deleted <> 1 Then
        If AcctNum$ = QPTrim$(IFEdit(1).AcctNum) And EntryT$ = QPTrim$(IFEdit(1).EType) And DDay = IFEdit(1).TRDATE Then
          'ReDim Preserve Trsort(1 To NumAcctTrans) As TrEditRecType       '|
          Trsort(NumAcctTrans).DrAmt = Trsort(NumAcctTrans).DrAmt + IFEdit(1).DrAmt
          Trsort(NumAcctTrans).CrAmt = Trsort(NumAcctTrans).CrAmt + IFEdit(1).CrAmt
          IFEdit(1).Deleted = 1
          Put IFEditFileNum, i, IFEdit(1)
        Else
          NumAcctTrans = NumAcctTrans + 1
          ReDim Preserve Trsort(1 To NumAcctTrans) As TrEditRecType
          Trsort(NumAcctTrans).TRDATE = IFEdit(1).TRDATE
          Trsort(NumAcctTrans).AcctNum = IFEdit(1).AcctNum
          Trsort(NumAcctTrans).AcctName = IFEdit(1).AcctName
          Trsort(NumAcctTrans).DrAmt = IFEdit(1).DrAmt
          Trsort(NumAcctTrans).CrAmt = IFEdit(1).CrAmt
          Trsort(NumAcctTrans).Desc = IFEdit(1).Desc
          Trsort(NumAcctTrans).LDesc = IFEdit(1).LDesc
          Trsort(NumAcctTrans).Ref = IFEdit(1).Ref
          Trsort(NumAcctTrans).EType = IFEdit(1).EType
          Trsort(NumAcctTrans).Src = IFEdit(1).Src
          AcctNum$ = QPTrim$(IFEdit(1).AcctNum)
          EntryT$ = QPTrim$(IFEdit(1).EType)
          DDay = IFEdit(1).TRDATE
          IFEdit(1).Deleted = 1
          Put IFEditFileNum, i, IFEdit(1)

        End If
        End If
      For cnt = 1 To NumIfTrans
      Get IFEditFileNum, cnt, IFEdit(1)
      If IFEdit(1).Deleted <> 1 Then
        If AcctNum$ = QPTrim$(IFEdit(1).AcctNum) And EntryT$ = QPTrim$(IFEdit(1).EType) And DDay = IFEdit(1).TRDATE Then
          'ReDim Preserve Trsort(1 To NumAcctTrans) As TrEditRecType       '|
          Trsort(NumAcctTrans).DrAmt = Trsort(NumAcctTrans).DrAmt + IFEdit(1).DrAmt
          Trsort(NumAcctTrans).CrAmt = Trsort(NumAcctTrans).CrAmt + IFEdit(1).CrAmt
          IFEdit(1).Deleted = 1
          Put IFEditFileNum, cnt, IFEdit(1)
         End If
        End If
       Next

    Next
  Dim SumEdLen As Integer
  SumEdLen = Len(GLSumEd)
  SumEditNum = FreeFile
  Open "GLSumEd.DAT" For Random Shared As SumEditNum Len = SumEdLen
  
  For cnt = 1 To NumAcctTrans
    GLSumEd.AcctRec = cnt
    GLSumEd.AcctNum = Trsort(cnt).AcctNum
    GLSumEd.TRDATE = Trsort(cnt).TRDATE
    GLSumEd.AcctName = Trsort(cnt).AcctName
    GLSumEd.DrAmt = Trsort(cnt).DrAmt
    GLSumEd.CrAmt = Trsort(cnt).CrAmt
    GLSumEd.Desc = Trsort(cnt).Desc
    GLSumEd.LDesc = Trsort(cnt).LDesc
    GLSumEd.Ref = Trsort(cnt).Ref
    GLSumEd.EType = Trsort(cnt).EType
    GLSumEd.Src = Trsort(cnt).Src
    Put #SumEditNum, cnt, GLSumEd
  Next
  Close
'  Dim FileHandle As Integer
'  Dim FileSize As Long
'  FileHandle = FreeFile
'  Open "GLTRXED.old" For Binary Shared As FileHandle
'  FileSize = LOF(FileHandle)
'  Close FileHandle
  'If FileSize > 0 Then
  KillFileD "GLTRXED.old"
  'End If
  SH_Rename "GLTRXED.DAT", "GLTRXED.old"
  SH_CopyFile "GLSumEd.DAT", "GLTRXED.DAT"
  MsgBox "Summarization Is Complete.", vbOKOnly, "Transactions Summarized"
'If Exist("GLTRXED.old") Then
Exit Sub

CancelExit:
  Exit Sub
End Sub

