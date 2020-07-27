VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Begin VB.Form frmBudgetPrepMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Budget Preparation Menu"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   Icon            =   "frmBudgetPrepMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   12225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin fpBtnAtlLibCtl.fpBtn cmdEnterEditBudPrep 
      Height          =   492
      Left            =   4308
      TabIndex        =   1
      Top             =   3336
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmBudgetPrepMenu.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintBgtWorksheet 
      Height          =   492
      Left            =   4308
      TabIndex        =   2
      Top             =   4056
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmBudgetPrepMenu.frx":0ABF
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExitMenu 
      Height          =   492
      Left            =   4296
      TabIndex        =   3
      Top             =   4776
      Width           =   3612
      _Version        =   131072
      _ExtentX        =   6371
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmBudgetPrepMenu.frx":0CB2
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BUDGET PREPARATION MENU"
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
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   1440
      Width           =   6852
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
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   132
      Left            =   2400
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
Attribute VB_Name = "frmBudgetPrepMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim BgtEdit As TrEditRecType
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLAcctidx As GLAcctIndexType
Dim GLAcct As GLAcctRecType
Private Temp_Class As Resize_Class


Private Sub cmdEnterEditBudPrep_Click()
  frmBudPrepOptions.Show
  Unload frmBudgetPrepMenu

End Sub

Private Sub cmdExitMenu_Click()
  frmBudgetMaintMenu.Show
  Unload Me
End Sub


Private Sub cmdPrintBgtWorksheet_Click()
'  'DeActivateControls frmBudgetMaintMenu
'  frmReportOpt.Show 1
'  If rptopt = 1 Then
'    Call PrnEditList
'  ElseIf rptopt = 2 Then
'    Call PrnEditList2
'  End If
  frmPrnBudPrepWork.Show
  Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitMenu_Click
      KeyCode = 0
      DoEvents
    Case Else:
  End Select
End Sub
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Me.HelpContextID = hlpBudgetPreparation
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
    If cmdExitMenu.Enabled = False Then
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

Private Sub PrnEditList()
  Dim BgtEditFileNum As Integer, NumEdTrans As Integer
  Dim LookFor As String
  Dim BgtEdLen As Integer
  Dim ReportFile As String, ToPrint As String, CrLF As String
  Dim Header As String, ErrMsg As String
  Dim PRNFileNum As Integer, cnt As Integer, Howmany As Integer
  Dim FundCode As String, PageNum As Integer, Newrp As String
  Dim CommaFmt As String, FundNum As String
  Dim TotDr As Double, TotCr As Double, TranCashTot As Double
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  PRNFileNum = FreeFile
  Newrp = "BGTREG"
  GetRPTName Newrp
  ReportFile$ = Newrp
  Open ReportFile$ For Output As #PRNFileNum
  CommaFmt$ = "###,###,###.##"
  
  
  BgtEdLen = Len(BgtEdit)
  BgtEditFileNum = FreeFile
  Open "BGTED.dat" For Random As BgtEditFileNum Len = BgtEdLen
  NumEdTrans = LOF(BgtEditFileNum) \ BgtEdLen
  
  For cnt = 1 To NumEdTrans
    Get BgtEditFileNum, cnt, BgtEdit
    If BgtEdit.Deleted = 0 Then
      

         '--First Line
         ToPrint$ = ""
         ToPrint$ = Format(DateAdd("d", (BgtEdit.TRDATE), "12-31-1979"), "mm/dd/yyyy")
         ToPrint$ = ToPrint$ + "~" + BgtEdit.Desc + "~" + BgtEdit.Ref
         ToPrint$ = ToPrint$ + "~" + BgtEdit.AcctNum + "~" + BgtEdit.AcctName
         ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, BgtEdit.DrAmt)
         ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, BgtEdit.CrAmt)
         Print #PRNFileNum, ToPrint$


         TotDr# = TotDr# + BgtEdit.DrAmt
         TotCr# = TotCr# + BgtEdit.CrAmt

      End If
   Next

  Close
   Load frmLoadingRpt
   ActivateControls frmBudgetPrepMenu
   If TotDr# <> TotCr# Then
     ARptEditList.Label9.Visible = True
     ARptEditList.Label9.Caption = "Entries Do Not Balance!"
   End If
   ARptEditList.RType = 2
   ARptEditList.totDebits = Using$(CommaFmt$, TotDr#)
   ARptEditList.totCredits = Using$(CommaFmt$, TotCr#)
   ARptEditList.txtDate = Now
   ARptEditList.txtTown = GLUserName$
   ARptEditList.Title = "Budget Register Report"
   ARptEditList.GetName2 ReportFile$
   ARptEditList.startrpt

Exit Sub

End Sub
Private Sub PrnEditList2()
  Dim BgtEditFileNum As Integer, NumEdTrans As Integer
  Dim MaxLines As Integer, LookFor As String
  Dim Linecnt As Integer, BgtEdLen As Integer
  Dim ReportFile As String, ToPrint As String, CrLF As String
  Dim FF As String, Header As String, ErrMsg As String
  Dim PRNFileNum As Integer, cnt As Integer, Howmany As Integer
  Dim FundCode As String, PageNum As Integer, Newrp As String
  Dim CommaFmt As String, FundNum As String
  Dim TotDr As Double, TotCr As Double, TranCashTot As Double
  GetAcctStruct GLUserName$, GLFundLen, GLAcctLen, GLDetLen
  PRNFileNum = FreeFile
  Newrp = "BGTREG"
  GetRPTName Newrp
  ReportFile$ = Newrp
  Open ReportFile$ For Output As #PRNFileNum
  CommaFmt$ = "###,###,###.##"
  
    '--Report Vars
  MaxLines = 55
  CrLF$ = Chr$(13) + Chr$(10)
  FF$ = Chr$(12)
  GoSub PrintBgtHeader
  
  BgtEdLen = Len(BgtEdit)
  BgtEditFileNum = FreeFile
  Open "BGTED.dat" For Random As BgtEditFileNum Len = BgtEdLen
  NumEdTrans = LOF(BgtEditFileNum) \ BgtEdLen
  
  For cnt = 1 To NumEdTrans
    Get BgtEditFileNum, cnt, BgtEdit
    If BgtEdit.Deleted = 0 Then
      

         '--First Line
         ToPrint$ = Space$(80)
         LSet ToPrint$ = Format(DateAdd("d", (BgtEdit.TRDATE), "12-31-1979"), "mm/dd/yyyy")
         Mid$(ToPrint$, 13) = BgtEdit.AcctNum
         Mid$(ToPrint$, 27) = BgtEdit.AcctName
         Mid$(ToPrint$, 50) = Using$(CommaFmt$, BgtEdit.DrAmt)
         Mid$(ToPrint$, 65) = Using$(CommaFmt$, BgtEdit.CrAmt)
         Print #PRNFileNum, ToPrint$
         Linecnt = Linecnt + 1
         If Linecnt > MaxLines Then
            Print #PRNFileNum, FF$
            GoSub PrintBgtHeader
         End If

         '--2nd Line
         ToPrint$ = Space$(80)
         Mid$(ToPrint$, 13) = BgtEdit.Desc
         Mid$(ToPrint$, 35) = BgtEdit.Ref
         Print #PRNFileNum, ToPrint$
         Linecnt = Linecnt + 1
         If Linecnt > MaxLines Then
            Print #PRNFileNum, FF$
            GoSub PrintBgtHeader
         End If

         '--3rd line is blank
         Print #PRNFileNum,
         Linecnt = Linecnt + 1
         If Linecnt > MaxLines Then
            Print #PRNFileNum, FF$
            GoSub PrintBgtHeader
         End If

         TotDr# = TotDr# + BgtEdit.DrAmt
         TotCr# = TotCr# + BgtEdit.CrAmt

      End If
   Next
  
      Print #PRNFileNum, String$(80, "-")

   ToPrint$ = Space$(80)
   LSet ToPrint$ = "File Totals"
   Mid$(ToPrint$, 50) = Using$(CommaFmt$, TotDr#)
   Mid$(ToPrint$, 65) = Using$(CommaFmt$, TotCr#)
   Print #PRNFileNum, ToPrint$
   Print #PRNFileNum, FF$
   
   If TotDr# <> TotCr# Then
     ToPrint$ = Space$(80)
     ErrMsg$ = "The Debits and Credits Do Not Balance, You May Wish To Correct Before Posting."
     LSet ToPrint$ = ErrMsg$
     Print #PRNFileNum, ToPrint$
  End If
   Close
   ActivateControls frmBudgetPrepMenu
   ViewPrint ReportFile$, "Budget Register Report"
   Kill ReportFile$
Exit Sub

PrintBgtHeader:
  PageNum = PageNum + 1
  Print #PRNFileNum, "Budget Register"
  Print #PRNFileNum, Tab(70); "Page: "; PageNum
  Print #PRNFileNum, "Date        G/L Account                                  Debit          Credit"
  Print #PRNFileNum, "            Description           Reference "
  Print #PRNFileNum, String$(80, "-")
  Linecnt = 5
Return
End Sub

