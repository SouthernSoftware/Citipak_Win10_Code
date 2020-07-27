VERSION 5.00
Begin VB.Form frmBudgetMaintMenu1 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Budget Maintenance Menu"
   ClientHeight    =   8868
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   12228
   Icon            =   "frmBudgetMaint1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12228
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBudgetPrep 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Budget Preparation &Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H8000000F&
      TabIndex        =   3
      Top             =   6120
      Width           =   3612
   End
   Begin VB.CommandButton cmdEnterEditBudget 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Enter/Edit Budget Transactions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   3240
      Width           =   3612
   End
   Begin VB.CommandButton cmdPrintBgtTrans 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Print Transactions Register"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      TabIndex        =   1
      Top             =   4080
      Width           =   3612
   End
   Begin VB.CommandButton cmdPostBgtEntries 
      Caption         =   "Post Entries to &Budget"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      TabIndex        =   2
      Top             =   4920
      Width           =   3612
   End
   Begin VB.CommandButton cmdExitBudgetMaintMenu 
      Caption         =   "E&xit Budget Maintenance Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4320
      TabIndex        =   4
      Top             =   6960
      Width           =   3612
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000E&
      X1              =   3720
      X2              =   8520
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BUDGET MAINTENANCE MENU"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   2640
      TabIndex        =   5
      Top             =   1440
      Width           =   6852
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H8000000E&
      Height          =   1092
      Left            =   1800
      Top             =   1080
      Width           =   8652
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H8000000E&
      Height          =   132
      Left            =   8880
      Top             =   2280
      Width           =   972
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Left            =   8880
      Top             =   2160
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000E&
      Height          =   132
      Left            =   2400
      Top             =   2280
      Width           =   972
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Left            =   2400
      Top             =   2160
      Width           =   972
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9000
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   9000
      X2              =   9720
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   2400
      Y2              =   8280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2520
      X2              =   3240
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1800
      Top             =   960
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   9000
      Top             =   2400
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2520
      Top             =   2400
      Width           =   732
   End
End
Attribute VB_Name = "frmBudgetMaintMenu1"
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

Private Sub cmdBudgetPrep_Click()
  frmBudPrepOptions.Show
  Unload frmBudgetMaintMenu
End Sub

Private Sub cmdEnterEditBudget_Click()
  Dim FileHandle As Integer, WhosOnFirst As String
  If Exist("BGTED.opn") Then
    FileHandle = FreeFile
    Open "BGTED.opn" For Input As FileHandle
    Line Input #FileHandle, WhosOnFirst$
    Close FileHandle
    MsgBox "The Budget Edit File Is In Use By: " + WhosOnFirst$, vbOKOnly, "File Not Accessible"
  Else
    FileHandle = FreeFile
    Open "BGTED.opn" For Output As FileHandle
    Print #FileHandle, ComputerName$
    Close FileHandle
    frmBudgetEntEdit.Show
    Unload frmBudgetMaintMenu
  End If
 
End Sub

Private Sub cmdExitBudgetMaintMenu_Click()
  frmGLMainMenu.Show
  Unload frmBudgetMaintMenu
End Sub

Private Sub cmdPostBgtEntries_Click()
  Dim FileHandle As Integer, WhosOnFirst As String
  If Exist("BGTED.opn") Then
    FileHandle = FreeFile
    Open "BGTED.opn" For Input As FileHandle
    Line Input #FileHandle, WhosOnFirst$
    Close FileHandle
    MsgBox "The Budget Edit File Is In Use By: " + WhosOnFirst$, vbOKOnly, "File Not Accessible"
  Else
    FileHandle = FreeFile
    Open "BGTED.opn" For Output As FileHandle
    Print #FileHandle, ComputerName$
    Close FileHandle
    frmPostBgt.Show 1
  End If
End Sub

Private Sub cmdPrintBgtTrans_Click()
  'DeActivateControls frmBudgetMaintMenu
  frmReportOpt.Show 1
  If rptopt = 1 Then
    Call PrnEditList
  ElseIf rptopt = 2 Then
    Call PrnEditList2
  End If
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      cmdExitBudgetMaintMenu_Click
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
End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitBudgetMaintMenu.Enabled = False Then
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
   ActivateControls frmBudgetMaintMenu
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
   ActivateControls frmBudgetMaintMenu
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

