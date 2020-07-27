VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUBPaymentMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payments, Deposits Menu"
   ClientHeight    =   8868
   ClientLeft      =   3924
   ClientTop       =   1884
   ClientWidth     =   12216
   Icon            =   "frmUBPaymentMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDelDep 
      Caption         =   "De&lete Deposit Transaction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      TabIndex        =   3
      Top             =   4500
      Width           =   4524
   End
   Begin VB.CommandButton cmdExitMenu 
      Caption         =   "E&xit to Previous Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      TabIndex        =   6
      Top             =   6456
      Width           =   4524
   End
   Begin VB.CommandButton cmdDelPay 
      Caption         =   "D&elete Payment Transaction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      TabIndex        =   2
      Top             =   3852
      Width           =   4524
   End
   Begin VB.CommandButton cmdPostPayments 
      Caption         =   "P&ost Transaction Entries"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      TabIndex        =   5
      Top             =   5808
      Width           =   4524
   End
   Begin VB.CommandButton cmdPrintJournal 
      Caption         =   "P&rint Transaction Journal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      TabIndex        =   4
      Top             =   5148
      Width           =   4524
   End
   Begin VB.CommandButton cmdDeposits 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Deposit Transaction Entry/Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      TabIndex        =   1
      Top             =   3192
      Width           =   4524
   End
   Begin VB.CommandButton cmdPayment 
      BackColor       =   &H008F8265&
      Caption         =   "&Payment Transaction Entry/Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3864
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   2544
      Width           =   4524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7154
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "11:33 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "4/16/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1788
      Top             =   744
      Width           =   8652
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   2508
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      X1              =   2508
      X2              =   3228
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   3348
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   2400
      X2              =   3360
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   2388
      X2              =   2388
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   0
      X1              =   3348
      X2              =   3348
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   8988
      Y1              =   2064
      Y2              =   7944
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   8988
      X2              =   9708
      Y1              =   7944
      Y2              =   7944
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   8868
      X2              =   9828
      Y1              =   2064
      Y2              =   2064
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      Index           =   3
      X1              =   8868
      X2              =   9828
      Y1              =   1944
      Y2              =   1944
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   8868
      X2              =   8868
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   9828
      X2              =   9828
      Y1              =   1944
      Y2              =   2064
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payments, Deposits Menu"
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
      Left            =   3540
      TabIndex        =   8
      Top             =   1104
      Width           =   5148
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2508
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   2388
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8988
      Top             =   2064
      Width           =   732
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   8868
      Top             =   1824
      Width           =   972
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1788
      Top             =   624
      Width           =   8652
   End
End
Attribute VB_Name = "frmUBPaymentMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim DefPayDate As String
Public Sub setstuff(dt As String)
DefPayDate = dt
End Sub

Private Sub cmdDelDep_Click()
  frmPaymentDelete.Wherefrom 2
  'Load frmPaymentDelete
  DoEvents
  frmPaymentDelete.Show

End Sub

Private Sub cmdDelPay_Click()
  frmPaymentDelete.Wherefrom 1
  'Load frmPaymentDelete
  DoEvents
  frmPaymentDelete.Show

End Sub

Private Sub cmdDeposits_Click()
  Dim FntSize As Integer
  If Not Exist("C:\RcptPrn.dat") Then
    ReDim MsgText(0 To 5) As String
    FntSize = frmMsgDialog.Label(1).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    MsgText(0) = "WARNING:"
    MsgText(1) = ""
    MsgText(2) = "RECEIPT SETUP FILE NOT FOUND!"
    MsgText(3) = "If you continue receipt printing"
    MsgText(4) = "will be disabled."
    MsgText(5) = "Receipt setup option is on CitiPak Main Menu."
    If GetOKorNot(MsgText()) Then
     UBLog "USER WANTS TO CONTINUE!"
    Else
     UBLog "USER ABORTED."
     Exit Sub
   End If
  End If
  frmDepositPayment.Wheretogo frmUBPaymentMenu, frmUBPaymentMenu, , DefPayDate
  DoEvents
  frmDepositPayment.Show
  'Unload frmUBPaymentMenu
End Sub

'Private Sub cmdCustConsumpRpt_Click()
'  frmCustEditLookUP.Caption = "Customer Consumption History"
'  frmCustEditLookUP.Label1.Caption = "Customer Consumption History"
'  frmCustEditLookUP.Wheretogo frmUBCustMenu, frmRptCustConsHist
'  DoEvents
'  frmCustEditLookUP.Show
'  Unload frmUBCustMenu
''  Load frmRptCustHistory
''  frmRptCustHistory.RptType = True
''  DoEvents
''  frmRptCustHistory.Caption = "Customer Consumption History"
''  frmRptCustHistory.Label1 = frmRptCustHistory.Caption
''  frmRptCustHistory.fpDetailFlag.Visible = False
''  frmRptCustHistory.DetailLabel.Visible = False
''  frmRptCustHistory.Wheretogo frmUBCustMenu, frmRptCustHistory
''  frmRptCustHistory.Show
''  Unload frmUBCustMenu
'End Sub
'
'Private Sub cmdCustTransRpt_Click()
'  frmCustEditLookUP.Caption = "Customer Transaction History"
'  frmCustEditLookUP.Label1.Caption = "Customer Transaction History"
'  frmCustEditLookUP.Wheretogo frmUBCustMenu, frmRptCustTranHist
'  DoEvents
'  frmCustEditLookUP.Show
'  Unload frmUBCustMenu
'
''  Load frmRptCustHistory
''  frmRptCustHistory.RptType = False
''  DoEvents
''  frmRptCustHistory.Caption = "Customer Transaction History"
''  frmRptCustHistory.Label1 = frmRptCustHistory.Caption
''  frmRptCustHistory.Wheretogo frmUBCustMenu, frmRptCustHistory
''  frmRptCustHistory.Show
''  Unload frmUBCustMenu
'End Sub
'
'Private Sub cmdDeleteCustomer_Click()
'  frmCustEditLookUP.Caption = "Customer Delete Search"
'  frmCustEditLookUP.Label1.Caption = "Customer Delete Search"
'  frmCustEditLookUP.Wheretogo frmUBCustMenu, frmCustDelete, , 1
'  'Load frmCustEditLookUP
'  DoEvents
'  frmCustEditLookUP.Show
'  Unload frmUBCustMenu
'
'End Sub
'
'Private Sub cmdEditCustomer_Click()
'  frmCustEditLookUP.Wheretogo frmUBCustMenu, frmCustAddEdit
''  Load frmCustEditLookUP
'  DoEvents
'  frmCustEditLookUP.Show
'  Unload frmUBCustMenu
'End Sub

Private Sub cmdExitMenu_Click()
  Load frmUBMainMenu
  DoEvents
  frmUBMainMenu.Show
  Unload Me
End Sub


Private Sub cmdPayment_Click()
  Dim FntSize As Integer
  If Not Exist("C:\RcptPrn.dat") Then
    ReDim MsgText(0 To 5) As String
    FntSize = frmMsgDialog.Label(1).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    MsgText(0) = "WARNING:"
    MsgText(1) = ""
    MsgText(2) = "RECEIPT SETUP FILE NOT FOUND!"
    MsgText(3) = "If you continue receipt printing"
    MsgText(4) = "will be disabled."
    MsgText(5) = "Receipt setup option is on CitiPak Main Menu."
    If GetOKorNot(MsgText()) Then
      UBLog "USER WANTS TO CONTINUE!"
    Else
      UBLog "USER ABORTED."
      Exit Sub
    End If
  End If
  frmPaymentEntry.Wheretogo frmUBPaymentMenu, frmUBPaymentMenu, , DefPayDate
  DoEvents
  frmPaymentEntry.Show
  'Unload frmUBPaymentMenu
End Sub

Private Sub cmdPostPayments_Click()
  Dim FntSize As Integer, PayBillName As String, PayDepoName As String
  'Dim OPERNUM As Integer
  
  PayBillName$ = UBPath$ + "UBPAY" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  PayDepoName$ = UBPath$ + "UBDEP" + QPTrim$(Str$(OPERNUM)) + ".DAT"

  If FileSize&(PayBillName$) <= 0 And FileSize&(PayDepoName$) <= 0 Then
  ReDim MsgText(0 To 5) As String
   
     frmMsgDialog.RetLabel = "-2"
     FntSize = frmMsgDialog.Label(2).FontSize
     
     frmMsgDialog.Label(2).FontSize = (FntSize + 2)
     MsgText(0) = "ERROR:"
     MsgText(1) = ""
     MsgText(2) = ""
     MsgText(3) = "NO UNPOSTED PAYMENT TRANSACTIONS!"
     MsgText(4) = ""
     MsgText(5) = ""
     GetOKorNot MsgText(), True
    GoTo Exitthis
  End If
  UBLog " IN: UB POST PAYMENTS,  OPER:" + Str$(OPERNUM)

  DoItFlag = False
    frmNoOperatorsWarning.Label(5).Caption = "Post Payment Transactions"
    Load frmNoOperatorsWarning
    frmNoOperatorsWarning.Show vbModal
    If Not DoItFlag Then
      GoTo Exitthis
    End If
  DeActivateControls Me
  PostPayments
  ActivateControls Me
  MsgBox "Posting Payments Completed.", vbOKOnly, "Procedure Complete"
Exitthis:

End Sub

Private Sub cmdPrintJournal_Click()
  Dim FntSize As Integer, PayFileName As String, DepFileName As String
  Dim Tot As Integer
  Tot = 0
  PayFileName$ = UBPath$ + "UBPAY" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  DepFileName$ = UBPath$ + "UBDEP" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  If Exist(PayFileName$) Then Tot = Tot + 1
  If Exist(DepFileName$) Then Tot = Tot + 1
  ReDim MsgText(0 To 5) As String
   If Tot < 1 Then
     frmMsgDialog.RetLabel = "-2"
     FntSize = frmMsgDialog.Label(2).FontSize
     
     frmMsgDialog.Label(2).FontSize = (FntSize + 2)
     MsgText(0) = "ERROR:"
     MsgText(1) = ""
     MsgText(2) = ""
     MsgText(3) = "NO UNPOSTED PAYMENT TRANSACTIONS!"
     MsgText(4) = ""
     MsgText(5) = ""
     GetOKorNot MsgText(), True
  Else
    frmReportOpt.Show 1
    DeActivateControls Me
    If rptopt > 0 Then
     PrintTransJournal rptopt
    Else
      ActivateControls Me
    End If
    
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  StatusBar1.Panels.Item(1).Text = TOWNNAME$
  'screenW = (Screen.Width / Screen.TwipsPerPixelX)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        UBLog "Closed via PaymentMenu by " + PWUser$
        CitiTerminate
      End If
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    Temp_Class.ResizeControls Me
   ' Me.Visible = True
   ' Me.SetFocus
  End If
  DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyHome
      cmdPayment.SetFocus
    Case vbKeyEnd
      cmdExitMenu.SetFocus
    Case Else:
  End Select
End Sub

Private Sub PrintTransJournal(rptopt As Integer) '(OPERNUM, PostDate$)
  Dim NumofRevs As Integer, UBSetupLen As Integer, RevCnt As Integer
  Dim InvRev As Integer, cnt As Long, x As Integer, Dash1 As String
  Dim UBCustRecLen As Integer, LastRev As Integer
  Dim TempRev As String, Operator As String, Page As Integer
  Dim PayFileName As String, DepFileName As String, PayJourName As String
  Dim Header As String, CMOperRecLen As Integer, UBPayRecLen As Integer
  Dim CMFile As Integer, NumRecs As Long, PayOKFlag As Boolean
  Dim DepOKFlag As Boolean, CustFile As Integer, TotalRecs As Long
  Dim RptHandle As Integer, PHandle As Integer, NumOfRecs As Long
  Dim DoneCnt As Integer, TaxExempt As Boolean, Pmnt As String
  Dim CustBook As Integer, TotalCash As Double, TotalCheck As Double
  Dim TotalAmount As Double, TotalChange As Double, TotalReceipts As Integer
  Dim RCnt As Integer, Diff As Double, Tax As Double, DepositTot As Double
  Dim GTotal As Double, TTax As Double, PostDate As String, ToPrint As String
  Dim Graph As Boolean, ReportSum1 As String, ReportSum2 As String
  Dim SumRpt1 As Integer, SumRpt2 As Integer, TotalChrge As Double
  Dim tmp As DistArrayType, SumPrnt As String, TotalChks As Integer
  NumofRevs = MaxRevsCnt
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  
  ToPrint$ = ""
  If rptopt = 1 Then
    Graph = True
  Else
    Graph = False
  End If
  FrmShowPctComp.Label1 = "Creating Payment Journal Report"
  FrmShowPctComp.Show

  PostDate$ = Format(Now, "mm/dd/yyyy")
'  If InStr(UBSetUpRec(1).UTILNAME, "HARRISBURG") > 0 Then
'    BlockClear
'    OK = MsgBox("UBSETUP", "FINONLY")
'    Select Case OK
'    Case 1
'      StatusFlag = False
'    Case Else
'      StatusFlag = True
'    End Select
'  Else
'    StatusFlag = False
'  End If
  ReDim RevText$(1 To MaxRevsCnt)

  ReDim TaxRates(1 To 15) As Single
  ReDim TaxAmt(1 To 15) As Double

  ReDim BookTotals(0 To 99) As BookTotalType

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))

  'IF StatusFlag THEN


  'IF INSTR(UBSetUp(1).UTILNAME, "AUTRY") > 0 THEN
  '  LptPort = 2
  'ELSE
  '  LPTPORT = 1
  'END IF

  LastRev = 15
   For cnt = 1 To MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(cnt).RevName)
    If Len(TempRev$) = 0 Then
      LastRev = cnt - 1
      Exit For
    Else
      RevText$(cnt) = TempRev$
      TaxRates(cnt) = UBSetUpRec(1).Revenues(cnt).TAXRATE
    End If
  Next

  ReDim RevAmts(1 To LastRev) As Double
  ReDim DepRevs(1 To LastRev) As Double

  'CursorOff
  Operator$ = QPTrim$(Str$(OPERNUM))
  FF$ = Chr$(12)
  Page = 0
  LineCnt = 0
  MaxLines = 55
  Dash1$ = String$(120, "-")
  PayFileName$ = UBPath$ + "UBPAY" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  DepFileName$ = UBPath$ + "UBDEP" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  PayJourName$ = UBPath$ + "UBPAY" + QPTrim$(Str$(OPERNUM)) + ".RPT"

  Header$ = "Utility Payment/Deposit Journal"

  ReDim CMOperRec(1) As CMOperRecType
  CMOperRecLen = Len(CMOperRec(1))

  ReDim UBPaymentRec(1) As UBPaymentRecType
  UBPayRecLen = Len(UBPaymentRec(1))

'  CMFile = FreeFile
'  Open "CMOPER.DAT" For Random Shared As CMFile Len = CMOperRecLen
'  NumRecs = LOF(CMFile) \ CMOperRecLen
'
'  For cnt = 1 To NumRecs
'    Get CMFile, cnt, CMOperRec(1)
'    If CMOperRec(1).OperatorNumber = OPERNUM Then
'      Operator$ = QPTrim$(CMOperRec(1).OperatorName)
'      Exit For
'    End If
'  Next
'  Close CMFile

  If Exist(PayFileName$) And FileSize(PayFileName$) > 0 Then
    PayOKFlag = True
  End If
  If Exist(DepFileName$) And FileSize(DepFileName$) > 0 Then
    DepOKFlag = True
  End If

  If Not DepOKFlag And Not PayOKFlag Then
'    BlockClear
'    DisplayUBScrn "NOPAYJUR"
'    QPrintRC Str$(OPERNUM), 12, 34, 79
'    WaitForAction
    GoTo ExitJournal
  End If

  CustFile = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  TotalRecs& = (FileSize(DepFileName$) + FileSize(PayFileName$)) \ UBPayRecLen
  RptHandle = FreeFile
  Open PayJourName$ For Output As RptHandle
  If Graph Then
    ReportSum1$ = UBPath$ + "UBPJSUM1.RPT"
    SumRpt1 = FreeFile
    Open ReportSum1$ For Output As SumRpt1
    ReportSum2$ = UBPath$ + "UBPJSUM2.RPT"
    SumRpt2 = FreeFile
    Open ReportSum2$ For Output As SumRpt2
  End If
  If Not Graph Then
    GoSub PrintRptHeader
  End If
  If PayOKFlag Then
    'FOpenS PayFileName$, PHandle
    PHandle = FreeFile
    Open PayFileName$ For Random Shared As PHandle Len = UBPayRecLen
    NumOfRecs& = LOF(PHandle) \ UBPayRecLen
    For cnt& = 1 To NumOfRecs&
      FrmShowPctComp.ShowPctComp cnt&, NumOfRecs&
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Close
        Exit Sub
      End If

      Get PHandle, cnt&, UBPaymentRec(1)
      Get CustFile, UBPaymentRec(1).CustAcct, UBCustRec(1)
'      If StatusFlag Then
'        If UBPaymentRec(1).Status <> "F" Then
'          GoTo OnlyFinalSkip
'        End If
'      End If

      GoSub GetCustBook
      DoneCnt = DoneCnt + 1
      If UBPaymentRec(1).TaxExempt = "Y" Then
        TaxExempt = True
      Else
        TaxExempt = False
      End If
      If Not Graph Then
        If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintRptHeader
        End If
      End If
      If UBPaymentRec(1).CASHAMT < 0 Then UBPaymentRec(1).CASHAMT = 0
      If UBPaymentRec(1).CHKAMT < 0 Then UBPaymentRec(1).CHKAMT = 0
      If Not Graph Then
        Print #RptHandle, Num2Date(UBPaymentRec(1).PAYDATE);
        Print #RptHandle, Tab(13); Using("#####", UBPaymentRec(1).CustAcct);
      Else
        ToPrint$ = Num2Date(UBPaymentRec(1).PAYDATE) + "~"
        ToPrint$ = ToPrint$ + Using("#####", UBPaymentRec(1).CustAcct) + "~"
      End If
      If OPERNUM = 99 Then
        Pmnt$ = " DFT"
      Else
        Pmnt$ = " PMT"
      End If

      'IF UBPaymentRec(1).Status = "F" THEN
      '  Pmnt$ = Pmnt$ + "*F"
      'END IF
      If Not Graph Then
        Print #RptHandle, Tab(20); UBPaymentRec(1).CustName; Pmnt$;
        Print #RptHandle, Tab(55); Using("######.##", UBPaymentRec(1).CASHAMT);
        If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
          Print #RptHandle, Tab(70); Using("######.##", 0);
          Print #RptHandle, Tab(84); Using("######.##", UBPaymentRec(1).CHKAMT);
        Else
          Print #RptHandle, Tab(70); Using("######.##", UBPaymentRec(1).CHKAMT);
          Print #RptHandle, Tab(84); Using("######.##", 0);
        End If
        Print #RptHandle, Tab(99); Using("######.##", Round#(Round#(UBPaymentRec(1).CHKAMT + UBPaymentRec(1).CASHAMT) - UBPaymentRec(1).Change));
        Print #RptHandle, Tab(112); Using("######.##", UBPaymentRec(1).Change)
      Else
        ToPrint$ = ToPrint$ + QPTrim(UBPaymentRec(1).CustName) + "~" + Pmnt$ + "~"
        ToPrint$ = ToPrint$ + Using("#####,#.##", UBPaymentRec(1).CASHAMT) + "~"
        If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
          ToPrint$ = ToPrint$ + "0~"
        End If
        ToPrint$ = ToPrint$ + Using("#####,#.##", UBPaymentRec(1).CHKAMT) + "~"
        If QPTrim(UBPaymentRec(1).TENDERTY) <> "Charge" Then
          ToPrint$ = ToPrint$ + "0~"
        End If
        ToPrint$ = ToPrint$ + Using("#####,#.##", Round#(Round#(UBPaymentRec(1).CHKAMT + UBPaymentRec(1).CASHAMT) - UBPaymentRec(1).Change)) + "~"
        ToPrint$ = ToPrint$ + Using("#####,#.##", UBPaymentRec(1).Change)
        Print #RptHandle, ToPrint$
        ToPrint$ = ""
      End If
      If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
        BookTotals(CustBook).Charge = Round#(BookTotals(CustBook).Charge + UBPaymentRec(1).CHKAMT)
        TotalChrge# = Round#(TotalChrge# + UBPaymentRec(1).CHKAMT)
      Else
        BookTotals(CustBook).Check = Round#(BookTotals(CustBook).Check + UBPaymentRec(1).CHKAMT)
        TotalCheck# = Round#(TotalCheck# + UBPaymentRec(1).CHKAMT)
        If UBPaymentRec(1).CHKAMT > 0 Then
          TotalChks = TotalChks + 1
        End If
      End If
      BookTotals(CustBook).Cash = Round#(BookTotals(CustBook).Cash + UBPaymentRec(1).CASHAMT)
      BookTotals(CustBook).Change = Round#(BookTotals(CustBook).Change + UBPaymentRec(1).Change)
      
      TotalCash# = Round#(TotalCash# + UBPaymentRec(1).CASHAMT)
      'TotalCheck# = Round#(TotalCheck# + UBPaymentRec(1).CHKAMT)
      TotalAmount# = Round#(TotalAmount# + UBPaymentRec(1).AMTPAID)
      TotalChange# = Round#(TotalChange# + UBPaymentRec(1).Change)
      TotalReceipts = TotalReceipts + 1

      LineCnt = LineCnt + 1

      For RCnt = 1 To LastRev
        If Not TaxExempt Then
          If TaxRates(RCnt) > 0 Then
            Diff# = Round#(UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 / (1 + TaxRates(RCnt)))
            Tax# = Round#(UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 - Diff#)
            TaxAmt(RCnt) = Round#(TaxAmt(RCnt) + Tax#)
            RevAmts(RCnt) = Round#(RevAmts(RCnt) + (UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 - Tax#))
          Else
            RevAmts(RCnt) = Round#(RevAmts(RCnt) + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
            'IF UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 < -10000 OR UBPaymentRec(1).PaidOwed(RCnt).AMTPD1 > 10000 THEN
            '  STOP
            'END IF
          End If
        Else
          RevAmts(RCnt) = Round#(RevAmts(RCnt) + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
        End If
      Next
      '*********************
      'ShowPctComp DoneCnt, TotalRecs&
OnlyFinalSkip:
    Next
    Close PHandle
  Else
    FrmShowPctComp.ShowPctComp 100, 100
  End If
    If DepOKFlag Then
    PHandle = FreeFile
    Open DepFileName$ For Random Shared As PHandle Len = UBPayRecLen
    NumOfRecs& = LOF(PHandle) \ UBPayRecLen
    For cnt& = 1 To NumOfRecs&
      Get PHandle, cnt&, UBPaymentRec(1)
      Get CustFile, UBPaymentRec(1).CustAcct, UBCustRec(1)
      GoSub GetCustBook

      DoneCnt = DoneCnt + 1
      If Not Graph Then
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintRptHeader
        End If
      End If
      If UBPaymentRec(1).CASHAMT < 0 Then UBPaymentRec(1).CASHAMT = 0
      If UBPaymentRec(1).CHKAMT < 0 Then UBPaymentRec(1).CHKAMT = 0
      If Not Graph Then
        Print #RptHandle, Num2Date(UBPaymentRec(1).PAYDATE);
        Print #RptHandle, Tab(13); Using("#####", UBPaymentRec(1).CustAcct);
        Print #RptHandle, Tab(20); UBPaymentRec(1).CustName; ; " DEP";
        Print #RptHandle, Tab(55); Using("######.##", UBPaymentRec(1).CASHAMT);
        If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
          Print #RptHandle, Tab(70); Using("######.##", 0);
          Print #RptHandle, Tab(84); Using("######.##", UBPaymentRec(1).CHKAMT);
        Else
          Print #RptHandle, Tab(70); Using("######.##", UBPaymentRec(1).CHKAMT);
          Print #RptHandle, Tab(84); Using("######.##", 0);
        End If
        Print #RptHandle, Tab(112); Using("######.##", UBPaymentRec(1).Change)
      Else
        ToPrint$ = Num2Date(UBPaymentRec(1).PAYDATE) + "~"
        ToPrint$ = ToPrint$ + Using("#####", UBPaymentRec(1).CustAcct) + "~"
        ToPrint$ = ToPrint$ + QPTrim(UBPaymentRec(1).CustName) + "~ DEP~"
        ToPrint$ = ToPrint$ + Using("#####,#.##", UBPaymentRec(1).CASHAMT) + "~"
        If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
          ToPrint$ = ToPrint$ + "0~"
        End If
        ToPrint$ = ToPrint$ + Using("#####,#.##", UBPaymentRec(1).CHKAMT) + "~"
        If QPTrim(UBPaymentRec(1).TENDERTY) <> "Charge" Then
          ToPrint$ = ToPrint$ + "0~"
        End If
        ToPrint$ = ToPrint$ + " ~"   'Using("#####,#.##", Round#(Round#(UBPaymentRec(1).CHKAMT + UBPaymentRec(1).CASHAMT) - UBPaymentRec(1).CHANGE)) + "~"
        ToPrint$ = ToPrint$ + Using("#####,#.##", UBPaymentRec(1).Change)
        Print #RptHandle, ToPrint$
        ToPrint$ = ""
      End If
      If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
        BookTotals(CustBook).Charge = Round#(BookTotals(CustBook).Charge + UBPaymentRec(1).CHKAMT)
      Else
        BookTotals(CustBook).Check = Round#(BookTotals(CustBook).Check + UBPaymentRec(1).CHKAMT)
      End If
      BookTotals(CustBook).Cash = Round#(BookTotals(CustBook).Cash + UBPaymentRec(1).CASHAMT)
      BookTotals(CustBook).Change = Round#(BookTotals(CustBook).Change + UBPaymentRec(1).Change)
      TotalChange# = Round#(TotalChange# + UBPaymentRec(1).Change)
      TotalCash# = Round#(TotalCash# + UBPaymentRec(1).CASHAMT)
      If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
        TotalChrge# = Round#(TotalChrge# + UBPaymentRec(1).CHKAMT)
      Else
        TotalCheck# = Round#(TotalCheck# + UBPaymentRec(1).CHKAMT)
      End If
      TotalReceipts = TotalReceipts + 1
      LineCnt = LineCnt + 1
      For RCnt = 1 To LastRev
        DepositTot# = Round#(DepositTot# + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
        DepRevs(RCnt) = Round#(DepRevs(RCnt) + UBPaymentRec(1).PaidOwed(RCnt).AMTPD1)
      Next
      'ShowPctComp DoneCnt, TotalRecs&
    Next
    Close PHandle
  Else
    FrmShowPctComp.ShowPctComp 100, 100
  End If

  GoSub PrintRptEnding

  Close
  'PrintRptFile Header$, PayJourName$, LPTPORT, -1, EntryPoint
  If Not Graph Then
    ViewPrint PayJourName$, Header$, True
    ActivateControls Me
  Else
    Load frmLoadingRpt
    frmLoadingRpt.setwherefrom frmUBPaymentMenu
    ARptPaymentJournal.Title = Header$
    ARptPaymentJournal.txtDate = Now
    ARptPaymentJournal.txtTown = TOWNNAME$
    ARptPaymentJournal.lblOperator = PWUser
    ARptPaymentJournal.GetName PayJourName$, ReportSum1$, ReportSum2$
    ARptPaymentJournal.startrpt
  End If
  'KillFile PayJourName$

ExitJournal:
  Exit Sub

GetCustBook:
  CustBook = Val(UBCustRec(1).Book)
  BookTotals(CustBook).Count = BookTotals(CustBook).Count + 1
Return

PrintRptHeader:
  If Not Graph Then
  Page = Page + 1
'  If StatusFlag Then
'    Print #RptHandle, "Utility Payment Receipts Journal (Final Only)"
'  Else

    Print #RptHandle, "Utility Payment/Deposit Receipts Journal"
'  End If
  Print #RptHandle, "Posting Date: "; PostDate$
  Print #RptHandle, "    Operator: "; PWUser; Tab(89); "Page #"; Page
  Print #RptHandle, ""
  Print #RptHandle, "       "; Tab(12); "        "; Tab(44); "           "; Tab(61); "            "; Tab(98); "Amount Paid"
  Print #RptHandle, " Date"; Tab(11); "Acct No      Customer"; Tab(60); "Cash"; Tab(74); "Check"; Tab(87); "Charge"; Tab(98); " on Account"; Tab(115); "Change"
  Print #RptHandle, Dash1$
  LineCnt = 6
 End If
 Return

PrintRptEnding:
  If Not Graph Then
    Print #RptHandle, Dash1$
    Print #RptHandle, "                  Totals: ";
    Print #RptHandle, Tab(54); Using("###,###.##", TotalCash#);
    Print #RptHandle, Tab(69); Using("###,###.##", TotalCheck#);
    Print #RptHandle, Tab(83); Using("###,###.##", TotalChrge#);
    Print #RptHandle, Tab(98); Using("###,###.##", TotalAmount#);
    Print #RptHandle, Tab(111); Using("###,###.##", TotalChange#)
    Print #RptHandle, "Total Number of Receipts: "; Using("##,###", TotalReceipts)
    Print #RptHandle, "  Total Number of Checks: "; Using("##,###", TotalChks)
    Print #RptHandle, FF$
  '  If StatusFlag Then
  '    Print #RptHandle, "Utility Payment Receipts Journal (Final Only)"
  '  Else
       Print #RptHandle, "Utility Payment/Deposit Receipts Journal"
  '  End If
    Print #RptHandle, "Revenue Summary"
    Print #RptHandle, ""
    Print #RptHandle, "    Revenue"; Tab(33); "Payments         Deposits               Tax"
    Print #RptHandle, Dash1$
    GTotal# = 0
    For RCnt = 1 To LastRev
      Print #RptHandle, Tab(5); RevText$(RCnt); Tab(31); Using("$###,###.##", RevAmts(RCnt)); Tab(48); Using("$###,###.##", DepRevs(RCnt)); Tab(66); Using("$###,###.##", TaxAmt(RCnt))
      GTotal# = Round#(GTotal# + RevAmts(RCnt))
      TTax# = Round#(TTax# + TaxAmt(RCnt))
    Next
    Print #RptHandle, Dash1$
    Print #RptHandle, "Revenue Totals:"; Tab(29); Using("$#,###,###.##", GTotal#); Tab(46); Using("$#,###,###.##", DepositTot#); Tab(64); Using("$#,###,###.##", TTax#)
    Print #RptHandle,             '"Deposit Total:"; TAB(35); USING "$$#####,#.##"; DepositTot#
    Print #RptHandle, "   Grand Total:"; Tab(29); Using("$#,###,###.##", Round#(GTotal# + DepositTot# + TTax#))
    Print #RptHandle, FF$
  '  If StatusFlag Then
  '    Print #RptHandle, "Utility Payment Receipts Journal (Final Only)"
  '  Else
      Print #RptHandle, "Utility Payment/Deposit Receipts Journal"
  '  End If
    Print #RptHandle, "Posting Date: "; PostDate$; "    Operator: "; Operator$
    Print #RptHandle, "Books Summary"
    Print #RptHandle, "Book#     Count            Cash          Check            Charge              Total            Change"
    Print #RptHandle, Dash1$
    For cnt = 0 To 99
      If BookTotals(cnt).Check <> 0 Or BookTotals(cnt).Cash <> 0 Or BookTotals(cnt).Charge <> 0 Then
        If cnt < 10 Then
          Print #RptHandle, Tab(2); Using("0#", cnt);
        Else
          Print #RptHandle, Tab(2); Using("##", cnt);
        End If
        Print #RptHandle, Tab(12); Using("####", BookTotals(cnt).Count);
        Print #RptHandle, Tab(22); Using("#######.##", BookTotals(cnt).Cash); Tab(37); Using("#######.##", BookTotals(cnt).Check); Tab(55); Using("#######.##", BookTotals(cnt).Charge); Tab(74); Using("#######.##", Round#(BookTotals(cnt).Cash + BookTotals(cnt).Check + BookTotals(cnt).Charge));
        Print #RptHandle, Tab(92); Using("#######.##", BookTotals(cnt).Change)
      End If
    Next
    Print #RptHandle, Dash1$
    Print #RptHandle, "Totals:"; Tab(12); Using("####", TotalReceipts);
    Print #RptHandle, Tab(22); Using("#######.##", TotalCash#);
    Print #RptHandle, Tab(37); Using("#######.##", TotalCheck#);
    Print #RptHandle, Tab(55); Using("#######.##", TotalChrge#);
    Print #RptHandle, Tab(74); Using("#######.##", Round#(TotalCash# + TotalCheck# + TotalChrge#));
    Print #RptHandle, Tab(92); Using("#######.##", TotalChange#)
    Print #RptHandle, FF$
  Else
    GTotal# = 0
    SumPrnt$ = ""
    For RCnt = 1 To LastRev
      SumPrnt$ = SumPrnt$ + RevText$(RCnt) + "~" + Using("$###,###.##", RevAmts(RCnt)) + "~" + Using("$###,###.##", DepRevs(RCnt)) + "~" + Using("$###,###.##", TaxAmt(RCnt)) + "~"
      GTotal# = Round#(GTotal# + RevAmts(RCnt))
      TTax# = Round#(TTax# + TaxAmt(RCnt))
      Print #SumRpt1, SumPrnt$
      SumPrnt$ = ""
    Next
    For cnt = 0 To 99
      If BookTotals(cnt).Check <> 0 Or BookTotals(cnt).Cash <> 0 Or BookTotals(cnt).Charge <> 0 Then
        If cnt < 10 Then
          SumPrnt$ = SumPrnt$ + Using("0#", cnt) + "~"
        Else
          SumPrnt$ = SumPrnt$ + Using("##", cnt) + "~"
        End If
        SumPrnt$ = SumPrnt$ + Using("####", BookTotals(cnt).Count) + "~"
        SumPrnt$ = SumPrnt$ + Using("#######.##", BookTotals(cnt).Cash) + "~"
        SumPrnt$ = SumPrnt$ + Using("#######.##", BookTotals(cnt).Check) + "~"
        SumPrnt$ = SumPrnt$ + Using("#######.##", BookTotals(cnt).Charge) + "~"
        SumPrnt$ = SumPrnt$ + Using("#######.##", Round#(BookTotals(cnt).Cash + BookTotals(cnt).Check + BookTotals(cnt).Charge)) + "~"
        SumPrnt$ = SumPrnt$ + Using("#######.##", BookTotals(cnt).Change)
        Print #SumRpt2, SumPrnt$
        SumPrnt$ = ""
      End If
    Next
  End If
  Return
End Sub

Private Sub PostPayments()
' OPERNUM , PostDate$
  Dim PayBillName As String, PayDepoName As String
  Dim UBCustRecLen As Integer, UBPayRecLen As Integer, UBTransRecLen As Integer
  Dim TranFile As Integer, CHandle As Integer, thandle As Integer
  Dim PHandle As Integer, NumPayRecs As Long, cnt As Long
  Dim RevAmts As Integer, NextTransRec As Long, TotalCustBalance As Double
  Dim CustChCnt As Integer
  
  PayBillName$ = UBPath$ + "UBPAY" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  PayDepoName$ = UBPath$ + "UBDEP" + QPTrim$(Str$(OPERNUM)) + ".DAT"
  FrmShowPctComp.Label1 = "Posting Deposit Transactions"
  FrmShowPctComp.Show

  UBLog "POSTING TRANSACTIONS START"

  ReDim UBTransRec(1) As UBTransRecType
  ReDim TUBTransRec(1) As UBTransRecType
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBPaymentRec(1) As UBPaymentRecType

  UBCustRecLen = Len(UBCustRec(1))
  UBPayRecLen = Len(UBPaymentRec(1))
  UBTransRecLen = Len(UBTransRec(1))

  TranFile = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As TranFile Len = UBTransRecLen
  Close TranFile


  UBLog "POSTING: DEPOSITS"
  If FileSize&(PayDepoName$) > 0 Then
    CHandle = FreeFile
    Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = UBCustRecLen
    thandle = FreeFile
    Open UBPath$ + "UBTRANS.DAT" For Random Shared As thandle Len = UBTransRecLen
    PHandle = FreeFile
    Open PayDepoName$ For Random Shared As PHandle Len = UBPayRecLen

    NumPayRecs& = LOF(PHandle) \ UBPayRecLen

    'ShowProcessingScrn "Posting Deposit Transactions"
    For cnt& = 1 To NumPayRecs&
      FrmShowPctComp.ShowPctComp cnt&, NumPayRecs&
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Close
        Exit Sub
      End If
      LSet UBTransRec(1) = TUBTransRec(1)
      Get PHandle, cnt&, UBPaymentRec(1) ',  UBPayRecLen
      Get CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1) ',  UBCustRecLen
      UBTransRec(1).TransDate = UBPaymentRec(1).PAYDATE
      UBTransRec(1).TransType = TranDepositPayment

      '022098 Added
      If Len(QPTrim$(UBPaymentRec(1).Desc)) = 0 Then
        UBTransRec(1).TransDesc = "DEPOSIT PAYMENT"
      Else
        UBTransRec(1).TransDesc = UBPaymentRec(1).Desc
        UBTransRec(1).BillMsg = "DEPOSIT PAYMENT"
      End If
      '^^This holds the Payment Description

      'UBTransRec(1)CustLocation = UBPaymentRec(1).CUSTACCT
      UBTransRec(1).OperatorNumber = OPERNUM
      UBTransRec(1).CustAcctNo = UBPaymentRec(1).CustAcct
      UBTransRec(1).CustStatus = UBCustRec(1).Status
      UBTransRec(1).Transamt = UBPaymentRec(1).AMTPAID
      UBTransRec(1).CheckAmount = UBPaymentRec(1).CHKAMT
      UBTransRec(1).CashAmount = UBPaymentRec(1).CASHAMT

      If UBTransRec(1).CheckAmount > 0 And UBTransRec(1).CashAmount > 0 Then
        UBTransRec(1).PayTypeCode = 3
      ElseIf UBTransRec(1).CashAmount > 0 Then
        UBTransRec(1).PayTypeCode = 1
      ElseIf UBTransRec(1).CheckAmount > 0 Then
        If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
          UBTransRec(1).PayTypeCode = 4
        Else
          UBTransRec(1).PayTypeCode = 2
        End If
      End If

      For RevAmts = 1 To MaxRevsCnt
        UBTransRec(1).RevAmt(RevAmts) = UBPaymentRec(1).PaidOwed(RevAmts).AMTPD1
      Next
   '05-05-97 FIX added run balance to deposit trans
   UBTransRec(1).RunBalance = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)

   UBTransRec(1).PrevTrans = UBCustRec(1).LastTrans
   NextTransRec& = (LOF(thandle) \ UBTransRecLen) + 1

   If NextTransRec& <= 0 Then
     NextTransRec& = 1
   End If

   Put thandle, NextTransRec&, UBTransRec(1) ',  UBTransRecLen

   'UBCustRec(1).DepositAmt = UBTransRec(1).TransAmt

   '04-14-98 Testing
   UBCustRec(1).DepositAmt = Round#(UBCustRec(1).DepositAmt + UBTransRec(1).Transamt)

   UBCustRec(1).LastTrans = NextTransRec&
   Put CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1) ',  UBCustRecLen
   'ShowPctComp cnt&, NumPayRecs&
 Next
    Close CHandle
    Close thandle
    Close PHandle

    KillFile PayDepoName$
  End If
  UBLog "POSTED:" + Str$(NumPayRecs&)
  '**********
  UBLog "POSTING: PAYMENTS"
  If FileSize&(PayBillName$) > 0 Then
    CHandle = FreeFile
    Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = UBCustRecLen
    thandle = FreeFile
    Open UBPath$ + "UBTRANS.DAT" For Random Shared As thandle Len = UBTransRecLen
    PHandle = FreeFile
    Open PayBillName$ For Random Shared As PHandle Len = UBPayRecLen

    NumPayRecs& = LOF(PHandle) \ UBPayRecLen
    FrmShowPctComp.Label1 = "Posting Payment Transactions"
    FrmShowPctComp.Show

    'ShowProcessingScrn "Posting Payment Transactions"
    For cnt& = 1 To NumPayRecs&
      FrmShowPctComp.ShowPctComp cnt&, NumPayRecs&
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Close
        Exit Sub
      End If

      LSet UBTransRec(1) = TUBTransRec(1)
      Get PHandle, cnt&, UBPaymentRec(1) ',  UBPayRecLen
      Get CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1) ',  UBCustRecLen
      UBTransRec(1).TransDate = UBPaymentRec(1).PAYDATE
      UBTransRec(1).TransType = TranBillPayment

      '052698 Added tax exempt flag to trans rec. For payment summary report
      UBTransRec(1).TaxExempt = UBPaymentRec(1).TaxExempt

      '022098 Added
      If Len(QPTrim$(UBPaymentRec(1).Desc)) = 0 Then
        UBTransRec(1).TransDesc = "BILLING-PAYMENT"
            Else
        UBTransRec(1).TransDesc = UBPaymentRec(1).Desc
        UBTransRec(1).BillMsg = "BILLING-PAYMENT"
      End If
      '^^This holds the Payment Description
      'UBTransRec(1)CustLocation = UBPaymentRec(1).CUSTACCT
      UBTransRec(1).OperatorNumber = OPERNUM
      UBTransRec(1).CustAcctNo = UBPaymentRec(1).CustAcct
      UBTransRec(1).CustStatus = UBCustRec(1).Status
      UBTransRec(1).Transamt = UBPaymentRec(1).AMTPAID
      UBTransRec(1).CheckAmount = UBPaymentRec(1).CHKAMT
      UBTransRec(1).CashAmount = UBPaymentRec(1).CASHAMT

      If UBTransRec(1).CheckAmount > 0 And UBTransRec(1).CashAmount > 0 Then
        UBTransRec(1).PayTypeCode = 3
      ElseIf UBTransRec(1).CashAmount > 0 Then
        UBTransRec(1).PayTypeCode = 1
      ElseIf UBTransRec(1).CheckAmount > 0 Then
        If QPTrim(UBPaymentRec(1).TENDERTY) = "Charge" Then
          UBTransRec(1).PayTypeCode = 4
        Else
          UBTransRec(1).PayTypeCode = 2
        End If
      End If

      'IF UBCustRec(1).PrevBalance > 0 THEN
      '050597 changed to zero if <> zero
      If UBCustRec(1).PrevBalance <> 0 Then
        If UBTransRec(1).Transamt >= UBCustRec(1).PrevBalance Then
          UBCustRec(1).PrevBalance = 0
        ElseIf UBTransRec(1).Transamt < UBCustRec(1).PrevBalance Then
          UBCustRec(1).PrevBalance = Round#(UBCustRec(1).PrevBalance - UBTransRec(1).Transamt)
        End If
      End If

      For RevAmts = 1 To MaxRevsCnt
        UBTransRec(1).RevAmt(RevAmts) = UBPaymentRec(1).PaidOwed(RevAmts).AMTPD1
        UBCustRec(1).CurrRevAmts(RevAmts) = Round#(UBCustRec(1).CurrRevAmts(RevAmts) - UBTransRec(1).RevAmt(RevAmts))
        'This is for previous bill distribution
        'UBCustRec(1).PrevRevAmts(RevAmts) = Round#(UBCustRec(1).PrevRevAmts(RevAmts) - UBTransRec(1).RevAmt(RevAmts))
        'IF UBCustRec(1).PrevRevAmts(RevAmts) < 0 THEN
        '  UBCustRec(1).PrevRevAmts(RevAmts) = 0
        'END IF
      Next
      TotalCustBalance# = 0
      For RevAmts = 1 To MaxRevsCnt
        TotalCustBalance# = Round#(TotalCustBalance# + UBCustRec(1).CurrRevAmts(RevAmts))
      Next
      UBCustRec(1).CurrBalance = Round#(TotalCustBalance# - UBCustRec(1).PrevBalance)
      '02-26-97 Was not adding prev bal
      UBTransRec(1).RunBalance = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
      UBTransRec(1).PrevTrans = UBCustRec(1).LastTrans
      UBTransRec(1).VoidFlag = False
      UBTransRec(1).FromCMFlag = False
      NextTransRec& = (LOF(thandle) \ UBTransRecLen) + 1
      If NextTransRec& <= 0 Then
        NextTransRec& = 1
      End If
      Put thandle, NextTransRec&, UBTransRec(1) ',  UBTransRecLen
      UBCustRec(1).LastTrans = NextTransRec&
      If Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance) = 0 Then
        If UBCustRec(1).Status = "B" Then
          UBCustRec(1).Status = "I"
          CustChCnt = CustChCnt + 1
          UBLog "POSTING: SET CUST STATUS to I. Acct:" + Str$(UBPaymentRec(1).CustAcct)
        End If
      End If
      Put CHandle, UBPaymentRec(1).CustAcct, UBCustRec(1) ',  UBCustRecLen
      'ShowPctComp cnt&, NumPayRecs&
    Next
    Close CHandle
    Close thandle
    Close PHandle
    KillFile PayBillName$
  End If
  UBLog "POSTED:" + Str$(NumPayRecs&)
  If CustChCnt > 0 Then
    UBLog "POSTING: CUST STATUS CHANGED:" + Str$(CustChCnt)
  End If
  UBLog "POSTING TRANSACTIONS FINISH"

'  BlockClear
'  DisplayUBScrn "UPDATEOK"
'  WaitForAction

  Erase UBTransRec, TUBTransRec, UBCustRec, UBPaymentRec

ExitPayPost:
  UBLog "OUT: UB POST PAYMENTS,  OPER:" + Str$(OPERNUM)
End Sub
'Public Function CheckPayDate(ValCheck As String)
'  Dim Month As Integer, Day As Integer, Year As Integer
'  Month = Val(Mid(ValCheck, 1, 2))
'  Day = Val(Mid(ValCheck, 4, 2))
'  Year = Val(Mid(ValCheck, 7, 4))
'  'Checks date if Blank then won't check for valid date
'  'and then checks each section, month, day and year
'  'if any section wrong then returns false value
'  If InStr(ValCheck, "_") <= 0 Then
'    If ((Month > 0) And (Month < 13)) Then
'      If Day > 0 And Day < 32 Then
'        If Year > 1979 And Year < 2099 Then
'          CheckValDate = True
'        End If
'      End If
'    End If
'  End If
'End Function

