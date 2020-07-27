VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUBPenaltyMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penalty Processing Menu"
   ClientHeight    =   8868
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   12216
   Icon            =   "frmUBPenaltyMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   12216
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExitPenaltyProcess 
      Caption         =   "E&xit Penalty Process Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3846
      TabIndex        =   4
      Top             =   5808
      Width           =   4524
   End
   Begin VB.CommandButton cmdEditPenalty 
      Caption         =   "&Edit Penalty Transactions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3846
      TabIndex        =   2
      Top             =   4308
      Width           =   4524
   End
   Begin VB.CommandButton cmdPostPenaltyTrans 
      Caption         =   "&Post Penalty Transactions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3846
      TabIndex        =   3
      Top             =   5064
      Width           =   4524
   End
   Begin VB.CommandButton cmdPenaltyReport 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Customer Penalty &Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3846
      TabIndex        =   1
      Top             =   3564
      Width           =   4524
   End
   Begin VB.CommandButton cmdCalcPenalties 
      BackColor       =   &H008F8265&
      Caption         =   "&Calculate Penalty Charges"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   3846
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   2808
      Width           =   4524
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
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
            TextSave        =   "11:12 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7154
            TextSave        =   "7/11/2003"
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
      Caption         =   "Penalty Processing Menu"
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
      TabIndex        =   5
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
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1788
      Top             =   624
      Width           =   8652
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
End
Attribute VB_Name = "frmUBPenaltyMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdCalcPenalties_Click()
  Load frmPenaltyCalculation
  DoEvents
  frmPenaltyCalculation.Show
  Unload frmUBPenaltyMenu
End Sub

Private Sub cmdEditPenalty_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
   If Not Exist("UBPENINF.DAT") Then
     frmMsgDialog.RetLabel = "-2"
     FntSize = frmMsgDialog.Label(2).FontSize
     frmMsgDialog.Label(2).FontSize = (FntSize + 2)
     MsgText(0) = "ERROR:"
     MsgText(1) = ""
     MsgText(2) = ""
     MsgText(3) = "NO UNPOSTED PENALTY TRANSACTIONS!"
     MsgText(4) = ""
     MsgText(5) = ""
     GetOKorNot MsgText(), True
  Else
    Load frmPenaltyEdit
    DoEvents
    frmPenaltyEdit.Show
    Unload frmUBPenaltyMenu
  End If
End Sub

Private Sub cmdExitPenaltyProcess_Click()
  Load frmUBBillingMenu
  DoEvents
  frmUBBillingMenu.Show
  Unload frmUBPenaltyMenu
End Sub


Private Sub cmdPenaltyReport_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
   If Not Exist("UBPENINF.DAT") Then
     frmMsgDialog.RetLabel = "-2"
     FntSize = frmMsgDialog.Label(2).FontSize
     
     frmMsgDialog.Label(2).FontSize = (FntSize + 2)
     MsgText(0) = "ERROR:"
     MsgText(1) = ""
     MsgText(2) = ""
     MsgText(3) = "NO UNPOSTED PENALTY TRANSACTIONS!"
     MsgText(4) = ""
     MsgText(5) = ""
     GetOKorNot MsgText(), True
  Else
    frmReportOpt.Show 1
    DeActivateControls Me
    If rptopt > 0 Then
     PenaltyReport rptopt
    End If
    ActivateControls Me
  End If
End Sub

Private Sub cmdPostPenaltyTrans_Click()
  Dim FntSize As Integer
  ReDim MsgText(0 To 5) As String
  If Not Exist("UBPENINF.DAT") Then
     frmMsgDialog.RetLabel = "-2"
     FntSize = frmMsgDialog.Label(2).FontSize
     frmMsgDialog.Label(2).FontSize = (FntSize + 2)
     frmMsgDialog.Label(3).FontSize = (FntSize + 2)
     MsgText(0) = "ERROR:"
     MsgText(1) = ""
     MsgText(2) = ""
     MsgText(3) = "NO UNPOSTED PENALTY TRANSACTIONS!"
     MsgText(4) = "NOTHING TO POST."
     MsgText(5) = ""
     GetOKorNot MsgText(), True
     GoTo Exitthis
  End If
 
  DoItFlag = False
    frmNoOperatorsWarning.Label(5).Caption = "Post Penalty Transactions"
    Load frmNoOperatorsWarning
    frmNoOperatorsWarning.Show vbModal
    If Not DoItFlag Then
      GoTo Exitthis
    End If
  DeActivateControls Me
  PostPenalties
  ActivateControls Me
  MsgBox "Posting Penalties Completed.", vbOKOnly, "Procedure Complete"
Exitthis:
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  StatusBar1.Panels.Item(1).Text = TownName$
  'screenW = (Screen.Width / Screen.TwipsPerPixelX)
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyHome
      'cmdAddCustomer.SetFocus
    Case vbKeyEnd
      'cmdExitCustomerMenu.SetFocus
    Case Else:
  End Select
End Sub

Private Sub PenaltyReport(rptopt As Integer)
  Dim FntSize As Integer, fmt As String, UBCustRecLen As Integer
  Dim PenFile As String, hand2 As Integer, UBTranRecLen As Integer
  Dim UBSetupLen As Integer, cnt As Integer, PHandle As Integer
  Dim CHandle As Integer, UBRpt As Integer, NumPenRec As Long
  Dim lcnt As Long, PCnt As Long, DCnt As Long, PenTotal As Double
  Dim ToPrint As String, Rptname As String, PenDate As String
  Dim BalType As String, RevSource As String, MinBal As String
  Dim PenPct As String, FlatAmt As String, which As String
  Dim BLable As String, Descrp As String, B1 As String, B2 As String
  Dim Source(1 To 15) As String
  ReDim MsgText(0 To 5) As String
  PenFile$ = UBPath$ + "UBPENTRN.DAT"
  FF$ = Chr$(12)
  'ReDim Source$(15)
  fmt$ = String$(80, "-")
  MaxLines = 55
  Rptname$ = UBPath$ + "UBPENTRN.RPT"
  ReDim PenaltyInfo(1) As PenaltyInfoType

  If Not Exist("UBPENINF.DAT") Then
     frmMsgDialog.RetLabel = "-2"
     FntSize = frmMsgDialog.Label(2).FontSize
     frmMsgDialog.Label(2).FontSize = (FntSize + 2)
     MsgText(0) = "ERROR:"
     MsgText(1) = ""
     MsgText(2) = ""
     MsgText(3) = "NO UNPOSTED PENALTY TRANSACTIONS!"
     MsgText(4) = "CAN NOT PRINT PENALTY REPORT"
     MsgText(5) = ""
     GetOKorNot MsgText(), True
     GoTo ExitPenReport
  End If
  FrmShowPctComp.Label1 = "Creating Penalty Report"
  FrmShowPctComp.Show , Me

  hand2 = FreeFile
  Open UBPath$ + "UBPENINF.DAT" For Random As hand2
  Get hand2, 1, PenaltyInfo(1)
  Close hand2
  ReDim UBCustRec(1) As NewUBCustRecType
  ReDim UBSetUpRec(1) As UBSetupRecType
  ReDim UBTranRec(1) As UBTransRecType

  UBCustRecLen = Len(UBCustRec(1))
  UBTranRecLen = Len(UBTranRec(1))

  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  TownName$ = QPTrim$(UBSetUpRec(1).UTILNAME)

  For cnt = 1 To 15
    Source$(cnt) = UBSetUpRec(1).Revenues(cnt).REVNAME
  Next
  PHandle = FreeFile
  Open PenFile$ For Random Shared As PHandle Len = UBTranRecLen

  CHandle = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As CHandle Len = UBCustRecLen

  UBRpt = FreeFile
  Open Rptname$ For Output As UBRpt

  NumPenRec& = LOF(PHandle) / UBTranRecLen
  If rptopt = 2 Then 'do text
    GoSub PenRptHeader
    For lcnt& = 1 To NumPenRec&
      FrmShowPctComp.ShowPctComp lcnt&, NumPenRec&
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        GoTo ExitPenReport
      End If

      Get PHandle, lcnt&, UBTranRec(1)
      'IF UBTranRec(1).ActiveFlag <> 0 THEN
      PCnt& = PCnt& + 1
      Get CHandle, UBTranRec(1).CustAcctNo, UBCustRec(1)
      Print #UBRpt, Using$("######", UBTranRec(1).CustAcctNo);
      Print #UBRpt, Tab(10); UBCustRec(1).Status; Tab(15); UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB;
      Print #UBRpt, Tab(30); UBCustRec(1).CustName;
      If UBTranRec(1).TransAmt = 0 Then
        Print #UBRpt, Tab(65); " DELETED"
        DCnt& = DCnt& + 1
      Else
        Print #UBRpt, Tab(65); Using$("#####.##", UBTranRec(1).TransAmt)
      End If
      PenTotal# = Round#(PenTotal# + UBTranRec(1).TransAmt)
      Linecnt = Linecnt + 1
      'END IF
      If Linecnt > MaxLines Then
        Print #UBRpt, FF$
        GoSub PenRptHeader
      End If
  '    If AskAbandonPrint% Then
  '      AbortFlag = True
  '      Exit For
  '    End If
      'ShowPctComp cnt&, NumPenRec&
    Next
    GoSub PenRptFooter
    Print #UBRpt, FF$
    GoSub PenRptParms
    Print #UBRpt, FF$
  Else
    For lcnt& = 1 To NumPenRec&
      FrmShowPctComp.ShowPctComp lcnt&, NumPenRec&
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        GoTo ExitPenReport
      End If
      Get PHandle, lcnt&, UBTranRec(1)
      'IF UBTranRec(1).ActiveFlag <> 0 THEN
      PCnt& = PCnt& + 1
      Get CHandle, UBTranRec(1).CustAcctNo, UBCustRec(1)
      ToPrint$ = Using$("######", UBTranRec(1).CustAcctNo) + "~"
      ToPrint$ = ToPrint$ + UBCustRec(1).Status + "~" + UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
      ToPrint$ = ToPrint$ + "~" + UBCustRec(1).CustName
      If UBTranRec(1).TransAmt = 0 Then
        ToPrint$ = ToPrint$ + "~ DELETED"
        DCnt& = DCnt& + 1
      Else
        ToPrint$ = ToPrint$ + "~" + Using$("#####.##", UBTranRec(1).TransAmt)
      End If
      PenTotal# = Round#(PenTotal# + UBTranRec(1).TransAmt)
      Linecnt = Linecnt + 1
      Print #UBRpt, ToPrint$
      ToPrint$ = ""
    Next
    GoSub PenParms
  End If
  Close
  If rptopt = 2 Then
    ViewPrint Rptname$, "Penalty Transaction Report"
  Else
    Load frmLoadingRpt
    ARptPenalties.txtDate = Now
    ARptPenalties.txtTown = TownName$
    ARptPenalties.Title = "Penalty Transaction Report"
    ARptPenalties.totamt = Using("$######.##", PenTotal#)
    ARptPenalties.totDel = DCnt&
    ARptPenalties.totTrans = PCnt&
    ARptPenalties.PenDate = PenDate
    ARptPenalties.BalType = BalType
    ARptPenalties.RevSource = RevSource
    ARptPenalties.MinBal = MinBal
    ARptPenalties.Flat = FlatAmt
    ARptPenalties.Percent = PenPct
    ARptPenalties.which = which
    ARptPenalties.B1 = B1
    ARptPenalties.B2 = B2
    ARptPenalties.Desc = Descrp
    ARptPenalties.GetName Rptname$
    ARptPenalties.startrpt
  End If
  Exit Sub

PenRptHeader:
  PageNo = PageNo + 1
  Print #UBRpt, "Penalty Transaction Report.  "; Tab(70); "Page: "; PageNo
  Print #UBRpt, TownName$
  Print #UBRpt, " Acct    Location       Customer                             Amount"
  Print #UBRpt, fmt$
  Linecnt = 4
  Return

PenRptFooter:
  Print #UBRpt, fmt$
  Print #UBRpt, " Transactions:"; PCnt&; Tab(22); "Deleted:"; DCnt&; Tab(42); "Penalty Total: "; Using("$######.##", PenTotal#)
  Linecnt = 4
  Return
PenParms:    'for graphics
  PenDate$ = Num2Date$(PenaltyInfo(1).PenDate)
  BalType = PenaltyInfo(1).ChargeOn
  RevSource = Source$(PenaltyInfo(1).RevSource)
  MinBal = Using("####.##", PenaltyInfo(1).MinBalance)
  If Len(QPTrim$(PenaltyInfo(1).GreatLess)) > 0 Then
    PenPct = Using("####.##", PenaltyInfo(1).PctCharge)
    FlatAmt = Using("####.##", PenaltyInfo(1).AmtCharge)
    Select Case PenaltyInfo(1).GreatLess
    Case "L"
      which = "LESS"
    Case "G"
      which = "GREATER"
    End Select
  Else
    If PenaltyInfo(1).PctCharge > 0 Then
      PenPct = Using("####.##", PenaltyInfo(1).PctCharge)
    ElseIf PenaltyInfo(1).AmtCharge > 0 Then
      FlatAmt = Using("####.##", PenaltyInfo(1).AmtCharge)
    End If
  End If
  If PenaltyInfo(1).CycLast > 0 Then
    B1 = "Cycle " + Using("##", PenaltyInfo(1).CycFirst)
    B2 = "Cycle " + Using("##", PenaltyInfo(1).CycLast)
  ElseIf PenaltyInfo(1).BookLast > 0 Then
    B1 = "Book " + Using("##", PenaltyInfo(1).BookFirst)
    B2 = "Book " + Using("##", PenaltyInfo(1).BookLast)
  End If
  Descrp = PenaltyInfo(1).PenDesc
  Return

PenRptParms:
  PageNo = PageNo + 1
  Print #UBRpt, "Penalty Calculation Parameters.  "; Tab(70); "Page: "; PageNo
  Print #UBRpt, TownName$
  Print #UBRpt, fmt$
  Print #UBRpt, "    Penalty Date: "; Num2Date$(PenaltyInfo(1).PenDate)
  Print #UBRpt, "    Balance Type: "; PenaltyInfo(1).ChargeOn
  Print #UBRpt, "  Revenue Source: "; Source$(PenaltyInfo(1).RevSource)
  Print #UBRpt, " Minimum Balance: "; Using("####.##", PenaltyInfo(1).MinBalance)
  If Len(QPTrim$(PenaltyInfo(1).GreatLess)) > 0 Then
    Print #UBRpt, " Penalty Percent: "; Using("######", PenaltyInfo(1).PctCharge)
    Print #UBRpt, "     Flat Amount: "; Using("####.##", PenaltyInfo(1).AmtCharge)
    Print #UBRpt, "    Whichever is: ";
    Select Case PenaltyInfo(1).GreatLess
    Case "L"
      Print #UBRpt, "LESS"
    Case "G"
      Print #UBRpt, "GREATER"
    End Select
  Else
    If PenaltyInfo(1).PctCharge > 0 Then
      Print #UBRpt, " Penalty Percent: "; Using("######", PenaltyInfo(1).PctCharge)
    ElseIf PenaltyInfo(1).AmtCharge > 0 Then
      Print #UBRpt, "     Flat Amount: "; Using("####.##", PenaltyInfo(1).AmtCharge)
    End If
  End If
  If PenaltyInfo(1).CycLast > 0 Then
    Print #UBRpt, "      From Cycle: "; Using("######", PenaltyInfo(1).CycFirst)
    Print #UBRpt, "      Thru Cycle: "; Using("######", PenaltyInfo(1).CycLast)
  ElseIf PenaltyInfo(1).BookLast > 0 Then
    Print #UBRpt, "       From Book: "; Using("######", PenaltyInfo(1).BookFirst)
    Print #UBRpt, "       Thru Book: "; Using("######", PenaltyInfo(1).BookLast)
  End If
  Print #UBRpt, "     Description: "; PenaltyInfo(1).PenDesc
  'PenDesc    AS STRING * 21
  'CycFirst   AS INTEGER
  'CycLast    AS INTEGER
  'BookFirst  AS INTEGER
  'BookLast   AS INTEGER
  'PenCnt     AS INTEGER

  Return

ExitPenReport:
End Sub
Private Sub PostPenalties()
  '01-15-99 Added penalty processing as seperate parts
  Dim UBCustRecLen As Integer, cnt As Integer
  Dim PenFile As String, hand2 As Integer, UBTranRecLen As Integer
  Dim UBSetupLen As Integer, PHandle As Integer, UBCust As Integer
  Dim CHandle As Integer, UBRpt As Integer, NumPenRec As Long
  Dim UBTran As Integer, NumOfTranRecs As Long, PenCnt As Long
  Dim PostedCnt As Long, RevCnt As Integer, PrevLastTrans As Long

  UBLog "IN: Post Penalty Transactions (PPT)"

  PenFile$ = UBPath$ + "UBPENTRN.DAT"

  ReDim Source$(15)
  ReDim UBSetUpRec(1) As UBSetupRecType
   LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  For cnt = 1 To 15
    Source$(cnt) = UBSetUpRec(1).Revenues(cnt).REVNAME
  Next
  FrmShowPctComp.Label1 = "Posting Penalty Transactions"
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show '1, Parent

  UBLog "START: Posting Penalty Transactions."

  ReDim PenaltyInfo(1) As PenaltyInfoType
  hand2 = FreeFile
  Open UBPath$ + "UBPENINF.DAT" For Random As hand2
  Get hand2, 1, PenaltyInfo(1)
  Close hand2

  ReDim UBTranRec(1) As UBTransRecType
  ReDim UBCustRec(1) As NewUBCustRecType

  UBTranRecLen = Len(UBTranRec(1))
  UBCustRecLen = Len(UBCustRec(1))

  UBCust = FreeFile
  Open UBPath$ + "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen

  PHandle = FreeFile
  Open PenFile$ For Random Shared As PHandle Len = UBTranRecLen

  UBTran = FreeFile
  Open UBPath$ + "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen

  NumOfTranRecs& = LOF(UBTran) \ UBTranRecLen
  NumPenRec& = LOF(PHandle) \ UBTranRecLen
  

  
  For PenCnt& = 1 To NumPenRec&
    FrmShowPctComp.ShowPctComp PenCnt&, NumPenRec&

    Get PHandle, PenCnt&, UBTranRec(1)
    If (UBTranRec(1).ActiveFlag And UBTranRec(1).TransAmt > 0) Then
      PostedCnt& = PostedCnt& + 1
      NumOfTranRecs& = NumOfTranRecs& + 1       'point to next trans to write
      Get UBCust, UBTranRec(1).CustAcctNo, UBCustRec(1)
      If UBCustRec(1).Status = "B" Then
        UBCustRec(1).LATEFEE = "N"
      End If
            '020399 Changed this back to the way it used to be
      UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + UBTranRec(1).TransAmt)
      'UBCustRec(1).PrevBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      UBTranRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      For RevCnt = 1 To MaxRevsCnt
        '        UBCustRec(1).PrevRevAmts(RevCnt) = Round#(UBCustRec(1).CurrRevAmts(RevCnt) + UBCustRec(1).PrevRevAmts(RevCnt))
        UBCustRec(1).CurrRevAmts(RevCnt) = Round#(UBCustRec(1).CurrRevAmts(RevCnt) + UBTranRec(1).RevAmt(RevCnt))
      Next
      UBTranRec(1).TransType = TranPenaltyCharge
      UBTranRec(1).TransDesc = PenaltyInfo(1).PenDesc
      UBTranRec(1).TransDate = PenaltyInfo(1).PenDate
      PrevLastTrans& = UBCustRec(1).LastTrans
      UBTranRec(1).PrevTrans = PrevLastTrans&
      UBCustRec(1).LastTrans = NumOfTranRecs&
      Put UBCust, UBTranRec(1).CustAcctNo, UBCustRec(1)
      Put UBTran, NumOfTranRecs&, UBTranRec(1)
    End If
    
  Next
  Close
  UBLog "  DONE: Posting Penalty Transactions."
  UBLog "POSTED:" + Str$(PostedCnt&) + " New Penalty Transactions."
  UBLog " Parameters:"
  UBRpt = FreeFile
  Open "UBLOG.DAT" For Append Shared As UBRpt
  Print #UBRpt, "   Penalty Date: "; Num2Date$(PenaltyInfo(1).PenDate)
  Print #UBRpt, "   Balance Type: "; PenaltyInfo(1).ChargeOn
  Print #UBRpt, " Revenue Source: "; Source$(PenaltyInfo(1).RevSource)
  Print #UBRpt, "Minimum Balance: "; Using("####.##", PenaltyInfo(1).MinBalance)
  If PenaltyInfo(1).PctCharge > 0 Then
    Print #UBRpt, "Penalty Percent: "; Using("######", PenaltyInfo(1).PctCharge)
  ElseIf PenaltyInfo(1).AmtCharge > 0 Then
    Print #UBRpt, "    Flat Amount: "; Using("####.##", PenaltyInfo(1).AmtCharge)
  End If
  If PenaltyInfo(1).CycLast > 0 Then
    Print #UBRpt, "     From Cycle: "; Using("######", PenaltyInfo(1).CycFirst)
    Print #UBRpt, "     Thru Cycle: "; Using("######", PenaltyInfo(1).CycLast)
  ElseIf PenaltyInfo(1).BookLast > 0 Then
    Print #UBRpt, "      From Book: "; Using("######", PenaltyInfo(1).BookFirst)
    Print #UBRpt, "      Thru Book: "; Using("######", PenaltyInfo(1).BookLast)
  End If
  Close
  
  KillFile UBPath$ + "UBPENINF.DAT"
  KillFile PenFile$

  UBLog "KILLED: UBPENINF.DAT  & " + PenFile$

ExitPenPost:
  UBLog "OUT: Penalty Posting." + CrLf$

End Sub
