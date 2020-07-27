VERSION 5.00
Begin VB.Form frmAPLdgUtilMenu 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AP Ledger Utilities"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   12195
   Icon            =   "frmAPLdgUtilMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   870
      TabIndex        =   5
      Top             =   2664
      Width           =   924
   End
   Begin VB.CommandButton cmdExitAPLdgMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "E&xit AP Ledger Menu"
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
      Left            =   4290
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5796
      Width           =   3612
   End
   Begin VB.CommandButton cmdPurgeOldHist 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Purge &Old History"
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
      Left            =   4290
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4956
      Width           =   3612
   End
   Begin VB.CommandButton cmdPrintLedger 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Print Ledger"
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
      Left            =   4290
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4116
      Width           =   3612
   End
   Begin VB.CommandButton cmdRelinkAP 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "&Re-Link Transactions"
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
      Left            =   4290
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3276
      Width           =   3612
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      Height          =   132
      Left            =   8850
      Top             =   2076
      Width           =   972
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000009&
      Height          =   132
      Left            =   2370
      Top             =   2076
      Width           =   972
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   8970
      X2              =   8970
      Y1              =   2196
      Y2              =   8076
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   8970
      X2              =   9690
      Y1              =   8076
      Y2              =   8076
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   2490
      X2              =   2490
      Y1              =   2196
      Y2              =   8076
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   2490
      X2              =   3210
      Y1              =   8076
      Y2              =   8076
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AP LEDGER UTILITIES"
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
      Left            =   4266
      TabIndex        =   4
      Top             =   1236
      Width           =   3660
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   1092
      Left            =   1770
      Top             =   876
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   0
      Left            =   2490
      Top             =   2196
      Width           =   732
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   5892
      Index           =   1
      Left            =   8970
      Top             =   2196
      Width           =   732
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Left            =   2370
      Top             =   1956
      Width           =   972
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   252
      Left            =   8850
      Top             =   1956
      Width           =   972
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1770
      Top             =   756
      Width           =   8652
   End
End
Attribute VB_Name = "frmAPLdgUtilMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim GLSetup As GLSetupRecType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim ApLedger As APLedger81RecType
Dim apvendor As VendorRecType
Dim APDist As APDistRecType
Dim x As Control
Dim Orphan As Integer

Private Sub cmdExitAPLdgMenu_Click()
  frmGLUtilMenu.Show
  Unload frmAPLdgUtilMenu
End Sub

Private Sub cmdPrintLedger_Click()
  frmReportOpt.Show 1
  If rptopt = 1 Then
    PrintLedger
  ElseIf rptopt = 2 Then
    PrintLedger2
  End If
End Sub

Private Sub cmdPurgeOldHist_Click()
  frmPurgeAPLedger.Show
  Unload frmAPLdgUtilMenu

End Sub

Private Sub cmdRelinkAP_Click()
  LinkLedg
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape:
      SendKeys "%X"
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExitAPLdgMenu.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        ClearInUse PWcnt
      End If
    End If
  End If
End Sub

Private Sub PrintLedger()
  Dim APLedgerFile As Integer, NumTran As Long, APLRecLen As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, APDRecLen As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, CommaFmtT As String
  Dim RptFile As Integer, RptFileName As String, CommaFmt As String
  Dim RunTotal As Double, cnt As Long, VendorName As String, LCnt As Long
  Dim ToPrint As String, NextDist As Long, DistAmt As Double, TDistAmt As Double
  Dim ThisRec As Long, BalMsg As String, Status As String, DistAA As Double
  Dim PDChkDate As String, PDChkNum As String
  Dim ToPrintG As String, ToPrintD As String
 ' On Local Error GoTo rpterror
  OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen
  OpenVendorFile VendorFile, NumVRecs
10:
  RptFile = FreeFile
  RptFileName$ = "apledger.prn"
  Open RptFileName$ For Output As RptFile
  CommaFmt$ = "#,###,###.##"
  CommaFmtT$ = "##,###,###,###.##"
  RunTotal# = 0
  FrmShowPctComp.Label1 = "Printing AP Ledger Report"
  FrmShowPctComp.Show , Me
  DeActivateControls frmAPLdgUtilMenu
  DoEvents
  For cnt& = 1 To NumTran&
 'Print Using; "Processing Ledger Record: #####"; cnt&;
30:
    Get APLedgerFile, cnt&, ApLedger
    If ApLedger.VRecNum > 0 Then
      Get VendorFile, ApLedger.VRecNum, apvendor
      VendorName$ = apvendor.VNAME
    Else
      VendorName$ = "Orphaned Transaction"
    End If
40:
    LCnt& = LCnt& + 1
    RunTotal# = Round#(RunTotal# + ApLedger.Amt)

    ToPrint$ = ""
    ToPrint$ = "Trans: " + Str$(cnt&)
    ToPrint$ = ToPrint$ + "~" + ApLedger.VendorCode
    ToPrint$ = ToPrint$ + "~" + VendorName$
    ToPrint$ = ToPrint$ + "~" + Format(DateAdd("d", (ApLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy")
    ToPrint$ = ToPrint$ + "~" + Trim(ApLedger.Bankcode)
    'Print #RptFile, ToPrint$
    
50:
    Select Case ApLedger.TRCode
       Case 1
          'MID$(ToPrint$, 5) = "Invoice " + APLedger.DOCNum
          GoSub PrintInv
       Case -1
          GoSub PrintInv
       Case 3
          'MID$(ToPrint$, 5) = "Check " + APLedger.DOCNum
          GoSub PrintChk
       Case -3
          GoSub PrintChk
       Case 4
          'MID$(ToPrint$, 5) = "Purchase Order  " + APLedger.DOCNum
          GoSub PrintPO
       Case -4
          GoSub PrintPO
       Case Else
          GoSub PrintOther
          'APLedger.TrCode = 4
          'PUT APLedgerFile, Cnt&, APLedger
    End Select
60:
    '--Now print the distribution
    NextDist& = ApLedger.FrstDist
    DistAmt# = 0
    If NextDist& > 0 Then  '--ignore checks, no distribution
       'Print #RptFile, Tab(40); "Accounting Distribution:"
       Do
68:
          Get APDistFile, NextDist&, APDist
          If APDist.DistAmt > -100000000000# Then
            DistAA = Round#(APDist.DistAmt)
            DistAmt# = Round#(DistAmt# + DistAA)
            TDistAmt# = Round#(TDistAmt# + DistAA)
          Else
            DistAA = 0
          End If
          ThisRec& = NextDist&
          NextDist& = APDist.NextDist
78:
          ToPrintD$ = ""
          ToPrintD$ = APDist.DistAcctNum
          ToPrintD$ = ToPrintD$ + "~" + Using(CommaFmt$, Str$(DistAA))
          ToPrintD$ = ToPrintD$ + "~" + Str$(APDist.APLedgerRec) + "/" + Str$(ThisRec&)
          Print #RptFile, ToPrint$ + "~" + ToPrintG$ + "~" + ToPrintD$

       Loop Until NextDist& = 0
80:
       If Round#(DistAmt#) <> Round#(ApLedger.Amt) Then
          BalMsg$ = "***** Error *****"
       Else
          BalMsg$ = ""
       End If
       ''Print #RptFile, Tab(30); "Total Distributed:"; Tab(49); Using(CommaFmtT$, Str$(DistAmt#)) + BalMsg$
    Else
      Print #RptFile, ToPrint$ + "~" + ToPrintG$ + "~ ~ ~ "
    End If
    'Print #RptFile, String$(78, "=")
    'Put count up doololly here
    FrmShowPctComp.ShowPctComp cnt&, NumTran&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmAPLdgUtilMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

  Next

  'Print #RptFile,
  'Print #RptFile, "Running Total: " + Using(CommaFmtT$, Str$(RunTotal#))
  'Print #RptFile,
 ' Print #RptFile, "Dist Amt:"; Using(CommaFmtT$, Str$(TDistAmt#))
  'PRINT #RptFile, "Tax Amt:"; Tax#
  'Print #RptFile, LCnt&
  Close
90:
 ' ViewPrint RptFileName$, "APLEDGER.PRN"
  ActivateControls frmAPLdgUtilMenu
'  SHELL "list APLEDGER.PRN"
  Load frmLoadingRpt
  ARptLedHist.totdist = Using(CommaFmtT$, Str$(TDistAmt#))
  ARptLedHist.TotRun = Using(CommaFmtT$, Str$(RunTotal#))
  ARptLedHist.totTrans = LCnt&
  ARptLedHist.txtDate = Now
  ARptLedHist.txtTown = GLUserName$
  ARptLedHist.GetName RptFileName$
  ARptLedHist.startrpt

Exit Sub
PrintInv:
  ToPrintG$ = ""
  ToPrintG$ = "Invoice " + QPTrim$(ApLedger.DOCNum) + " 1099-" + QPTrim$(ApLedger.Get1099)
  ToPrintG$ = ToPrintG$ + "~" + "Total Amt: " + Using(CommaFmt$, Str$(ApLedger.Amt))
  'Print #RptFile, ToPrint$

  
  ToPrintG$ = ToPrintG$ + "~" + "Tr Date: " + Format(DateAdd("d", (ApLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy")
  ToPrintG$ = ToPrintG$ + "~" + "Due Date: " + Format(DateAdd("d", (ApLedger.DueDate), "12-31-1979"), "mm/dd/yyyy")
  ToPrintG$ = ToPrintG$ + "~" + "G/L Date: " + Format(DateAdd("d", (ApLedger.GLDistDate), "12-31-1979"), "mm/dd/yyyy")
  'Print #RptFile, ToPrint$

  
  Select Case ApLedger.PAYCODE
    Case 1
       Status$ = "Status: Open"
       PDChkDate$ = " "
       PDChkNum$ = " "
    Case 3
       Status$ = "Status: Paid"
       PDChkDate$ = "Check Date: " + Format(DateAdd("d", (ApLedger.PDCheckDate), "12-31-1979"), "mm/dd/yyyy")
       PDChkNum$ = "Check Num: " + Str$(ApLedger.PDCheckNum)
    Case Else

       Status$ = "Status: Invalid Pay Code"
       PDChkDate$ = Format(DateAdd("d", (ApLedger.PDCheckDate), "12-31-1979"), "mm/dd/yyyy")
       PDChkNum$ = Str$(ApLedger.PDCheckNum)
  End Select
  
  Select Case ApLedger.TRCode
    Case -1
      Status$ = "Status: VOIDED"
      'PDChkDate$ = "ON: " + Format(DateAdd("d", (ApLedger.), "12-31-1979"), "mm/dd/yyyy")
    Case Else
  End Select
  ToPrintG$ = ToPrintG$ + "~" + Status$
  ToPrintG$ = ToPrintG$ + "~" + PDChkDate$  'Num2Date(APLedger.PdCheckDate)
  ToPrintG$ = ToPrintG$ + "~" + PDChkNum$
  'Print #RptFile, ToPrint$

Return


PrintChk:
  ToPrintG$ = ""
  ToPrintG$ = "Check " + ApLedger.DOCNum
  ToPrintG$ = ToPrintG$ + "~" + "Check Amt: " + Using(CommaFmt$, Str$(ApLedger.Amt))
  ToPrintG$ = ToPrintG$ + "~" + "Dated: " + Format(DateAdd("d", (ApLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy")
  ToPrintG$ = ToPrintG$ + "~~"
  
  Select Case ApLedger.TRCode
    Case 3
       Status$ = "Status: Paid"
       PDChkDate$ = Format(DateAdd("d", (ApLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy")
    Case -3
       Status$ = "Status: VOIDED"
       PDChkDate$ = Format(DateAdd("d", (ApLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy")
    Case Else
 End Select
 
 ToPrintG$ = ToPrintG$ + "~" + Status$
 ToPrintG$ = ToPrintG$ + "~" + PDChkDate$ + "~"
  

  'LSET ToPrint$ = ""
  'MID$(ToPrint$, 29) = "Due Date: " + Num2Date$(APLedger.DueDate)
  'MID$(ToPrint$, 55) = "G/L Date: " + Num2Date$(APLedger.GLDistDate)
  'PRINT #RptFile, ToPrint$

  'LSET ToPrint$ = ""
  'SELECT CASE APLedger.PayCode
  '  CASE 1
  '     Status$ = "Status: Open"
  '     PdChkDate$ = " "
  '     PdChkNum$ = " "
  '  CASE 3
  '     Status$ = "Status: Paid"
  '     PdChkDate$ = "Check Date: " + Num2Date(APLedger.PdCheckDate)
  '     PdChkNum$ = "Check Num: " + STR$(APLedger.PdCheckNum)
  '  CASE ELSE
  '     Status$ = "Status: Invalid Pay Code"
  '     PdChkDate$ = Num2Date(APLedger.PdCheckDate)
  '     PdChkNum$ = STR$(APLedger.PdCheckNum)
  'END SELECT
  'MID$(ToPrint$, 5) = Status$
  'MID$(ToPrint$, 27) = PdChkDate$ 'Num2Date(APLedger.PdCheckDate)
  'MID$(ToPrint$, 54) = PdChkNum$
  'PRINT #RptFile, ToPrint$

Return


PrintPO:
  ToPrintG$ = ""
  ToPrintG$ = "Purchase Order  " + ApLedger.DOCNum
  ToPrintG$ = ToPrintG$ + "~" + "Total Amt: " + Using(CommaFmt$, Str$(ApLedger.Amt))
  ToPrintG$ = ToPrintG$ + "~" + "PO Date: " + Format(DateAdd("d", (ApLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy")
  ToPrintG$ = ToPrintG$ + "~~"

  'LSET ToPrint$ = ""
  'MID$(ToPrint$, 5) = "Tr Date: " + Num2Date$(APLedger.TrDate)
  'MID$(ToPrint$, 29) = "Due Date: " + Num2Date$(APLedger.DueDate)
  'MID$(ToPrint$, 55) = "G/L Date: " + Num2Date$(APLedger.GLDistDate)
  'PRINT #RptFile, ToPrint$
  '
  'LSET ToPrint$ = ""
  Select Case ApLedger.PAYCODE
    Case 4
       Status$ = "Status: Open"
       PDChkDate$ = " "
       PDChkNum$ = " "
    Case -4
       Status$ = "Status: Paid"
       PDChkDate$ = "Check Date: " + Format(DateAdd("d", (ApLedger.PDCheckDate), "12-31-1979"), "mm/dd/yyyy")
       PDChkNum$ = "Check Num: " + Str$(ApLedger.PDCheckNum)
    Case Else
       Status$ = "Status: Invalid Pay Code"
       PDChkDate$ = Format(DateAdd("d", (ApLedger.PDCheckDate), "12-31-1979"), "mm/dd/yyyy")
       PDChkNum$ = Str$(ApLedger.PDCheckNum)
  End Select
  ToPrintG$ = ToPrintG$ + "~" + Status$
  ToPrintG$ = ToPrintG$ + "~" + PDChkDate$  'Num2Date(APLedger.PdCheckDate)
  ToPrintG$ = ToPrintG$ + "~" + PDChkNum$
  'Print #RptFile, ToPrint$

Return


PrintOther:
 ToPrintG$ = ""
 ToPrintG$ = "**Unknown Tr Code**" + Str$(ApLedger.TRCode) + ApLedger.DOCNum
 ToPrintG$ = ToPrintG$ + "~~~~~~~"
 'Print #RptFile, ToPrint$
Return

CancelExit:
  Exit Sub
  
rpterror:
  If Err > 0 Then
    Unload frmLoadingRpt
    MsgBox "Error Code Was " + Err.Description + Str$(Err) + " (prnledg - Line:" & Erl & ")"
  End If
  Close
  Exit Sub

End Sub
Private Sub PrintLedger2()
  Dim APLedgerFile As Integer, NumTran As Long, APLRecLen As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, APDRecLen As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, CommaFmtT As String
  Dim RptFile As Integer, RptFileName As String, CommaFmt As String
  Dim RunTotal As Double, cnt As Long, VendorName As String, LCnt As Long
  Dim ToPrint As String, NextDist As Long, DistAmt As Double, TDistAmt As Double
  Dim ThisRec As Long, BalMsg As String, Status As String, DistAA As Double
  Dim PDChkDate As String, PDChkNum As String
  OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen
  OpenVendorFile VendorFile, NumVRecs

  RptFile = FreeFile
  RptFileName$ = "apledger.prn"
  Open RptFileName$ For Output As RptFile
  CommaFmt$ = "#,###,###.##"
  CommaFmtT$ = "##,###,###,###.##"
  RunTotal# = 0
  FrmShowPctComp.Label1 = "Printing AP Ledger Report"
  FrmShowPctComp.Show , Me
  DeActivateControls frmAPLdgUtilMenu
  DoEvents
  For cnt& = 1 To NumTran&
 'Print Using; "Processing Ledger Record: #####"; cnt&;

    Get APLedgerFile, cnt&, ApLedger
    If ApLedger.VRecNum > 0 Then
      Get VendorFile, ApLedger.VRecNum, apvendor
      VendorName$ = apvendor.VNAME
    Else
      VendorName$ = "Orphaned Transaction"
    End If

    LCnt& = LCnt& + 1
   ' If ApLedger.Amt < 0 Then Stop 'ApLedger.Amt = 0 ' Stop
    'ApLedger.Amt = 0
    RunTotal# = Round#(RunTotal# + ApLedger.Amt)

    ToPrint$ = Space$(80)
    Mid$(ToPrint$, 2) = "Trans: " + Str$(cnt&)
    Mid$(ToPrint$, 15) = ApLedger.VendorCode
    Mid$(ToPrint$, 27) = VendorName$
    Mid$(ToPrint$, 60) = Format(DateAdd("d", (ApLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy")
    Mid$(ToPrint$, 76) = Trim(ApLedger.Bankcode)
    Print #RptFile, ToPrint$
    LSet ToPrint$ = ""

    Select Case ApLedger.TRCode
       Case 1
          'MID$(ToPrint$, 5) = "Invoice " + APLedger.DOCNum
          GoSub PrintInv
       Case -1
          GoSub PrintInv
       Case 3
          'MID$(ToPrint$, 5) = "Check " + APLedger.DOCNum
          GoSub PrintChk
       Case -3
          GoSub PrintChk
       Case 4
          'MID$(ToPrint$, 5) = "Purchase Order  " + APLedger.DOCNum
          GoSub PrintPO
       Case -4
          GoSub PrintPO
       Case Else
          GoSub PrintOther
          'APLedger.TrCode = 4
          'PUT APLedgerFile, Cnt&, APLedger
    End Select

    '--Now print the distribution
    NextDist& = ApLedger.FrstDist
    DistAmt# = 0
    If NextDist& > 0 Then  '--ignore checks, no distribution
       Print #RptFile, Tab(40); "Accounting Distribution:"
       Do
          Get APDistFile, NextDist&, APDist
          If APDist.DistAmt > -100000000000# Then
         ''' Stop
            DistAA = Round#(APDist.DistAmt)
            DistAmt# = Round#(DistAmt# + DistAA)
            TDistAmt# = Round#(TDistAmt# + DistAA)
          Else
            DistAA = 0
          End If
          ThisRec& = NextDist&
          NextDist& = APDist.NextDist

          ToPrint$ = Space$(80)
          Mid$(ToPrint$, 40) = APDist.DistAcctNum
          Mid$(ToPrint$, 54) = Using(CommaFmt$, Str$(DistAA))
          Mid$(ToPrint$, 67) = Str$(APDist.APLedgerRec) + "/" + Str$(ThisRec&)
          Print #RptFile, ToPrint$

       Loop Until NextDist& = 0
       Print #RptFile, Tab(54); "------------"
       If DistAmt# <> Round#(ApLedger.Amt) Then
          BalMsg$ = "***** Error *****"
       Else
          BalMsg$ = ""
       End If
       Print #RptFile, Tab(30); "Total Distributed:"; Tab(49); Using(CommaFmtT$, Str$(DistAmt#)) + BalMsg$

    End If
    Print #RptFile, String$(78, "=")
    'Put count up doololly here
    FrmShowPctComp.ShowPctComp cnt&, NumTran&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmAPLdgUtilMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

  Next

  Print #RptFile,
  Print #RptFile, "Running Total: " + Using(CommaFmtT$, Str$(RunTotal#))
  Print #RptFile,
  Print #RptFile, "Dist Amt:"; Using(CommaFmtT$, Str$(TDistAmt#))
  'PRINT #RptFile, "Tax Amt:"; Tax#
  Print #RptFile, LCnt&
  Close
  
  ViewPrint RptFileName$, "APLEDGER.PRN"
  ActivateControls frmAPLdgUtilMenu
'  SHELL "list APLEDGER.PRN"

Exit Sub
PrintInv:
  Mid$(ToPrint$, 5) = "Invoice " + QPTrim$(ApLedger.DOCNum) + "  1099-" + ApLedger.Get1099
  Mid$(ToPrint$, 54) = "Total Amt: " + Using(CommaFmt$, Str$(ApLedger.Amt))
  Print #RptFile, ToPrint$

  LSet ToPrint$ = ""
  Mid$(ToPrint$, 5) = "Tr Date: " + Format(DateAdd("d", (ApLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy")
  Mid$(ToPrint$, 29) = "Due Date: " + Format(DateAdd("d", (ApLedger.DueDate), "12-31-1979"), "mm/dd/yyyy")
  Mid$(ToPrint$, 55) = "G/L Date: " + Format(DateAdd("d", (ApLedger.GLDistDate), "12-31-1979"), "mm/dd/yyyy")
  Print #RptFile, ToPrint$

  LSet ToPrint$ = ""
  Select Case ApLedger.PAYCODE
    Case 1
       Status$ = "Status: Open"
       PDChkDate$ = " "
       PDChkNum$ = " "
    Case 3
       Status$ = "Status: Paid"
       PDChkDate$ = "Check Date: " + Format(DateAdd("d", (ApLedger.PDCheckDate), "12-31-1979"), "mm/dd/yyyy")
       PDChkNum$ = "Check Num: " + Str$(ApLedger.PDCheckNum)
    Case Else

       Status$ = "Status: Invalid Pay Code"
       PDChkDate$ = Format(DateAdd("d", (ApLedger.PDCheckDate), "12-31-1979"), "mm/dd/yyyy")
       PDChkNum$ = Str$(ApLedger.PDCheckNum)
  End Select
  LSet ToPrint$ = ""
  Select Case ApLedger.TRCode
    Case -1
      Status$ = "Status: VOIDED"
      'PDChkDate$ = "ON: " + Format(DateAdd("d", (ApLedger.), "12-31-1979"), "mm/dd/yyyy")
    Case Else
  End Select
  Mid$(ToPrint$, 5) = Status$
  Mid$(ToPrint$, 27) = PDChkDate$ 'Num2Date(APLedger.PdCheckDate)
  Mid$(ToPrint$, 54) = PDChkNum$
  Print #RptFile, ToPrint$

Return


PrintChk:
  LSet ToPrint$ = ""
  Mid$(ToPrint$, 5) = "Check " + ApLedger.DOCNum
  Mid$(ToPrint$, 29) = "Dated: " + Format(DateAdd("d", (ApLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy")
  Mid$(ToPrint$, 54) = "Check Amt: " + Using(CommaFmt$, Str$(ApLedger.Amt))
   Print #RptFile, ToPrint$
  
  Select Case ApLedger.TRCode
    Case 3
       Status$ = "Status: Paid"
       PDChkDate$ = Format(DateAdd("d", (ApLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy")
    Case -3
       Status$ = "Status: VOIDED"
       PDChkDate$ = Format(DateAdd("d", (ApLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy")
    Case Else
 End Select
 LSet ToPrint$ = ""
 Mid$(ToPrint$, 5) = Status$
 Mid$(ToPrint$, 27) = PDChkDate$
  Print #RptFile, ToPrint$

  'LSET ToPrint$ = ""
  'MID$(ToPrint$, 29) = "Due Date: " + Num2Date$(APLedger.DueDate)
  'MID$(ToPrint$, 55) = "G/L Date: " + Num2Date$(APLedger.GLDistDate)
  'PRINT #RptFile, ToPrint$

  'LSET ToPrint$ = ""
  'SELECT CASE APLedger.PayCode
  '  CASE 1
  '     Status$ = "Status: Open"
  '     PdChkDate$ = " "
  '     PdChkNum$ = " "
  '  CASE 3
  '     Status$ = "Status: Paid"
  '     PdChkDate$ = "Check Date: " + Num2Date(APLedger.PdCheckDate)
  '     PdChkNum$ = "Check Num: " + STR$(APLedger.PdCheckNum)
  '  CASE ELSE
  '     Status$ = "Status: Invalid Pay Code"
  '     PdChkDate$ = Num2Date(APLedger.PdCheckDate)
  '     PdChkNum$ = STR$(APLedger.PdCheckNum)
  'END SELECT
  'MID$(ToPrint$, 5) = Status$
  'MID$(ToPrint$, 27) = PdChkDate$ 'Num2Date(APLedger.PdCheckDate)
  'MID$(ToPrint$, 54) = PdChkNum$
  'PRINT #RptFile, ToPrint$

Return


PrintPO:
  LSet ToPrint$ = ""
  Mid$(ToPrint$, 5) = "Purchase Order  " + ApLedger.DOCNum
  Mid$(ToPrint$, 29) = "PO Date: " + Format(DateAdd("d", (ApLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy")
  Mid$(ToPrint$, 54) = "Total Amt: " + Using(CommaFmt$, Str$(ApLedger.Amt))
  Print #RptFile, ToPrint$
  LSet ToPrint$ = ""
  Mid$(ToPrint$, 5) = "Dept  " + Str(ApLedger.DeptNumb)
  Print #RptFile, ToPrint$

  'LSET ToPrint$ = ""
  'MID$(ToPrint$, 5) = "Tr Date: " + Num2Date$(APLedger.TrDate)
  'MID$(ToPrint$, 29) = "Due Date: " + Num2Date$(APLedger.DueDate)
  'MID$(ToPrint$, 55) = "G/L Date: " + Num2Date$(APLedger.GLDistDate)
  'PRINT #RptFile, ToPrint$
  '
  'LSET ToPrint$ = ""
  Select Case ApLedger.PAYCODE
    Case 4
       Status$ = "Status: Open"
       PDChkDate$ = " "
       PDChkNum$ = " "
    Case -4
       Status$ = "Status: Paid"
       PDChkDate$ = "Check Date: " + Format(DateAdd("d", (ApLedger.PDCheckDate), "12-31-1979"), "mm/dd/yyyy")
       PDChkNum$ = "Check Num: " + Str$(ApLedger.PDCheckNum)
    Case Else
       Status$ = "Status: Invalid Pay Code"
       PDChkDate$ = Format(DateAdd("d", (ApLedger.PDCheckDate), "12-31-1979"), "mm/dd/yyyy")
       PDChkNum$ = Str$(ApLedger.PDCheckNum)
  End Select
  Mid$(ToPrint$, 5) = Status$
  Mid$(ToPrint$, 27) = PDChkDate$ 'Num2Date(APLedger.PdCheckDate)
  Mid$(ToPrint$, 54) = PDChkNum$
  Print #RptFile, ToPrint$

Return


PrintOther:
 LSet ToPrint$ = ""
 Mid$(ToPrint$, 5) = "**Unknown Tr Code**" + Str$(ApLedger.TRCode) + ApLedger.DOCNum
 Print #RptFile, ToPrint$
Return

CancelExit:
  Exit Sub
End Sub

Private Sub LinkLedg()
  Call MainLog("LinkLedger Started.")
  RelinkLedger2Vendor
  RelinkDist2Trans
  If Orphan > 0 Then
    MsgBox "Ledger/Vendor Link encountered orphans! View Log For Details.", vbOKOnly, "Errors Found"
    Call MainLog("Link orphans.")
  Else
    MsgBox "Linking Operation complete.", vbOKOnly, "Operation Complete"
    Call MainLog("LinkLedger Complete.")
  End If
End Sub
Public Function RelinkLedger2Vendor()
  Dim APLedgerFile As Integer, NumTran As Long, APLRecLen As Integer
  Dim VendorFile As Integer, NumVRecs As Integer
  Dim LogFile As Integer, LogFileName As String, ToPrint As String
  Dim cnt As Long, Prev As Long, Ophan As Integer, VRecNum As Integer
'Open Percent thingy here
  FrmShowPctComp.Label1 = "Linking A/P Databases."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DeActivateControls frmAPLdgUtilMenu
  DoEvents
  Orphan = 0
   OpenVendorFile VendorFile, NumVRecs
   OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen

   LogFile = FreeFile
   LogFileName$ = "APLINK.LOG"
   Open LogFileName$ For Output As LogFile
   ToPrint$ = Space$(80)
   ToPrint$ = "Linking operations began on " + Date$ + " " + Time$
   Print #LogFile, ToPrint$

   'CommaFmt$ = "#######,.##"

   '--Reset the vendor trans pointers to 0.
   'PRINT "Initializing Vendor File."
   For cnt& = 1 To NumVRecs
      Get VendorFile, cnt&, apvendor
         apvendor.FrstTran = 0
         apvendor.LastTran = 0
      Put VendorFile, cnt&, apvendor
   Next

   'Print "APLedger Records:"; NumTran&

   '--Relink Transactions to Vendor
   For cnt& = 1 To NumTran&
    FrmShowPctComp.ShowPctComp cnt&, NumTran&

      Get APLedgerFile, cnt&, ApLedger

      '--reset next transaction pointer to 0
      ApLedger.NextTrans = 0
      Put APLedgerFile, cnt&, ApLedger

      VRecNum = ApLedger.VRecNum   'GetVendorRec(apledger.VendorCode)
      'VRECNUM = GetVendorRec(APLedger.VendorCode)

      If VRecNum > 0 And VRecNum <= NumVRecs Then

         Get VendorFile, VRecNum, apvendor

         If apvendor.FrstTran > 0 Then
            '--Vendor has previous transactions..
            '--Remember the last transaction for this vendor
            Prev& = apvendor.LastTran

            '--In the vendor record...
            '--Set the Last Trans pointer to this record
            apvendor.LastTran = cnt&
            Put VendorFile, VRecNum, apvendor

            '--In the apledger record...
            '--Set the Last trans pointer in the prev trans
            '--to point to this record
            Get APLedgerFile, Prev&, ApLedger
            ApLedger.NextTrans = cnt&
            Put APLedgerFile, Prev&, ApLedger
         Else
            '--First Trans for this vendor
            '--set both pointers to this ledger record
            apvendor.FrstTran = cnt&
            apvendor.LastTran = cnt&
            Put VendorFile, VRecNum, apvendor
        End If
      Else
        Orphan = Orphan + 1
        GoSub LogAPLOrphan
      End If

     ' Print "Processed Record: "; cnt&

   Next

   ToPrint$ = Space$(80)
   ToPrint$ = "Linking operations completed on " + Date$ + " " + Time$
   Print #LogFile, ToPrint$
   If Ophan > 1 Then
     ToPrint$ = Space$(80)
     ToPrint$ = "Orphaned Transactions: " + Str$(Ophan)
   Else
     ToPrint$ = Space$(80)
     ToPrint$ = "No Orphaned Transactions. "
   End If
   Print #LogFile, ToPrint$
   Close
ActivateControls frmAPLdgUtilMenu

Exit Function

LogAPLOrphan:
  Ophan = Orphan + 1
  Print #LogFile, "Rec: " + Str$(cnt&) + " Orphan: " + Str$(ApLedger.VRecNum)
Return
CancelExit:
Exit Function

End Function


Public Sub RelinkDist2Trans()
  Dim APLedgerFile As Integer, NumTran As Long, APLRecLen As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, APDRecLen As Integer
  Dim cnt As Long, Prev As Long
  
  FrmShowPctComp.Label1 = "Linking Distributions to Ledger."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DeActivateControls frmAPLdgUtilMenu
  DoEvents

   OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen


   OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

   'PRINT "Initializing Ledger Database."
   For cnt& = 1 To NumTran&
      Get APLedgerFile, cnt&, ApLedger
      ''''If ApLedger.FrstDist < 0 Or ApLedger.LastDist < 0 Then Stop
      ApLedger.FrstDist = 0
      ApLedger.LastDist = 0
      Put APLedgerFile, cnt&, ApLedger
   Next

   For cnt& = 1 To NumDistRecs&
'Put % thing here
     FrmShowPctComp.ShowPctComp cnt&, NumDistRecs&
      '--Assume no one else will follow.
      Get APDistFile, cnt&, APDist
      APDist.NextDist = 0
      Put APDistFile, cnt&, APDist

'      IF APLdgrDist.APLedgerRec > NumTran& THEN
'        STOP
'      END IF

'      IF APLdgrDist.APLedgerRec > 3555 AND APLdgrDist.APLedgerRec < 26905 THE
'        STOP
'      END IF
      '--Get the parent record

'      If APDist.APLedgerRec > NumTran& Or APDist.APLedgerRec <= 0 Then 'Stop
'        APDist.APLedgerRec = -1
'        Put APDistFile, cnt&, APDist
'        GoTo SkipHere
'      End If
      Get APLedgerFile, APDist.APLedgerRec, ApLedger

      If ApLedger.FrstDist > 0 Then
         '--We're not the first one here, so let us not forget those who have
         '--come before us
         Prev& = ApLedger.LastDist

         '--This is now the new last distribution
         '--Update Last Dist pointer in apledger to this rec
         ApLedger.LastDist = cnt&
         Put APLedgerFile, APDist.APLedgerRec, ApLedger

         '--Get the former last distribution
         '--and tell it that this rec is the next one
         Get APDistFile, Prev&, APDist
         APDist.NextDist = cnt&
         Put APDistFile, Prev&, APDist

      Else
         '--Virgin territory. we're now first and last
         ApLedger.FrstDist = cnt&
         ApLedger.LastDist = cnt&
         Put APLedgerFile, APDist.APLedgerRec, ApLedger
      End If
SkipHere:
   Next

   Close

   'PRINT "Press any key to continue."
   'K$ = INPUT$(1)
ActivateControls frmAPLdgUtilMenu
End Sub


Private Sub Command1_Click()
'FixoneVendcode ''FixJohnsonDist2Trans
'SpecialVoidPOTrans
'Fixvoidpoamt 'Fixonetilda
'  FixAcctandVend4New
'  RelinkDist2Trans4SB
' setallvendstoact
'PrintLedgerTransNums
RelinkDist2Trans4SB
'setallvendstoact
'RelinkDist2Trans4SB
End Sub



Private Sub LinkLedg4HB()
  Call MainLog("LinkLedger Started.")
  RelinkLedger2Vendor4HarryBurg
  RelinkDist2Trans
  If Orphan > 0 Then
    MsgBox "Ledger/Vendor Link encountered orphans! View Log For Details.", vbOKOnly, "Errors Found"
    Call MainLog("Link orphans.")
  Else
    MsgBox "Linking Operation complete.", vbOKOnly, "Operation Complete"
    Call MainLog("LinkLedger Complete.")
  End If
End Sub

Public Function RelinkLedger2Vendor4HarryBurg()
  Dim APLedgerFile As Integer, NumTran As Long, APLRecLen As Integer
  Dim VendorFile As Integer, NumVRecs As Integer
  Dim LogFile As Integer, LogFileName As String, ToPrint As String
  Dim cnt As Long, Prev As Long, Ophan As Integer, VRecNum As Integer
'Open Percent thingy here
  FrmShowPctComp.Label1 = "Fixing Trans links to 0 for new start."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DeActivateControls frmAPLdgUtilMenu
  DoEvents
  Orphan = 0
   OpenVendorFile VendorFile, NumVRecs
   OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen

   LogFile = FreeFile
   LogFileName$ = "APLINK.LOG"
   Open LogFileName$ For Output As LogFile
   ToPrint$ = Space$(80)
   ToPrint$ = "Linking operations began on " + Date$ + " " + Time$
   Print #LogFile, ToPrint$

   'CommaFmt$ = "#######,.##"

   '--Reset the vendor trans pointers to 0.
   'PRINT "Initializing Vendor File."
   For cnt& = 1 To NumVRecs
      Get VendorFile, cnt&, apvendor
         apvendor.FrstTran = 0
         apvendor.LastTran = 0
         
      Put VendorFile, cnt&, apvendor
   Next

   'Print "APLedger Records:"; NumTran&
''This is only if need to relink
   '--Relink Transactions to Vendor
   For cnt& = 1 To NumTran&
    FrmShowPctComp.ShowPctComp cnt&, NumTran&

      Get APLedgerFile, cnt&, ApLedger

      '--reset next transaction pointer to 0
      ApLedger.NextTrans = 0
      Put APLedgerFile, cnt&, ApLedger
      'this is for fix for harryburg
      If ApLedger.VRecNum = 1108 Then ApLedger.VRecNum = 521
      VRecNum = ApLedger.VRecNum   'GetVendorRec(apledger.VendorCode)
      'VRECNUM = GetVendorRec(APLedger.VendorCode)

      If VRecNum > 0 And VRecNum <= NumVRecs Then

         Get VendorFile, VRecNum, apvendor

         If apvendor.FrstTran > 0 Then
            '--Vendor has previous transactions..
            '--Remember the last transaction for this vendor
            Prev& = apvendor.LastTran

            '--In the vendor record...
            '--Set the Last Trans pointer to this record
            apvendor.LastTran = cnt&
            Put VendorFile, VRecNum, apvendor

            '--In the apledger record...
            '--Set the Last trans pointer in the prev trans
            '--to point to this record
            Get APLedgerFile, Prev&, ApLedger
            ApLedger.NextTrans = cnt&
            Put APLedgerFile, Prev&, ApLedger
         Else
            '--First Trans for this vendor
            '--set both pointers to this ledger record
            apvendor.FrstTran = cnt&
            apvendor.LastTran = cnt&
            Put VendorFile, VRecNum, apvendor
        End If
      Else
        Orphan = Orphan + 1
        GoSub LogAPLOrphan
      End If

     ' Print "Processed Record: "; cnt&

   Next

   ToPrint$ = Space$(80)
   ToPrint$ = "Linking operations completed on " + Date$ + " " + Time$
   Print #LogFile, ToPrint$
   If Ophan > 1 Then
     ToPrint$ = Space$(80)
     ToPrint$ = "Orphaned Transactions: " + Str$(Ophan)
   Else
     ToPrint$ = Space$(80)
     ToPrint$ = "No Orphaned Transactions. "
   End If
   Print #LogFile, ToPrint$
   Close
ActivateControls frmAPLdgUtilMenu

Exit Function

LogAPLOrphan:
  Ophan = Orphan + 1
  Print #LogFile, "Rec: " + Str$(cnt&) + " Orphan: " + Str$(ApLedger.VRecNum)
Return
CancelExit:
Exit Function

End Function



Public Sub RelinkDist2Trans4SB()
  Dim APLedgerFile As Integer, NumTran As Long, APLRecLen As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, APDRecLen As Integer
  Dim cnt As Long, Prev As Long, CntD As Long, keepup As Long
  keepup = 0

   OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen
    Get APLedgerFile, 19654, ApLedger
      'ApLedger.NextTrans = 19654
      'Put APLedgerFile, 19653, ApLedger
      'Get APLedgerFile, 19654, ApLedger
      ApLedger.VRecNum = 124
      Put APLedgerFile, 19654, ApLedger
   'OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

   'PRINT "Initializing Ledger Database."
   'For cnt& = 1 To NumTran&
'      Get APLedgerFile, cnt&, ApLedger
'      If ApLedger.FrstDist <= 0 Or ApLedger.LastDist <= 0 Then
'        keepup = keepup + 1
'        GoTo skipthis
'      End If
'      For CntD& = ApLedger.FrstDist To ApLedger.LastDist
'        Get APDistFile, CntD&, APDist
'        APDist.APLedgerRec = cnt&
'        Put APDistFile, CntD&, APDist
'      Next
'skipthis:
'   Next
'
'   If keepup > 0 Then
'   For cnt& = 1 To NumDistRecs&
'     Get APDistFile, cnt&, APDist
'      If APDist.APLedgerRec <= 0 Then Stop
'
'    Next
'
'   End If
   Close
End Sub

Public Function setallvendstoact()
  Dim VendorFile As Integer, NumVRecs As Integer
  Dim LogFile As Integer, LogFileName As String, ToPrint As String
  Dim cnt As Long, Prev As Long, Ophan As Integer, VRecNum As Integer
'Open Percent thingy here
  FrmShowPctComp.Label1 = "Fixing A/P Vendors."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DeActivateControls frmAPLdgUtilMenu
  DoEvents
  Orphan = 0
   OpenVendorFile VendorFile, NumVRecs
   For cnt& = 1 To NumVRecs
      Get VendorFile, cnt&, apvendor
        If QPTrim(apvendor.vnum) = "" Then
           apvendor.ActiveFlag = 1
          Put VendorFile, cnt&, apvendor
        End If
   Next
   Close
   Unload FrmShowPctComp
   ActivateControls frmAPLdgUtilMenu

End Function
Public Function FixoneVendcode()
  Dim VendorFile As Integer, NumVRecs As Integer
  Dim LogFile As Integer, LogFileName As String, ToPrint As String
  Dim cnt As Long, Prev As Long, Ophan As Integer, VRecNum As Integer
'Open Percent thingy here
  FrmShowPctComp.Label1 = "Fixing A/P Vendor."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DeActivateControls frmAPLdgUtilMenu
  DoEvents
  Orphan = 0
   OpenVendorFile VendorFile, NumVRecs
   'For cnt& = 1 To NumVRecs
      Get VendorFile, 446, apvendor
         apvendor.vnum = "B & J OLD"
      Put VendorFile, 446, apvendor
   'Next
   Close
   Unload FrmShowPctComp
   ActivateControls frmAPLdgUtilMenu

End Function
Public Function FixoneVendTransLink()
  Dim VendorFile As Integer, NumVRecs As Integer
  Dim LogFile As Integer, LogFileName As String, ToPrint As String
  Dim cnt As Long, Prev As Long, Ophan As Integer, VRecNum As Integer
'Open Percent thingy here
'  FrmShowPctComp.Label1 = "Fixing A/P Vendor."
'  FrmShowPctComp.cmdCancel.Enabled = False
'  FrmShowPctComp.Show , Me
'  DeActivateControls frmAPLdgUtilMenu
  DoEvents
  Orphan = 0
   OpenVendorFile VendorFile, NumVRecs
   For cnt& = 1 To NumVRecs
      Get VendorFile, cnt, apvendor
      If QPTrim(apvendor.vnum) = "FERGUSON" Then
        MsgBox ("Transnum - " + CStr(cnt&)), vbOKOnly
        
         apvendor.LastTran = 19654
        Put VendorFile, cnt, apvendor
      End If
   Next
   Close
'   Unload FrmShowPctComp
'   ActivateControls frmAPLdgUtilMenu

End Function
Private Sub FixAcctandVend4New()
  Dim VendorFile As Integer, NumVRecs As Integer
  Dim cnt As Long, VRecNum As Integer, CntA As Long
  Dim First As Long, Last As Long, RecNo As Long, AcctRecNum As Integer
  Dim GLAcctFile As Integer, NumAccts As Integer
  Dim AcctIdxFileNum As Integer, NumAIdxRecs As Integer, LookFor As String
  Dim GLAcct As GLAcctRecType
'Open Percent thingy here
  FrmShowPctComp.Label1 = "Fixing Trans links to 0 for new start."
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , Me
  DeActivateControls frmAPLdgUtilMenu
  DoEvents
  Orphan = 0
   OpenVendorFile VendorFile, NumVRecs

   For cnt& = 1 To NumVRecs
      Get VendorFile, cnt&, apvendor
         apvendor.FrstTran = 0
         apvendor.LastTran = 0
         apvendor.CurrBal = 0
         apvendor.FrstPO = 0
         apvendor.LastPO = 0
      Put VendorFile, cnt&, apvendor
   Next
  Close
   OpenAcctFile GLAcctFile, NumAccts
   Lock GLAcctFile

   FrmShowPctComp.Label1 = "Initializing account file."
   FrmShowPctComp.cmdCancel.Enabled = False
   FrmShowPctComp.Show , Me
   DoEvents

   '-Set the pointers in the account file to zero
   For cnt = 1 To NumAccts
      FrmShowPctComp.ShowPctComp cnt, NumAccts
      Get GLAcctFile, cnt, GLAcct
      GLAcct.FrstTran = 0
      GLAcct.Bal = 0
      GLAcct.LastTran = 0
      GLAcct.FrstBTran = 0
      GLAcct.Bgt = 0
      GLAcct.LastBTran = 0
      GLAcct.Encumb = 0
      GLAcct.FrstPTran = 0
      GLAcct.LastPTran = 0

      Put GLAcctFile, cnt, GLAcct
  Next          'Process next transaction
   Unlock GLAcctFile

   Close
   
  MsgBox "This is done.", vbOKOnly, "Complete"
  ActivateControls frmAPLdgUtilMenu
End Sub
'Public Sub Fixvoidpoamt()
'  Dim APLedgerFile As Integer, NumTran As Long, APLRecLen As Integer
'  Dim cnt As Long
'  OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen
'  'For cnt = 1 To NumTran&
'
'    Get APLedgerFile, 21313, ApLedger
'    If ApLedger.Amt = 0 Then Stop
''    If InStr(1, ApLedger.DOCNum, "~") Then
''      'Stop
''      ApLedger.DOCNum = QTilStrip$(ApLedger.DOCNum)
''      Put APLedgerFile, cnt, ApLedger
''    End If
''    If InStr(1, ApLedger.Comment, "~") Then
''      'Stop
''      ApLedger.Comment = QTilStrip$(ApLedger.Comment)
''      Put APLedgerFile, cnt, ApLedger
''    End If
'' Next
'
'    Close
'    MsgBox "Done", vbOKOnly
'End Sub

Public Sub Fixonetilda()
  Dim APLedgerFile As Integer, NumTran As Long, APLRecLen As Integer
  Dim cnt As Long
  OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen
  For cnt = 1 To NumTran&

    Get APLedgerFile, cnt, ApLedger
    If InStr(1, ApLedger.DOCNum, "~") Then
      'Stop
      ApLedger.DOCNum = QTilStrip$(ApLedger.DOCNum)
      Put APLedgerFile, cnt, ApLedger
    End If
    If InStr(1, ApLedger.Comment, "~") Then
      'Stop
      ApLedger.Comment = QTilStrip$(ApLedger.Comment)
      Put APLedgerFile, cnt, ApLedger
    End If
 Next
   
    Close
    MsgBox "Done", vbOKOnly
End Sub

Public Function QTilStrip$(Desc$)
  Dim x As String, DashPos As Integer
   x$ = QPTrim$(Desc$)  '(Form$(AcctNum, 0))
   Do
      DashPos = InStr(x$, "~")
      If DashPos > 0 Then
         x$ = Left$(x$, DashPos - 1) + Mid$(x$, DashPos + 1)
      End If
    Loop While DashPos

    QTilStrip$ = x$

End Function
'Private Sub UndeleteOne()
'  Dim VendorFile As Integer, NumVRecs As Integer
'  Dim cnt As Long, VRecNum As Integer
'   OpenVendorFile VendorFile, NumVRecs
'      Get VendorFile, 489, apvendor
'         apvendor.DelFlag = False
'         apvendor.ActiveFlag = 0
'      Put VendorFile, 489, apvendor
'   Close
'  MsgBox "Done", vbOKOnly
'End Sub
Private Sub SpecialVoidPOTrans()
  Dim LdRecLen As Integer, DistRecLEn As Integer, POLogFileName As String
  Dim APLedgerFile As Integer, NumTrans As Long, IFRec As Integer
  Dim POIFFile As String, GLIFRecLen As Integer, GLIFFile As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, NextDist As Long
  Dim AcctNum As String, BadAcct As Integer, ReportFile As String
 Dim EncAcct As String
 
  Dim ApLedger(1) As APLedger81RecType
  ReDim DistRec(1) As APDistRecType
  LdRecLen = Len(ApLedger(1))
  DistRecLEn = Len(DistRec(1))
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  GLGetEncAcct EncAcct

    
    
  POIFFile$ = "POVDIF.DAT"
  KillFile POIFFile$
  ReDim GLifRec(1) As GLTransRecType
  GLIFRecLen = Len(GLifRec(1))
  GLIFFile = FreeFile
  Open POIFFile$ For Random As GLIFFile Len = GLIFRecLen
  OpenAPDistFile APDistFile, NumDistRecs&, DistRecLEn
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  Get APLedgerFile, 29706, ApLedger(1)
  NextDist& = ApLedger(1).FrstDist
  Close APLedgerFile

  Do Until NextDist& = 0
    Get APDistFile, NextDist&, DistRec(1)
    IFRec = IFRec + 1
    'make sure distribution hasn't been liquidated
    If (QPTrim(DistRec(1).DistStat)) <> "L" Then
    '--Make Debit side of entry

      GLifRec(1).Src = "VD" + Format(Date$, "mmddyy")
      AcctNum$ = Left$(DistRec(1).DistAcctNum, GLFundLen) + EncAcct$
      GLifRec(1).AcctNum = AcctNum$
      GLifRec(1).TRDATE = DateDiff("d", "12/31/1979", Date$)
      GLifRec(1).Desc = "CANCELLED PO"            'APLedger(1).PONum
      GLifRec(1).Ref = ApLedger(1).PONum
      GLifRec(1).CrAmt = 0
      GLifRec(1).DrAmt = DistRec(1).DistAmt
      Put GLIFFile, IFRec, GLifRec(1)
    
      IFRec = IFRec + 1
      'AcctNum$ = LEFT$(DistRec(1).DistAcctNum, FundLen) + APAcct$
      'GLIFRec(1).AcctNum = AcctNum$
      GLifRec(1).AcctNum = DistRec(1).DistAcctNum
      GLifRec(1).CrAmt = DistRec(1).DistAmt
      GLifRec(1).DrAmt = 0
      Put GLIFFile, IFRec, GLifRec(1)
    Else
      GLifRec(1).Src = "VD" + Format(Date$, "mmddyy")
      AcctNum$ = Left$(DistRec(1).DistAcctNum, GLFundLen) + EncAcct$
      GLifRec(1).AcctNum = AcctNum$
      GLifRec(1).TRDATE = DateDiff("d", "12/31/1979", Date$)
      GLifRec(1).Desc = "CANCELLED PO"            'APLedger(1).PONum
      GLifRec(1).Ref = ApLedger(1).PONum
      GLifRec(1).CrAmt = 0
      GLifRec(1).DrAmt = 0
      Put GLIFFile, IFRec, GLifRec(1)
  
      IFRec = IFRec + 1
      'AcctNum$ = LEFT$(DistRec(1).DistAcctNum, FundLen) + APAcct$
      'GLIFRec(1).AcctNum = AcctNum$
      GLifRec(1).AcctNum = DistRec(1).DistAcctNum
      GLifRec(1).CrAmt = 0
      GLifRec(1).DrAmt = 0
      Put GLIFFile, IFRec, GLifRec(1)

    End If
    NextDist& = DistRec(1).NextDist

  Loop

  Close

  GLPost2PO POIFFile$, BadAcct, frmAPLdgUtilMenu, False
  If BadAcct <> 0 Then
    '--Couldn't find an account.
    '--Account was possibly deleted after entry made?
      MsgBox "Errors Were Found, DO NOT CONTINUE!! Contact Software Support.", vbOKOnly, "Errors"
      ReportFile$ = "TempLog.PRN"
      frmReportOpt.Show 1
      If rptopt = 1 Then
        ARptErrorLog.GetName ReportFile$
        ARptErrorLog.startrpt
      ElseIf rptopt = 2 Then
        ViewPrint ReportFile$, "Error Log"
      End If
'      frmAPLdgUtilMenu.Show
'      Unload frmPOCancel
'      frmPOProcessMenu.Show
      Exit Sub

  End If
  GLPost2PO POIFFile$, BadAcct%, frmAPLdgUtilMenu, True
  If BadAcct <> 0 Then                  'posting problem
      MsgBox "Error, One or more transactions were not posted. Make sure the printer is ready and Press a Key to View Log.", vbOKOnly, "Posting Error"
      POLogFileName = "POlog.dat"
      ReportFile$ = "POlog.dat"
      frmReportOpt.Show 1
      If rptopt = 1 Then
        ARptErrorLog.GetName ReportFile$
        ARptErrorLog.startrpt
      ElseIf rptopt = 2 Then
        ViewPrint ReportFile$, "Posting Log"
      End If
   End If
  
  OpenAPLedgerFile APLedgerFile, NumTrans, LdRecLen
  Get APLedgerFile, 29706, ApLedger(1)
  ApLedger(1).TRCode = -4
  Put APLedgerFile, 29706, ApLedger(1)
  Close APLedgerFile

  KillFile POIFFile$
MsgBox "Cancel Purchase Order Complete.", vbOKOnly, "Completed"
End Sub

Private Function GLGetEncAcct(EncAcct As String)
  Dim GLSetup As GLSetupRecType, SetUpRecLen As Integer, SetupFile As Integer
  SetUpRecLen = Len(GLSetup)
  SetupFile = FreeFile
  Open "GLSetup.DAT" For Random Access Read Write Shared As SetupFile Len = SetUpRecLen
  Get SetupFile, 1, GLSetup
  EncAcct = QPTrim(GLSetup.EncAcct)
  Close SetupFile
End Function
Private Sub GLPost2PO(FileName$, BadTrans%, formname As Form, go4it As Boolean)
  Dim TrRecLen As Integer, File2Post As Integer, Num2Post As Integer
  Dim AcctFileNum As Integer, NumAccts As Integer, Log As String
  Dim TransFileNum As Integer, NumTrans As Long, cnt As Integer
  Dim POLogFileName As String, POLogFile As Integer, RecNum As Integer
  Dim Posted As Integer, Prev As Long, TransPosted As Integer
  Dim PRNFile As Integer, ReportFile As String
  Dim Acct As GLAcctRecType
  Dim Tran2Post As GLTransRecType        '--Dim a buffer for the edit file
 ' On Local Error GoTo ItsBroke
  Dim POTrans As GLTransRecType

  TrRecLen = Len(Tran2Post)              'Determine the rec length
  File2Post = FreeFile                   'Get a handle on the Interface file
  Open FileName$ For Random As File2Post Len = TrRecLen
  Num2Post = LOF(File2Post) \ TrRecLen   'Find the num of transactions

  OpenAcctFile AcctFileNum, NumAccts     'Open & lock GL files
   'LOCK AcctFileNum

  OpenPOTransFile TransFileNum, NumTrans&
   'LOCK TransFileNum

   '--update the posting log file
  If go4it = True Then
    POLogFileName$ = "GLUTIL.LOG"
    POLogFile = FreeFile
    Open POLogFileName$ For Append As POLogFile
    Print #POLogFile, "Purchase Order initiated on " + Date$ + " @ " + Time$
    Log$ = Space$(132)
    'set correct Title for screen
    FrmShowPctComp.Label1 = "Posting PO Transactions"
  Else
    PRNFile = FreeFile
    ReportFile$ = "TempLog.PRN"
    Open ReportFile$ For Output As #PRNFile
    Print #PRNFile, "PO Verification initiated on " + Date$ + " @ " + Time$
    'set correct screen title
    FrmShowPctComp.Label1 = "Checking Accounts"
  End If
  FrmShowPctComp.cmdCancel.Enabled = False
  FrmShowPctComp.Show , formname
  DoEvents
  For cnt = 1 To Num2Post                'Start processing transactions
    FrmShowPctComp.ShowPctComp cnt, Num2Post
    Get File2Post, cnt, Tran2Post

    RecNum = AcctFind(Tran2Post.AcctNum)   'Verify account is in G/L
    'Use recnum = 0 to test error log on verification
    'RecNum = 0
    If RecNum > 0 Then                  'if valid acct then proceed
    '''''If cnt = 25 Then Stop
         'tell user what's going on
        ' QPrintRC " Posting Account Number: ", 25, 1, 112
        ' QPrintRC Tran2Post.AcctNum, 25, 26, 112
'skip this part if posting potrans from invoices update acct.encumb there
'with correct amt from po not invoice.
         
         Get AcctFileNum, RecNum, Acct    'Get the account
         If Left$(Tran2Post.Src, 2) <> "AP" Then

         '--Update encumbrace field
         Select Case Acct.Typ
            Case "A", "E"                 'asset, exp accts
               Acct.Encumb = Round#(Acct.Encumb + Tran2Post.DrAmt - Tran2Post.CrAmt)
               If go4it = True Then
                 Put AcctFileNum, RecNum, Acct
               End If
            Case "L", "R"                 'liab, rev accts
               Acct.Encumb = Round#(Acct.Encumb + Tran2Post.CrAmt - Tran2Post.DrAmt)
               If go4it = True Then
                 Put AcctFileNum, RecNum, Acct
               End If
         End Select
         End If
         NumTrans& = NumTrans& + 1          'increment record pointer

         Get TransFileNum, NumTrans&, POTrans

         POTrans.AcctNum = Tran2Post.AcctNum 'Assign editfile to trans history
         POTrans.TRDATE = Tran2Post.TRDATE
         POTrans.Desc = Tran2Post.Desc
         POTrans.LDesc = Tran2Post.LDesc
         POTrans.CrAmt = Tran2Post.CrAmt
         POTrans.DrAmt = Tran2Post.DrAmt
         POTrans.Ref = Tran2Post.Ref
         POTrans.Src = Tran2Post.Src
         POTrans.NextTran = 0
         If go4it = True Then
           Put TransFileNum, NumTrans&, POTrans
         End If
         Posted = Posted + 1
         '---------------------------------Start linking here
         If Acct.FrstPTran = 0 Then        'if first trans for this acct,
            Acct.FrstPTran = NumTrans&      'assign first & last pointers to
            Acct.LastPTran = NumTrans&      'this transaction
            If go4it = True Then
              Put AcctFileNum, RecNum, Acct
            End If
         Else                             'otherwise
                                          'in the account file..
            Prev& = Acct.LastPTran             'remember the prev trans pointe
            Acct.LastPTran = NumTrans&        'reset last trans to this trans
            If go4it = True Then
              Put AcctFileNum, RecNum, Acct
            End If
                                          'In the POTrans file...
            Get TransFileNum, Prev&, POTrans    'Get the last transaction
            POTrans.NextTran = NumTrans&       'reset pointer to this trans
            If go4it = True Then
              Put TransFileNum, Prev&, POTrans
            End If
         End If

         TransPosted = TransPosted + 1

      Else                                'Account NOT found!
         BadTrans = BadTrans + 1          'Pass info back to caller
         '--how about an error log here.
         If go4it = True Then
            GoSub LogPOPostErr
         Else
            GoSub LogTempErr
         End If

      End If

   Next

   'UNLOCK AcctFileNum
   'UNLOCK TransFileNum
  If go4it = True Then
    If BadTrans = 0 Then
      Print #POLogFile, ("No Posting Errors. Posted Transaction Count: " + Using("####", TransPosted))
      Print #POLogFile, String$(80, "-")
    End If
  Else
    If BadTrans = 0 Then
      Print #PRNFile, ("No Errors Found. Transaction Count :" + Using$("####", TransPosted))
    End If
  End If
  Close AcctFileNum
  Close TransFileNum
  Close File2Post
  Close POLogFile
  Close
'Clean up editfile in calling program in case not posted
Exit Sub

POGotErr:
   Select Case Err
      Case 70
         Close
         MsgBox "Another user has the file locked, Please try again later.", vbOKOnly, "Access Denied"
         Exit Sub
      Case Else
   End Select
Return
LogTempErr:
   Print #PRNFile, "Error: Unpostable Transaction "
   Print #PRNFile, "Record Number  :"; Str$(cnt)
   Print #PRNFile, "Account Number :"; Tran2Post.AcctNum
   Print #PRNFile, "Date           :"; Format(DateAdd("d", (Tran2Post.TRDATE), "12-31-1979"), "mm/dd/yyyy")
   Print #PRNFile, "Description    :"; Tran2Post.Desc
   Print #PRNFile, "Debit          :"; Str$(Tran2Post.CrAmt)
   Print #PRNFile, "Credit         :"; Str$(Tran2Post.DrAmt)
   Print #PRNFile, "**********************"

Return

LogPOPostErr:
   Print #POLogFile, "Error: Unposted Transaction "
   Print #POLogFile, "Record Number  :"; Str$(cnt)
   Print #POLogFile, "Account Number :"; Tran2Post.AcctNum
   Print #POLogFile, "Date           :"; Format(DateAdd("d", (Tran2Post.TRDATE), "12-31-1979"), "mm/dd/yyyy")
   Print #POLogFile, "Description    :"; Tran2Post.Desc
   Print #POLogFile, "Debit          :"; Str$(Tran2Post.CrAmt)
   Print #POLogFile, "Credit         :"; Str$(Tran2Post.DrAmt)
   Print #POLogFile, "********************"

Return
ItsBroke:
  BadTrans = BadTrans + 1
  Print #PRNFile, "Error *** Call Software Support***"
  Print #PRNFile, "Record Number :"; Str$(cnt); Tran2Post.AcctNum
  Print #PRNFile, "Error Code"; Str(Err.Number)
  Resume Next

End Sub

Public Sub FixJohnsonDist2Trans()
  Dim APLedgerFile As Integer, NumTran As Long, APLRecLen As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, APDRecLen As Integer
  Dim cnt As Long, Prev As Long
  
'  FrmShowPctComp.Label1 = "Linking Distributions to Ledger."
'  FrmShowPctComp.cmdCancel.Enabled = False
'  FrmShowPctComp.Show , Me
  DeActivateControls frmAPLdgUtilMenu
  DoEvents

  ' OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen


   OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

   'PRINT "Initializing Ledger Database."
'   For cnt& = 1 To NumTran&
'      Get APLedgerFile, cnt&, ApLedger
'      ''''If ApLedger.FrstDist < 0 Or ApLedger.LastDist < 0 Then Stop
'      ApLedger.FrstDist = 0
'      ApLedger.LastDist = 0
'      Put APLedgerFile, cnt&, ApLedger
'   Next

 '  For cnt& = 1 To NumDistRecs&
'Put % thing here
 '    FrmShowPctComp.ShowPctComp cnt&, NumDistRecs&
      '--Assume no one else will follow.
      Get APDistFile, 32885, APDist
        APDist.NextDist = 0
        APDist.APLedgerRec = 21249
        APDist.DistAcctNum = "15-0002-638"
        APDist.DistAmt = 2409.61
      Put APDistFile, 32885, APDist
      Get APDistFile, 32886, APDist
        APDist.NextDist = 0
        APDist.APLedgerRec = 21250
        APDist.DistAcctNum = "10-0004-614"
        APDist.DistAmt = 654.99
      Put APDistFile, 32886, APDist
      Get APDistFile, 32887, APDist
        APDist.NextDist = 0
        APDist.APLedgerRec = 21251
        APDist.DistAcctNum = "15-0002-646"
        APDist.DistAmt = 1281.5
      Put APDistFile, 32887, APDist
      Get APDistFile, 32888, APDist
        APDist.NextDist = 32889
        APDist.APLedgerRec = 21252
        APDist.DistAcctNum = "15-0002-650"
        APDist.DistAmt = 5931#
      Put APDistFile, 32888, APDist
      Get APDistFile, 32889, APDist
        APDist.NextDist = 32890
        APDist.APLedgerRec = 21252
        APDist.DistAcctNum = "10-0003-650"
        APDist.DistAmt = 988.5

      Put APDistFile, 32889, APDist
'      IF APLdgrDist.APLedgerRec > NumTran& THEN
'        STOP
'      END IF

'      IF APLdgrDist.APLedgerRec > 3555 AND APLdgrDist.APLedgerRec < 26905 THE
'        STOP
'      END IF
      '--Get the parent record

'      If APDist.APLedgerRec > NumTran& Or APDist.APLedgerRec <= 0 Then 'Stop
'        APDist.APLedgerRec = -1
'        Put APDistFile, cnt&, APDist
'        GoTo SkipHere
'      End If
'      Get APLedgerFile, APDist.APLedgerRec, ApLedger
'
'      If ApLedger.FrstDist > 0 Then
'         '--We're not the first one here, so let us not forget those who have
'         '--come before us
'         Prev& = ApLedger.LastDist
'
'         '--This is now the new last distribution
'         '--Update Last Dist pointer in apledger to this rec
'         ApLedger.LastDist = cnt&
'         Put APLedgerFile, APDist.APLedgerRec, ApLedger
'
'         '--Get the former last distribution
'         '--and tell it that this rec is the next one
'         Get APDistFile, Prev&, APDist
'         APDist.NextDist = cnt&
'         Put APDistFile, Prev&, APDist
'
'      Else
'         '--Virgin territory. we're now first and last
'         ApLedger.FrstDist = cnt&
'         ApLedger.LastDist = cnt&
'         Put APLedgerFile, APDist.APLedgerRec, ApLedger
'      End If
'SkipHere:
'   Next

   Close

   'PRINT "Press any key to continue."
   'K$ = INPUT$(1)
ActivateControls frmAPLdgUtilMenu
End Sub

Private Sub PrintLedgerTransNums()
  Dim APLedgerFile As Integer, NumTran As Long, APLRecLen As Integer
  Dim APDistFile As Integer, NumDistRecs As Long, APDRecLen As Integer
  Dim VendorFile As Integer, NumVRecs As Integer, CommaFmtT As String
  Dim RptFile As Integer, RptFileName As String, CommaFmt As String
  Dim RunTotal As Double, cnt As Long, VendorName As String, LCnt As Long
  Dim ToPrint As String, NextDist As Long, DistAmt As Double, TDistAmt As Double
  Dim ThisRec As Long, BalMsg As String, Status As String, DistAA As Double
  Dim PDChkDate As String, PDChkNum As String
  OpenAPLedgerFile APLedgerFile, NumTran&, APLRecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen
  OpenVendorFile VendorFile, NumVRecs

  RptFile = FreeFile
  RptFileName$ = "apledger.prn"
  Open RptFileName$ For Output As RptFile
  CommaFmt$ = "#,###,###.##"
  CommaFmtT$ = "##,###,###,###.##"
  RunTotal# = 0
  FrmShowPctComp.Label1 = "Printing AP Ledger Report"
  FrmShowPctComp.Show , Me
  DeActivateControls frmAPLdgUtilMenu
  DoEvents
  For cnt& = 1 To NumTran&
 'Print Using; "Processing Ledger Record: #####"; cnt&;

    Get APLedgerFile, cnt&, ApLedger
    If ApLedger.VRecNum > 0 Then
      Get VendorFile, ApLedger.VRecNum, apvendor
      VendorName$ = apvendor.VNAME
    Else
      VendorName$ = "Orphaned Transaction"
    End If

    LCnt& = LCnt& + 1

    ToPrint$ = Space$(80)
    Mid$(ToPrint$, 2) = "Trans: " + Str$(cnt&)
    Mid$(ToPrint$, 15) = ApLedger.VendorCode
    Mid$(ToPrint$, 27) = VendorName$
    Mid$(ToPrint$, 60) = Format(DateAdd("d", (ApLedger.TRDATE), "12-31-1979"), "mm/dd/yyyy")
    
    Print #RptFile, ToPrint$
    LSet ToPrint$ = ""

    Mid$(ToPrint$, 15) = Str(ApLedger.FrstDist)
    Mid$(ToPrint$, 27) = Str(ApLedger.LastDist)
    Mid$(ToPrint$, 60) = Str(ApLedger.NextTrans)
    Print #RptFile, ToPrint$
    LSet ToPrint$ = ""
    
    Print #RptFile, String$(78, "=")
    'Put count up doololly here
    FrmShowPctComp.ShowPctComp cnt&, NumTran&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmAPLdgUtilMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

  Next    '--Now print the distribution

       Print #RptFile, Tab(40); "Accounting Distribution:"
  For cnt& = 1 To NumDistRecs&
   Get APDistFile, cnt&, APDist
'          If APDist.DistAmt > -100000000000# Then
'         ''' Stop
'            DistAA = Round#(APDist.DistAmt)
'            DistAmt# = Round#(DistAmt# + DistAA)
'            TDistAmt# = Round#(TDistAmt# + DistAA)
'          Else
'            DistAA = 0
'          End If
'          ThisRec& = NextDist&
'          NextDist& = APDist.NextDist
'
          ToPrint$ = Space$(80)
          Mid$(ToPrint$, 20) = Str$(ThisRec&)
          Mid$(ToPrint$, 40) = APDist.DistAcctNum
          Mid$(ToPrint$, 54) = Using(CommaFmt$, Str$(APDist.DistAmt))
          Mid$(ToPrint$, 67) = Str$(APDist.APLedgerRec)
          Print #RptFile, ToPrint$
'


'    End If
    Print #RptFile, String$(78, "=")
    'Put count up doololly here
    FrmShowPctComp.ShowPctComp cnt&, NumDistRecs&
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmAPLdgUtilMenu
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

  Next

  Print #RptFile, LCnt&
  Close
  
  ViewPrint RptFileName$, "APLEDGER.PRN"
  ActivateControls frmAPLdgUtilMenu
'  SHELL "list APLEDGER.PRN"

Exit Sub




CancelExit:
  Exit Sub
End Sub
