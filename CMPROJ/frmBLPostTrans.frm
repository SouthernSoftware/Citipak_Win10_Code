VERSION 5.00
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmBLPostTrans 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Post Transactions"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLPostTrans.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "ESC &Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   3637
      TabIndex        =   2
      ToolTipText     =   "Press to exit this screen."
      Top             =   6824
      Width           =   1884
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "F10 P&ost"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   6133
      TabIndex        =   1
      ToolTipText     =   "Press to post pending penalty assessments."
      Top             =   6809
      Width           =   1884
   End
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   3030
      Left            =   1891
      TabIndex        =   0
      Top             =   2789
      Width           =   7935
      _Version        =   196609
      _ExtentX        =   13996
      _ExtentY        =   5345
      _StockProps     =   70
      Caption         =   $"frmBLPostTrans.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   192
      AlignTextH      =   1
      AlignTextV      =   1
      Caption         =   $"frmBLPostTrans.frx":09F5
      ForeColor       =   8454143
      Picture         =   "frmBLPostTrans.frx":0B20
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSACTIONS POSTING"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2293
      TabIndex        =   3
      Top             =   1552
      Width           =   7068
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1501
      Top             =   1405
      Width           =   8652
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   3480
      Left            =   1681
      Top             =   2534
      Width           =   8415
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1501
      Top             =   1357
      Width           =   8652
   End
End
Attribute VB_Name = "frmBLPostTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      Call cmdExit_Click
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF10:
      Call cmdPost_Click
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLPostTrans.")
      Call Terminate
      End
    End If
  End If
End Sub
Private Sub cmdPost_Click()
  Dim PayHandle As Integer
  Dim PayRec As AREditPaymentRecType
  Dim NumOfPayRecs As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfCodeRecs As Integer
  Dim x As Integer
  Dim CustSrchIdxRec As CustSearchNameIdxType
  Dim SearchHandle As Integer
  Dim NumOfIdxRecs As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim TransRec As ARTransRecType
  Dim TransHandle As Integer
  Dim NumOfTransRecs As Long
  Dim NextTransRec As Long
  Dim cnt As Long
  Dim CRec$
  Dim SCnt As Long
  Dim ARCode$
  Dim NewBalance#
  Dim CustRecd As Integer
  Dim Prev&
  Dim ThisNumOfCodes As Integer
  Dim NextCode As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim LoopStop As Integer
  
  On Error GoTo ERRORSTUFF
  If Exist("artownsu.dat") Then
    OpenTownFile TownHandle
    Get TownHandle, 1, TownRec
    Close TownHandle
  End If
  
  OpenCustFile CustHandle
  NumOfCustRecs = LOF(CustHandle) / Len(CustRec)
  
  OpenPayFile PayHandle, OPERNUM 'file saved from TransEntry
  NumOfPayRecs = LOF(PayHandle) / Len(PayRec)

  ' See if any records to post
  OpenTransFile TransHandle
  NumOfTransRecs = LOF(TransHandle) / Len(TransRec)
  NextTransRec = NumOfTransRecs + 1
  
  Do
    ThisNumOfCodes = 1
    cnt = cnt + 1
    Get PayHandle, cnt, PayRec
    CRec = Val(PayRec.CustNumber)
    If CRec > 0 Then
      Get CustHandle, Val(PayRec.CustNumber), CustRec
      'Set New Balance
      NewBalance# = CustRec.AcctBal - PayRec.Amount

      ' Post Transaction Record First
      TransRec.CustomerNumber = PayRec.CustNumber
      TransRec.TransDate = PayRec.TranDate
      TransRec.TransAmount = PayRec.Amount
      TransRec.TransType = 2               ' Type 2 = Payment
      TransRec.TransDesc = QPTrim$(PayRec.DESC)
      TransRec.CashAmount = PayRec.Amount
      TransRec.ChkAmount = 0
      TransRec.BalanceAfterTrans = NewBalance#
      TransRec.ExtraRoom = ""
      TransRec.Posted2GL = "N"
      TransRec.CatCodeRec1 = GetCatRecNum(QPTrim$(CustRec.BILLCAT1)) 'CatRecord(1) 'record # for category code
      TransRec.CatCodeRec2 = GetCatRecNum(QPTrim$(CustRec.BILLCAT2)) 'CatRecord(2)
      TransRec.CatCodeRec3 = GetCatRecNum(QPTrim$(CustRec.BILLCAT3)) '
      TransRec.CatCodeRec4 = GetCatRecNum(QPTrim$(CustRec.BILLCAT4)) '
      TransRec.CatCodeRec5 = GetCatRecNum(QPTrim$(CustRec.BILLCAT5)) '
      If PayRec.LICPAID1 > 0 Or PayRec.LICPAID2 > 0 Or PayRec.LICPAID3 > 0 Or PayRec.LICPAID4 > 0 Or PayRec.LICPAID5 > 0 Or PayRec.ISSPAID > 0 Then
        TransRec.DetailTransType = 210
      End If
      TransRec.CatLicAmt1 = OldRound(PayRec.LICPAID1)
      TransRec.CatLicAmt2 = OldRound(PayRec.LICPAID2)
      TransRec.CatLicAmt3 = OldRound(PayRec.LICPAID3)
      TransRec.CatLicAmt4 = OldRound(PayRec.LICPAID4)
      TransRec.CatLicAmt5 = OldRound(PayRec.LICPAID5)
      TransRec.CatLicBal1 = OldRound(PayRec.LICDUE1 - PayRec.LICPAID1)
      TransRec.CatLicBal2 = OldRound(PayRec.LICDUE2 - PayRec.LICPAID2)
      TransRec.CatLicBal3 = OldRound(PayRec.LICDUE3 - PayRec.LICPAID3)
      TransRec.CatLicBal4 = OldRound(PayRec.LICDUE4 - PayRec.LICPAID4)
      TransRec.CatLicBal5 = OldRound(PayRec.LICDUE5 - PayRec.LICPAID5)
      TransRec.FeeAmt = 0
      TransRec.PenAmt = OldRound(PayRec.PENPAID)
      TransRec.IssAmt = OldRound(PayRec.ISSPAID)
      TransRec.LicAmt = OldRound(PayRec.LICPAID1 + PayRec.LICPAID2 + PayRec.LICPAID3 + PayRec.LICPAID4 + PayRec.LICPAID5)
      TransRec.IssBal = OldRound(PayRec.ISSDUE - PayRec.ISSPAID)
      TransRec.LicBal = OldRound(PayRec.LICDUE - PayRec.LICPAID)
      TransRec.PenBal = OldRound(PayRec.PENDUE - PayRec.PENPAID)
      If PayRec.PENPAID > 0 Then
        If TransRec.DetailTransType = 210 Then
          TransRec.DetailTransType = 211
        Else
          TransRec.DetailTransType = 201
        End If
      End If
      TransRec.NextTrans = 0
      Put TransHandle, NextTransRec, TransRec
      'Update Customer
      CustRecd = Val(PayRec.CustNumber)
      Get CustHandle, CustRecd, CustRec
      CustRec.IssueLicense = PayRec.ISSUELIC
      CustRec.AcctBal = OldRound(CustRec.AcctBal - PayRec.Amount)
      CustRec.LicBal = OldRound(CustRec.LicBal - PayRec.LICPAID)
      CustRec.FeeLicBal1 = OldRound(PayRec.LICDUE1 - PayRec.LICPAID1)
      CustRec.FeeLicBal2 = OldRound(PayRec.LICDUE2 - PayRec.LICPAID2)
      CustRec.FeeLicBal3 = OldRound(PayRec.LICDUE3 - PayRec.LICPAID3)
      CustRec.FeeLicBal4 = OldRound(PayRec.LICDUE4 - PayRec.LICPAID4)
      CustRec.FeeLicBal5 = OldRound(PayRec.LICDUE5 - PayRec.LICPAID5)
      CustRec.FeeLicPay1 = PayRec.LICPAID1
      CustRec.FeeLicPay2 = PayRec.LICPAID2
      CustRec.FeeLicPay3 = PayRec.LICPAID3
      CustRec.FeeLicPay4 = PayRec.LICPAID4
      CustRec.FeeLicPay5 = PayRec.LICPAID5
      CustRec.PenBal = OldRound(CustRec.PenBal - PayRec.PENPAID)
      CustRec.IssuanceFee = PayRec.ISSueFEE
      CustRec.IssuanceBal = OldRound(CustRec.IssuanceBal - PayRec.ISSPAID)
      CustRec.IssuancePay = PayRec.ISSPAID
      If PayRec.SetFee = "Y" Then
        CustRec.FeeAmt = PayRec.Amount
      End If

      Put CustHandle, CustRecd, CustRec

      If CustRec.FirstTrans = 0 Then
        CustRec.FirstTrans = NextTransRec
        CustRec.LastTrans = NextTransRec
        Put CustHandle, CustRecd, CustRec
      Else
        Prev& = CustRec.LastTrans
        CustRec.LastTrans = NextTransRec
        Put CustHandle, CustRecd, CustRec
        Get TransHandle, Prev&, TransRec
        TransRec.NextTrans = NextTransRec
        Put TransHandle, Prev&, TransRec
      End If
      NextTransRec = NextTransRec + 1
    End If

  Loop Until cnt > NumOfPayRecs
  Close

  KillFile "AREDPY" + Str$(OPERNUM) + ".DAT"
  MainLog ("Transaction entries posted. " + "AREDPY" + Str$(OPERNUM) + ".DAT" + " deleted.")
  
  frmBLMessageBoxJr.Label1.Caption = "Posting completed successfully!"
  frmBLMessageBoxJr.Label1.FontSize = 16
  frmBLMessageBoxJr.Label1.Top = 900
  frmBLMessageBoxJr.Show vbModal
  
  frmBLEnterPayments.Show
  DoEvents
  Unload frmBLPostTrans
  
  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLPostTrans", "cmdPost_Click", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
  

End Sub

Private Sub cmdExit_Click()
  frmBLEnterPayments.Show
  DoEvents
  Unload frmBLPostTrans
End Sub

