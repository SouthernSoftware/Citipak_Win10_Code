VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmVATaxPayPost 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Payment Post"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   Icon            =   "frmVATaxPayPost.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EditLib.fpText fptxtOperator 
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   2520
      Width           =   3135
      _Version        =   196608
      _ExtentX        =   5530
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPost 
      Height          =   492
      Left            =   6480
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1812
      _Version        =   131072
      _ExtentX        =   3196
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmVATaxPayPost.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   3120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1812
      _Version        =   131072
      _ExtentX        =   3196
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmVATaxPayPost.frx":0AA6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMinTax 
      Height          =   492
      Left            =   4320
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   3012
      _Version        =   131072
      _ExtentX        =   5313
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ButtonDesigner  =   "frmVATaxPayPost.frx":0C82
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Press F10 To Post."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   4538
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Press ESC To Exit."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   4538
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   2055
      Left            =   1995
      Top             =   3338
      Width           =   7575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "                                                                                      Ready to Post Tax Payment Transactions? "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2055
      Left            =   1995
      TabIndex        =   1
      Top             =   3338
      Width           =   7575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Payment Post"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3795
      TabIndex        =   0
      Top             =   1560
      Width           =   4020
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   840
      Left            =   2265
      Top             =   1320
      Width           =   7020
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   990
      Left            =   2280
      Top             =   1200
      Width           =   7020
   End
End
Attribute VB_Name = "frmVATaxPayPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim MinTax As Double
  Dim MinOpt As Integer
  Dim MinNames() As String
  Dim MinAmts() As Double
  Dim MinCnt As Integer
  Dim TaxYear As Integer
  Public ThisBillType$
Private Sub cmdExit_Click()
  frmVATaxPayMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdPost_Click()
  If ThisBillType = "R" Or ThisBillType = "B" Then
    Call PostReal
  ElseIf ThisBillType = "P" Then
    Call PostPers
  End If
End Sub

Private Sub PostReal()
  Dim Oper$
  Dim TaxPaymentRec As TaxPaymentRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  Dim PayListRec As RealPayListType
  Dim LHandle As Integer
  Dim NumOfLRecs As Integer
  Dim TaxCustRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim TaxTranRec As TaxTransactionType
  Dim TaxTranHandle As Integer
  Dim NumOfTaxTranRecs As Long
  Dim PayTranRec As TaxTransactionType
  Dim PayTranHandle As Integer
  Dim NumOfPayTranRecs As Long
  Dim EmptyPay As TaxTransactionType
  Dim cnt&, TotalPaid#
  Dim ThisListRec&
  Dim NextTransRec&
  Dim One As Integer, x As Integer
  Dim AHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  Oper$ = QPTrim$(Str$(OperNum))
  If Not Exist("TAXRCPR" + Oper$ + ".DAT") Then Exit Sub
  OpenTempRealPayFile PHandle, OperNum ' is the same as open TaxCPRFileName
  NumOfPRecs = LOF(PHandle) / Len(TaxPaymentRec)
  
  OpenRealPayListFile LHandle, OperNum 'is the same as open TaxLOPFileName
  NumOfLRecs = LOF(LHandle) / Len(PayListRec)
  
  OpenTaxCustFile CHandle, NumOfCRecs
  OpenTaxTransFile TaxTranHandle, NumOfTaxTranRecs
  
  For cnt& = 1 To NumOfPRecs 'NumOfPRecs = number of customers paying
    Get PHandle, cnt&, TaxPaymentRec
    ThisListRec& = TaxPaymentRec.LastPayRec
    Do While ThisListRec& > 0
      Get LHandle, ThisListRec&, PayListRec 'listrec is the list of bills being paid (some could be
      'multiple tags for a single customer
      'get paylist rec
      Get CHandle, TaxPaymentRec.CustAcct, TaxCustRec
      'get cust rec
      
      If PayListRec.BillRec < 0 Then
        GoSub PrePay
        GoTo SkipThisRec
      End If
      
      Get TaxTranHandle, PayListRec.BillRec, TaxTranRec
      'get bill trans this payrec is for
      TotalPaid# = 0
      PayTranRec = EmptyPay
      'make a new clean payment trans
'      TotalPaid# = OldRound#(PayListRec.DiscAmt + PayListRec.Principle1 + PayListRec.Interest1 + PayListRec.Collection + PayListRec.LateList)
      TotalPaid# = OldRound#(PayListRec.Principle1 + PayListRec.Interest1 + PayListRec.Collection + PayListRec.LateList + PayListRec.Penalty) '4/22/05
      TotalPaid# = OldRound(TotalPaid# + PayListRec.OptRev1 + PayListRec.OptRev2 + PayListRec.OptRev3 + PayListRec.PrePayAmt)
      If TotalPaid# = 0 Then
        GoTo SkipThisRec
      End If
      'PayTranRec = the new record for tax transaction records
      PayTranRec.TransDate = TaxPaymentRec.PayDate
      If PayListRec.PrePayAmt > 0 Then
        PayTranRec.TranType = 21 'overpay and bill pay combined
      Else
        PayTranRec.TranType = 2
      End If
'      PayTranRec.Revenue.Principle1Pd = OldRound(PayListRec.Principle1 + PayListRec.DiscAmt)
      PayTranRec.Revenue.Principle1Pd = OldRound(PayListRec.Principle1) '4/22/05
      PayTranRec.Revenue.InterestPd = PayListRec.Interest1
      PayTranRec.Revenue.CollectionPd = PayListRec.Collection
      PayTranRec.Revenue.LateListPd = PayListRec.LateList
      PayTranRec.Revenue.PenaltyPd = PayListRec.Penalty
      PayTranRec.Revenue.RevOpt1Pd = PayListRec.OptRev1
      PayTranRec.Revenue.RevOpt2Pd = PayListRec.OptRev2
      PayTranRec.Revenue.RevOpt3Pd = PayListRec.OptRev3
      PayTranRec.Revenue.Principle2Pd = 0
      PayTranRec.Revenue.Principle3Pd = 0
      PayTranRec.Revenue.Principle4Pd = 0
      PayTranRec.Revenue.Principle5Pd = 0
      PayTranRec.CustPin = TaxCustRec.PIN
      PayTranRec.DiscXDate = TaxTranRec.DiscXDate
      PayTranRec.RealPin = QPTrim$(TaxTranRec.RealPin)
      PayTranRec.PersPin = ""
      PayTranRec.Posted2GL = "N"
      PayTranRec.TaxYear = TaxTranRec.TaxYear
      PayTranRec.DiscAmt = PayListRec.DiscAmt
      PayTranRec.OperNum = OperNum
      PayTranRec.Amount = TotalPaid#
      If QPTrim$(PayListRec.Description) = "" Then
        PayTranRec.Description = TaxTranRec.Description
      Else
        PayTranRec.Description = QPTrim$(PayListRec.Description) + " " + ParseBillNum$(TaxTranRec.Description)
      End If
      PayTranRec.CustomerRec = TaxPaymentRec.CustAcct
      PayTranRec.LastTrans = TaxCustRec.LastTrans
      PayTranRec.BelongTo = PayListRec.BillRec
      PayTranRec.Revenue.PrePaidAmt = PayListRec.PrePayAmt
      PayTranRec.Revenue.PrePaidUsed = 0
      PayTranRec.Revenue.PrePaidBal = OldRound(GetOverPayBalance(TaxPaymentRec.CustAcct, "N") + PayTranRec.Revenue.PrePaidAmt)
      PayTranRec.InternalPin = TaxTranRec.InternalPin
      PayTranRec.BillType = TaxPaymentRec.BillType
      'TaxTranRec is the update to the existing tax record
'      TaxTranRec.Revenue.Principle1Pd = OldRound#(TaxTranRec.Revenue.Principle1Pd + PayListRec.Principle1 + PayListRec.DiscAmt)
      TaxTranRec.Revenue.Principle1Pd = OldRound#(TaxTranRec.Revenue.Principle1Pd + PayListRec.Principle1) '4/22/05
      TaxTranRec.Revenue.InterestPd = OldRound#(TaxTranRec.Revenue.InterestPd + PayListRec.Interest1)
      TaxTranRec.Revenue.CollectionPd = OldRound#(TaxTranRec.Revenue.CollectionPd + PayListRec.Collection)
      TaxTranRec.Revenue.LateListPd = OldRound#(TaxTranRec.Revenue.LateListPd + PayListRec.LateList)
      TaxTranRec.Revenue.PenaltyPd = OldRound#(TaxTranRec.Revenue.PenaltyPd + PayListRec.Penalty)
      TaxTranRec.Revenue.RevOpt1Pd = OldRound#(TaxTranRec.Revenue.RevOpt1Pd + PayListRec.OptRev1)
      TaxTranRec.Revenue.RevOpt2Pd = OldRound#(TaxTranRec.Revenue.RevOpt2Pd + PayListRec.OptRev2)
      TaxTranRec.Revenue.RevOpt3Pd = OldRound#(TaxTranRec.Revenue.RevOpt3Pd + PayListRec.OptRev3)
      TaxTranRec.Revenue.Future1Pd = OldRound#(TaxTranRec.DiscAmt + PayListRec.DiscAmt)
      TaxTranRec.DiscAmt = OldRound#(TaxTranRec.DiscAmt + PayListRec.DiscAmt)
      
      Put TaxTranHandle, PayListRec.BillRec, TaxTranRec
      NextTransRec& = (LOF(TaxTranHandle) \ Len(TaxTranRec)) + 1

      Put TaxTranHandle, NextTransRec&, PayTranRec
      TaxCustRec.LastTrans = NextTransRec&
      Put CHandle, TaxPaymentRec.CustAcct, TaxCustRec

SkipThisRec:
      ThisListRec& = PayListRec.PrevListRec
    Loop
  Next

  Close

  For x = 1 To 5
    KillFile ("TAXRCPR" + Oper$ + ".DAT")
    KillFile ("TAXRLOP" + Oper$ + ".DAT")
  Next x
  
  If Exist("TAXRCPR" + Oper$ + ".DAT") = True Or Exist("TAXRLOP" + Oper$ + ".DAT") = True Then '11/30/06
    One = 1
    AHandle = FreeFile
    Open "rpayposterror" + Oper$ + ".dat" For Output As AHandle
    Print #AHandle, One
    Close AHandle
    Call TaxMsg(800, "ERROR: POSTING WAS SUCCESSFUL but temporary payment files were not deleted properly. Please call Southern Software @ 1-800-842-8190 for assistance.")
    MainLog ("ERROR: A real payment error has taken place because payment batch files were not deleted. User was notified.")
  Else
    Call Savemsg(900, "Real payment posting has completed successfully.")
    MainLog ("Real payment post completed successfully.")
  End If
  
  If ThisBillType = "B" Then
    Call PostPers
  Else
    Call cmdExit_Click
  End If
  
  Exit Sub

PrePay:
  TotalPaid# = 0
  PayTranRec = EmptyPay
  'make a new clean payment trans
  TotalPaid# = OldRound#(PayListRec.DiscAmt + PayListRec.Principle1 + PayListRec.Interest1 + PayListRec.Collection + PayListRec.LateList)
  TotalPaid# = OldRound(TotalPaid# + PayListRec.OptRev1 + PayListRec.OptRev2 + PayListRec.OptRev3 + PayListRec.PrePayAmt)
  If TotalPaid# = 0 Then
    GoTo SkipThisRec
  End If
  'PayTranRec = the new record for tax transaction records
  PayTranRec.TransDate = TaxPaymentRec.PayDate
  PayTranRec.TranType = 22 'overpay only
  PayTranRec.Revenue.Principle1Pd = OldRound(PayListRec.Principle1 + PayListRec.DiscAmt)
  PayTranRec.Revenue.InterestPd = PayListRec.Interest1
  PayTranRec.Revenue.CollectionPd = PayListRec.Collection
  PayTranRec.Revenue.LateListPd = PayListRec.LateList
  PayTranRec.Revenue.PenaltyPd = PayListRec.Penalty
  PayTranRec.Revenue.RevOpt1Pd = PayListRec.OptRev1
  PayTranRec.Revenue.RevOpt2Pd = PayListRec.OptRev2
  PayTranRec.Revenue.RevOpt3Pd = PayListRec.OptRev3
  PayTranRec.Revenue.Principle2Pd = 0
  PayTranRec.Revenue.Principle3Pd = 0
  PayTranRec.Revenue.Principle4Pd = 0
  PayTranRec.Revenue.Principle5Pd = 0
  PayTranRec.CustPin = TaxCustRec.PIN
  PayTranRec.DiscXDate = TaxTranRec.DiscXDate
  PayTranRec.RealPin = " " 'QPTrim$(TaxTranRec.RealPin)'added 1/29/08
  PayTranRec.PersPin = ""
  PayTranRec.Posted2GL = "N"
  PayTranRec.TaxYear = TaxYear
  PayTranRec.DiscAmt = PayListRec.DiscAmt
  PayTranRec.OperNum = OperNum
  PayTranRec.Amount = TotalPaid#
  If QPTrim$(PayListRec.Description) = "" Then
    PayTranRec.Description = "Prepay"
  Else
    PayTranRec.Description = QPTrim$(PayListRec.Description)
  End If
  PayTranRec.CustomerRec = TaxPaymentRec.CustAcct
  PayTranRec.LastTrans = TaxCustRec.LastTrans
  PayTranRec.BelongTo = 0
  PayTranRec.Revenue.PrePaidAmt = PayListRec.PrePayAmt
  PayTranRec.Revenue.PrePaidUsed = 0
  PayTranRec.Revenue.PrePaidBal = OldRound(GetOverPayBalance(TaxPaymentRec.CustAcct, "N") + PayTranRec.Revenue.PrePaidAmt)
  PayTranRec.BillType = TaxPaymentRec.BillType
  NextTransRec& = (LOF(TaxTranHandle) \ Len(TaxTranRec)) + 1
  Put TaxTranHandle, NextTransRec&, PayTranRec
  
  TaxCustRec.LastTrans = NextTransRec&
  Put CHandle, TaxPaymentRec.CustAcct, TaxCustRec

  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayPost", "cmdPostReal_Click", Erl)
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

Private Sub PostPers()
  Dim Oper$
  Dim TaxPaymentRec As TaxPaymentRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  Dim PayListRec As PersPayListType
  Dim LHandle As Integer
  Dim NumOfLRecs As Integer
  Dim TaxCustRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim TaxTranRec As TaxTransactionType
  Dim TaxTranHandle As Integer
  Dim NumOfTaxTranRecs As Long
  Dim PayTranRec As TaxTransactionType
  Dim PayTranHandle As Integer
  Dim NumOfPayTranRecs As Long
  Dim EmptyPay As TaxTransactionType
  Dim cnt&, TotalPaid#
  Dim ThisListRec&
  Dim NextTransRec&
  Dim One As Integer, x As Integer
  Dim AHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  Oper$ = QPTrim$(Str$(OperNum))
  If Not Exist("TAXPCPR" + Oper$ + ".DAT") Then Exit Sub
  OpenTempPersPayFile PHandle, OperNum ' is the same as open TaxCPRFileName
  NumOfPRecs = LOF(PHandle) / Len(TaxPaymentRec)
  
  OpenPersPayListFile LHandle, OperNum 'is the same as open TaxLOPFileName
  NumOfLRecs = LOF(LHandle) / Len(PayListRec)
  
  OpenTaxCustFile CHandle, NumOfCRecs
  OpenTaxTransFile TaxTranHandle, NumOfTaxTranRecs
  
  For cnt& = 1 To NumOfPRecs 'NumOfPRecs = number of customers paying
    Get PHandle, cnt&, TaxPaymentRec
    ThisListRec& = TaxPaymentRec.LastPayRec
    Do While ThisListRec& > 0
      Get LHandle, ThisListRec&, PayListRec 'listrec is the list of bills being paid (some could be
      'multiple tags for a single customer
      'get paylist rec
      Get CHandle, TaxPaymentRec.CustAcct, TaxCustRec
      'get cust rec
      
      If PayListRec.BillRec < 0 Then
        GoSub PrePay
        GoTo SkipThisRec
      End If
      
      Get TaxTranHandle, PayListRec.BillRec, TaxTranRec
      'get bill trans this payrec is for
      TotalPaid# = 0
      PayTranRec = EmptyPay
      'make a new clean payment trans
'      TotalPaid# = OldRound#(PayListRec.DiscAmt + PayListRec.Principle1 + PayListRec.Interest1 + PayListRec.Collection + PayListRec.LateList)
      TotalPaid# = OldRound#(PayListRec.Personal + PayListRec.MachTools + PayListRec.MerchCap + PayListRec.FarmEquip) '4/22/05
      TotalPaid# = OldRound(TotalPaid# + PayListRec.MobHomes + PayListRec.Interest + PayListRec.Penalty + PayListRec.PrePayAmt)
      TotalPaid# = OldRound(TotalPaid# + PayListRec.Opt1 + PayListRec.Opt2 + PayListRec.Opt3) 'added 11/6/06
      If TotalPaid# = 0 Then
        GoTo SkipThisRec
      End If
      'PayTranRec = the new record for tax transaction records
      PayTranRec.TransDate = TaxPaymentRec.PayDate
      If PayListRec.PrePayAmt > 0 Then
        PayTranRec.TranType = 21 'overpay and bill pay combined
      Else
        PayTranRec.TranType = 2
      End If
'      PayTranRec.Revenue.Principle1Pd = OldRound(PayListRec.Principle1 + PayListRec.DiscAmt)
      PayTranRec.Revenue.Principle1Pd = PayListRec.Personal '4/22/05
      PayTranRec.Revenue.Principle2Pd = PayListRec.MachTools
      PayTranRec.Revenue.Principle3Pd = PayListRec.MerchCap
      PayTranRec.Revenue.Principle4Pd = PayListRec.FarmEquip
      PayTranRec.Revenue.Principle5Pd = PayListRec.MobHomes
      PayTranRec.Revenue.InterestPd = PayListRec.Interest
      PayTranRec.Revenue.PenaltyPd = PayListRec.Penalty
      PayTranRec.Revenue.CollectionPd = 0
      PayTranRec.Revenue.LateListPd = 0
      PayTranRec.Revenue.RevOpt1Pd = PayListRec.Opt1
      PayTranRec.Revenue.RevOpt2Pd = PayListRec.Opt2
      PayTranRec.Revenue.RevOpt3Pd = PayListRec.Opt3
      PayTranRec.CustPin = TaxCustRec.PIN
      PayTranRec.DiscXDate = TaxTranRec.DiscXDate
      PayTranRec.RealPin = ""
      PayTranRec.PersPin = QPTrim$(TaxTranRec.PersPin)
      PayTranRec.Posted2GL = "N"
      PayTranRec.TaxYear = TaxTranRec.TaxYear
      PayTranRec.DiscAmt = PayListRec.DiscAmt
      PayTranRec.OperNum = OperNum
      PayTranRec.Amount = TotalPaid#
      If QPTrim$(PayListRec.Description) = "" Then
        PayTranRec.Description = TaxTranRec.Description
      Else
        PayTranRec.Description = QPTrim$(PayListRec.Description) + " " + ParseBillNum$(TaxTranRec.Description)
      End If
      PayTranRec.CustomerRec = TaxPaymentRec.CustAcct
      PayTranRec.LastTrans = TaxCustRec.LastTrans
      PayTranRec.BelongTo = PayListRec.BillRec
      PayTranRec.Revenue.PrePaidAmt = PayListRec.PrePayAmt
      PayTranRec.Revenue.PrePaidUsed = 0
      PayTranRec.Revenue.PrePaidBal = OldRound(GetOverPayBalance(TaxPaymentRec.CustAcct, "N") + PayTranRec.Revenue.PrePaidAmt)
      PayTranRec.InternalPin = TaxTranRec.InternalPin
      PayTranRec.BillType = TaxPaymentRec.BillType
      'TaxTranRec is the update to the existing tax record
'      TaxTranRec.Revenue.Principle1Pd = OldRound#(TaxTranRec.Revenue.Principle1Pd + PayListRec.Principle1 + PayListRec.DiscAmt)
      TaxTranRec.Revenue.Principle1Pd = OldRound#(TaxTranRec.Revenue.Principle1Pd + PayListRec.Personal) '4/22/05
      TaxTranRec.Revenue.Principle2Pd = OldRound#(TaxTranRec.Revenue.Principle2Pd + PayListRec.MachTools)
      TaxTranRec.Revenue.Principle3Pd = OldRound#(TaxTranRec.Revenue.Principle3Pd + PayListRec.MerchCap)
      TaxTranRec.Revenue.Principle4Pd = OldRound#(TaxTranRec.Revenue.Principle4Pd + PayListRec.FarmEquip)
      TaxTranRec.Revenue.Principle5Pd = OldRound#(TaxTranRec.Revenue.Principle5Pd + PayListRec.MobHomes)
      TaxTranRec.Revenue.InterestPd = OldRound#(TaxTranRec.Revenue.InterestPd + PayListRec.Interest)
      TaxTranRec.Revenue.PenaltyPd = OldRound#(TaxTranRec.Revenue.PenaltyPd + PayListRec.Penalty)
      TaxTranRec.Revenue.Future1Pd = OldRound#(TaxTranRec.DiscAmt + PayListRec.DiscAmt)
      
      TaxTranRec.DiscAmt = OldRound#(TaxTranRec.DiscAmt + PayListRec.DiscAmt)
      TaxTranRec.Revenue.RevOpt1Pd = OldRound#(TaxTranRec.Revenue.RevOpt1Pd + PayListRec.Opt1)
      TaxTranRec.Revenue.RevOpt2Pd = OldRound#(TaxTranRec.Revenue.RevOpt2Pd + PayListRec.Opt2)
      TaxTranRec.Revenue.RevOpt3Pd = OldRound#(TaxTranRec.Revenue.RevOpt3Pd + PayListRec.Opt3)
      Put TaxTranHandle, PayListRec.BillRec, TaxTranRec
      NextTransRec& = (LOF(TaxTranHandle) \ Len(TaxTranRec)) + 1

      Put TaxTranHandle, NextTransRec&, PayTranRec
      TaxCustRec.LastTrans = NextTransRec&
      Put CHandle, TaxPaymentRec.CustAcct, TaxCustRec

SkipThisRec:
      ThisListRec& = PayListRec.PrevListRec
    Loop
  Next

  Close

  For x = 1 To 5
    KillFile ("TAXPCPR" + Oper$ + ".DAT")
    KillFile ("TAXPLOP" + Oper$ + ".DAT")
  Next x
  
  If Exist("TAXPCPR" + Oper$ + ".DAT") = True Or Exist("TAXPLOP" + Oper$ + ".DAT") = True Then '11/30/06
    One = 1
    AHandle = FreeFile
    Open "ppayposterror" + Oper$ + ".dat" For Output As AHandle
    Print #AHandle, One
    Close AHandle
    Call TaxMsg(800, "ERROR: POSTING WAS SUCCESSFUL but temporary payment files were not deleted properly. Please call Southern Software @ 1-800-842-8190 for assistance.")
    MainLog ("ERROR: A personal payment error has taken place because payment batch files were not deleted. User was notified.")
  Else
    Call Savemsg(900, "Personal payment posting has completed successfully.")
    MainLog ("Personal payment post completed successfully.")
  End If
  
  Call cmdExit_Click
  
  Exit Sub

PrePay:
  TotalPaid# = 0
  PayTranRec = EmptyPay
  'make a new clean payment trans
  TotalPaid# = OldRound#(PayListRec.DiscAmt + PayListRec.Personal + PayListRec.MachTools + PayListRec.MerchCap + PayListRec.FarmEquip)
  TotalPaid# = OldRound(TotalPaid# + PayListRec.MobHomes + PayListRec.Interest + PayListRec.Penalty + PayListRec.PrePayAmt)
  If TotalPaid# = 0 Then
    GoTo SkipThisRec
  End If
  'PayTranRec = the new record for tax transaction records
  PayTranRec.TransDate = TaxPaymentRec.PayDate
  PayTranRec.TranType = 22 'overpay only
  PayTranRec.Revenue.Principle1Pd = OldRound(PayListRec.Personal + PayListRec.DiscAmt)
  PayTranRec.Revenue.Principle2Pd = PayListRec.MachTools
  PayTranRec.Revenue.Principle3Pd = PayListRec.MerchCap
  PayTranRec.Revenue.Principle4Pd = PayListRec.FarmEquip
  PayTranRec.Revenue.Principle5Pd = PayListRec.MobHomes
  PayTranRec.Revenue.InterestPd = PayListRec.Interest
  PayTranRec.Revenue.PenaltyPd = PayListRec.Penalty
  PayTranRec.Revenue.CollectionPd = 0
  PayTranRec.Revenue.LateListPd = 0
  PayTranRec.Revenue.RevOpt1Pd = 0
  PayTranRec.Revenue.RevOpt2Pd = 0
  PayTranRec.Revenue.RevOpt3Pd = 0
  PayTranRec.CustPin = TaxCustRec.PIN
  PayTranRec.DiscXDate = TaxTranRec.DiscXDate
  PayTranRec.RealPin = ""
  PayTranRec.PersPin = QPTrim$(TaxTranRec.PersPin)
  PayTranRec.Posted2GL = "N"
  PayTranRec.TaxYear = TaxYear
  PayTranRec.DiscAmt = PayListRec.DiscAmt
  PayTranRec.OperNum = OperNum
  PayTranRec.Amount = TotalPaid#
  If QPTrim$(PayListRec.Description) = "" Then
    PayTranRec.Description = "Prepay"
  Else
    PayTranRec.Description = QPTrim$(PayListRec.Description)
  End If
  PayTranRec.CustomerRec = TaxPaymentRec.CustAcct
  PayTranRec.LastTrans = TaxCustRec.LastTrans
  PayTranRec.BelongTo = 0
  PayTranRec.Revenue.PrePaidAmt = PayListRec.PrePayAmt
  PayTranRec.Revenue.PrePaidUsed = 0
  PayTranRec.Revenue.PrePaidBal = OldRound(GetOverPayBalance(TaxPaymentRec.CustAcct, "N") + PayTranRec.Revenue.PrePaidAmt)
  PayTranRec.BillType = TaxPaymentRec.BillType
  NextTransRec& = (LOF(TaxTranHandle) \ Len(TaxTranRec)) + 1
  Put TaxTranHandle, NextTransRec&, PayTranRec
  
  TaxCustRec.LastTrans = NextTransRec&
  Put CHandle, TaxPaymentRec.CustAcct, TaxCustRec

  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayPost", "cmdPostPers_Click", Erl)
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
Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxPayPost.")
      Call Terminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
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
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdPost_Click
      KeyCode = 0
  End Select

End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim THandle As Integer
  Dim PayRec As TaxPaymentRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer
  Dim TownName$
  
  On Error GoTo ERRORSTUFF
  
  fptxtOperator.Text = "Operator #" + CStr(OperNum)
  OpenTaxSetUpFile THandle
  Get THandle, 1, TaxMasterRec
  Close THandle
  
  TownName = QPTrim$(TaxMasterRec.Name)
  TaxYear = TaxMasterRec.RTaxYear
  
  Close
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayOperEntry", "LoadMe", Erl)
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
