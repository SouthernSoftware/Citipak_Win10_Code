Attribute VB_Name = "NCTX_Common"
Option Explicit
Public Const NCTAXData = "NCTAXData\"

Public Const NCTaxPropFile = "NCTAXData\TAXPROP.DAT"
Public Const NCTaxPersFile = "NCTAXData\TAXPERS.DAT"
Public Const NCTaxCustFile = "NCTAXData\TAXCUST.DAT"
Public Const NCMortCodeName = "NCTAXData\TAXMORT.DAT"
Public Const NCPersOptSearch = "NCTAXData\TXPOPTSH.DAT"
Public Const NCRealOptSearch = "NCTAXData\TXROPTSH.DAT"
Public Const NCCustOptSearch = "NCTAXData\TXCOPTSH.DAT"
Public Const NCTaxTownships = "NCTAXData\TXTWNSHP.DAT"
Public Const NCTaxSetupName = "NCTAXData\TAXSETUP.DAT"
Public Const NCLaserPersItemBill = "NCTAXData\LSRPITEM.DAT"
Public Const NCLaserRealItemBill = "NCTAXData\LSRRITEM.DAT"
Public Const NCTMessageName = "NCTAXData\TAXMESS.DAT"
Public Const NCTaxBillRealName = "NCTAXData\TAXBILRLSR.DAT"
Public Const NCTaxBillPersName = "NCTAXData\TAXBILPLSR.DAT"
Public Const NCLateLtrText = "NCTAXData\TXLATLTR.DAT"
Public Const NCTaxRateTableFile = "NCTAXData\TXRTTBLS.DAT"
Public Const NCTxGLInterBill = "NCTAXData\TAXGLBAC.DAT"
Public Const NCTxGLInterPay = "NCTAXData\TAXGLACT.DAT"
Public Const NCTaxTransFile = "NCTAXData\TAXTRANS.DAT"
Public Const NCTaxBillFile = "NCTAXData\TAXTBILL.DAT"
Public Const NCTaxBill1Name = "NCTaxData\TaxBil1.DAT"

Public Sub NCOpenTaxCustFile(TaxCustHandle As Integer, NumOfTaxCustRec As Long)
  Dim TaxCustLen As Integer
  Dim TaxCustRec As TaxCustType
  TaxCustLen = Len(TaxCustRec)
  TaxCustHandle = FreeFile
  Open NCTaxCustFile For Random Shared As TaxCustHandle Len = TaxCustLen
  NumOfTaxCustRec = LOF(TaxCustHandle) / Len(TaxCustRec)
End Sub

Public Sub NCOpenRealPropFile(RealPropHandle As Integer, NumOfRealProp As Long)
  Dim RealPropLen As Integer
  Dim RealPropRec As NCPropertyRecType
  RealPropLen = Len(RealPropRec)
  RealPropHandle = FreeFile
  Open NCTaxPropFile For Random Shared As RealPropHandle Len = RealPropLen
  NumOfRealProp = LOF(RealPropHandle) / Len(RealPropRec)
End Sub

Public Sub NCOpenPersPropFile(PersPropHandle As Integer, NumOfPersProp As Long)
  Dim PersPropLen As Integer
  Dim PersPropRec As NCPersonalRecType
  PersPropLen = Len(PersPropRec)
  PersPropHandle = FreeFile
  Open NCTaxPersFile For Random Shared As PersPropHandle Len = PersPropLen
  NumOfPersProp = LOF(PersPropHandle) / Len(PersPropRec)
End Sub

Public Sub NCOpenMortCodeFile(MortCodeHandle As Integer, NumOfMortCodes As Integer)
  Dim MortCodeLen As Integer
  Dim MortCodeRec As MortCodeRecType
  MortCodeLen = Len(MortCodeRec)
  MortCodeHandle = FreeFile
  Open NCMortCodeName For Random Shared As MortCodeHandle Len = Len(MortCodeRec)
  NumOfMortCodes = LOF(MortCodeHandle) / Len(MortCodeRec)
End Sub

Public Sub NCOpenPersOptSearchFile(POSHandle As Integer, NumOfPOSFiles As Long)
  Dim POSRecLen As Integer
  Dim POSRec As OptPersIdxType
  POSRecLen = Len(POSRec)
  POSHandle = FreeFile
  Open NCPersOptSearch For Random Shared As POSHandle Len = POSRecLen
  NumOfPOSFiles = LOF(POSHandle) / POSRecLen
End Sub

Public Sub NCOpenRealOptSearchFile(ROSHandle As Integer, NumOfROSFiles As Long)
  Dim ROSRecLen As Integer
  Dim ROSRec As NCOptRealIdxType
  ROSRecLen = Len(ROSRec)
  ROSHandle = FreeFile
  Open NCRealOptSearch For Random Shared As ROSHandle Len = ROSRecLen
  NumOfROSFiles = LOF(ROSHandle) / ROSRecLen
End Sub

Public Sub NCOpenCustOptSearchFile(COSHandle As Integer, NumOfCOSFiles As Long)
  Dim COSRecLen As Integer
  Dim COSRec As OptCustIdxType
  COSRecLen = Len(COSRec)
  COSHandle = FreeFile
  Open NCCustOptSearch For Random Shared As COSHandle Len = COSRecLen
  NumOfCOSFiles = LOF(COSHandle) / COSRecLen
End Sub

Public Sub NCOpenTownshipFile(TownshipHandle As Integer, NumOfTownships As Integer)
  Dim TownshipLen As Integer
  Dim TownshipRec As TownshipType
  TownshipLen = Len(TownshipRec)
  TownshipHandle = FreeFile
  Open NCTaxTownships For Random Shared As TownshipHandle Len = TownshipLen
  NumOfTownships = LOF(TownshipHandle) / Len(TownshipRec)
End Sub

Public Sub NCOpenTaxSetUpFile(TaxSetUpHandle As Integer)
  Dim TaxSetUpLen As Integer
  Dim TaxSetUp As TaxMasterType
  TaxSetUpLen = Len(TaxSetUp)
  TaxSetUpHandle = FreeFile
  Open NCTaxSetupName For Random Shared As TaxSetUpHandle Len = TaxSetUpLen
End Sub
Public Sub NCOpenTaxMessage(MessHandle As Integer, MsgCnt As Integer)
  Dim MessLen As Integer
  Dim MessRec As TaxMessRecType
  MessLen = Len(MessRec)
  MessHandle = FreeFile
  Open NCTMessageName For Random Shared As MessHandle Len = MessLen
  MsgCnt = LOF(MessHandle) / Len(MessRec)
End Sub
Public Sub NCOpenTaxBillFile(TaxBillHandle As Integer, NumOfTaxBills As Long)
  Dim TaxBillLen As Integer
  Dim TaxBillRec As NCTaxBillType
  TaxBillLen = Len(TaxBillRec)
  TaxBillHandle = FreeFile
  Open NCTaxBillFile For Random Shared As TaxBillHandle Len = Len(TaxBillRec)
  NumOfTaxBills = LOF(TaxBillHandle) / Len(TaxBillRec)
End Sub

Public Sub NCOpenLateLtrFile(LateHandle As Integer)
  Dim LateRecLen As Integer
  Dim LateRec As TAXLateLetterType
  LateRecLen = Len(LateRec)
  LateHandle = FreeFile
  Open NCLateLtrText For Random Shared As LateHandle Len = LateRecLen
End Sub

Public Sub NCOpenTaxRateTables(RateTablesHandle As Integer, RateTablesCnt As Integer)
  Dim RateTablesLen As Integer
  Dim RateTablesRec As NCOptRevRateTablesType
  RateTablesLen = Len(RateTablesRec)
  RateTablesHandle = FreeFile
  Open NCTaxRateTableFile For Random Shared As RateTablesHandle Len = RateTablesLen
  RateTablesCnt = LOF(RateTablesHandle) / Len(RateTablesRec)
End Sub

Public Sub NCOpenTaxGLInterPay(TaxGLIntHandle As Integer)
  Dim TaxGLIntLen As Integer
  Dim TaxGLIntRec As NCTaxAcctsType
  TaxGLIntLen = Len(TaxGLIntRec)
  TaxGLIntHandle = FreeFile
  Open NCTxGLInterPay For Random Shared As TaxGLIntHandle Len = Len(TaxGLIntRec)
End Sub
Public Sub NCOpenTaxGLInterBill(TaxGLIntHandle As Integer)
  Dim TaxGLIntLen As Integer
  Dim TaxGLIntRec As NCTaxAcctsType
  TaxGLIntLen = Len(TaxGLIntRec)
  TaxGLIntHandle = FreeFile
  Open NCTxGLInterBill For Random Shared As TaxGLIntHandle Len = Len(TaxGLIntRec)
End Sub

Public Sub NCOpenTaxTransFile(TaxTransHandle As Integer, NumOfTaxTransRecs As Long)
  Dim TaxTransLen As Integer
  Dim TaxTransRate As NCTaxTransactionType
  TaxTransLen = Len(TaxTransRate)
  TaxTransHandle = FreeFile
  Open NCTaxTransFile For Random Shared As TaxTransHandle Len = TaxTransLen
  NumOfTaxTransRecs = LOF(TaxTransHandle) / Len(TaxTransRate)
End Sub

Public Sub NCOpenTxBill1File(TxBill1Handle As Integer)
  Dim TxBill1RecLen As Integer
  Dim TxBill1Rec As NCTxBill1DefaultsType
  TxBill1RecLen = Len(TxBill1Rec)
  TxBill1Handle = FreeFile
  Open NCTaxBill1Name For Random Shared As TxBill1Handle Len = TxBill1RecLen
End Sub

Public Function NCGetCustBalance(RecNo&, TaxYear As Integer) As Double
  Dim TaxTran As NCTaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCustRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim PrevTranRec&
  Dim GTOwed#
  Dim TPaid#
  Dim GTPaid#
  
  If RecNo = 0 Then
    NCGetCustBalance = 0
    Exit Function
  End If
'  If RecNo = 16 Then Stop
  NCOpenTaxCustFile CHandle, NumOfCRecs
  Get CHandle, RecNo, TaxCustRec
  Close CHandle

  NCOpenTaxTransFile THandle, NumOfTRecs
  Get THandle, 1, TaxTran
  
  PrevTranRec& = TaxCustRec.LastTrans
  GTOwed = 0
  TPaid = 0
  GTPaid = 0
  Dim Cnt As Integer
'  TaxYear = 2005
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get THandle, PrevTranRec&, TaxTran
      TaxTran.OperNum = TaxTran.OperNum
      If TaxTran.TaxYear = TaxYear Then GoTo MoveAlong 'if we only want to get
      'the balance for all old bills then by entering the current tax year
      'we can send just that data
'      If TaxTran.BelongTo = 16892 And TaxTran.TranType = 2 Then Stop
'      Debug.Print CStr(PrevTranRec)
'      If PrevTranRec& = 1650 Then Stop
      Select Case TaxTran.TranType
      Case 1    'bill
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 2    'payment
'        TPaid# = OldRound#(TPaid# + TaxTran.Amount)
'        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        TaxTran.Revenue.Principle1Pd = TaxTran.Revenue.Principle1Pd
        TaxTran.Revenue.InterestPd = TaxTran.Revenue.InterestPd
        TaxTran.Revenue.CollectionPd = TaxTran.Revenue.CollectionPd
        TaxTran.Revenue.LateListPd = TaxTran.Revenue.LateListPd
        TaxTran.Revenue.Principle1 = TaxTran.Revenue.Principle1
        TaxTran.Revenue.Interest = TaxTran.Revenue.Interest
        TaxTran.Revenue.Collection = TaxTran.Revenue.Collection
        TaxTran.Revenue.LateList = TaxTran.Revenue.LateList
        TaxTran.Revenue.PrePaidAmt = TaxTran.Revenue.PrePaidAmt
        TaxTran.Revenue.PrePaidUsed = TaxTran.Revenue.PrePaidUsed
       
        TaxTran.BelongTo = TaxTran.BelongTo
        TaxTran.TaxYear = TaxTran.TaxYear
        TaxTran.CustomerRec = TaxTran.CustomerRec
      Case 3    'release
        GTOwed# = OldRound#(GTOwed# - TaxTran.Amount)
      Case 4    'interest
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 6    'collect/add cost
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 7    'adjust paid down
        If TaxTran.CustPin = 0 Then
          TPaid# = OldRound#(TPaid# + TaxTran.Amount)
          GTPaid# = OldRound#(GTPaid# + TaxTran.Amount)
        Else
          TPaid# = OldRound#(TPaid# - TaxTran.Amount)
          GTPaid# = OldRound#(GTPaid# - TaxTran.Amount)
        End If
      Case 8    'misc cost
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 9    'credit applied at billing
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 13 'adjust bill down
        GTOwed# = OldRound(GTOwed# - TaxTran.Amount)
      Case 14 'adjust bill up
        GTOwed# = OldRound(GTOwed# + TaxTran.Amount)
      Case 21    'payment plus overpayment
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 22    'overpayment only
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
      Case 10    'adjust pay down affecting credit balance
        TPaid# = OldRound(TPaid - TaxTran.Amount)
        GTPaid# = OldRound(GTPaid - TaxTran.Amount)
      Case 11    'adjust prepay down
        TPaid# = OldRound(TPaid - TaxTran.Amount)
        GTPaid# = OldRound(GTPaid - TaxTran.Amount)
      Case 12    'refund total prepay balance
        TPaid# = OldRound(TPaid - TaxTran.Amount)
        GTPaid# = OldRound(GTPaid - TaxTran.Amount)
      Case 24    'adjust bill up affecting credit balance
        GTOwed# = OldRound(GTOwed# + TaxTran.Amount)
      Case Else
'        BillType$ = "?????"
      End Select
MoveAlong:
      PrevTranRec& = TaxTran.LastTrans
    Loop

    NCGetCustBalance# = OldRound#(GTOwed# - GTPaid#)
  Else
    NCGetCustBalance# = 0
  End If

  Close THandle

End Function


