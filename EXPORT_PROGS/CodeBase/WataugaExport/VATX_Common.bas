Attribute VB_Name = "VATX_Common"
Option Explicit
Public Const VATAXData = "vataxdata\"
Public Const TaxCustFile = "VATAXData\TAXCUST.DAT"
Public Const TaxPropFile = "VATAXData\TAXPROP.DAT"
Public Const TaxPersFile = "VATAXData\TAXPERS.DAT"
Public Const MortCodeName = "VATAXData\TAXMORT.DAT"
Public Const PersOptSearch = "VATAXData\TXPOPTSH.DAT"
Public Const RealOptSearch = "VATAXData\TXROPTSH.DAT"
Public Const CustOptSearch = "VATAXData\TXCOPTSH.DAT"
Public Const TaxTownships = "VATAXData\TXTWNSHP.DAT"
Public Const TaxSetupName = "VATAXData\TAXSETUP.DAT"
Public Const LaserPersItemBill = "VATAXData\LSRPITEM.DAT"
Public Const LaserRealItemBill = "VATAXData\LSRRITEM.DAT"
Public Const TMessageName = "VATAXData\TAXMESS.DAT"
Public Const TaxBillRealName = "VATAXData\TAXBILRLSR.DAT"
Public Const TaxBillPersName = "VATAXData\TAXBILPLSR.DAT"
Public Const LateLtrText = "VATAXData\TXLATLTR.DAT"
Public Const TaxRateTableFile = "VATAXData\TXRTTBLS.DAT"
Public Const TaxPenRateTblFile = "VATAXData\TXPENRTB.DAT"
Public Const TxRGLInterBill = "VATAXData\TAXRGLBAC.DAT"
Public Const TxPGLInterBill = "VATAXData\TAXPGLBAC.DAT"
Public Const TxRGLInterPay = "VATAXData\TAXRGLACT.DAT"
Public Const TxPGLInterPay = "VATAXData\TAXPGLACT.DAT"
Public Const TaxTransFile = "VATAXData\TAXTRANS.DAT"
Public Const TaxBillPostDateFile = "VATAXData\TXBLPSTDTE.DAT"
Public Sub OpenTaxCustFile(TaxCustHandle As Integer, NumOfTaxCustRec As Long)
  Dim TaxCustLen As Integer
  Dim TaxCustRec As TaxCustType
  TaxCustLen = Len(TaxCustRec)
  TaxCustHandle = FreeFile
  Open TaxCustFile For Random Shared As TaxCustHandle Len = TaxCustLen
  NumOfTaxCustRec = LOF(TaxCustHandle) / Len(TaxCustRec)
End Sub

Public Sub OpenRealPropFile(RealPropHandle As Integer, NumOfRealProp As Long)
  Dim RealPropLen As Integer
  Dim RealPropRec As PropertyRecType
  RealPropLen = Len(RealPropRec)
  RealPropHandle = FreeFile
  Open TaxPropFile For Random Shared As RealPropHandle Len = RealPropLen
  NumOfRealProp = LOF(RealPropHandle) / Len(RealPropRec)
End Sub

Public Sub OpenPersPropFile(PersPropHandle As Integer, NumOfPersProp As Long)
  Dim PersPropLen As Integer
  Dim PersPropRec As PersonalRecType
  PersPropLen = Len(PersPropRec)
  PersPropHandle = FreeFile
  Open TaxPersFile For Random Shared As PersPropHandle Len = PersPropLen
  NumOfPersProp = LOF(PersPropHandle) / Len(PersPropRec)
End Sub

Public Sub OpenMortCodeFile(MortCodeHandle As Integer, NumOfMortCodes As Integer)
  Dim MortCodeLen As Integer
  Dim MortCodeRec As MortCodeRecType
  MortCodeLen = Len(MortCodeRec)
  MortCodeHandle = FreeFile
  Open MortCodeName For Random Shared As MortCodeHandle Len = Len(MortCodeRec)
  NumOfMortCodes = LOF(MortCodeHandle) / Len(MortCodeRec)
End Sub

Public Sub OpenPersOptSearchFile(POSHandle As Integer, NumOfPOSFiles As Long)
  Dim POSRecLen As Integer
  Dim POSRec As OptPersIdxType
  POSRecLen = Len(POSRec)
  POSHandle = FreeFile
  Open PersOptSearch For Random Shared As POSHandle Len = POSRecLen
  NumOfPOSFiles = LOF(POSHandle) / POSRecLen
End Sub

Public Sub OpenRealOptSearchFile(ROSHandle As Integer, NumOfROSFiles As Long)
  Dim ROSRecLen As Integer
  Dim ROSRec As OptRealIdxType
  ROSRecLen = Len(ROSRec)
  ROSHandle = FreeFile
  Open RealOptSearch For Random Shared As ROSHandle Len = ROSRecLen
  NumOfROSFiles = LOF(ROSHandle) / ROSRecLen
End Sub

Public Sub OpenCustOptSearchFile(COSHandle As Integer, NumOfCOSFiles As Long)
  Dim COSRecLen As Integer
  Dim COSRec As OptCustIdxType
  COSRecLen = Len(COSRec)
  COSHandle = FreeFile
  Open CustOptSearch For Random Shared As COSHandle Len = COSRecLen
  NumOfCOSFiles = LOF(COSHandle) / COSRecLen
End Sub

Public Sub OpenTownshipFile(TownshipHandle As Integer, NumOfTownships As Integer)
  Dim TownshipLen As Integer
  Dim TownshipRec As TownshipType
  TownshipLen = Len(TownshipRec)
  TownshipHandle = FreeFile
  Open TaxTownships For Random Shared As TownshipHandle Len = TownshipLen
  NumOfTownships = LOF(TownshipHandle) / Len(TownshipRec)
End Sub

Public Sub OpenTaxSetUpFile(TaxSetUpHandle As Integer)
  Dim TaxSetUpLen As Integer
  Dim TaxSetUp As TaxMasterType
  TaxSetUpLen = Len(TaxSetUp)
  TaxSetUpHandle = FreeFile
  Open TaxSetupName For Random Shared As TaxSetUpHandle Len = TaxSetUpLen
End Sub

Public Sub OpenLaserRealItemized(LsrItmHandle As Integer)
  Dim LsrItmLen As Integer
  Dim LsrItmRec As TxBillLaserItemized
  LsrItmLen = Len(LsrItmRec)
  LsrItmHandle = FreeFile
  Open LaserRealItemBill For Random Shared As LsrItmHandle Len = LsrItmLen
End Sub
Public Sub OpenLaserPersItemized(LsrItmHandle As Integer)
  Dim LsrItmLen As Integer
  Dim LsrItmRec As TxBillLaserItemized
  LsrItmLen = Len(LsrItmRec)
  LsrItmHandle = FreeFile
  Open LaserPersItemBill For Random Shared As LsrItmHandle Len = LsrItmLen
End Sub

Public Sub OpenTaxMessage(MessHandle As Integer, MsgCnt As Integer)
  Dim MessLen As Integer
  Dim MessRec As TaxMessRecType
  MessLen = Len(MessRec)
  MessHandle = FreeFile
  Open TMessageName For Random Shared As MessHandle Len = MessLen
  MsgCnt = LOF(MessHandle) / Len(MessRec)
End Sub

Public Sub OpenTxBillPersFile(TxBill1Handle As Integer)
  Dim TxBill1RecLen As Integer
  Dim TxBill1Rec As TxBillLaser1DefaultsType
  TxBill1RecLen = Len(TxBill1Rec)
  TxBill1Handle = FreeFile
  Open TaxBillPersName For Random Shared As TxBill1Handle Len = TxBill1RecLen
End Sub
Public Sub OpenTxBillRealFile(TxBill1Handle As Integer)
  Dim TxBill1RecLen As Integer
  Dim TxBill1Rec As TxBillLaser1DefaultsType
  TxBill1RecLen = Len(TxBill1Rec)
  TxBill1Handle = FreeFile
  Open TaxBillRealName For Random Shared As TxBill1Handle Len = TxBill1RecLen
End Sub

Public Sub OpenLateLtrFile(LateHandle As Integer)
  Dim LateRecLen As Integer
  Dim LateRec As TAXLateLetterType
  LateRecLen = Len(LateRec)
  LateHandle = FreeFile
  Open LateLtrText For Random Shared As LateHandle Len = LateRecLen
End Sub

Public Sub OpenTaxRateTables(RateTablesHandle As Integer, RateTablesCnt As Integer)
  Dim RateTablesLen As Integer
  Dim RateTablesRec As OptRevRateTablesType
  RateTablesLen = Len(RateTablesRec)
  RateTablesHandle = FreeFile
  Open TaxRateTableFile For Random Shared As RateTablesHandle Len = RateTablesLen
  RateTablesCnt = LOF(RateTablesHandle) / Len(RateTablesRec)
End Sub

Public Sub OpenTaxPenRateTbls(PRateHandle As Integer, NumOfPRRecs As Integer)
  Dim PRateLen As Integer
  Dim PRateRec As PenaltyRateTablesType
  PRateLen = Len(PRateRec)
  PRateHandle = FreeFile
  Open TaxPenRateTblFile For Random Shared As PRateHandle Len = PRateLen
  NumOfPRRecs = LOF(PRateHandle) / PRateLen
End Sub

Public Sub OpenRTaxGLInterBill(TaxGLIntHandle As Integer)
  Dim TaxGLIntLen As Integer
  Dim TaxGLIntRec As TaxRAcctsType
  TaxGLIntLen = Len(TaxGLIntRec)
  TaxGLIntHandle = FreeFile
  Open TxRGLInterBill For Random Shared As TaxGLIntHandle Len = Len(TaxGLIntRec)
End Sub

Public Sub OpenPTaxGLInterBill(TaxGLIntHandle As Integer)
  Dim TaxGLIntLen As Integer
  Dim TaxGLIntRec As TaxPAcctsType
  TaxGLIntLen = Len(TaxGLIntRec)
  TaxGLIntHandle = FreeFile
  Open TxPGLInterBill For Random Shared As TaxGLIntHandle Len = Len(TaxGLIntRec)
End Sub

Public Sub OpenRTaxGLInterPay(TaxGLIntHandle As Integer)
  Dim TaxGLIntLen As Integer
  Dim TaxGLIntRec As TaxRAcctsType
  TaxGLIntLen = Len(TaxGLIntRec)
  TaxGLIntHandle = FreeFile
  Open TxRGLInterPay For Random Shared As TaxGLIntHandle Len = Len(TaxGLIntRec)
End Sub

Public Sub OpenPTaxGLInterPay(TaxGLIntHandle As Integer)
  Dim TaxGLIntLen As Integer
  Dim TaxGLIntRec As TaxPAcctsType
  TaxGLIntLen = Len(TaxGLIntRec)
  TaxGLIntHandle = FreeFile
  Open TxPGLInterPay For Random Shared As TaxGLIntHandle Len = Len(TaxGLIntRec)
End Sub

Public Sub OpenTaxTransFile(TaxTransHandle As Integer, NumOfTaxTransRecs As Long)
  Dim TaxTransLen As Integer
  Dim TaxTransRate As TaxTransactionType
  TaxTransLen = Len(TaxTransRate)
  TaxTransHandle = FreeFile
  Open TaxTransFile For Random Shared As TaxTransHandle Len = TaxTransLen
  NumOfTaxTransRecs = LOF(TaxTransHandle) / Len(TaxTransRate)
End Sub
Public Sub OpenPersPostedReprintFile(PRHandle As Integer, NumOfPRRecs As Long, ThisName$)
  Dim PRRecLen As Integer
  Dim PRRec As VAPPTaxBillType
  PRRecLen = Len(PRRec)
  PRHandle = FreeFile
  Open ThisName For Random Shared As PRHandle Len = PRRecLen
  NumOfPRRecs = LOF(PRHandle) / PRRecLen
End Sub
Public Sub OpenRealPostedReprintFile(PRHandle As Integer, NumOfPRRecs As Long, ThisName$)
  Dim PRRecLen As Integer
  Dim PRRec As VARETaxBillType
  PRRecLen = Len(PRRec)
  PRHandle = FreeFile
  Open ThisName For Random Shared As PRHandle Len = PRRecLen
  NumOfPRRecs = LOF(PRHandle) / PRRecLen
End Sub
Public Sub OpenBillPostDateFile(BillPostDateHandle As Integer, NumOfBillPostDateFiles As Long)
  Dim BillPostDateLen As Integer
  Dim BillPostDateRec As TaxBillPostDateType
  BillPostDateLen = Len(BillPostDateRec)
  BillPostDateHandle = FreeFile
  Open TaxBillPostDateFile For Random Shared As BillPostDateHandle Len = BillPostDateLen
  NumOfBillPostDateFiles = LOF(BillPostDateHandle) / BillPostDateLen
End Sub

Public Function ParseBillNum$(Text$)
  Dim BillNum$
  Dim BNumLen As Integer
  Dim thischar$
  Dim GoodPos As Integer
  Dim cnt As Integer
  
  BillNum$ = QPTrim$(Text$)
  BNumLen = Len(BillNum$)
  If BNumLen > 0 Then
    For cnt = BNumLen To 1 Step -1
      thischar$ = Mid$(BillNum$, cnt, 1)
      If InStr("0123456789", thischar$) <= 0 Then
        Exit For
      End If
    Next
    GoodPos = cnt + 1
    BillNum$ = Mid$(BillNum$, GoodPos)
  End If
  If Not IsNumeric(BillNum$) Then
    BillNum = "-911"
  End If
  ParseBillNum$ = BillNum$
End Function

Public Function GetCustBalance(RecNo&, TaxYear As Integer) As Double
  Dim TaxTran As TaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TaxCustRec As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim PrevTranRec&
  Dim GTOwed#
  Dim TPaid#
  Dim GTPaid#
  Dim ThisAmt$
  
  If RecNo = 0 Then
    GetCustBalance = 0
    Exit Function
  End If
  
  OpenTaxCustFile CHandle, NumOfCRecs
  Get CHandle, RecNo, TaxCustRec
  Close CHandle

  OpenTaxTransFile THandle, NumOfTRecs
 
  PrevTranRec& = TaxCustRec.LastTrans
  GTOwed = 0
  TPaid = 0
  GTPaid = 0
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get THandle, PrevTranRec&, TaxTran
      If TaxTran.TaxYear = TaxYear Then GoTo MoveAlong 'if we only want to get
      'the balance for all old bills then by entering the current tax year
      'we can send just that data
'      TaxTran.BelongTo = TaxTran.BelongTo
      If InStr(CStr(TaxTran.Amount), "E") Then TaxTran.Amount = 0
      If InStr(CStr(TaxTran.DiscAmt), "E") Then TaxTran.DiscAmt = 0
      Select Case TaxTran.TranType
      Case 1    'bill
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 2    'payment
        TPaid# = OldRound#(TPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        GTPaid# = OldRound#(GTPaid# + TaxTran.Amount + TaxTran.DiscAmt)
        TaxTran.Revenue.Principle1Pd = TaxTran.Revenue.Principle1Pd
        TaxTran.Revenue.Principle2Pd = TaxTran.Revenue.Principle2Pd
        TaxTran.Revenue.Principle3Pd = TaxTran.Revenue.Principle3Pd
        TaxTran.Revenue.Principle4Pd = TaxTran.Revenue.Principle4Pd
        TaxTran.Revenue.Principle5Pd = TaxTran.Revenue.Principle5Pd
        TaxTran.Revenue.PenaltyPd = TaxTran.Revenue.PenaltyPd
        TaxTran.Revenue.Future1Pd = TaxTran.Revenue.Future1Pd
        TaxTran.Revenue.Future2Pd = TaxTran.Revenue.Future2Pd
        TaxTran.Revenue.Future1 = TaxTran.Revenue.Future1
        TaxTran.Revenue.Future2 = TaxTran.Revenue.Future2
        TaxTran.Revenue.InterestPd = TaxTran.Revenue.InterestPd
        TaxTran.Revenue.CollectionPd = TaxTran.Revenue.CollectionPd
        TaxTran.Revenue.LateListPd = TaxTran.Revenue.LateListPd
        TaxTran.Revenue.Principle1 = TaxTran.Revenue.Principle1
        TaxTran.Revenue.Principle2 = TaxTran.Revenue.Principle2
        TaxTran.Revenue.Principle3 = TaxTran.Revenue.Principle3
        TaxTran.Revenue.Principle4 = TaxTran.Revenue.Principle4
        TaxTran.Revenue.Principle5 = TaxTran.Revenue.Principle5
        TaxTran.Revenue.Penalty = TaxTran.Revenue.Penalty
        TaxTran.Revenue.Interest = TaxTran.Revenue.Interest
        TaxTran.Revenue.Collection = TaxTran.Revenue.Collection
        TaxTran.Revenue.LateList = TaxTran.Revenue.LateList
        TaxTran.BelongTo = TaxTran.BelongTo
        TaxTran.TaxYear = TaxTran.TaxYear
        TaxTran.CustomerRec = TaxTran.CustomerRec
        TaxTran.LastTrans = TaxTran.LastTrans
      Case 3    'release
        GTOwed# = OldRound#(GTOwed# - TaxTran.Amount)
      Case 4    'interest
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
      Case 5    'penalty
        GTOwed# = OldRound#(GTOwed# + TaxTran.Amount)
        TaxTran.DiscAmt = TaxTran.DiscAmt
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
      Case 30    'PPTRA removal transaction
        GTOwed# = OldRound(GTOwed# + TaxTran.PPTRARmvl)
      Case Else
'        BillType$ = "?????"
      End Select
MoveAlong:
      PrevTranRec& = TaxTran.LastTrans
'      Debug.Print "Cust " + CStr(GTOwed#) + " " + CStr(GTPaid#) + " " + CStr(OldRound#(GTOwed# - GTPaid#))
    Loop

    GetCustBalance# = OldRound#(GTOwed# - GTPaid#)
  Else
    GetCustBalance# = 0
  End If

  Close THandle

End Function

