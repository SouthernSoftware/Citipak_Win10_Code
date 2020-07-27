DEFINT A-Z
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB OpenTaxPropFile (NumOfPropRecs%, PropTaxFile%)
DECLARE SUB OpenTaxPersFile (NumOfPersRecs%, PersTaxFile%)
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE SUB CustomerListing ()
DECLARE SUB TAXCustomerMenu ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ClearBack ()
DECLARE SUB SendDist2GL ()
DECLARE SUB ClearScrn ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Round# (b#)
  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB WaitForAction ()
  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
   '$INCLUDE: 'lctax.bi'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TBCustRec(1) AS TBRECust
  DIM SHARED TBRECust(1) AS TBRECust
  DIM SHARED TBRECustBal(1) AS TBREBAL
  DIM SHARED TBCustRec1(1) AS TBPPCust
  DIM SHARED TaxTran(1) AS TaxTransactionType

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
   TaxTranRecLen = LEN(TaxTran(1))
   OpenTaxPropFile NumOfPropRecs, PropTaxFile
   'OpenTaxPersFile NumOfPersRecs, PersTaxFile

    OPEN "TBRECUST.DAT" FOR RANDOM AS #11 LEN = LEN(TBCustRec(1))
    OPEN "TBREBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #12 LEN = LEN(TBRECustBal(1))
    TaxFile = FREEFILE
    OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
    NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))

   FOR x = 1 TO LOF(11) / LEN(TBCustRec(1))
   GET 11, x, TBCustRec(1)
   GET 11, x, TBRECust(1)
   GET 12, x, TBRECustBal(1)
   Record! = x

  'PersonalValue# = TBCustRec(1).mh + TBCustRec(1).fm + TBCustRec(1).ls + TBCustRec(1).op + TBCustRec(1).me
  'PersonalValue# = INT((PersonalValue# * 100) + .5) / 100

   RealValue# = TBCustRec(1).LandValue + TBCustRec(1).BldgValue
   RealValue# = INT((RealValue# * 100) + .5) / 100


'    OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
'    NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
    PRINT RTRIM$(TBCustRec(1).CustName)
    kk = INSTR(RTRIM$(TBCustRec(1).CustName), " ")
   IF kk > 0 THEN
     LNM$ = LEFT$(TBCustRec(1).CustName, kk - 1)
     FME$ = MID$(TBCustRec(1).CustName, kk, LEN(TBCustRec(1).CustName) - kk)
    ELSE
     LNM$ = RTRIM$(TBCustRec(1).CustName)
     FME$ = ""
   END IF
   TaxCustRec(1).FName = LTRIM$(FME$)
   TaxCustRec(1).LName = LTRIM$(LNM$)
   TaxCustRec(1).SName = LTRIM$(LNM$)
   TaxCustRec(1).acct = Record!
   TaxCustRec(1).CountyAcct = 0
   TaxCustRec(1).Addr1 = TBCustRec(1).Address1
   TaxCustRec(1).Addr2 = TBCustRec(1).Address2
   TaxCustRec(1).City = TBCustRec(1).Address3
   TaxCustRec(1).State = ""
   TaxCustRec(1).Zip = TBCustRec(1).Zipcode
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "Y"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = x
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("01-01-1999")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1

   GOSUB UpdateReal

   'Set Trans & Bills to Zero
    LastTrans& = 0
    BillTran& = 0

    IF TBRECust(1).CurYrTaxAmount > 0 THEN
     TaxTran = FREEFILE
     OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
     TaxTran(1).TransDate = Date2Num%("10-01-1998")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = "R"
     TaxTran(1).Amount = TBRECust(1).CurYrTaxAmount
     TaxTran(1).Revenue.Principle1 = TBRECust(1).CurYrTaxAmount
     TaxTran(1).Revenue.Interest = TBRECust(1).CurYrIntAmount
     TaxTran(1).Revenue.Penalty = TBRECust(1).CurYrPenAmount
     TaxTran(1).Description = "Tax Bill"
     TaxTran(1).Posted2Gl = "Y"
     TaxTran(1).CustomerRec = Record!
     TaxTran(1).LastTrans = 0
     TaxTran(1).BelongTo = 0
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     BillTran& = NextTran&
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, Record!, TaxCustRec(1)
      TaxCustRec(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, Record!, TaxCustRec(1)
   END IF

     'Check Interest
    IF TBRECust(1).CurYrIntAmount > 0 THEN
     TaxTran(1).TransDate = Date2Num%("10-01-1999")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = "R"
     TaxTran(1).Amount = TBRECust(1).CurYrIntAmount
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBRECust(1).CurYrIntAmount
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2Gl = "Y"
     TaxTran(1).CustomerRec = Record!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, Record!, TaxCustRec(1)
      TaxCustRec(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, Record!, TaxCustRec(1)
    END IF
     'Check Penalty
    IF TBRECust(1).CurYrPenAmount > 0 THEN
     TaxTran(1).TransDate = Date2Num%("12-08-1998")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = "R"
     TaxTran(1).Amount = TBRECust(1).CurYrPenAmount
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBRECust(1).CurYrPenAmount
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2Gl = "Y"
     TaxTran(1).CustomerRec = Record!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, Record!, TaxCustRec(1)
      TaxCustRec(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, Record!, TaxCustRec(1)
    END IF

     CLOSE TaxTran
    


  'Update Tax File if Past Bal >0

     TaxTran = FREEFILE
     OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen

    FOR ll = 1 TO 10
      ThisYear = (1998 - ll)
    IF TBRECustBal(1).TaxAmt(ll) THEN
     Yr$ = LTRIM$(STR$(ThisYear))
     TaxTran(1).TransDate = Date2Num%(("07-01-" + Yr$))
     TaxTran(1).TaxYear = ThisYear
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = "R"
     TaxTran(1).Amount = TBRECustBal(1).TaxAmt(ll)
     TaxTran(1).Revenue.Principle1 = TBRECustBal(1).TaxAmt(ll)
     TaxTran(1).Revenue.Interest = TBRECustBal(1).IntAmt(ll)
     TaxTran(1).Revenue.Penalty = TBRECustBal(1).PenAmt(ll)
     TaxTran(1).Description = "Tax Bill"
     TaxTran(1).Posted2Gl = "Y"
     TaxTran(1).CustomerRec = Record!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = 0
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     BillTran& = NextTran&
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, Record!, TaxCustRec(1)
      TaxCustRec(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, Record!, TaxCustRec(1)
    END IF
    IF TBRECustBal(1).IntAmt(ll) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("07-01-1999")
     TaxTran(1).TaxYear = ThisYear
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = "R"
     TaxTran(1).Amount = TBRECustBal(1).IntAmt(ll)
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBRECustBal(1).IntAmt(ll)
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2Gl = "Y"
     TaxTran(1).CustomerRec = Record!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, Record!, TaxCustRec(1)
      TaxCustRec(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, Record!, TaxCustRec(1)
    END IF
    IF TBRECustBal(1).PenAmt(ll) > 0 THEN
     TaxTran(1).TransDate = Date2Num%(("12-05-" + Yr$))
     TaxTran(1).TaxYear = ThisYear
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = "R"
     TaxTran(1).Amount = TBRECustBal(1).PenAmt(ll)
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBRECustBal(1).PenAmt(ll)
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2Gl = "Y"
     TaxTran(1).CustomerRec = Record!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, Record!, TaxCustRec(1)
      TaxCustRec(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, Record!, TaxCustRec(1)
    END IF
    NEXT ll
    CLOSE TaxTran

Nextone:
   NEXT x
   CLOSE
   EXIT SUB



UpdateReal:
    RERecord& = LOF(PropTaxFile) / LEN(PropertyRec(1)) + 1
    PropertyRec(1).REALPIN = TBCustRec(1).MapNumber
    PropertyRec(1).PROPDATE = Date2Num%("03-01-99")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = ""
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PROPSIZE = VAL(TBCustRec(1).Acres)
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = TBCustRec(1).LandValue
    PropertyRec(1).EXMPSENI = TBCustRec(1).BldgValue
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = TBCustRec(1).PropDesc1
    PropertyRec(1).PROPNOT2 = TBCustRec(1).PropDesc2
    PropertyRec(1).PROPNOT3 = TBCustRec(1).PropDesc3
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 1998
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    PUT PropTaxFile, RERecord&, PropertyRec(1)
    RETURN

END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, PersTaxFile)
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

