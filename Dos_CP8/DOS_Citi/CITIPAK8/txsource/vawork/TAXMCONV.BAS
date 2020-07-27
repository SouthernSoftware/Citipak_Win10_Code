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
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TBCustRec(1) AS TBCustType
  DIM SHARED TBCust(1) AS TBCustType
  DIM SHARED TBCust1(1) AS TBCust1Type
  DIM SHARED TaxTran(1) AS TaxTransactionType

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing

  CLS

  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

  TaxTranRecLen = LEN(TaxTran(1))

  OPEN "TBCUST.DAT" FOR RANDOM AS #11 LEN = 512
  OPEN "TBCust1.Dat" FOR RANDOM AS #12 LEN = 256

  FOR MasterLoop = 1 TO 2
  'Process Real First

   FOR x = 1 TO LOF(11) / 512
   GET 11, x, TBCustRec(1)
   GET 11, x, TBCust(1)
   GET 12, x, TBCust1(1)
   IF VAL(TBCust(1).Acct) = 0 THEN GOTO NextOne
   GOSUB figurebalance

   IF MasterLoop = 1 THEN
    RealValue# = TBCustRec(1).pland + TBCustRec(1).pbldg
    RealValue# = INT((RealValue# * 100) + .5) / 100
    IF RealValue# = 0 THEN GOTO NextOne
   END IF
   
   IF MasterLoop = 2 THEN
    PersonalValue# = TBCustRec(1).MH + TBCustRec(1).fm + TBCustRec(1).ls + TBCustRec(1).op + TBCustRec(1).me
    PersonalValue# = INT((PersonalValue# * 100) + .5) / 100
    IF PersonalValue# = 0 AND CurBal# = 0 AND PastDue# = 0 THEN GOTO NextOne
   END IF


   IF MasterLoop = 1 THEN
    CLOSE TaxFile
    OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
    NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
   END IF
   IF MasterLoop = 2 THEN
    CLOSE TaxFile
    OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
    NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
   END IF
   
    Record! = NumOfTaxRecs + 1
    PRINT LTRIM$(TBCustRec(1).firstname + " " + TBCustRec(1).LName)
    TName$ = LTRIM$(TBCustRec(1).firstname + " " + TBCustRec(1).LName)
    IF LEN(TName$) > 0 THEN
    
   TaxCustRec(1).FName = LTRIM$(TBCustRec(1).firstname)
   TaxCustRec(1).LName = LTRIM$(TBCustRec(1).LName)
   TaxCustRec(1).SName = LTRIM$(TBCustRec(1).LName)
   
   TaxCustRec(1).Acct = Record!
   TaxCustRec(1).CountyAcct = 0
   TaxCustRec(1).Addr1 = TBCustRec(1).Addr1
   TaxCustRec(1).Addr2 = TBCustRec(1).Addr2
   TaxCustRec(1).City = TBCustRec(1).City
   TaxCustRec(1).State = TBCustRec(1).State
   TaxCustRec(1).Zip = TBCustRec(1).Zip
   TaxCustRec(1).HPHONE = TBCustRec(1).PHONE
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = LEFT$(TBCustRec(1).SSN, 9)
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "Y"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = Record!
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("01-01-1998")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1

   IF MasterLoop = 1 THEN
    GOSUB UpdateReal
    IF CurBal# > 0 THEN GOSUB UpdateCurTrans
    IF PastDue# > 0 THEN GOSUB UpdatePastTrans
   END IF
   

   IF MasterLoop = 2 THEN
    GOSUB updatepersonal
    IF CurBal# > 0 THEN GOSUB UpdateCurTrans
    IF PastDue# > 0 THEN GOSUB UpdatePastTrans
   END IF


NextOne:
   END IF
   NEXT x 'loop thru data files

   NEXT MasterLoop


   CLOSE
   EXIT SUB


updatepersonal:
   PersRec(1).PersVal = TBCustRec(1).MH
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
upd1:
   PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
   PersRec(1).PROPDATE = Date2Num%("03-01-1999")
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = TBCustRec(1).pVin
   PersRec(1).DESC2 = TBCustRec(1).PMake
   PersRec(1).DESC3 = ""
   PersRec(1).Desc4 = ""
   IF LEN(LTRIM$(TBCustRec(1).pVin)) > 0 THEN
    PersRec(1).Desc5 = "Y"
   ELSE
    PersRec(1).Desc5 = "N"
   END IF
   PersRec(1).CustPin = Record!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 1998
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
   PUT PersTaxFile, PRecord&, PersRec(1)
   RETURN

UpdateReal:
    RERecord& = LOF(PropTaxFile) / LEN(PropertyRec(1)) + 1
    PropertyRec(1).REALPIN = TBCustRec(1).map
    PropertyRec(1).PROPDATE = Date2Num%("03-01-99")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).map = ""
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PROPSIZE = VAL(TBCustRec(1).plot)
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = TBCustRec(1).pland
    PropertyRec(1).EXMPSENI = TBCustRec(1).pbldg
    PropertyRec(1).EXMPOTHR = TBCustRec(1).homeexp
    PropertyRec(1).PROPNOT1 = TBCustRec(1).PDesc
    PropertyRec(1).PROPNOT2 = RIGHT$(TBCustRec(1).PDesc, 8)
    PropertyRec(1).PROPNOT3 = TBCustRec(1).SubDiv
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 1998
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    PUT PropTaxFile, RERecord&, PropertyRec(1)
    RETURN

UpdateCurTrans:
     TaxTran = FREEFILE
     
     OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
   IF Round#((TBCust(1).CTax)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1998")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust(1).CTax
     TaxTran(1).Revenue.Principle1 = TBCust(1).CTax
     TaxTran(1).Revenue.Interest = TBCust(1).CInterest
     TaxTran(1).Revenue.Penalty = TBCust(1).CLate
     TaxTran(1).Revenue.Collection = TBCust(1).CAdv
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust(1).CNotice)
     TaxTran(1).Posted2GL = "Y"
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
    IF Round#((TBCust(1).CInterest)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("09-01-1998")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust(1).CInterest
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust(1).CInterest
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
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
    IF Round#((TBCust(1).CLate)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("09-01-1998")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust(1).CLate
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust(1).CLate
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
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
    RETURN



UpdatePastTrans:
 'Must Loop Thru Each of 10 Years

     TaxTran = FREEFILE
     OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
     LastTrans& = TaxCustRec(1).LastTrans
   IF Round#((TBCust1(1).ptaxamt0)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1997")
     TaxTran(1).TaxYear = 1997
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt0
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt0
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt0
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice0)
     TaxTran(1).Posted2GL = "Y"
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

     'Check Interest
    IF Round#((TBCust1(1).pintamt0)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1998")
     TaxTran(1).TaxYear = 1997
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt0
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt0
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
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
    IF Round#((TBCust1(1).plateamt0)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1998")
     TaxTran(1).TaxYear = 1997
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt0
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
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

   'Check Advertising
    IF Round#((TBCust1(1).padvamt0)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1998")
     TaxTran(1).TaxYear = 1997
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt0
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt0
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
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


   IF Round#((TBCust1(1).ptaxamt1)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("10-01-1996")
     TaxTran(1).TaxYear = 1996
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt1
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt1
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt1
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt1
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt1
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice1)
     TaxTran(1).Posted2GL = "Y"
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

     'Check Interest
    IF Round#((TBCust1(1).pintamt1)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1996
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt1
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt1
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
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
    IF Round#((TBCust1(1).plateamt1)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1996")
     TaxTran(1).TaxYear = 1996
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt1
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt1
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
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
   'Check Advertising
    IF Round#((TBCust1(1).padvamt1)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1997")
     TaxTran(1).TaxYear = 1996
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt1
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt1
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
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


   IF Round#((TBCust1(1).ptaxamt2)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1995")
     TaxTran(1).TaxYear = 1995
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt2
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt2
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt2
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt2
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt2
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice2)
     TaxTran(1).Posted2GL = "Y"
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

     'Check Interest
    IF Round#((TBCust1(1).pintamt2)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1995
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt2
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt2
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
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
    IF Round#((TBCust1(1).plateamt2)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1995")
     TaxTran(1).TaxYear = 1995
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt2
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt2
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Posted2GL = "Y"
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

   'Check Advertising
    IF Round#((TBCust1(1).padvamt2)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1996")
     TaxTran(1).TaxYear = 1995
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt2
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt2
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
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

   IF Round#((TBCust1(1).ptaxamt3)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1994")
     TaxTran(1).TaxYear = 1994
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt3
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt3
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt3
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt3
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt3
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice3)
     TaxTran(1).Posted2GL = "Y"
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

     'Check Interest
    IF Round#((TBCust1(1).pintamt3)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1994
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt3
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt3
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
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
    IF Round#((TBCust1(1).plateamt3)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1994")
     TaxTran(1).TaxYear = 1994
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt3
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt3
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
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

   'Check Advertising
    IF Round#((TBCust1(1).padvamt3)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1995")
     TaxTran(1).TaxYear = 1994
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt3
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt3
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
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

   IF Round#((TBCust1(1).ptaxamt4)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1993")
     TaxTran(1).TaxYear = 1993
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt4
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt4
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt4
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt4
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt4
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice4)
     TaxTran(1).Posted2GL = "Y"
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

     'Check Interest
    IF Round#((TBCust1(1).pintamt4)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1993
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt4
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt4
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
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
    IF Round#((TBCust1(1).plateamt4)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1993")
     TaxTran(1).TaxYear = 1993
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt4
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt4
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
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
   'Check Advertising
    IF Round#((TBCust1(1).padvamt4)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1994")
     TaxTran(1).TaxYear = 1993
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt4
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt4
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
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

   IF Round#((TBCust1(1).ptaxamt5)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("10-01-1992")
     TaxTran(1).TaxYear = 1992
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt5
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt5
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt5
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt5
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt5
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice5)
     TaxTran(1).Posted2GL = "Y"
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

     'Check Interest
    IF Round#((TBCust1(1).pintamt5)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1992
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt5
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt5
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
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
    IF Round#((TBCust1(1).plateamt5)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1992")
     TaxTran(1).TaxYear = 1992
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt5
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt5
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
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
   'Check Advertising
    IF Round#((TBCust1(1).padvamt5)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1993")
     TaxTran(1).TaxYear = 1992
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt5
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt5
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
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




   IF Round#((TBCust1(1).ptaxamt6)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1991")
     TaxTran(1).TaxYear = 1991
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt6
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt6
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt6
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt6
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt6
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice6)
     TaxTran(1).Posted2GL = "Y"
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

     'Check Interest
    IF Round#((TBCust1(1).pintamt6)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1991
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt6
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt6
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
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
    IF Round#((TBCust1(1).plateamt6)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1991")
     TaxTran(1).TaxYear = 1991
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt6
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt6
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
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
   'Check Advertising
    IF Round#((TBCust1(1).padvamt6)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1992")
     TaxTran(1).TaxYear = 1991
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt6
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt6
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
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


   IF Round#((TBCust1(1).ptaxamt7)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1990")
     TaxTran(1).TaxYear = 1990
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt7
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt7
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt7
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt7
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt7
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice7)
     TaxTran(1).Posted2GL = "Y"
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

     'Check Interest
    IF Round#((TBCust1(1).pintamt7)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1990
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt7
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt7
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
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
    IF Round#((TBCust1(1).plateamt7)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1990")
     TaxTran(1).TaxYear = 1990
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt7
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt7
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
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
   'Check Advertising
    IF Round#((TBCust1(1).padvamt7)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1991")
     TaxTran(1).TaxYear = 1990
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt7
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt7
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
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


   IF Round#((TBCust1(1).ptaxamt8)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("10-01-1989")
     TaxTran(1).TaxYear = 1989
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt8
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt8
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt8
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt8
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt8
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice8)
     TaxTran(1).Posted2GL = "Y"
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

     'Check Interest
    IF Round#((TBCust1(1).pintamt8)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1989
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt8
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt8
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
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
    IF Round#((TBCust1(1).plateamt8)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1989")
     TaxTran(1).TaxYear = 1989
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt8
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt8
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
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

   'Check Advertising
    IF Round#((TBCust1(1).padvamt8)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1990")
     TaxTran(1).TaxYear = 1989
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt8
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt8
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
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


   IF Round#((TBCust1(1).ptaxamt9)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1988")
     TaxTran(1).TaxYear = 1988
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt9
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt9
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt9
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt9
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt9
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice9)
     TaxTran(1).Posted2GL = "Y"
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

     'Check Interest
    IF Round#((TBCust1(1).pintamt9)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1988
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt9
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).pintamt9
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
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
    IF Round#((TBCust1(1).plateamt9)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1988")
     TaxTran(1).TaxYear = 1988
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt9
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt9
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
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
   'Check Advertising
    IF Round#((TBCust1(1).padvamt9)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1989")
     TaxTran(1).TaxYear = 1988
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt9
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt9
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
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
   CLOSE TaxTran 'Close Transaction File Here
   RETURN


figurebalance:
     CurBal# = 0: PastDue# = 0
     CurBal# = TBCust(1).CLate + TBCust(1).CTax + TBCust(1).CInterest + TBCust(1).CAdv
     CurBal# = Round#(CurBal#)
     IF CurBal# > 99999999 THEN CurBal# = 0
     PastDue# = (TBCust1(1).ptaxamt0 + TBCust1(1).pintamt0 + TBCust1(1).padvamt0 + TBCust1(1).pcolamt0 + TBCust1(1).plateamt0)
     PastDue# = PastDue# + (TBCust1(1).ptaxamt1 + TBCust1(1).pintamt1 + TBCust1(1).padvamt1 + TBCust1(1).pcolamt1 + TBCust1(1).plateamt1)
     PastDue# = PastDue# + (TBCust1(1).ptaxamt2 + TBCust1(1).pintamt2 + TBCust1(1).padvamt2 + TBCust1(1).pcolamt2 + TBCust1(1).plateamt2)
     PastDue# = PastDue# + (TBCust1(1).ptaxamt3 + TBCust1(1).pintamt3 + TBCust1(1).padvamt3 + TBCust1(1).pcolamt3 + TBCust1(1).plateamt3)
     PastDue# = PastDue# + (TBCust1(1).ptaxamt4 + TBCust1(1).pintamt4 + TBCust1(1).padvamt4 + TBCust1(1).pcolamt4 + TBCust1(1).plateamt4)
     PastDue# = PastDue# + (TBCust1(1).ptaxamt5 + TBCust1(1).pintamt5 + TBCust1(1).padvamt5 + TBCust1(1).pcolamt5 + TBCust1(1).plateamt5)
     PastDue# = PastDue# + (TBCust1(1).ptaxamt6 + TBCust1(1).pintamt6 + TBCust1(1).padvamt6 + TBCust1(1).pcolamt6 + TBCust1(1).plateamt6)
     PastDue# = PastDue# + (TBCust1(1).ptaxamt7 + TBCust1(1).pintamt7 + TBCust1(1).padvamt7 + TBCust1(1).pcolamt7 + TBCust1(1).plateamt7)
     PastDue# = PastDue# + (TBCust1(1).ptaxamt8 + TBCust1(1).pintamt8 + TBCust1(1).padvamt8 + TBCust1(1).pcolamt8 + TBCust1(1).plateamt8)
     PastDue# = PastDue# + (TBCust1(1).ptaxamt9 + TBCust1(1).pintamt9 + TBCust1(1).padvamt9 + TBCust1(1).pcolamt9 + TBCust1(1).plateamt9)
     PastDue# = Round#(PastDue#)
     IF PastDue# > 9999999 THEN PastDue# = 0
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

