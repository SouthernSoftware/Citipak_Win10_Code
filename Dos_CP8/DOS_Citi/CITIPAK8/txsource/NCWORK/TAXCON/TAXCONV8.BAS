DEFINT A-Z
'MAIN OLD TO NEW CONVERSION (NORTH CAROLINA W/BALANCES)
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB OpenTaxPropFile (NumOfPropRecs%, PropTaxFile%)
DECLARE SUB OpenTaxPersFile (NumOfPersRecs%, PersTaxFile%)
DECLARE SUB DisplayTaxScrn (ScrnName$)
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
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Round# (B#)
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
  '$INCLUDE: 'TaxConv.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TBCust(1) AS TBCustType
  DIM SHARED TBCust1(1) AS TBCust1Type
  'DIM SHARED TBCust2(1) AS TBCust2Type
  DIM SHARED TaxTran(1) AS TaxTransactionType


  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing

  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile


   OPEN "TBCust.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 512
   OPEN "TBCust1.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #12 LEN = 256

   TaxTranRecLen = LEN(TaxTran(1))

   Record! = 0
1 FOR x = 1 TO LOF(11) / 512
   GET 11, x, TBCust(1)
   GET 12, x, TBCust1(1)


   
     IF VAL(TBCust(1).Acct) > 0 THEN

     GOSUB FigureBalance

     TotalReal# = TBCust(1).pland + TBCust(1).PBldg
     TotalReal# = INT((TotalReal# * 100) + .5) / 100
     TotalPers# = TBCust(1).LS + TBCust(1).MH + TBCust(1).FM
     TotalPers# = INT((TotalPers# * 100) + .5) / 100

     IF CURBAL# > 0 OR PastDue# > 0 OR TotalReal# > 0 OR TotalPers# > 0 THEN

   

    OpenTaxCustFile NumOfTaxRecs, TaxFile
     
    Record! = NumOfTaxRecs + 1

   'Split Name Here
   TaxCustRec(1).Acct = Record!
   TaxCustRec(1).FName = TBCust(1).FirstName
   TaxCustRec(1).LName = TBCust(1).LName
   TaxCustRec(1).SName = TBCust(1).LName
   TaxCustRec(1).Addr1 = TBCust(1).Addr1
   TaxCustRec(1).Addr2 = TBCust(1).Addr2
   TaxCustRec(1).City = TBCust(1).City
   TaxCustRec(1).State = TBCust(1).State
   TaxCustRec(1).Zip = LEFT$(TBCust(1).Zip, 5)
   TaxCustRec(1).HPHONE = TBCust(1).phone
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = LEFT$(TBCust(1).SSN, 3) + MID$(TBCust(1).SSN, 5, 2) + MID$(TBCust(1).SSN, 8, 4)
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "N"
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).Employer = ""
    TaxCustRec(1).Bankrupt = "N"
    TaxCustRec(1).TownShip = ""
    PUT TaxFile, Record!, TaxCustRec(1)

    IF TotalReal# > 0 THEN
     GOSUB UpdateReal
    END IF

    IF TotalPers# > 0 THEN
     GOSUB UpdatePers
    END IF

    BillType$ = "C"

    IF CURBAL# > 0 THEN
     GOSUB UpdateCurTrans
    END IF

    IF PastDue# > 0 THEN
     GOSUB UpdatePastTrans
    END IF
    
    END IF      ' Check for Valid Acct
    END IF      'Check Value or Balance
    CLOSE TaxFile
   NEXT x
   
   CLOSE
   EXIT SUB

UpdateCurTrans:
     TaxTran = FREEFILE
     OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
   IF INT(((TBCust(1).CTax) * 100) + .5) / 100 > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-2000")
     TaxTran(1).TaxYear = 2000
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust(1).CTax
     TaxTran(1).Revenue.Principle1 = TBCust(1).CTax
     TaxTran(1).Revenue.Interest = TBCust(1).CInt
     TaxTran(1).Revenue.Penalty = TBCust(1).CLate
     TaxTran(1).Revenue.Collection = TBCust(1).CAdv
     TaxTran(1).DESCRIPTION = "Tax Bill" + STR$(TBCust(1).CNotice)
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
    IF INT(((TBCust(1).CInt * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-07-2001")
     TaxTran(1).TaxYear = 2000
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust(1).CInt
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust(1).CInt
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Interest"
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
    IF INT(((TBCust(1).CLate * 100 + .5) / 100)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("09-01-2000")
     TaxTran(1).TaxYear = 2000
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust(1).CLate
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust(1).CLate
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Penalty"
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
   IF INT(((TBCust1(1).PTaxAmt0 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1999
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PTaxAmt0
     TaxTran(1).Revenue.Principle1 = TBCust1(1).PTaxAmt0
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt0
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt0
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt0
     TaxTran(1).DESCRIPTION = "Tax Bill" + STR$(TBCust1(1).PTaxNotice0)
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
    IF INT(((TBCust1(1).PIntAmt0 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1999
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PIntAmt0
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt0
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Interest"
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
    IF INT(((TBCust1(1).PLateAmt0 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1999
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PLateAmt0
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Penalty"
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
    IF INT(((TBCust1(1).PAdvAmt0 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1999")
     TaxTran(1).TaxYear = 1999
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PAdvAmt0
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt0
     TaxTran(1).DESCRIPTION = "Ad Cost"
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


   IF INT(((TBCust1(1).PTaxAmt1 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("10-01-1998")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PTaxAmt1
     TaxTran(1).Revenue.Principle1 = TBCust1(1).PTaxAmt1
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt1
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt1
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt1
     TaxTran(1).DESCRIPTION = "Tax Bill" + STR$(TBCust1(1).PTaxNotice1)
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
    IF INT(((TBCust1(1).PIntAmt1 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1998")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PIntAmt1
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt1
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Interest"
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
    IF INT(((TBCust1(1).PLateAmt1 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1998")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PLateAmt1
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt1
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Penalty"
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
    IF INT(((TBCust1(1).PAdvAmt1 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1998")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PAdvAmt1
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt1
     TaxTran(1).DESCRIPTION = "Ad Cost"
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


   IF INT(((TBCust1(1).PTaxAmt2 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1997")
     TaxTran(1).TaxYear = 1997
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PTaxAmt2
     TaxTran(1).Revenue.Principle1 = TBCust1(1).PTaxAmt2
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt2
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt2
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt2
     TaxTran(1).DESCRIPTION = "Tax Bill" + STR$(TBCust1(1).PTaxNotice2)
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
    IF INT(((TBCust1(1).PIntAmt2 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1997")
     TaxTran(1).TaxYear = 1997
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PIntAmt2
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt2
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Interest"
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
    IF INT(((TBCust1(1).PLateAmt2 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1997")
     TaxTran(1).TaxYear = 1997
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PLateAmt2
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt2
     TaxTran(1).DESCRIPTION = "Penalty"
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
    IF INT(((TBCust1(1).PAdvAmt2 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1997")
     TaxTran(1).TaxYear = 1997
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PAdvAmt2
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt2
     TaxTran(1).DESCRIPTION = "Ad Cost"
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

   IF INT(((TBCust1(1).PTaxAmt3 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1996")
     TaxTran(1).TaxYear = 1996
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PTaxAmt3
     TaxTran(1).Revenue.Principle1 = TBCust1(1).PTaxAmt3
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt3
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt3
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt3
     TaxTran(1).DESCRIPTION = "Tax Bill" + STR$(TBCust1(1).PTaxNotice3)
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
    IF INT(((TBCust1(1).PIntAmt3 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1996")
     TaxTran(1).TaxYear = 1996
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PIntAmt3
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt3
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Interest"
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
    IF INT(((TBCust1(1).PLateAmt3 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1996")
     TaxTran(1).TaxYear = 1996
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PLateAmt3
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt3
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Penalty"
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
    IF INT(((TBCust1(1).PAdvAmt3 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1996")
     TaxTran(1).TaxYear = 1996
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PAdvAmt3
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt3
     TaxTran(1).DESCRIPTION = "Ad Cost"
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

   IF INT(((TBCust1(1).PTaxAmt4 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1995")
     TaxTran(1).TaxYear = 1995
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PTaxAmt4
     TaxTran(1).Revenue.Principle1 = TBCust1(1).PTaxAmt4
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt4
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt4
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt4
     TaxTran(1).DESCRIPTION = "Tax Bill" + STR$(TBCust1(1).PTaxNotice4)
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
    IF INT(((TBCust1(1).PIntAmt4 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1995")
     TaxTran(1).TaxYear = 1995
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PIntAmt4
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt4
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Interest"
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
    IF INT(((TBCust1(1).PLateAmt4 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1995")
     TaxTran(1).TaxYear = 1995
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PLateAmt4
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt4
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Penalty"
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
    IF INT(((TBCust1(1).PAdvAmt4 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1995")
     TaxTran(1).TaxYear = 1995
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PAdvAmt4
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt4
     TaxTran(1).DESCRIPTION = "Ad Cost"
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

   IF INT(((TBCust1(1).PTaxAmt5 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("10-01-1994")
     TaxTran(1).TaxYear = 1994
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PTaxAmt5
     TaxTran(1).Revenue.Principle1 = TBCust1(1).PTaxAmt5
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt5
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt5
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt5
     TaxTran(1).DESCRIPTION = "Tax Bill" + STR$(TBCust1(1).PTaxNotice5)
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
    IF INT(((TBCust1(1).PIntAmt5 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1994")
     TaxTran(1).TaxYear = 1994
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PIntAmt5
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt5
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Interest"
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
    IF INT(((TBCust1(1).PLateAmt5 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1994")
     TaxTran(1).TaxYear = 1994
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PLateAmt5
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt5
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Penalty"
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
    IF INT(((TBCust1(1).PAdvAmt5 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1994")
     TaxTran(1).TaxYear = 1994
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PAdvAmt5
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt5
     TaxTran(1).DESCRIPTION = "Ad Cost"
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




   IF INT(((TBCust1(1).PTaxAmt6 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1993")
     TaxTran(1).TaxYear = 1993
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PTaxAmt6
     TaxTran(1).Revenue.Principle1 = TBCust1(1).PTaxAmt6
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt6
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt6
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt6
     TaxTran(1).DESCRIPTION = "Tax Bill" + STR$(TBCust1(1).PTaxNotice6)
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
    IF INT(((TBCust1(1).PIntAmt6 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1993")
     TaxTran(1).TaxYear = 1993
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PIntAmt6
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt6
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Interest"
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
    IF INT(((TBCust1(1).PLateAmt6 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1993")
     TaxTran(1).TaxYear = 1993
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PLateAmt6
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt6
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Penalty"
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
    IF INT(((TBCust1(1).PAdvAmt6 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1993")
     TaxTran(1).TaxYear = 1993
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PAdvAmt6
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt6
     TaxTran(1).DESCRIPTION = "Ad Cost"
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


   IF INT(((TBCust1(1).PTaxAmt7 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1992")
     TaxTran(1).TaxYear = 1992
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PTaxAmt7
     TaxTran(1).Revenue.Principle1 = TBCust1(1).PTaxAmt7
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt7
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt7
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt7
     TaxTran(1).DESCRIPTION = "Tax Bill" + STR$(TBCust1(1).PTaxNotice7)
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
    IF INT(((TBCust1(1).PIntAmt7 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1992")
     TaxTran(1).TaxYear = 1992
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PIntAmt7
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt7
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Interest"
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
    IF INT(((TBCust1(1).PLateAmt7 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1992")
     TaxTran(1).TaxYear = 1992
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PLateAmt7
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt7
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Penalty"
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
    IF INT(((TBCust1(1).PAdvAmt7 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1992")
     TaxTran(1).TaxYear = 1992
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PAdvAmt7
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt7
     TaxTran(1).DESCRIPTION = "Ad Cost"
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


   IF INT(((TBCust1(1).PTaxAmt8 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("10-01-1991")
     TaxTran(1).TaxYear = 1991
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PTaxAmt8
     TaxTran(1).Revenue.Principle1 = TBCust1(1).PTaxAmt8
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt8
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt8
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt8
     TaxTran(1).DESCRIPTION = "Tax Bill" + STR$(TBCust1(1).PTaxNotice8)
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
    IF INT(((TBCust1(1).PIntAmt8 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1991")
     TaxTran(1).TaxYear = 1991
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PIntAmt8
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt8
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Interest"
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
    IF INT(((TBCust1(1).PLateAmt8 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1991")
     TaxTran(1).TaxYear = 1991
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PLateAmt8
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt8
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Penalty"
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
    IF INT(((TBCust1(1).PAdvAmt8 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1991")
     TaxTran(1).TaxYear = 1991
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PAdvAmt8
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt8
     TaxTran(1).DESCRIPTION = "Ad Cost"
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


   IF INT(((TBCust1(1).PTaxAmt9 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1990")
     TaxTran(1).TaxYear = 1990
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PTaxAmt9
     TaxTran(1).Revenue.Principle1 = TBCust1(1).PTaxAmt9
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt9
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt9
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt9
     TaxTran(1).DESCRIPTION = "Tax Bill" + STR$(TBCust1(1).PTaxNotice9)
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
    IF INT(((TBCust1(1).PIntAmt9 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1990")
     TaxTran(1).TaxYear = 1990
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PIntAmt9
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = TBCust1(1).PIntAmt9
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Interest"
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
    IF INT(((TBCust1(1).PLateAmt9 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1990")
     TaxTran(1).TaxYear = 1990
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PLateAmt9
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).PLateAmt9
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).DESCRIPTION = "Penalty"
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
    IF INT(((TBCust1(1).PAdvAmt9 * 100) + .5) / 100) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1990")
     TaxTran(1).TaxYear = 1990
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).PAdvAmt9
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.Interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).PAdvAmt9
     TaxTran(1).DESCRIPTION = "Ad Cost"
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
  

UpdateReal:
    PropertyRec(1).REALPIN = TBCust(1).Acct
    PropertyRec(1).PROPDATE = -32767
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = TBCust(1).pmap
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = ""
    PropertyRec(1).PROPSIZE = TBCust(1).psize
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = TBCust(1).pland + TBCust(1).PBldg
    PropertyRec(1).EXMPSENI = TBCust(1).ME
    PropertyRec(1).EXMPOTHR = TBCust(1).homeexp + TBCust(1).UseExp
    PropertyRec(1).PROPNOT1 = TBCust(1).pdesc
    PropertyRec(1).PROPNOT2 = ""
    PropertyRec(1).PROPNOT3 = ""
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 0
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN

UpdatePers:
   PersRec(1).PROPPIN = "1"
   PersRec(1).PROPDATE = -32767
   PersRec(1).PersVal = TBCust(1).LS + TBCust(1).MH + TBCust(1).FM
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = ""
   PersRec(1).DESC2 = ""
   PersRec(1).DESC3 = ""
   PersRec(1).Desc4 = ""
   PersRec(1).Desc5 = ""
   PersRec(1).CustPin = Record!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 0
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   NextRec = (LOF(PersTaxFile) / LEN(PersRec(1))) + 1
   PUT PersTaxFile, NextRec, PersRec(1)
   RETURN

FigureBalance:
     CURBAL# = 0: PastDue# = 0
     CURBAL# = TBCust(1).CLate + TBCust(1).CTax + TBCust(1).CInt + TBCust(1).CAdv
     CURBAL# = INT(((CURBAL# * 100) + .5) / 100)
     IF CURBAL# > 99999999 THEN CURBAL# = 0
     PastDue# = (TBCust1(1).PTaxAmt0 + TBCust1(1).PIntAmt0 + TBCust1(1).PAdvAmt0 + TBCust1(1).PColAmt0 + TBCust1(1).PLateAmt0)
     PastDue# = PastDue# + (TBCust1(1).PTaxAmt1 + TBCust1(1).PIntAmt1 + TBCust1(1).PAdvAmt1 + TBCust1(1).PColAmt1 + TBCust1(1).PLateAmt1)
     PastDue# = PastDue# + (TBCust1(1).PTaxAmt2 + TBCust1(1).PIntAmt2 + TBCust1(1).PAdvAmt2 + TBCust1(1).PColAmt2 + TBCust1(1).PLateAmt2)
     PastDue# = PastDue# + (TBCust1(1).PTaxAmt3 + TBCust1(1).PIntAmt3 + TBCust1(1).PAdvAmt3 + TBCust1(1).PColAmt3 + TBCust1(1).PLateAmt3)
     PastDue# = PastDue# + (TBCust1(1).PTaxAmt4 + TBCust1(1).PIntAmt4 + TBCust1(1).PAdvAmt4 + TBCust1(1).PColAmt4 + TBCust1(1).PLateAmt4)
     PastDue# = PastDue# + (TBCust1(1).PTaxAmt5 + TBCust1(1).PIntAmt5 + TBCust1(1).PAdvAmt5 + TBCust1(1).PColAmt5 + TBCust1(1).PLateAmt5)
     PastDue# = PastDue# + (TBCust1(1).PTaxAmt6 + TBCust1(1).PIntAmt6 + TBCust1(1).PAdvAmt6 + TBCust1(1).PColAmt6 + TBCust1(1).PLateAmt6)
     PastDue# = PastDue# + (TBCust1(1).PTaxAmt7 + TBCust1(1).PIntAmt7 + TBCust1(1).PAdvAmt7 + TBCust1(1).PColAmt7 + TBCust1(1).PLateAmt7)
     PastDue# = PastDue# + (TBCust1(1).PTaxAmt8 + TBCust1(1).PIntAmt8 + TBCust1(1).PAdvAmt8 + TBCust1(1).PColAmt8 + TBCust1(1).PLateAmt8)
     PastDue# = PastDue# + (TBCust1(1).PTaxAmt9 + TBCust1(1).PIntAmt9 + TBCust1(1).PAdvAmt9 + TBCust1(1).PColAmt9 + TBCust1(1).PLateAmt9)
     PastDue# = INT(((PastDue# * 100) + .5) / 100)
     IF PastDue# > 9999999 THEN PastDue# = 0
     RETURN

END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
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

