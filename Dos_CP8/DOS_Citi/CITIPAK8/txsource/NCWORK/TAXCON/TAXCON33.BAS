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
  '$INCLUDE: 'PROPAbst.BI'
  
  
  DIM SHARED TAXCUSTREC(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TBCust(1) AS TBCustType
  DIM SHARED TBCust1(1) AS TBCust1Type
  DIM SHARED TBCust2(1) AS TBCust2Type
  DIM SHARED TaxTran(1) AS TaxTransactionType


  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing

  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  OpenTaxCustFile NumOfTaxRecs, TaxFile

   OPEN "TBCust.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 512
   OPEN "TBCust1.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #12 LEN = 256
   TaxTranRecLen = LEN(TaxTran(1))

   
1 FOR X = 1 TO LOF(11) / 512
   GET 11, X, TBCust(1)
   GET 12, X, TBCust1(1)
   RECORD! = X
   
     IF VAL(TBCust(1).Acct) > 0 THEN

     GOSUB FigureBalance

     TotalReal# = TBCust(1).pland + TBCust(1).pbldg
     TotalReal# = INT((TotalReal# * 100) + .5) / 100
     TotalPers# = TBCust(1).ls + TBCust(1).MH + TBCust(1).fm + TBCust(1).homeexp + TBCust(1).useexp
     TotalPers# = INT((TotalPers# * 100) + .5) / 100

    
 ' Split Name Here
   TAXCUSTREC(1).Acct = RECORD!
   TAXCUSTREC(1).FName = TBCust(1).firstnamE
   TAXCUSTREC(1).LName = TBCust(1).LName
   TAXCUSTREC(1).SName = TBCust(1).LName
   TAXCUSTREC(1).Addr1 = TBCust(1).Addr1
   TAXCUSTREC(1).Addr2 = TBCust(1).Addr2
   TAXCUSTREC(1).CITY = TBCust(1).CITY
   TAXCUSTREC(1).State = TBCust(1).State
   TAXCUSTREC(1).ZIP = LEFT$(TBCust(1).ZIP, 5)
   TAXCUSTREC(1).HPHONE = TBCust(1).Phone
   TAXCUSTREC(1).WPHONE = ""
   TAXCUSTREC(1).CSSN = LEFT$(TBCust(1).SSN, 3) + MID$(TBCust(1).SSN, 5, 2) + MID$(TBCust(1).SSN, 8, 4)
   TAXCUSTREC(1).SSSN = ""
   TAXCUSTREC(1).Active = "Y"
   TAXCUSTREC(1).interest = "Y"
   TAXCUSTREC(1).TaxExempt = "N"
    TAXCUSTREC(1).Penalty = "N"
    TAXCUSTREC(1).LastTrans = 0
    TAXCUSTREC(1).FirstPropRec = 0
    TAXCUSTREC(1).FirstPersRec = 0
    TAXCUSTREC(1).Pin = RECORD!
    TAXCUSTREC(1).Deleted = 0
    TAXCUSTREC(1).FileVer = 8
    PUT TaxFile, RECORD!, TAXCUSTREC(1)

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


    ELSE
     TAXCUSTREC(1).Deleted = -1
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF      'Check VALID ACCT

   NEXT X
   
   CLOSE
   EXIT SUB

UpdateCurTrans:
     TaxTran = FREEFILE
     OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
   IF Round#((TBCust(1).CTax)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1999
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust(1).CTax
     TaxTran(1).Revenue.Principle1 = TBCust(1).CTax
     TaxTran(1).Revenue.interest = TBCust(1).CInterest
     TaxTran(1).Revenue.Penalty = TBCust(1).CLate
     TaxTran(1).Revenue.Collection = TBCust(1).CAdv
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust(1).CNotice)
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = 0
     TaxTran(1).BelongTo = 0
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     BillTran& = NextTran&
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

     'Check Interest
    IF Round#((TBCust(1).CInterest)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("09-01-1999")
     TaxTran(1).TaxYear = 1999
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust(1).CInterest
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust(1).CInterest
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF
     'Check Penalty
    IF Round#((TBCust(1).CLate)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("09-01-1999")
     TaxTran(1).TaxYear = 1999
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust(1).CLate
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Penalty = TBCust(1).CLate
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF
    CLOSE TaxTran
    RETURN



UpdatePastTrans:
 'Must Loop Thru Each of 10 Years

     TaxTran = FREEFILE
     OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
     LastTrans& = TAXCUSTREC(1).LastTrans
   IF Round#((TBCust1(1).ptaxamt0)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1998")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt0
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt0
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt0
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice0)
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = 0
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     BillTran& = NextTran&
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

     'Check Interest
    IF Round#((TBCust1(1).pintamt0)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt0
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt0
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

   'Check Penalty
    IF Round#((TBCust1(1).plateamt0)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt0
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF

   'Check Advertising
    IF Round#((TBCust1(1).padvamt0)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1999")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt0
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt0
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF


   IF Round#((TBCust1(1).ptaxamt1)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("10-01-1997")
     TaxTran(1).TaxYear = 1997
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt1
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt1
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt1
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt1
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt1
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice1)
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = 0
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     BillTran& = NextTran&
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

     'Check Interest
    IF Round#((TBCust1(1).pintamt1)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1997
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt1
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt1
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

   'Check Penalty
    IF Round#((TBCust1(1).plateamt1)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1997
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt1
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt1
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF
   'Check Advertising
    IF Round#((TBCust1(1).padvamt1)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1998")
     TaxTran(1).TaxYear = 1997
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt1
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt1
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF


   IF Round#((TBCust1(1).ptaxamt2)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1996")
     TaxTran(1).TaxYear = 1996
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt2
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt2
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt2
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt2
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt2
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice2)
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = 0
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     BillTran& = NextTran&
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

     'Check Interest
    IF Round#((TBCust1(1).pintamt2)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1996
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt2
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt2
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

   'Check Penalty
    IF Round#((TBCust1(1).plateamt2)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1996
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt2
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt2
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF

   'Check Advertising
    IF Round#((TBCust1(1).padvamt2)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1997")
     TaxTran(1).TaxYear = 1996
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt2
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt2
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF

   IF Round#((TBCust1(1).ptaxamt3)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1995")
     TaxTran(1).TaxYear = 1995
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt3
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt3
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt3
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt3
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt3
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice3)
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = 0
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     BillTran& = NextTran&
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

     'Check Interest
    IF Round#((TBCust1(1).pintamt3)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1995
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt3
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt3
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

   'Check Penalty
    IF Round#((TBCust1(1).plateamt3)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1995
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt3
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt3
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF

   'Check Advertising
    IF Round#((TBCust1(1).padvamt3)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1996")
     TaxTran(1).TaxYear = 1995
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt3
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt3
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF

   IF Round#((TBCust1(1).ptaxamt4)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1994")
     TaxTran(1).TaxYear = 1994
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt4
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt4
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt4
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt4
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt4
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice4)
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = 0
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     BillTran& = NextTran&
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

     'Check Interest
    IF Round#((TBCust1(1).pintamt4)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1994
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt4
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt4
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

   'Check Penalty
    IF Round#((TBCust1(1).plateamt4)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1994
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt4
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt4
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF
   'Check Advertising
    IF Round#((TBCust1(1).padvamt4)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1995")
     TaxTran(1).TaxYear = 1994
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt4
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt4
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF

   IF Round#((TBCust1(1).ptaxamt5)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("10-01-1993")
     TaxTran(1).TaxYear = 1993
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt5
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt5
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt5
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt5
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt5
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice5)
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = 0
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     BillTran& = NextTran&
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

     'Check Interest
    IF Round#((TBCust1(1).pintamt5)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1993
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt5
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt5
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

   'Check Penalty
    IF Round#((TBCust1(1).plateamt5)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1993
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt5
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt5
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF
   'Check Advertising
    IF Round#((TBCust1(1).padvamt5)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1994")
     TaxTran(1).TaxYear = 1993
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt5
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt5
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF




   IF Round#((TBCust1(1).ptaxamt6)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1992")
     TaxTran(1).TaxYear = 1992
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt6
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt6
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt6
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt6
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt6
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice6)
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = 0
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     BillTran& = NextTran&
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

     'Check Interest
    IF Round#((TBCust1(1).pintamt6)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1992
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt6
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt6
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

   'Check Penalty
    IF Round#((TBCust1(1).plateamt6)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1992
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt6
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt6
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF
   'Check Advertising
    IF Round#((TBCust1(1).padvamt6)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1993")
     TaxTran(1).TaxYear = 1992
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt6
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt6
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF


   IF Round#((TBCust1(1).ptaxamt7)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1991")
     TaxTran(1).TaxYear = 1991
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt7
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt7
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt7
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt7
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt7
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice7)
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = 0
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     BillTran& = NextTran&
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

     'Check Interest
    IF Round#((TBCust1(1).pintamt7)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1991
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt7
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt7
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

   'Check Penalty
    IF Round#((TBCust1(1).plateamt7)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1991
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt7
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt7
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF
   'Check Advertising
    IF Round#((TBCust1(1).padvamt7)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1992")
     TaxTran(1).TaxYear = 1991
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt7
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt7
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF


   IF Round#((TBCust1(1).ptaxamt8)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("10-01-1990")
     TaxTran(1).TaxYear = 1990
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt8
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt8
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt8
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt8
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt8
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice8)
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = 0
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     BillTran& = NextTran&
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

     'Check Interest
    IF Round#((TBCust1(1).pintamt8)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1990
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt8
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt8
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

   'Check Penalty
    IF Round#((TBCust1(1).plateamt8)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1990
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt8
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt8
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF

   'Check Advertising
    IF Round#((TBCust1(1).padvamt8)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1991")
     TaxTran(1).TaxYear = 1990
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt8
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt8
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF


   IF Round#((TBCust1(1).ptaxamt9)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1989")
     TaxTran(1).TaxYear = 1989
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).ptaxamt9
     TaxTran(1).Revenue.Principle1 = TBCust1(1).ptaxamt9
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt9
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt9
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt9
     TaxTran(1).Description = "Tax Bill" + STR$(TBCust1(1).ptaxnotice9)
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = 0
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     BillTran& = NextTran&
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

     'Check Interest
    IF Round#((TBCust1(1).pintamt9)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1989
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).pintamt9
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust1(1).pintamt9
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Interest"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
    END IF

   'Check Penalty
    IF Round#((TBCust1(1).plateamt9)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1989
     TaxTran(1).TranType = 5
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).plateamt9
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Penalty = TBCust1(1).plateamt9
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Penalty"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF
   'Check Advertising
    IF Round#((TBCust1(1).padvamt9)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("01-01-1990")
     TaxTran(1).TaxYear = 1989
     TaxTran(1).TranType = 6
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust1(1).padvamt9
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Collection = TBCust1(1).padvamt9
     TaxTran(1).Description = "Ad Cost"
     TaxTran(1).Posted2GL = "Y"
     TaxTran(1).CustomerRec = RECORD!
     TaxTran(1).LastTrans = LastTrans&
     TaxTran(1).BelongTo = BillTran&
     TaxTran(1).DMVSubmitted = ""
     TaxTran(1).DMVBatch = 0
     TaxTran(1).Padding = ""
     NextTran& = (LOF(TaxTran) / LEN(TaxTran(1))) + 1
     PUT TaxTran, NextTran&, TaxTran(1)
     GET TaxFile, RECORD!, TAXCUSTREC(1)
      TAXCUSTREC(1).LastTrans = NextTran&
      LastTrans& = NextTran&
     PUT TaxFile, RECORD!, TAXCUSTREC(1)
   END IF


    CLOSE TaxTran 'Close Transaction File Here
    RETURN
  

UpdateReal:
    PropertyRec(1).REALPIN = TBCust(1).Pin
    PropertyRec(1).PROPDATE = Date2Num("01-01-2000")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = TBCust(1).MAP
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = ""
    PropertyRec(1).PROPSIZE = VAL(TBCust(1).psize)
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = TBCust(1).pland + TBCust(1).pbldg
    PropertyRec(1).EXMPSENI = TBCust(1).me
    PropertyRec(1).EXMPOTHR = TBCust(1).homeexp1
    PropertyRec(1).PROPNOT1 = TBCust(1).pdesc
    PropertyRec(1).PROPNOT2 = ""
    PropertyRec(1).PROPNOT3 = TBCust(1).SubDiv
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = RECORD!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 0
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN

UpdatePers:
   PersRec(1).PROPPIN = "1"
   PersRec(1).PROPDATE = Date2Num("01-01-2000")
   PersRec(1).PersVal = TotalPers#
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
   PersRec(1).CustPin = RECORD!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 0
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   NextRec = (LOF(PersTaxFile) / LEN(PersRec(1))) + 1
   PUT PersTaxFile, NextRec, PersRec(1)
   RETURN

FigureBalance:
     CURBAL# = 0: PastDue# = 0
     CURBAL# = TBCust(1).CLate + TBCust(1).CTax + TBCust(1).CInterest + TBCust(1).CAdv
     CURBAL# = Round#(CURBAL#)
     IF CURBAL# > 99999999 THEN CURBAL# = 0
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
  OPEN "TAXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TAXCUSTREC(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TAXCUSTREC(1))
  
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

