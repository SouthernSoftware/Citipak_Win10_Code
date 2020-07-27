DEFINT A-Z
'Early OLD TO NEW CONVERSION (NORTH CAROLINA W/BALANCES) (Calabash)
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
  DIM SHARED TBCust(1) AS tbcusttype
  DIM SHARED TBCust1(1) AS TBcust1type
  DIM SHARED TBCust2(1) AS TBcust1type
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

   
1 FOR x = 1 TO LOF(11) / 512
   GET 11, x, TBCust(1)
   GET 12, x, TBCust2(1)
    RECORD! = x
    IF VAL(TBCust(1).Acct) = 0 THEN
      TAXCUSTREC(1).Deleted = -1
      TAXCUSTREC(1).FileVer = 8
      PUT TaxFile, RECORD!, TAXCUSTREC(1)
    ELSE
     GOSUB FigureBalance

     TotalReal# = TBCust(1).PLand + TBCust(1).PBldg
     TotalReal# = INT((TotalReal# * 100) + .5) / 100
     TotalPers# = TBCust(1).MH + TBCust(1).FM + TBCust(1).LS
     TotalPers# = INT((TotalPers# * 100) + .5) / 100

  
  
 ' Split Name Here
   TAXCUSTREC(1).Acct = RECORD!
   TAXCUSTREC(1).FName = UCASE$(TBCust(1).firstnamE)
   TAXCUSTREC(1).LName = UCASE$(TBCust(1).LName)
   TAXCUSTREC(1).SName = UCASE$(TBCust(1).LName)
   TAXCUSTREC(1).Addr1 = UCASE$(TBCust(1).Addr1)
   TAXCUSTREC(1).Addr2 = UCASE$(TBCust(1).Addr2)
   TAXCUSTREC(1).CITY = UCASE$(TBCust(1).CITY)
   TAXCUSTREC(1).State = UCASE$(TBCust(1).State)
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
   TAXCUSTREC(1).PIN = RECORD!
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


   END IF
  

   NEXT x
   
   CLOSE
   EXIT SUB

UpdateCurTrans:
     TaxTran = FREEFILE
     OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
   IF Round#((TBCust(1).CTax)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("09-01-1999")
     TaxTran(1).TaxYear = 1999
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust(1).CTax
     TaxTran(1).Revenue.Principle1 = TBCust(1).CTax
     TaxTran(1).Revenue.interest = 0
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Tax Bill"
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
    CLOSE TaxTran
    RETURN



UpdatePastTrans:
 'Must Loop Thru Each of 10 Years

     TaxTran = FREEFILE
     OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
     LastTrans& = TAXCUSTREC(1).LastTrans
   IF Round#((TBCust2(1).ptaxamt0)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1998")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).ptaxamt0
     TaxTran(1).Revenue.Principle1 = TBCust2(1).ptaxamt0
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt0
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Tax Bill"
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
    IF Round#((TBCust2(1).pintamt0)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1998
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).pintamt0
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt0
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


   IF Round#((TBCust2(1).ptaxamt1)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("10-01-1997")
     TaxTran(1).TaxYear = 1997
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).ptaxamt1
     TaxTran(1).Revenue.Principle1 = TBCust2(1).ptaxamt1
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt1
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Tax Bill"
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
    IF Round#((TBCust2(1).pintamt1)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1997
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).pintamt1
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt1
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


   IF Round#((TBCust2(1).ptaxamt2)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1996")
     TaxTran(1).TaxYear = 1996
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).ptaxamt2
     TaxTran(1).Revenue.Principle1 = TBCust2(1).ptaxamt2
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt2
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Tax Bill"
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
    IF Round#((TBCust2(1).pintamt2)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1996
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).pintamt2
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt2
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

   IF Round#((TBCust2(1).ptaxamt3)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1995")
     TaxTran(1).TaxYear = 1995
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).ptaxamt3
     TaxTran(1).Revenue.Principle1 = TBCust2(1).ptaxamt3
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt3
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Tax Bill"
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
    IF Round#((TBCust2(1).pintamt3)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1995
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).pintamt3
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt3
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


   IF Round#((TBCust2(1).ptaxamt4)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1994")
     TaxTran(1).TaxYear = 1994
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).ptaxamt4
     TaxTran(1).Revenue.Principle1 = TBCust2(1).ptaxamt4
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt4
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Tax Bill"
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
    IF Round#((TBCust2(1).pintamt4)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1994
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).pintamt4
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt4
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

   IF Round#((TBCust2(1).ptaxamt5)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("10-01-1993")
     TaxTran(1).TaxYear = 1993
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).ptaxamt5
     TaxTran(1).Revenue.Principle1 = TBCust2(1).ptaxamt5
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt5
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Tax Bill"
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
    IF Round#((TBCust2(1).pintamt5)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1993
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).pintamt5
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt5
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



   IF Round#((TBCust2(1).ptaxamt6)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1992")
     TaxTran(1).TaxYear = 1992
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).ptaxamt6
     TaxTran(1).Revenue.Principle1 = TBCust2(1).ptaxamt6
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt6
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Tax Bill"
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
    IF Round#((TBCust2(1).pintamt6)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1992
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).pintamt6
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt6
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


   IF Round#((TBCust2(1).ptaxamt7)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1991")
     TaxTran(1).TaxYear = 1991
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).ptaxamt7
     TaxTran(1).Revenue.Principle1 = TBCust2(1).ptaxamt7
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt7
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Tax Bill"
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
    IF Round#((TBCust2(1).pintamt7)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1991
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).pintamt7
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt7
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


   IF Round#((TBCust2(1).ptaxamt8)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("10-01-1990")
     TaxTran(1).TaxYear = 1990
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).ptaxamt8
     TaxTran(1).Revenue.Principle1 = TBCust2(1).ptaxamt8
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt8
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Tax Bill"
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
    IF Round#((TBCust2(1).pintamt8)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1990
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).pintamt8
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt8
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


   IF Round#((TBCust2(1).ptaxamt9)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1989")
     TaxTran(1).TaxYear = 1989
     TaxTran(1).TranType = 1
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).ptaxamt9
     TaxTran(1).Revenue.Principle1 = TBCust2(1).ptaxamt9
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt9
     TaxTran(1).Revenue.Penalty = 0
     TaxTran(1).Revenue.Collection = 0
     TaxTran(1).Description = "Tax Bill"
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
    IF Round#((TBCust2(1).pintamt9)) > 0 THEN
     TaxTran(1).TransDate = Date2Num%("08-01-1999")
     TaxTran(1).TaxYear = 1989
     TaxTran(1).TranType = 4
     TaxTran(1).BillType = BillType$
     TaxTran(1).Amount = TBCust2(1).pintamt9
     TaxTran(1).Revenue.Principle1 = 0
     TaxTran(1).Revenue.interest = TBCust2(1).pintamt9
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
    CLOSE TaxTran 'Close Transaction File Here
    RETURN
  

UpdateReal:
    PropertyRec(1).REALPIN = TBCust(1).Acct
    PropertyRec(1).PROPDATE = Date2Num("11-01-1999")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = ""
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = ""
    PropertyRec(1).PROPSIZE = 0
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = TBCust(1).PLand + TBCust(1).PBldg
    PropertyRec(1).EXMPSENI = TBCust(1).UseExp
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = TBCust(1).PDESC
    PropertyRec(1).PROPNOT2 = ""
    PropertyRec(1).PROPNOT3 = "Map/Block# " + TBCust(1).MAP
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
   IF TBCust(1).PBldg + TBCust(1).PLand = 0 THEN
    PersRec(1).PROPPIN = TBCust(1).Acct
   ELSE
    PersRec(1).PROPPIN = "1"
   END IF
   PersRec(1).PROPDATE = Date2Num("11-01-1999")
   PersRec(1).PersVal = PersValue#
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = TBCust(1).UseExp
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
     CURBAL# = TBCust(1).CTax
     CURBAL# = Round#(CURBAL#)
     IF CURBAL# > 99999999 THEN CURBAL# = 0
     PastDue# = (TBCust2(1).ptaxamt0 + TBCust2(1).pintamt0)
     PastDue# = PastDue# + (TBCust2(1).ptaxamt1 + TBCust2(1).pintamt1)
     PastDue# = PastDue# + (TBCust2(1).ptaxamt2 + TBCust2(1).pintamt2)
     PastDue# = PastDue# + (TBCust2(1).ptaxamt3 + TBCust2(1).pintamt3)
     PastDue# = PastDue# + (TBCust2(1).ptaxamt4 + TBCust2(1).pintamt4)
     PastDue# = PastDue# + (TBCust2(1).ptaxamt5 + TBCust2(1).pintamt5)
     PastDue# = PastDue# + (TBCust2(1).ptaxamt6 + TBCust2(1).pintamt6)
     PastDue# = PastDue# + (TBCust2(1).ptaxamt7 + TBCust2(1).pintamt7)
     PastDue# = PastDue# + (TBCust2(1).ptaxamt8 + TBCust2(1).pintamt8)
     PastDue# = PastDue# + (TBCust2(1).ptaxamt9 + TBCust2(1).pintamt9)
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

