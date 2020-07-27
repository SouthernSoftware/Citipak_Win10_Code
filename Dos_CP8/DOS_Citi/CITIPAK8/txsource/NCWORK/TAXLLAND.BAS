DEFINT A-Z
'Conversion for Leland NC From County 2003

'$INCLUDE: 'DefCnf.BI'
DECLARE FUNCTION Date2Num% (TheDate$)
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
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Round# (B#)

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
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  DIM a$(120)
  
  CALL KillFile("TAXPERS.DAT")

  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  DateKey = Date2Num%(DATE$)
  
initfiles:
  FOR Cnt& = 1 TO NumOfPropRecs
    GET PropTaxFile, Cnt&, PropertyRec(1)
    PropertyRec(1).PROPVALU = 0
    PropertyRec(1).EXMPSENI = 0
    PUT PropTaxFile, Cnt&, PropertyRec(1)
  NEXT Cnt&
  
  OPEN "LELAND5.TXT" FOR INPUT AS #11
  CLS
  PRINT "Importing Tax Information. . ."
getsomemore:
  IF EOF(11) THEN
    CLOSE
    LOCATE 12, 1
    PRINT "Processing Complete."
    CALL KillFile("taxlland.exe")
    END
  END IF
  LINE INPUT #11, a$
  IF MID$(a$, 319, 1) <> " " THEN
    GOTO getsomemore
  END IF
  
  c = c + 1
  
  'IF c = 525 THEN STOP
  
  acct$ = LEFT$(a$, 8)
  'IF acct$ = "10479950" THEN STOP
  
  LOCATE 3, 1
  PRINT acct$, c;
  
  N$ = MID$(a$, 9, 35)
  a1$ = QPTrim$(MID$(a$, 44, 25))
  a2$ = QPTrim$(MID$(a$, 69, 25))
  IF LEN(a1$) = 0 THEN
    a1$ = a2$
    a2$ = ""
  END IF
  CY$ = MID$(a$, 94, 18)
  ST$ = MID$(a$, 112, 2)
  ZP$ = MID$(a$, 114, 5)
  Pin$ = MID$(a$, 119, 15)
  D1$ = MID$(a$, 134, 20)
  d2$ = MID$(a$, 154, 20)
  d3$ = MID$(a$, 174, 10)
  BK$ = MID$(a$, 184, 4)
  PG$ = MID$(a$, 188, 4)
  RealVal$ = MID$(a$, 228, 9)
  
  
  IF VAL(MID$(a$, 319, 1)) > 0 THEN RealVal$ = "0": STOP
  r# = r# + VAL(RealVal$)
  L$ = MID$(a$, 237, 1)
  SZ$ = MID$(a$, 238, 7)
  SIZE! = VAL(SZ$)
  SIZE! = SIZE! / 100
  PersVal$ = MID$(a$, 299, 9)
  'TOP
  p# = p# + VAL(PersVal$)
  SrCit$ = MID$(a$, 309, 6)
  IF VAL(SrCit$) > VAL(RealVal$) THEN SrCit$ = RealVal$
  d# = d# + VAL(SrCit$)
  Record = Record + 1
  'Split Name Here
  IF c >= 175 THEN
    kk = INSTR(N$, " ")
    IF kk > 0 THEN
      LN$ = LEFT$(N$, kk - 1)
      SN$ = LEFT$(N$, kk - 1)
      FF1$ = MID$(N$, kk + 1, LEN(N$) - (kk + 1))
    ELSE
      LN$ = QPTrim$(N$)
      SN$ = QPTrim$(N$)
      FF1$ = ""
    END IF
  ELSE
    LN$ = QPTrim$(N$)
    SN$ = QPTrim$(N$)
    FF1$ = ""
    
  END IF
  
  
  'Search for Existing Record
  'Pin Search
  'IF QPTrim$(Pin$) = "037OB018" THEN STOP
  
  FOR Cnt& = 1 TO NumOfPropRecs
    GET PropTaxFile, Cnt&, PropertyRec(1)
    'IF QPTrim$(PropertyRec(1).RealPin) = "037OB018" AND QPTrim$(Pin$) = "037OB018" THEN STOP

    IF QPTrim$(PropertyRec(1).RealPin) = QPTrim$(Pin$) AND PropertyRec(1).CustPin > 0 THEN
      CustRecord! = PropertyRec(1).CustPin
      
      GET TaxFile, CustRecord!, TaxCustRec(1)
      IF VAL(acct$) = TaxCustRec(1).CountyAcct AND DateKey <> TaxCustRec(1).OpenDate THEN
        TaxCustRec(1).OpenDate = Date2Num%(DATE$)
        TaxCustRec(1).FName = FF1$
        TaxCustRec(1).LName = LN$
        TaxCustRec(1).SName = LN$
        TaxCustRec(1).Addr1 = a1$
        TaxCustRec(1).Addr2 = a2$
        TaxCustRec(1).City = CY$
        TaxCustRec(1).State = ST$
        TaxCustRec(1).Zip = ZP$
        TaxCustRec(1).FirstPropRec = 0
        TaxCustRec(1).FirstPersRec = 0
        PUT TaxFile, CustRecord!, TaxCustRec(1)
        Record = CustRecord!
        IF VAL(RealVal$) > 0 THEN
          PropertyRec(1).RealPin = Pin$
          PropertyRec(1).PROPDATE = Date2Num%(DATE$)
          PropertyRec(1).GISPOS = ""
          PropertyRec(1).MAP = BK$
          PropertyRec(1).BLOCK = PG$
          PropertyRec(1).LOTNUMB = ""
          PropertyRec(1).LOTACRE = L$
          PropertyRec(1).PROPSIZE = SIZE!
          PropertyRec(1).PROPDISC = "N"
          PropertyRec(1).LATELIST = "N"
          PropertyRec(1).MORTCODE = ""
          PropertyRec(1).PROPVALU = VAL(RealVal$)
          PropertyRec(1).EXMPSENI = VAL(SrCit$)
          PropertyRec(1).EXMPOTHR = 0
          PropertyRec(1).PROPNOT1 = D1$
          PropertyRec(1).PROPNOT2 = d2$
          PropertyRec(1).PROPNOT3 = d3$
          PropertyRec(1).CustPin = Record
          PropertyRec(1).NextRec = 0
          PropertyRec(1).LastYrPrinted = 0
          PropertyRec(1).Deleted = 0
          PropertyRec(1).Blank = ""
          PUT PropTaxFile, Cnt&, PropertyRec(1)
          TaxCustRec(1).FirstPropRec = NextRecord
          PUT TaxFile, Record, TaxCustRec(1)
        END IF
        IF VAL(PersVal$) > 0 THEN
          GOSUB UpdatePers
        END IF
        GOTO getsomemore
      END IF
    END IF
  NEXT Cnt&
zz:
  
  
  '**************************
  IF VAL(PersVal$) > 0 AND VAL(RealVal$) = 0 THEN
    FOR Cnt& = 1 TO NumOfTaxRecs
      GET TaxFile, Cnt&, TaxCustRec(1)
      
      IF VAL(acct$) = TaxCustRec(1).CountyAcct THEN
        TaxCustRec(1).OpenDate = Date2Num%(DATE$)
        TaxCustRec(1).FName = FF1$
        TaxCustRec(1).LName = LN$
        TaxCustRec(1).SName = LN$
        TaxCustRec(1).Addr1 = a1$
        TaxCustRec(1).Addr2 = a2$
        TaxCustRec(1).City = CY$
        TaxCustRec(1).State = ST$
        TaxCustRec(1).Zip = ZP$
        TaxCustRec(1).FirstPropRec = 0
        TaxCustRec(1).FirstPersRec = 0
        PUT TaxFile, Cnt&, TaxCustRec(1)
        Record = Cnt&
        GOSUB UpdatePers
        GOTO getsomemore
      END IF
    NEXT Cnt&
  END IF
  '*****************************
  
  'Add Account Here
  Record = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1
  TaxCustRec(1).acct = Record
  TaxCustRec(1).OpenDate = Date2Num%(DATE$)
  TaxCustRec(1).FName = FF1$
  TaxCustRec(1).LName = LN$
  TaxCustRec(1).SName = LN$
  TaxCustRec(1).Addr1 = a1$
  TaxCustRec(1).Addr2 = a2$
  TaxCustRec(1).City = CY$
  TaxCustRec(1).State = ST$
  TaxCustRec(1).CountyAcct = VAL(acct$)
  TaxCustRec(1).CountyAcctString = acct$
  TaxCustRec(1).Zip = ZP$
  TaxCustRec(1).HPhone = ""
  TaxCustRec(1).WPhone = ""
  TaxCustRec(1).CSSN = ""
  TaxCustRec(1).SSSN = ""
  TaxCustRec(1).Active = "Y"
  TaxCustRec(1).Interest = "Y"
  TaxCustRec(1).TaxExempt = "N"
  TaxCustRec(1).Penalty = "N"
  TaxCustRec(1).LastTrans = 0
  TaxCustRec(1).FirstPropRec = 0
  TaxCustRec(1).FirstPersRec = 0
  TaxCustRec(1).Pin = Record
  TaxCustRec(1).Deleted = 0
  TaxCustRec(1).FileVer = 8
  PUT TaxFile, Record, TaxCustRec(1)
  
  IF VAL(RealVal$) > 0 THEN
    GOSUB UpdateReal
  END IF
  IF VAL(PersVal$) > 0 THEN
    GOSUB UpdatePers
  END IF
  GOTO getsomemore
  
  CLOSE
  EXIT SUB
  

UpdateReal:
  IF VAL(RealVal$) = 0 THEN RETURN
  PropertyRec(1).RealPin = Pin$
  PropertyRec(1).PROPDATE = Date2Num%(DATE$)
  PropertyRec(1).GISPOS = ""
  PropertyRec(1).MAP = BK$
  PropertyRec(1).BLOCK = PG$
  PropertyRec(1).LOTNUMB = ""
  PropertyRec(1).LOTACRE = L$
  PropertyRec(1).PROPSIZE = SIZE!
  PropertyRec(1).PROPDISC = "N"
  PropertyRec(1).LATELIST = "N"
  PropertyRec(1).MORTCODE = ""
  PropertyRec(1).PROPVALU = VAL(RealVal$)
  PropertyRec(1).EXMPSENI = VAL(SrCit$)
  PropertyRec(1).EXMPOTHR = 0
  PropertyRec(1).PROPNOT1 = D1$
  PropertyRec(1).PROPNOT2 = d2$
  PropertyRec(1).PROPNOT3 = d3$
  PropertyRec(1).CustPin = Record
  PropertyRec(1).NextRec = 0
  PropertyRec(1).LastYrPrinted = 0
  PropertyRec(1).Deleted = 0
  PropertyRec(1).Blank = ""
  NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
  PUT PropTaxFile, NextRecord, PropertyRec(1)
  TaxCustRec(1).FirstPropRec = NextRecord
  PUT TaxFile, Record, TaxCustRec(1)
  
  RETURN

UpdatePers:
  SrCit$ = MID$(a$, 319, 6)
  'IF VAL(SrCit$) > 0 THEN STOP
  IF VAL(PersVal$) = 0 THEN RETURN
  PersRec(1).PROPPIN = "1"
  PersRec(1).PROPDATE = Date2Num%(DATE$)
  PersRec(1).PersVal = VAL(PersVal$)
  PersRec(1).MHVALUE = 0
  PersRec(1).MCVALUE = 0
  PersRec(1).CVALUE = 0
  PersRec(1).MTVALUE = 0
  IF VAL(SrCit$) > 0 THEN
    PersRec(1).EXMPSENI = VAL(SrCit$)
  ELSE
    PersRec(1).EXMPSENI = 0
  END IF
  
  PersRec(1).EXMPOTHR = 0
  PersRec(1).DISCOV = "N"
  PersRec(1).LATELIST = "N"
  PersRec(1).DESC1 = ""
  PersRec(1).DESC2 = ""
  PersRec(1).DESC3 = ""
  PersRec(1).DESC4 = ""
  PersRec(1).DESC5 = ""
  PersRec(1).CustPin = Record
  PersRec(1).NextRec = 0
  PersRec(1).LastYrPrinted = 0
  PersRec(1).Deleted = 0
  PersRec(1).Blank = ""
  NextRec = (LOF(PersTaxFile) / LEN(PersRec(1))) + 1
  PUT PersTaxFile, NextRec, PersRec(1)
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

