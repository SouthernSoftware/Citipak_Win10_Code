DEFINT A-Z
'Real Estate Conversion from Text File Information  PEN GAP VA 2002
'Kill taxprop.dat before running
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
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED REALRec(1) AS PropertyRecType
  
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
DIM A$(39)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile

   CLS
   OPEN "PGREAL.TXT" FOR INPUT AS #11
1 LINE INPUT #11, A1$
  LINE INPUT #11, A2$
  LINE INPUT #11, A3$
  LINE INPUT #11, A4$
  LINE INPUT #11, A5$

    SPIN$ = LEFT$(A1$, 20)
    REALAMT$ = MID$(A1$, 69, 8)
    REALAMT1$ = ""
    FOR JJ = 1 TO LEN(REALAMT$)
     IF MID$(REALAMT$, JJ, 1) <> "," THEN
     REALAMT1$ = REALAMT1$ + MID$(REALAMT$, JJ, 1)
     END IF
    NEXT JJ
    PERSAMT$ = MID$(A1$, 78, 11)
    PERSAMT1$ = ""
    FOR JJ = 1 TO LEN(PERSAMT$)
     IF MID$(PERSAMT$, JJ, 1) <> "," THEN
     PERSAMT1$ = PERSAMT1$ + MID$(PERSAMT$, JJ, 1)
     END IF
    NEXT JJ

   Real# = VAL(REALAMT1$)
   Pers# = VAL(PERSAMT1$)
   tReal# = tReal# + Real#
   tPers# = tPers# + Pers#
   LOCATE 5, 1
   PRINT SPIN$
   'IF VAL(SPIN$) = 0 THEN STOP
   PRINT "REAL: "; USING "##########,#.##"; tReal#
   PRINT "BLDG: "; USING "##########,#.##"; tPers#
   PRINT " TOT: "; USING "##########,#.##"; tReal# + tPers#
   'SLEEP 1

   

   'Find Exiting Acct
   FOR Cnt& = 1 TO LOF(TaxFile) / LEN(TaxCustRec(1))
   GET TaxFile, Cnt&, TaxCustRec(1)
   IF TaxCustRec(1).CountyAcct = VAL(MID$(A1$, 123, 10)) THEN
   Record! = Cnt&
   GOSUB UpdateReal

   GOTO 1
   END IF
   NEXT Cnt&


   'Add New Acct Here
   'Decode Name Here
   anam$ = MID$(A2$, 1, 40)
   KK = INSTR(anam$, " ")
   IF KK = 0 THEN
    LN$ = anam$
    FM$ = ""
    ELSE
    LN$ = LEFT$(anam$, KK - 1)
    FM$ = LTRIM$(MID$(anam$, KK + 1, LEN(anam$) - KK))
   END IF

   SN$ = LN$
   LN$ = UCASE$(LN$)
   FM$ = UCASE$(FM$)
   SN$ = UCASE$(SN$)

   Record! = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1

   TaxCustRec(1).FNAME = FM$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SName = SN$
    
   TaxCustRec(1).Acct = Record!
   TaxCustRec(1).Addr1 = UCASE$(LEFT$(A3$, 40))
   IF LEN(RTRIM$(LEFT$(A5$, 10))) = 0 THEN

   'FIND ZIP
   'ASSUME FIRST CHARACTER 0 OR GREATER IS ZIP START
   FOR JJ = 1 TO 40
    IF ASC(MID$(A4$, JJ, 1)) >= 48 AND ASC(MID$(A4$, JJ, 1)) <= 57 THEN
     ZIPSTART = JJ
     EXIT FOR
    END IF
   NEXT JJ



   TaxCustRec(1).Addr2 = ""
   TaxCustRec(1).CITY = UCASE$(LEFT$(A4$, ZIPSTART - 1))
   TaxCustRec(1).STATE = ""




   TaxCustRec(1).ZIP = MID$(A4$, ZIPSTART, 5)

   ELSE
   TaxCustRec(1).Addr2 = UCASE$(MID$(A4$, 83, 30))
   TaxCustRec(1).CITY = UCASE$(MID$(A$, 113, 30))
   TaxCustRec(1).STATE = ""
   TaxCustRec(1).ZIP = MID$(A$, 144, 5)
    END IF



   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).CountyAcctString = ""
   TaxCustRec(1).CountyAcct = VAL(MID$(A1$, 123, 10))
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
   TaxCustRec(1).OPENDATE = Date2Num%("10-22-2002")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB UpdateReal

NEXTONE:
   GOTO 1
   CLOSE
   EXIT SUB


UpdateReal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PropertyRec(1).RealPin = SPIN$
    PropertyRec(1).PropDate = Date2Num%("10-22-2002")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = ""
    PropertyRec(1).Block = ""
    PropertyRec(1).LotNumb = ""
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PROPSIZE = 0
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = Real#
    PropertyRec(1).EXMPSENI = Pers#
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = MID$(A1$, 41, 25)
    PropertyRec(1).PROPNOT2 = MID$(A4$, 41, 25)
    PropertyRec(1).PROPNOT3 = MID$(A5$, 41, 25)

    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 2001
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    PUT PropTaxFile, RERecord&, PropertyRec(1)
    RETURN



END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, REALTaxFile)
  REALTaxFile = FREEFILE
  OPEN "TAXREAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #REALTaxFile LEN = LEN(REALRec(1))
  NumOfPersRecs = LOF(REALTaxFile) / LEN(REALRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

