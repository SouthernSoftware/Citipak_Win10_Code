DEFINT A-Z
'Real Estate Conversion from Text File Information  CHILHOWIE VA 2003
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
DIM a$(36)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile

   CLS
   OPEN "re2003c.dat" FOR INPUT AS #11

1 LINE INPUT #11, a$
   class = VAL(MID$(a$, 359, 1))
   IF class = 1 OR class = 3 OR class = 4 OR class = 5 THEN
    ELSE
    GOTO 1
   END IF

   Luse# = VAL(MID$(a$, 360, 10))
   Real# = VAL(MID$(a$, 371, 10))
   Bldg# = VAL(MID$(a$, 381, 10))
   

   IF Luse# > 0 THEN
    Real# = Luse#
    ELSE
    'Use Real Value as computed
   END IF

   TReal# = TReal# + Real#
   TPers# = TPers# + Bldg#


   Cacct$ = QPTrim$(MID$(a$, 1, 10))
   
   
   PRINT "REAL: "; USING "##########,#.##"; TReal#
   PRINT "BLDG: "; USING "##########,#.##"; TPers#
   PRINT " TOT: "; USING "##########,#.##"; TReal# + TPers# - Te#
  
   'Find Exiting Acct

                                   
   FOR Snt& = 1 TO NumOfTaxRecs
   GET TaxFile, Snt&, TaxCustRec(1)
    IF QPTrim$(TaxCustRec(1).CountyAcctString) = Cacct$ THEN
     Record! = Snt&
       TaxCustRec(1).OPENDATE = Date2Num%("10-14-2003")
      ANAM$ = RTRIM$(MID$(a$, 13, 35))
      KK = INSTR(ANAM$, ",")
      IF KK > 0 THEN
         LN$ = LEFT$(ANAM$, KK - 1)
         FM$ = LTRIM$(MID$(ANAM$, KK + 1, LEN(ANAM$) - KK))
       ELSE
         LN$ = ANAM$
         FM$ = ""
      END IF
      
   SN$ = LN$
   LN$ = UCASE$(LN$)
   FM$ = UCASE$(FM$)
   SN$ = UCASE$(SN$)
      
   TaxCustRec(1).FName = FM$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SName = SN$

     
     IF LEN(RTRIM$(MID$(a$, 83, 30))) = 0 THEN
       TaxCustRec(1).Addr1 = MID$(a$, 113, 30)
       TaxCustRec(1).Addr2 = ""
       City$ = MID$(a$, 142, 30)
     ELSE
       TaxCustRec(1).Addr1 = MID$(a$, 83, 30)
       TaxCustRec(1).Addr2 = MID$(a$, 113, 30)
      City$ = MID$(a$, 142, 30)
     END IF
     City$ = QPTrim$(City$)
     State$ = RIGHT$(City$, 2)
     City$ = LEFT$(City$, LEN(City$) - 2)

     
     TaxCustRec(1).City = City$
     TaxCustRec(1).State = State$
     TaxCustRec(1).Zip = MID$(a$, 174, 5)
     PUT TaxFile, Snt&, TaxCustRec(1)
     GOSUB updatereal
     GOTO 1
    END IF
   NEXT Snt&
   
   'Add New Acct Here
   'Decode Name Here
   ANAM$ = RTRIM$(MID$(a$, 13, 35))
   KK = INSTR(ANAM$, ",")

   IF KK > 0 THEN
    LN$ = LEFT$(ANAM$, KK - 1)
    FM$ = LTRIM$(MID$(ANAM$, KK + 1, LEN(ANAM$) - KK))
    ELSE
    LN$ = ANAM$
    FM$ = ""
   END IF
   

   SN$ = LN$
   LN$ = UCASE$(LN$)
   FM$ = UCASE$(FM$)
   SN$ = UCASE$(SN$)

   Record! = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1

   TaxCustRec(1).FName = FM$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SName = SN$

     
     IF LEN(RTRIM$(MID$(a$, 83, 30))) = 0 THEN
       TaxCustRec(1).Addr1 = MID$(a$, 113, 30)
       TaxCustRec(1).Addr2 = ""
       City$ = MID$(a$, 142, 30)
     ELSE
       TaxCustRec(1).Addr1 = MID$(a$, 83, 30)
       TaxCustRec(1).Addr2 = MID$(a$, 113, 30)
      City$ = MID$(a$, 142, 30)
     END IF
     City$ = QPTrim$(City$)
     State$ = RIGHT$(City$, 2)
     City$ = LEFT$(City$, LEN(City$) - 2)




     TaxCustRec(1).Addr2 = ""
     TaxCustRec(1).City = City$
     TaxCustRec(1).State = State$
     TaxCustRec(1).Zip = MID$(a$, 174, 5)
     TaxCustRec(1).CountyAcctString = Cacct$
    
    
   TaxCustRec(1).Acct = Record!

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
   TaxCustRec(1).PIN = Record!
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("10-14-2003")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB updatereal
NEXTONE:
   GOTO 1
   CLOSE
   EXIT SUB


updatereal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PropertyRec(1).RealPin = QPTrim$(MID$(a$, 285, 45))
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = ""
    PropertyRec(1).Block = ""
    PropertyRec(1).LotNumb = ""
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PropSize = 0
    PropertyRec(1).PROPVALU = Real#
    PropertyRec(1).EXMPSENI = Bldg#
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PropNot1 = MID$(a$, 185, 25)
    PropertyRec(1).PropNot2 = MID$(a$, 210, 25)
    PropertyRec(1).PropNot3 = MID$(a$, 235, 25)
    PropertyRec(1).PropDate = Date2Num%("10-14-2003")
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 2002
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

