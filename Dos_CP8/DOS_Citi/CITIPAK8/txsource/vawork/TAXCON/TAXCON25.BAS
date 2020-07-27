DEFINT A-Z
'Real Estate Conversion from Text File Information ELKTON REAL 2003
'KILL TAXPROP AND RUN THE RELINK BEFORE STARTING CONVERSION
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NUMOFTAXRECS%, TAXFILE%)
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
  
  
  DIM SHARED TAXCUSTREC(1) AS TaxCustType
  DIM SHARED TaxTrans(1) AS TaxTransactionType
  DIM SHARED PROPERTYRec(1) AS PropertyRecType
  DIM SHARED REALRec(1) AS PropertyRecType
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS

  OpenTaxCustFile NUMOFTAXRECS, TAXFILE
  OpenTaxPropFile NumOfPropRecs, PropTaxFile




        CLS
        OPEN "ELKRE.TXT" FOR INPUT AS #11
1 LINE INPUT #11, a$
        t = t + 1
        DESC$ = MID$(a$, 183, 30)
        DESC$ = RTRIM$(DESC$)
        LNME$ = MID$(a$, 33, 5)
        LNME$ = RTRIM$(LNME$)
        LandVal$ = MID$(a$, 297, 10)
        BldgVal$ = MID$(a$, 307, 10)
        
        TLand# = TLand# + VAL(LandVal$)
        TBldg# = TBldg# + VAL(BldgVal$)
       
        NM$ = MID$(a$, 33, 34)
        ADD1$ = MID$(a$, 68, 34)
        ADD2$ = MID$(a$, 103, 35)
        Zip$ = MID$(a$, 139, 5)
        city$ = MID$(a$, 148, 35)
        PDESC$ = MID$(a$, 183, 30)

                          
             
'Search For Record
 FOR Cnt& = 1 TO LOF(TAXFILE) / LEN(TAXCUSTREC(1))
  GET TAXFILE, Cnt&, TAXCUSTREC(1)
  IF RTRIM$(TAXCUSTREC(1).CountyAcctString) = RTRIM$(MID$(a$, 14, 16)) AND TAXCUSTREC(1).FirstPropRec = 0 AND RTRIM$(ADD2$) = RTRIM$(TAXCUSTREC(1).ADDR1) THEN
   Record! = Cnt&
   kk = INSTR(NM$, " ")
   IF kk = 0 THEN
    LN$ = NM$
    fm$ = ""
    ELSE
    LN$ = LEFT$(NM$, kk - 1)
    fm$ = LTRIM$(MID$(NM$, kk + 1, LEN(NM$) - kk))
   END IF
    SN$ = LN$
    LN$ = UCASE$(LN$)
    fm$ = UCASE$(fm$)
    SN$ = UCASE$(SN$)

    kk1 = INSTR(city$, "  ")
    IF kk1 > 0 THEN city$ = LEFT$(city$, kk1 - 1)

    TAXCUSTREC(1).FName = fm$
    TAXCUSTREC(1).LNAME = LN$
    TAXCUSTREC(1).SName = SN$
   IF LEN(RTRIM$(ADD1$)) = 0 THEN
    TAXCUSTREC(1).ADDR1 = UCASE$(ADD2$)
    TAXCUSTREC(1).ADDR2 = ""
    ELSE
    TAXCUSTREC(1).ADDR1 = UCASE$(ADD1$)
    TAXCUSTREC(1).ADDR2 = UCASE$(ADD2$)
   END IF

   TAXCUSTREC(1).city = UCASE$(city$)
   TAXCUSTREC(1).State = "VA"
   TAXCUSTREC(1).Zip = Zip$
   TAXCUSTREC(1).CSSN = MID$(a$, 10, 4)
   TAXCUSTREC(1).TownShip = "Y"
   TAXCUSTREC(1).CountyAcctString = RTRIM$(MID$(a$, 14, 16))
   TAXCUSTREC(1).FirstPropRec = 1
   TAXCUSTREC(1).OPENDATE = Date2Num%("10-27-2003")
   PUT TAXFILE, Record!, TAXCUSTREC(1)
   GOSUB UpdateReal
  GOTO 1
  END IF
 
  
  
NEXT Cnt&

'Add Here CUSTOMER NOT LOCATED


  
   'Decode Name Here
   Record! = LOF(TAXFILE) / LEN(TAXCUSTREC(1)) + 1
   

   kk = INSTR(NM$, " ")
   IF kk = 0 THEN
    LN$ = NM$
    fm$ = ""
    ELSE
    LN$ = LEFT$(NM$, kk - 1)
    fm$ = LTRIM$(MID$(NM$, kk + 1, LEN(NM$) - kk))
   END IF
    SN$ = LN$
    LN$ = UCASE$(LN$)
    fm$ = UCASE$(fm$)
    SN$ = UCASE$(SN$)

    kk1 = INSTR(city$, "  ")
    IF kk1 > 0 THEN city$ = LEFT$(city$, kk1 - 1)

    Record! = (LOF(TAXFILE) / LEN(TAXCUSTREC(1))) + 1
    TAXCUSTREC(1).FName = fm$
    TAXCUSTREC(1).LNAME = LN$
    TAXCUSTREC(1).SName = SN$
    TAXCUSTREC(1).Acct = Record!
   
   IF LEN(RTRIM$(ADD1$)) = 0 THEN
    TAXCUSTREC(1).ADDR1 = UCASE$(ADD2$)
    TAXCUSTREC(1).ADDR2 = ""
    ELSE
    TAXCUSTREC(1).ADDR1 = UCASE$(ADD1$)
    TAXCUSTREC(1).ADDR2 = UCASE$(ADD2$)
   END IF

   TAXCUSTREC(1).city = UCASE$(city$)
   TAXCUSTREC(1).State = "VA"
   TAXCUSTREC(1).Zip = Zip$
   TAXCUSTREC(1).HPHONE = ""
   TAXCUSTREC(1).WPHONE = ""
   TAXCUSTREC(1).CSSN = MID$(a$, 10, 4)
   TAXCUSTREC(1).SSSN = ""
   TAXCUSTREC(1).CountyAcctString = RTRIM$(MID$(a$, 14, 16))
   TAXCUSTREC(1).CountyAcct = 0
   TAXCUSTREC(1).TownShip = "Y"
   TAXCUSTREC(1).Active = "Y"
   TAXCUSTREC(1).Interest = "Y"
   TAXCUSTREC(1).TaxExempt = "N"
   TAXCUSTREC(1).Penalty = "Y"
   TAXCUSTREC(1).LastTrans = 0
   TAXCUSTREC(1).FirstPropRec = 0
   TAXCUSTREC(1).FirstPersRec = 0
   TAXCUSTREC(1).PIN = Record!
   TAXCUSTREC(1).Deleted = 0
   TAXCUSTREC(1).FileVer = 8
   TAXCUSTREC(1).OPENDATE = Date2Num%("10-27-2003")
   PUT TAXFILE, Record!, TAXCUSTREC(1)
   PropNumb = 1
   GOSUB UpdateReal
   GOTO 1

NEXTONE:
   
   CLOSE
   EXIT SUB


UpdateReal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PROPERTYRec(1).REALPIN = "B" + LTRIM$(MID$(a$, 10, 4))
    PROPERTYRec(1).PROPDATE = Date2Num%(DATE$)
    PROPERTYRec(1).GISPOS = ""
    PROPERTYRec(1).MAP = ""
    PROPERTYRec(1).BLOCK = ""
    PROPERTYRec(1).LOTNUMB = ""
    PROPERTYRec(1).LOTACRE = "A"
    PROPERTYRec(1).PROPSIZE = VAL(MID$(a$, 287, 9)) / 100
    PROPERTYRec(1).PROPDISC = "N"
    PROPERTYRec(1).LATELIST = "N"
    PROPERTYRec(1).MORTCODE = ""
    PROPERTYRec(1).PROPVALU = VAL(LandVal$)
    PROPERTYRec(1).EXMPSENI = VAL(BldgVal$)
    PROPERTYRec(1).EXMPOTHR = 0
    PROPERTYRec(1).PROPNOT1 = DESC$
    PROPERTYRec(1).PROPNOT2 = ""
    PROPERTYRec(1).PROPNOT3 = ""
    
    PROPERTYRec(1).CustPin = Record!
    PROPERTYRec(1).NextRec = 0
    PROPERTYRec(1).LastYrPrinted = 2000
    PROPERTYRec(1).Deleted = 0
    PROPERTYRec(1).Blank = ""
    PUT PropTaxFile, RERecord&, PROPERTYRec(1)
    RETURN
  
 END SUB

SUB OpenTaxCustFile (NUMOFTAXRECS, TAXFILE)
  
  TAXFILE = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TAXFILE LEN = LEN(TAXCUSTREC(1))
  NUMOFTAXRECS = LOF(TAXFILE) / LEN(TAXCUSTREC(1))
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, REALTaxFile)
  REALTaxFile = FREEFILE
  OPEN "TAXREAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #REALTaxFile LEN = LEN(REALRec(1))
  NumOfPersRecs = LOF(REALTaxFile) / LEN(REALRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PROPERTYRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PROPERTYRec(1))
END SUB

