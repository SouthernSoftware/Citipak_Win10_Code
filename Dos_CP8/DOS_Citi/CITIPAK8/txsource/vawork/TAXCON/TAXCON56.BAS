DEFINT A-Z
'Real Estate Conversion from Text File Information  Honaker VA 2001
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

   OPEN "thre.dat" FOR INPUT AS #11
1 LINE INPUT #11, a$
  
   nme$ = RTRIM$(LEFT$(a$, 35))
   add1$ = RTRIM$(MID$(a$, 36, 30))
   add2$ = RTRIM$(MID$(a$, 66, 30))
   Add3$ = RTRIM$(MID$(a$, 96, 30))

   IF LEN(Add3$) > 0 THEN
    City$ = LEFT$(Add3$, LEN(Add3$) - 2)
   ELSE
    City$ = ""
   END IF

   zp$ = MID$(a$, 127, 5)

   LandVal# = VAL(MID$(a$, 132, 10))
   BldgVal# = VAL(MID$(a$, 142, 10))
   UseVal# = VAL(MID$(a$, 152, 10))
   TotVal# = VAL(MID$(a$, 162, 11))
   acres$ = MID$(a$, 173, 7)
   map$ = MID$(a$, 180, 45)
   class$ = MID$(a$, 226, 1)

   

  tl# = tl# + LandVal#
  tB# = tB# + BldgVal#

   

   PRINT nme$
   PRINT "REAL: "; USING "##########,#.##"; tl#
   PRINT "BLDG: "; USING "##########,#.##"; tB#
   PRINT " TOT: "; USING "##########,#.##"; tB# + tl#
   
   'Add New Acct Here
   'Decode Name Here
   anam$ = nme$
   kk = INSTR(anam$, " ")
   IF kk = 0 THEN
    LN$ = anam$
    FM$ = ""
    ELSE
    LN$ = LEFT$(anam$, kk - 1)
    FM$ = LTRIM$(MID$(anam$, kk + 1, LEN(anam$) - kk))
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
   IF LEN(RTRIM$(add1$)) = 0 THEN
   TaxCustRec(1).Addr1 = add2$
   TaxCustRec(1).Addr2 = ""
   TaxCustRec(1).City = City$
   ELSE
   TaxCustRec(1).Addr1 = add1$
   TaxCustRec(1).Addr2 = add2$
   TaxCustRec(1).City = City$

   END IF
   TaxCustRec(1).STATE = RIGHT$(RTRIM$(Add3$), 2)
   TaxCustRec(1).ZIP = zp$
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).CountyAcctString = RTRIM$(map$)
   TaxCustRec(1).CountyAcct = 0
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
   TaxCustRec(1).OPENDATE = Date2Num%("01-01-2001")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   GOSUB updatereal

NEXTONE:
   GOTO 1
   CLOSE
   EXIT SUB


updatereal:
    RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
    PropertyRec(1).RealPin = RTRIM$(map$)
    PropertyRec(1).PropDate = Date2Num%("07-01-2000")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).map = ""
    PropertyRec(1).Block = ""
    PropertyRec(1).LotNumb = ""
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PROPSIZE = VAL(acres$)
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = LandVal#
    PropertyRec(1).EXMPSENI = BldgVal#
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = "Class Code " + class$
    PropertyRec(1).PROPNOT2 = ""
    PropertyRec(1).PROPNOT3 = ""

    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 2000
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    PUT PropTaxFile, RERecord&, PropertyRec(1)
    RETURN
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

