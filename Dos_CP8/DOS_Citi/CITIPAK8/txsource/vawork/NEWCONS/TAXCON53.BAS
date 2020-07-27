DEFINT A-Z
'Real Estate Conv from Text File Information  Saltville Va from excel .csv file 2003
'real estate from smyth cty
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
  DIM SHARED RealRec(1) AS PropertyRecType
  
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing

  CLS
  DIM a$(-1 TO 112)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  CLS
  
  OPEN "re2004.csv" FOR INPUT AS #11
1 FOR cc = 0 TO 34
   INPUT #11, a$(cc)
   PRINT a$(cc), cc
  NEXT cc
   PRINT a$(1)
   'SLEEP 3
    
   total1# = total1# + VAL(a$(27))
   total2# = total2# + VAL(a$(28))
   total3# = total3# + VAL(a$(29))

   BldgVal# = VAL(a$(27))
   IF VAL(a$(28)) > 0 THEN
    LandVal# = VAL(a$(28))
    ELSE
    LandVal# = VAL(a$(29))
   END IF
   
  PRINT total1#, total2#, total3#

'FIND THE ACCOUNT NOW
  FOR Srch& = 1 TO NumOfTaxRecs
   GET TaxFile, Srch&, TaxCustRec(1)
  IF "SC" + a$(6) = RTRIM$(TaxCustRec(1).CountyAcctString) THEN
    'Account Located Update It Now
     Record! = Srch&
 'First Update Name and Address
     nm$ = UCASE$(a$(10))
    IF INSTR(nm$, "INC") > 0 THEN
     LN$ = UCASE$(a$(10))
     FM$ = ""
     SN$ = LN$
    ELSE
    kk = INSTR(nm$, " ")
    IF kk > 0 THEN
     LN$ = LEFT$(a$(10), kk - 1)
     FM$ = RIGHT$(a$(10), (LEN(a$(10)) - kk))
     SN$ = LN$
      ELSE
     LN$ = UCASE$(a$(10))
     FM$ = ""
     SN$ = LN$
    END IF
   END IF

   TaxCustRec(1).FNAME = LTRIM$(FM$)
   TaxCustRec(1).LName = LTRIM$(LN$)
   TaxCustRec(1).SName = LTRIM$(SN$)
   IF LEN(RTRIM$(a$(14))) = 0 THEN
    TaxCustRec(1).Addr1 = LTRIM$(UCASE$(a$(12)))
    TaxCustRec(1).Addr2 = ""
    ELSE
    TaxCustRec(1).Addr1 = LTRIM$(UCASE$(a$(12)))
    TaxCustRec(1).Addr2 = LTRIM$(UCASE$(a$(13)))
   END IF
   TaxCustRec(1).City = LTRIM$(UCASE$(a$(14)))
   City$ = RTRIM$(TaxCustRec(1).City)
   State$ = RIGHT$(City$, 2)
   IF LEN(City$) > 0 THEN City$ = LEFT$(City$, LEN(City$) - 2)
   TaxCustRec(1).City = City$
   TaxCustRec(1).State = State$
   TaxCustRec(1).ZIP = a$(15)
   TaxCustRec(1).OPENDATE = Date2Num%(DATE$)
   PUT TaxFile, Record!, TaxCustRec(1)
   GOSUB Updatereal
   GOTO 1
  END IF
   NEXT Srch&

   'New Account to A
  'Decode Name Here
     nm$ = UCASE$(a$(10))
    IF INSTR(nm$, "INC") > 0 THEN
     LN$ = UCASE$(a$(10))
     FM$ = ""
     SN$ = LN$
    ELSE
    kk = INSTR(nm$, ",")
    IF kk > 0 THEN
     LN$ = LEFT$(a$(10), kk - 1)
     FM$ = RIGHT$(a$(10), (LEN(a$(10)) - kk))
     SN$ = LN$
   ELSE
     LN$ = UCASE$(a$(10))
     FM$ = ""
     SN$ = LN$
    END IF
   END IF

   Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1

   TaxCustRec(1).FNAME = LTRIM$(FM$)
   TaxCustRec(1).LName = LTRIM$(LN$)
   TaxCustRec(1).SName = LTRIM$(SN$)
   IF LEN(RTRIM$(a$(12))) = 0 THEN
    TaxCustRec(1).Addr1 = LTRIM$(UCASE$(a$(13)))
    TaxCustRec(1).Addr2 = ""
    TaxCustRec(1).City = LTRIM$(UCASE$(a$(14)))
    ELSE
    TaxCustRec(1).Addr1 = LTRIM$(UCASE$(a$(12)))
    TaxCustRec(1).Addr2 = LTRIM$(UCASE$(a$(13)))
    TaxCustRec(1).City = LTRIM$(UCASE$(a$(14)))
   END IF
   
   City$ = RTRIM$(TaxCustRec(1).City)
   State$ = RIGHT$(City$, 2)
   City$ = LEFT$(City$, LEN(City$) - 2)
   TaxCustRec(1).City = City$
   TaxCustRec(1).State = State$
   TaxCustRec(1).ZIP = a$(15)
   
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).CountyAcctString = "SC" + a$(6)
   TaxCustRec(1).CountyAcct = VAL(a$(6))
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
   TaxCustRec(1).OPENDATE = Date2Num%("11-05-2003")
   PUT TaxFile, Record!, TaxCustRec(1)
   GOSUB Updatereal
   GOTO 1
   
   STOP
   EXIT SUB


Updatereal:
    RERecord& = LOF(PropTaxFile) / LEN(RealRec(1)) + 1
    PropertyRec(1).RealPin = a$(33)
    PropertyRec(1).PROPDATE = Date2Num%("11-05-2003")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).map = a$(33)
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PROPSIZE = 0
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = LandVal#
    PropertyRec(1).EXMPSENI = BldgVal#
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = a$(23)
    PropertyRec(1).PROPNOT2 = a$(24)
    PropertyRec(1).PROPNOT3 = a$(25)
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
  OPEN "TAXREAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #REALTaxFile LEN = LEN(RealRec(1))
  NumOfPersRecs = LOF(REALTaxFile) / LEN(RealRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

