  DEFINT A-Z
  'Real Estate Conv from Text File Information  Pen Gap Va from excel .csv file
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
  ON ERROR RESUME NEXT
  CLS
  DIM a$(19)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  CLS
  
  OPEN "real.csv" FOR INPUT SHARED AS #11
  
1
  FOR cc = 1 TO 19
    INPUT #11, a$(cc)
    IF ERR = 62 THEN
      CLOSE
      PRINT "Updated:"; UPCnt, "New:"; NWCnt
      PRINT "Done"
      END
    ELSEIF ERR <> 0 THEN
      PRINT "Error:"; ERR
      CLOSE
      STOP
    END IF

    'PRINT a$(cc), cc
  NEXT cc

  DDCnt = DDCnt + 1
  IF VAL(a$(1)) = 0 THEN
    GOTO 1
  END IF
  PRINT a$(1), a$(3)
  'SLEEP 1
  TOTAL1# = TOTAL1# + VAL(a$(15))
  TOTAL2# = TOTAL2# + VAL(a$(16))
  
  'FIND THE ACCOUNT NOW
  FOR Srch& = 1 TO NumOfTaxRecs
    GET TaxFile, Srch&, TaxCustRec(1)
    IF VAL(a$(1)) = TaxCustRec(1).CountyAcct THEN
      'Account Located Update It Now
      UPCnt = UPCnt + 1
      Record! = Srch&
      'First Update Name and Address
      nm$ = UCASE$(a$(3))
      IF INSTR(nm$, "INC") > 0 THEN
        LN$ = UCASE$(a$(3))
        FM$ = ""
        SN$ = LN$
      ELSE
        kk = INSTR(nm$, " ")
        IF kk > 0 THEN
          LN$ = LEFT$(a$(3), kk - 1)
          FM$ = RIGHT$(a$(3), (LEN(a$(3)) - kk))
          SN$ = LN$
        ELSE
          LN$ = UCASE$(a$(3))
          FM$ = ""
          SN$ = LN$
        END IF
      END IF
      
      TaxCustRec(1).FNAME = LTRIM$(FM$)
      TaxCustRec(1).LName = LTRIM$(LN$)
      TaxCustRec(1).SName = LTRIM$(SN$)
      TaxCustRec(1).Addr1 = LTRIM$(UCASE$(a$(5)))
      TaxCustRec(1).Addr2 = LTRIM$(UCASE$(a$(6)))
      TaxCustRec(1).City = LTRIM$(UCASE$(a$(7)))
      City$ = RTRIM$(TaxCustRec(1).City)
      State$ = RIGHT$(City$, 2)
      IF LEN(City$) > 2 THEN
        City$ = LEFT$(City$, LEN(City$) - 2)
      END IF
      TaxCustRec(1).City = City$
      TaxCustRec(1).State = State$
      TaxCustRec(1).ZIP = a$(8)
      TaxCustRec(1).OPENDATE = Date2Num%("11-23-2004")
      PUT TaxFile, Record!, TaxCustRec(1)
      GOSUB Updatereal
      GOTO 1
    END IF
  NEXT Srch&
  
  'New Account to Add
  NWCnt = NWCnt + 1
  'Decode Name Here
  nm$ = UCASE$(a$(3))
  IF INSTR(nm$, "INC") > 0 THEN
    LN$ = UCASE$(a$(3))
    FM$ = ""
    SN$ = LN$
  ELSE
    kk = INSTR(nm$, " ")
    IF kk > 0 THEN
      LN$ = LEFT$(a$(3), kk - 1)
      FM$ = RIGHT$(a$(3), (LEN(a$(3)) - kk))
      SN$ = LN$
    ELSE
      LN$ = UCASE$(a$(3))
      FM$ = ""
      SN$ = LN$
    END IF
  END IF
  
  Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1
  
  TaxCustRec(1).FNAME = LTRIM$(FM$)
  TaxCustRec(1).LName = LTRIM$(LN$)
  TaxCustRec(1).SName = LTRIM$(SN$)
  TaxCustRec(1).Addr1 = LTRIM$(UCASE$(a$(5)))
  TaxCustRec(1).Addr2 = LTRIM$(UCASE$(a$(6)))
  TaxCustRec(1).City = LTRIM$(UCASE$(a$(7)))
  City$ = RTRIM$(TaxCustRec(1).City)
  State$ = RIGHT$(City$, 2)
  City$ = LEFT$(City$, LEN(City$) - 2)
  TaxCustRec(1).City = City$
  TaxCustRec(1).State = State$
  TaxCustRec(1).ZIP = a$(8)
  
  TaxCustRec(1).HPHONE = ""
  TaxCustRec(1).WPHONE = ""
  TaxCustRec(1).CSSN = ""
  TaxCustRec(1).SSSN = ""
  TaxCustRec(1).CountyAcctString = a$(1)
  TaxCustRec(1).CountyAcct = VAL(a$(1))
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
  TaxCustRec(1).OPENDATE = Date2Num%("10-23-2004")
  PUT TaxFile, Record!, TaxCustRec(1)
  GOSUB Updatereal
NEXTONE:
  GOTO 1
  CLOSE
  STOP
  EXIT SUB
  
  
Updatereal:
  RERecord& = LOF(PropTaxFile) / LEN(RealRec(1)) + 1
  PropertyRec(1).RealPin = a$(10)
  PropertyRec(1).PROPDATE = Date2Num%("11-23-2004")
  PropertyRec(1).GISPOS = ""
  PropertyRec(1).map = ""       'a$(10)
  PropertyRec(1).BLOCK = ""     'a$(16)
  PropertyRec(1).LOTNUMB = ""
  PropertyRec(1).LOTACRE = "A"
  PropertyRec(1).PROPSIZE = VAL(a$(19))
  PropertyRec(1).PROPDISC = "N"
  PropertyRec(1).LATELIST = "N"
  PropertyRec(1).MORTCODE = ""
  PropertyRec(1).PROPVALU = VAL(a$(15))
  PropertyRec(1).EXMPSENI = VAL(a$(16))
  PropertyRec(1).EXMPOTHR = 0
  PropertyRec(1).PROPNOT1 = a$(11)
  PropertyRec(1).PROPNOT2 = a$(12)
  PropertyRec(1).PROPNOT3 = a$(13)
  PropertyRec(1).CustPin = Record!
  PropertyRec(1).NextRec = 0
  PropertyRec(1).LastYrPrinted = 2003
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

