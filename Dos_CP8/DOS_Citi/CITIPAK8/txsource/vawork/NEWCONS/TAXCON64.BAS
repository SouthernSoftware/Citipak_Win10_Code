DEFINT A-Z
'Convert Master PP Records from TXT File (Chilhowi 2000)
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
DECLARE FUNCTION Round# (b#)
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
  
  
  'New Market Va Tax PP Conversion of Master (Customers)
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  
  DIM SHARED taxcustrec(1) AS TaxCustType
  DIM SHARED PersRec(1) AS PersonalRecType
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  ON ERROR RESUME NEXT
  CLS
  DIM a$(55)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  
  OPEN "CPPMAST.DAT" FOR INPUT AS #11
  CLS
  b = 0
1
  LINE INPUT #11, a$
  IF ERR = 62 THEN
    CLOSE
    PRINT "Done."
    END
  END IF
  acct$ = MID$(a$, 6, 8)
  Bus$ = MID$(a$, 18, 1)
  ANAME$ = MID$(a$, 19, 35)
  add1$ = MID$(a$, 54, 30)
  add2$ = MID$(a$, 84, 30)
  add3$ = MID$(a$, 114, 30)
  zip$ = MID$(a$, 145, 5)
  SS$ = MID$(a$, 174, 9)
  SS1$ = MID$(a$, 184, 9)
  PRINT acct$, ANAME$
  PRINT STRING$(79, "-")
  
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(taxcustrec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(taxcustrec(1))
  FOR Srch! = 1 TO NumOfTaxRecs
    GET TaxFile, Srch!, taxcustrec(1)
    
    IF VAL(acct$) = taxcustrec(1).CountyAcct THEN
      GOSUB UpdateRecord
      CLOSE TaxFile
      GOTO 1
    END IF
  NEXT Srch!
  
  
  'Must Add Record : Could Not Find
  
  'Decode Name Here
  nm$ = UCASE$(ANAME$)
  IF INSTR(nm$, "INC") > 0 THEN
    LN$ = UCASE$(ANAME$)
    FM$ = ""
    SN$ = LN$
  ELSE
    kk = INSTR(nm$, " ")
    IF kk > 0 THEN
      LN$ = LEFT$(ANAME$, kk - 1)
      FM$ = RIGHT$(ANAME$, (LEN(ANAME$) - kk))
      SN$ = LN$
    ELSE
      LN$ = UCASE$(ANAME$)
      FM$ = ""
      SN$ = LN$
    END IF
    
  END IF
  
  
  Record! = NumOfTaxRecs + 1
  
  
  
  taxcustrec(1).FNAME = LTRIM$(FM$)
  taxcustrec(1).LName = LTRIM$(LN$)
  taxcustrec(1).SName = LTRIM$(SN$)
  taxcustrec(1).acct = Record!
  IF LEN(RTRIM$(add3$)) = 0 THEN
    taxcustrec(1).Addr1 = add1$
    taxcustrec(1).Addr2 = ""
    taxcustrec(1).City = add2$
  ELSE
    taxcustrec(1).Addr1 = add1$
    taxcustrec(1).Addr2 = add2$
    taxcustrec(1).City = add3$
  END IF
  
  taxcustrec(1).STATE = ""
  taxcustrec(1).zip = zip$
  taxcustrec(1).CSSN = SS$
  taxcustrec(1).SSSN = SS1$
  taxcustrec(1).Active = "Y"
  taxcustrec(1).Interest = "Y"
  taxcustrec(1).TaxExempt = "N"
  taxcustrec(1).Penalty = "Y"
  taxcustrec(1).CountyAcct = VAL(acct$)
  taxcustrec(1).LastTrans = 0
  taxcustrec(1).FirstPropRec = 0
  taxcustrec(1).FirstPersRec = 0
  taxcustrec(1).PIN = Record!
  taxcustrec(1).Deleted = 0
  taxcustrec(1).FileVer = 8
  taxcustrec(1).OPENDATE = Date2Num%("09-23-2004")
  PUT TaxFile, Record!, taxcustrec(1)
  CLOSE TaxFile
NextOne:
  GOTO 1
  CLOSE
  EXIT SUB
  
  
UpdateRecord:
  taxcustrec(1).CountyAcct = VAL(acct$)
  PUT TaxFile, Srch!, taxcustrec(1)
  RETURN
END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(taxcustrec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(taxcustrec(1))
  
END SUB

