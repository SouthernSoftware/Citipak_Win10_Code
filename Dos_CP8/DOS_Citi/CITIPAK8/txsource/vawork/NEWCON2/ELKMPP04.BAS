  DEFINT A-Z
'Convert ELKTON PP Records 2003 Master PP
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
  
  
  'Stephens City Va Tax PP Conversion of Master (Customers)
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TaxCnvrt.bi'
  
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TaxTrans(1) AS TaxTransactionType
  
  
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  CLS
  ON ERROR RESUME NEXT
  DIM a$(26)
  OpenTaxCustFile NumOfTaxRecs, TaxFile

  CLS
  
  REDIM Elk(1) AS ElktonPPType
'  OPEN "ELKPP.TXT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 200
  OPEN "ELKMPP04.TXT" FOR INPUT AS #11

Top:
  FOR cnt = 1 TO 10
    INPUT #11, a$(cnt)
    IF ERR = 62 THEN
      CLOSE
      PRINT "Done"
      END
    END IF
  NEXT

  'FOR cnt& = 1 TO LOF(11) / 200
  '  GET 11, cnt&, Elk(1)
    'Account$ = Elk(1).AcctNumber
    Account$ = a$(1)
    PRINT a$(1)
    PRINT a$(2)
    PRINT STRING$(79, "-")
    
    'FIRST LOOK FOR EXISTING ACCOUNT TO UPDATE
    Found = 0
    FOR Snt& = 1 TO NumOfTaxRecs
      GET TaxFile, Snt&, TaxCustRec(1)
      IF VAL(TaxCustRec(1).CountyAcctString) = VAL(a$(1)) THEN
        kk = INSTR(a$(2), " ")
        IF kk > 0 THEN
          TaxCustRec(1).FName = LTRIM$(MID$(a$(2), kk + 1, (LEN(a$(2)) - kk)))
          TaxCustRec(1).LName = LTRIM$(LEFT$(a$(2), kk - 1))
          TaxCustRec(1).SName = LTRIM$(LEFT$(a$(2), kk - 1))
        ELSE
          TaxCustRec(1).FName = ""
          TaxCustRec(1).LName = LTRIM$(a$(2))
          TaxCustRec(1).SName = LTRIM$(a$(2))
        END IF
        'string out zip
        City$ = a$(5)
        City$ = LEFT$(City$, LEN(City$) - 5)
        City$ = QPTrim$(City$)
        St$ = RIGHT$(City$, 2)
        City$ = LEFT$(City$, LEN(City$) - 2)
        City$ = QPTrim$(City$)
        
        Add1$ = a$(4)
        Add1$ = QPTrim$(Add1$)
        'Add2$ = Elk(1).Add2
        'Add2$ = QPTrim$(Add2$)
        IF LEN(Add1$) = 0 THEN
          Add1$ = Add2$
          Add2$ = ""
        END IF
        
        TaxCustRec(1).Addr1 = Add1$
        'TaxCustRec(1).Addr2 = Add2$
        TaxCustRec(1).City = City$
        TaxCustRec(1).State = St$
        TaxCustRec(1).Zip = LEFT$(a$(6), 5)
        TaxCustRec(1).CountyAcctString = a$(1)
        'strip out any hypens
        SS1$ = ""
        FOR lll = 1 TO LEN(a$(9))
          IF MID$(a$(9), lll, 1) <> "-" THEN
            SS1$ = SS1$ + MID$(a$(9), lll, 1)
          END IF
        NEXT lll
        SS2$ = ""
'        FOR lll = 1 TO LEN(a$(10))
'          IF MID$(a$(10), lll, 1) <> "-" THEN
'            SS2$ = SS2$ + MID$(a$(10), lll, 1)
'          END IF
'        NEXT lll
        
        TaxCustRec(1).SSSN = SS1$
        css$ = "000" + a$(8)
        css$ = RIGHT$(css$, 4)
        TaxCustRec(1).CSSN = css$

        TaxCustRec(1).TownShip = css$
        TaxCustRec(1).OPENDATE = Date2Num%("11-17-2004")
        PUT TaxFile, Snt&, TaxCustRec(1)
        GOTO Nextone
      END IF
    NEXT Snt&
    GOSUB UpdateCust
Nextone:
  GOTO Top
  'NEXT cnt&
  CLOSE
  EXIT SUB
  
UpdateCust:
  Record! = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1
  kk = INSTR(a$(2), " ")
  IF kk > 0 THEN
    TaxCustRec(1).FName = LTRIM$(MID$(a$(2), kk + 1, (LEN(a$(2)) - kk)))
    TaxCustRec(1).LName = LTRIM$(LEFT$(a$(2), kk - 1))
    TaxCustRec(1).SName = LTRIM$(LEFT$(a$(2), kk - 1))
    
  ELSE
    TaxCustRec(1).FName = ""
    TaxCustRec(1).LName = LTRIM$(a$(2))
    TaxCustRec(1).SName = LTRIM$(a$(2))
    
  END IF
  
  'string out zip
  'City$ = a$(5)
  City$ = a$(5)
  City$ = LEFT$(City$, LEN(City$) - 5)
  City$ = QPTrim$(City$)
  St$ = RIGHT$(City$, 2)
  City$ = LEFT$(City$, LEN(City$) - 2)
  City$ = QPTrim$(City$)


  Add1$ = a$(4)
  Add1$ = QPTrim$(Add1$)
  'Add2$ = Elk(1).Add2
  'Add2$ = QPTrim$(Add2$)
  IF LEN(Add1$) = 0 THEN
    Add1$ = Add2$
    Add2$ = ""
  END IF
  
  
  TaxCustRec(1).Acct = Record!
  'TaxCustRec(1).CountyAcctString = Elk(1).AcctNumber
  TaxCustRec(1).CountyAcctString = a$(1)

  TaxCustRec(1).Addr1 = Add1$
'  TaxCustRec(1).Addr2 = Add2$
  TaxCustRec(1).City = City$
  TaxCustRec(1).State = St$
  TaxCustRec(1).Zip = LEFT$(a$(6), 5)

  TaxCustRec(1).HPHONE = ""
  TaxCustRec(1).WPHONE = ""
  TaxCustRec(1).TownShip = css$
  
  'strip out any hypens

  SS1$ = ""
  FOR lll = 1 TO LEN(a$(9))
    IF MID$(a$(9), lll, 1) <> "-" THEN
      SS1$ = SS1$ + MID$(a$(9), lll, 1)
    END IF
  NEXT lll
  'SS2$ = ""
  'FOR lll = 1 TO LEN(a$(10))
  '  IF MID$(a$(10), lll, 1) <> "-" THEN
  '    SS2$ = SS2$ + MID$(a$(10), lll, 1)
  '  END IF
  'NEXT lll
  
  TaxCustRec(1).SSSN = SS1$
  css$ = "000" + a$(8)
  TaxCustRec(1).CSSN = css$

  TaxCustRec(1).TownShip = css$

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
  TaxCustRec(1).OPENDATE = Date2Num%("11-17-2004")
  PUT TaxFile, Record!, TaxCustRec(1)
  RETURN
  
  
END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
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

