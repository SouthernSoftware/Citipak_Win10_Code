DECLARE SUB MakeData ()
  DEFINT A-Z
  'Real Estate Conversion from Text File Information STUART VA 2003
  
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
  MakeData
  
  END

SUB MakeData
  ON ERROR RESUME NEXT
  CLS
  DIM a$(36)

  CALL KillFile("taxprop.dat")
  CALL KillFile("taxpers.dat")

  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  'First Initialize Property Values on File

  'FOR cnt& = 1 TO NumOfPropRecs
  '  GET PropTaxFile, cnt&, PropertyRec(1)
  '  PropertyRec(1).PROPVALU = 0
  '  PropertyRec(1).EXMPSENI = 0
  '  PropertyRec(1).EXMPOTHR = 0
  '  'PUT PropTaxFile, cnt&, PropertyRec(1)
  'NEXT cnt&
  '
  
  CLS
  OPEN "STUARTre.txt" FOR INPUT AS #11
  record! = NumOfTaxRecs
  
BackToTop:
  
  FOR C = 1 TO 23
    INPUT #11, a$(C)
    IF ERR = 62 THEN
      CLOSE
      PRINT
      PRINT "Done."
      CALL KillFile("taxcon51.exe")
      END
    ELSEIF ERR > 0 THEN
      CLOSE
      PRINT "Unexpected error! Code:"; ERR
      END

    END IF
  NEXT C
  
  DidCnt = DidCnt + 1

  PRINT a$(1), a$(2), a$(3)
  PRINT STRING$(79, "-")
  IF VAL(a$(20)) > 70 THEN
    GOTO BackToTop
  ELSE
    P# = P# + VAL(a$(22))
    PR# = PR# + VAL(a$(23))
  END IF
  
  'FIND THE ACCOUNT FIRST
  spin$ = RTRIM$(a$(10)) + RTRIM$(a$(11)) + RTRIM$(a$(12)) + RTRIM$(a$(13)) + RTRIM$(a$(14)) + RTRIM$(a$(15)) + RTRIM$(a$(16))
  PRINT spin$, a$(22), a$(23)
  
  FOR cnt& = 1 TO LOF(PropTaxFile) / LEN(PropertyRec(1))
    GET PropTaxFile, cnt&, PropertyRec(1)
    IF RTRIM$(PropertyRec(1).PROPNOT1) = RTRIM$(a$(17)) AND RTRIM$(PropertyRec(1).PROPNOT2) = RTRIM$(a$(18)) AND LEN(a$(17)) > 0 AND PropertyRec(1).CustPin <> 0 THEN
      GOSUB UPDATEVALUES
      record! = PropertyRec(1).CustPin
      GET TaxFile, record!, TaxCustRec(1)
      TaxCustRec(1).OPENDATE = Date2Num%(DATE$)
      PUT TaxFile, record!, TaxCustRec(1)
      
      GOTO BackToTop
    END IF
  NEXT cnt&
  
  'Decode Name Here
  IF VAL(a$(20)) = 4 THEN
    LN$ = a$(1)
    FM$ = ""
    SN$ = a$(1)
  ELSE
    KK = INSTR(a$(1), " ")
    LN$ = LEFT$(a$(1), KK - 1)
    FM$ = LTRIM$(MID$(a$(1), KK + 1, LEN(a$(1)) - KK))
    SN$ = LN$
  END IF
  
  LN$ = UCASE$(LN$)
  FM$ = UCASE$(FM$)
  SN$ = UCASE$(SN$)
  
  TaxCustRec(1).FNAME = FM$
  TaxCustRec(1).LName = LN$
  TaxCustRec(1).SName = SN$
  
  TaxCustRec(1).Acct = record!
  TaxCustRec(1).Addr1 = UCASE$(a$(5))
  TaxCustRec(1).Addr2 = UCASE$(a$(6))
  TaxCustRec(1).CITY = UCASE$(a$(7))
  TaxCustRec(1).STATE = UCASE$(a$(8))
  TaxCustRec(1).ZIP = a$(9)
  TaxCustRec(1).HPHONE = ""
  TaxCustRec(1).WPHONE = ""
  TaxCustRec(1).CSSN = ""
  TaxCustRec(1).SSSN = ""
  TaxCustRec(1).CountyAcctString = ""
  TaxCustRec(1).CountyAcct = 0
  TaxCustRec(1).Active = "Y"
  TaxCustRec(1).Interest = "Y"
  TaxCustRec(1).TaxExempt = "N"
  TaxCustRec(1).Penalty = "Y"
  TaxCustRec(1).LastTrans = 0
  TaxCustRec(1).FirstPropRec = 0
  TaxCustRec(1).FirstPersRec = 0
  TaxCustRec(1).PIN = record!
  TaxCustRec(1).Deleted = 0
  TaxCustRec(1).FileVer = 8
  TaxCustRec(1).OPENDATE = Date2Num%(DATE$)
  record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1
  PUT TaxFile, record!, TaxCustRec(1)
  PropNumb = 1
  GOSUB updatereal
  
NEXTONE:
  GOTO BackToTop
  CLOSE
  EXIT SUB
  
  
updatereal:
  RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
  PropertyRec(1).REALPIN = spin$
  PropertyRec(1).PROPDATE = Date2Num%(DATE$)
  PropertyRec(1).GISPOS = ""
  PropertyRec(1).MAP = ""
  PropertyRec(1).BLOCK = ""
  PropertyRec(1).LOTNUMB = ""
  PropertyRec(1).LOTACRE = "A"
  PropertyRec(1).PROPSIZE = VAL(a$(21))
  PropertyRec(1).PROPDISC = "N"
  PropertyRec(1).LATELIST = "N"
  PropertyRec(1).MORTCODE = ""
  PropertyRec(1).PROPVALU = VAL(a$(22))
  PropertyRec(1).EXMPSENI = VAL(a$(23))
  PropertyRec(1).EXMPOTHR = 0
  PropertyRec(1).PROPNOT1 = a$(17)
  PropertyRec(1).PROPNOT2 = a$(18)
  PropertyRec(1).PROPNOT3 = "Class " + a$(20)
  
  
  PropertyRec(1).CustPin = record!
  PropertyRec(1).NextRec = 0
  PropertyRec(1).LastYrPrinted = 2004
  PropertyRec(1).Deleted = 0
  PropertyRec(1).Blank = ""

  PUT PropTaxFile, RERecord&, PropertyRec(1)
  RETURN
  
UPDATEVALUES:
  PropertyRec(1).PROPDATE = Date2Num%(DATE$)
  PropertyRec(1).LOTACRE = "A"
  PropertyRec(1).PROPSIZE = VAL(a$(21))
  
  IF VAL(a$(21)) > 70 THEN
    
  ELSE
    PropertyRec(1).PROPVALU = PropertyRec(1).PROPVALU + VAL(a$(22))
    PropertyRec(1).EXMPSENI = PropertyRec(1).EXMPSENI + VAL(a$(23))
  END IF
  
  PropertyRec(1).PROPNOT1 = a$(17)
  PropertyRec(1).PROPNOT2 = a$(18)
  PropertyRec(1).PROPNOT3 = "Class " + a$(20)
  PropertyRec(1).LastYrPrinted = 2004
  PUT PropTaxFile, cnt&, PropertyRec(1)
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

