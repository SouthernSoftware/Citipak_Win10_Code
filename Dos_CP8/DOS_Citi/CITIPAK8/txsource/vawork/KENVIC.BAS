  DEFINT A-Z
  'Convert PP Records from txt victoria
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
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  CLS
  ON ERROR RESUME NEXT
  DIM a$(34)
  CALL KillFile("taxpers.dat")

  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  OPEN "luntpp.txt" FOR INPUT AS #11
  CustRecLen = LEN(TaxCustRec(1))
  'NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))

  CLS
  b = 0
1
  'DO UNTIL EOF(11)
    FOR cc = 1 TO 23
      INPUT #11, a$(cc)
      IF ERR = 62 THEN
        CLOSE
        PRINT "Done"
        END
      ELSEIF ERR > 0 THEN
        CLOSE
        PRINT "Error:"; ERR
        END
      END IF
    NEXT

    PRINT a$(1), a$(6), a$(15), a$(12)

'    SELECT CASE a$(12)
'    CASE "TP", "BP", "MH"
'    CASE ELSE
'      STOP
'
'    END SELECT
    'IF a$(12) <> "TP" OR a$(12) <> "BP" OR a$(12) <> "MH" THEN
    '  STOP
    'END IF

    PRINT STRING$(79, "-")
  'LOOP
  'CLOSE : END
'  GOTO 1
  
  Vtype$ = a$(12)
  
  'strip out - in ss#    primary
  s1$ = ""
  ll = LEN(a$(4))
  FOR a = 1 TO ll
    IF MID$(a$(4), a, 1) = "-" THEN
    ELSE
      s1$ = s1$ + MID$(a$(4), a, 1)
    END IF
  NEXT a
  
  'strip out - in ss#   secondary
  s2$ = ""
  ll = LEN(a$(5))
  FOR a = 1 TO ll
    IF MID$(a$(5), a, 1) = "-" THEN
    ELSE
      s2$ = s2$ + MID$(a$(5), a, 1)
    END IF
  NEXT a
  
  'MAKE SURE ADDRESS IS IN a$(7)
  IF LEN(QPTrim$(a$(7))) = 0 THEN
    a$(7) = a$(8)
    a$(8) = ""
  END IF
  
  'FIX ZIP TO 5 CHAR
  a$(11) = LEFT$(a$(11), 5)
  
 ' IF NOT (Vtype$ = "A" OR Vtype$ = "B" OR Vtype$ = "C" OR Vtype$ = "E" OR Vtype$ = "H" OR Vtype$ = "I" OR Vtype$ = "M" OR Vtype$ = "O" OR Vtype$ = "P" OR Vtype$ = "Q" OR Vtype$ = "R" OR Vtype$ = "T" OR Vtype$ = "W") THEN
 '   GOTO 1
 ' END IF
  
  'First Search for Existing Acct
  NumOfTaxRecs = LOF(TaxFile) / CustRecLen

  FOR Snt& = 1 TO NumOfTaxRecs
    GET TaxFile, Snt&, TaxCustRec(1)
    IF TaxCustRec(1).CountyAcct = VAL(a$(1)) THEN
      kk = INSTR(a$(6), " ")
      IF kk > 0 THEN
        lnm$ = LEFT$(a$(6), kk - 1)
        F$ = RIGHT$(a$(6), LEN(a$(6)) - kk)
      ELSE
        lnm$ = a$(6)
        F$ = ""
      END IF
      TaxCustRec(1).FNAME = LTRIM$(F$)
      TaxCustRec(1).LName = LTRIM$(lnm$)
      TaxCustRec(1).SName = LTRIM$(lnm$)
      TaxCustRec(1).Addr1 = a$(7)
      TaxCustRec(1).Addr2 = a$(8)
      TaxCustRec(1).City = a$(9)
      TaxCustRec(1).STATE = a$(10)
      TaxCustRec(1).Zip = a$(11)
      TaxCustRec(1).CSSN = s1$
      TaxCustRec(1).SSSN = s2$
      
      TaxCustRec(1).OPENDATE = Date2Num%("11-16-2005")
      PUT TaxFile, Snt&, TaxCustRec(1)
      
      'If Found then Update Property Only
      Record! = Snt&
      GOSUB UpdatePersProperty
      GOTO NextOne
    END IF
  NEXT Snt&
  
  'If Not Found then Add Customer & Property
  
  Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1
  
  kk = INSTR(a$(6), " ")
  IF kk > 0 THEN
    lnm$ = LEFT$(a$(6), kk - 1)
    F$ = RIGHT$(a$(6), LEN(a$(6)) - kk)
  ELSE
    lnm$ = a$(6)
    F$ = ""
  END IF
  TaxCustRec(1).FNAME = LTRIM$(F$)
  TaxCustRec(1).LName = LTRIM$(lnm$)
  TaxCustRec(1).SName = LTRIM$(lnm$)
  TaxCustRec(1).Acct = Record!
  TaxCustRec(1).Addr1 = a$(7)
  TaxCustRec(1).Addr2 = a$(8)
  TaxCustRec(1).City = a$(9)
  TaxCustRec(1).STATE = a$(10)
  TaxCustRec(1).Zip = a$(11)
  TaxCustRec(1).HPHONE = ""
  TaxCustRec(1).WPHONE = ""
  TaxCustRec(1).CSSN = s1$
  TaxCustRec(1).SSSN = s2$
  TaxCustRec(1).Active = "Y"
  TaxCustRec(1).Interest = "Y"
  TaxCustRec(1).TaxExempt = "N"
  TaxCustRec(1).Penalty = "Y"
  TaxCustRec(1).CountyAcct = VAL(a$(1))
  TaxCustRec(1).CountyAcctString = a$(1)
  TaxCustRec(1).LastTrans = 0
  TaxCustRec(1).FirstPropRec = 0
  TaxCustRec(1).FirstPersRec = 0
  TaxCustRec(1).PIN = Record!
  TaxCustRec(1).Deleted = 0
  TaxCustRec(1).FileVer = 8
  TaxCustRec(1).OPENDATE = Date2Num%("11-16-2005")
  PUT TaxFile, Record!, TaxCustRec(1)
  PropNumb = 1
  GOSUB UpdatePersProperty
  
NextOne:
  GOTO 1
  CLOSE
  EXIT SUB
  
UpdatePersProperty:
  
  PersRec(1).MCVALUE = 0
  PersRec(1).MTVALUE = 0
  PersRec(1).MHVALUE = 0
  PersRec(1).PERSVAL = 0
  PersRec(1).CValue = 0
  
  IF Vtype$ = "MT" THEN
    PersRec(1).MTVALUE = VAL(a$(15))
    GOTO upd1
  END IF
  IF Vtype$ = "MH" THEN
    PersRec(1).MHVALUE = VAL(a$(15))
    GOTO upd1
  END IF

'  IF Vtype$ = "BP" THEN
'    PersRec(1).MCVALUE = VAL(a$(15))
'    GOTO upd1
'  END IF
  
  PersRec(1).PERSVAL = VAL(a$(15))

upd1:
  TVal# = TVal# + VAL(a$(15))
  
'  mth = VAL(a$(19))
 ' MTHS$ = LTRIM$(STR$(mth))
  
 ' IF LEN(MTHS$) = 1 THEN MTHS$ = " " + MTHS$
  PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
  PersRec(1).PROPDATE = Date2Num%("11-27-2005")
  PersRec(1).CValue = 0
  PersRec(1).EXMPSENI = 0
  PersRec(1).EXMPOTHR = 0
  PersRec(1).DISCOV = "N"
  PersRec(1).LATELIST = "N"
  PersRec(1).DESC1 = a$(14)
  PersRec(1).DESC2 = a$(21)
  PersRec(1).DESC3 = a$(20)
  PersRec(1).DESC4 = a$(22)
  'IF a$(34) = "Y" THEN
  IF VAL(a$(18)) > 0 THEN
    PersRec(1).DESC5 = "Y" + MTHS$ + STRING$(24, " ") + "2005"
  ELSE
    PersRec(1).DESC5 = "N" + MTHS$ + STRING$(24, " ") + "2005"
  END IF
  
  
  
  PersRec(1).CustPin = Record!
  PersRec(1).NextRec = 0
  PersRec(1).LastYrPrinted = 0
  PersRec(1).Deleted = 0
  PersRec(1).Blank = ""
  PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
  PUT PersTaxFile, PRecord&, PersRec(1)
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

