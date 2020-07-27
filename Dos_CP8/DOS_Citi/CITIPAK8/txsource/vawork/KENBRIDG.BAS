DEFINT A-Z
'Real Estate Conv from Text File Information  Kenbridge Information 2003
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
  
  TYPE TBREKENCust
    AccountNumber     AS STRING * 5
    MapNumber         AS STRING * 14
    District          AS STRING * 2
    CustName          AS STRING * 35
    Address1          AS STRING * 35
    Address2          AS STRING * 35
    Address3          AS STRING * 35
    Zipcode           AS STRING * 5
    PropDesc1         AS STRING * 25
    PropDesc2         AS STRING * 25
    PropDesc3         AS STRING * 25
    PropDesc4         AS STRING * 25
    Zone              AS STRING * 2
    ClassCode         AS STRING * 2
    Tract             AS STRING * 18
    SqFtSize          AS STRING * 8
    Acres             AS STRING * 12
    Comment1          AS STRING * 20
    Comment2          AS STRING * 20
    Comment3          AS STRING * 20
    LandValue         AS STRING * 10
    BldgValue         AS STRING * 10
    LandUseValue      AS STRING * 10
    MPCurYrBillNumber AS STRING * 7             'Municipality Bills (Victoria/Kenbridge)
    MPCurYrTaxAmount  AS STRING * 10
  END TYPE
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PROPERTYRec(1) AS PropertyRecType
  DIM SHARED RealRec(1) AS PropertyRecType

  REDIM SHARED KNRECustRec(1) AS TBREKENCust
  
  STACK 5000
  PRINT LEN(KNRECustRec(1))
 ' END

  BalanceListing
  
  END

SUB BalanceListing
  ON ERROR RESUME NEXT
  CLS
  DIM Txt$(112)

  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  CLS
  
  GOSUB InitYear
  OPEN "kenbrdre.txt" FOR INPUT AS #11
  
1 LINE INPUT #11, a$
  IF ERR = 62 THEN
    CLOSE
    LOCATE 20, 1
    PRINT "Taxed:"; TTx#; " Building:"; TOTal2#; " Land:"; ToTAL1#
    END
  ELSEIF ERR > 0 THEN
    CLOSE
    PRINT "Error:"; ERR
    END
  END IF

  ReadCnt = ReadCnt + 1
  TxtLen = LEN(a$)
  IF TxtLen <> 415 THEN
    STOP
  END IF

  REDIM KNRECustRec(1)  AS TBREKENCust

  CLS
  c = c + 1

  PRINT a$
  
  Txt$(6) = MID$(a$, 22, 35)      'Name
  Txt$(4) = LTRIM$(MID$(a$, 1, 5))                'Acct #
  
  'Txt$(18) = MID$(a$, 224, 17)    'map #

  'Txt$(18) = MID$(a$, 242, 20)    'map #
  Txt$(18) = MID$(a$, 7, 13)    'map #
  
  Txt$(7) = MID$(a$, 57, 35)      'Address 1
  Txt$(8) = MID$(a$, 92, 35)      'Address 2

  Txt$(9) = MID$(a$, 127, 35)     'Address 3
  Txt$(9) = QPTrim$(Txt$(9))
  
  'State$ = RIGHT$(Txt$(9), 2)

  a9Len = LEN(Txt$(9))
  IF a9Len > 0 THEN
    FOR Cnt = a9Len TO 1 STEP -1
      SChr$ = MID$(Txt$(9), Cnt, 1)
      IF SChr$ = " " THEN
        Txt$(9) = LEFT$(Txt$(9), Cnt)   'strip Zip
        Txt$(9) = QPTrim$(Txt$(9))
        State$ = RIGHT$(Txt$(9), 2)
        Txt$(9) = LEFT$(Txt$(9), LEN(Txt$(9)) - 2)
        EXIT FOR
      END IF
    NEXT
  END IF

  Txt$(10) = MID$(a$, 162, 5)     'Zip
  
  Txt$(30) = QPTrim$(MID$(a$, 269, 2))      'Class

  Txt$(15) = MID$(a$, 167, 25)    '>
  Txt$(16) = MID$(a$, 192, 25)    '-NOTES
  Txt$(17) = MID$(a$, 217, 25)    '>
  
  Txt$(32) = MID$(a$, 369, 10)     'land value
  Txt$(33) = MID$(a$, 379, 10)     'bldg value
  
'  IF VAL(Txt$(30)) <> 2 THEN
    ToTAL1# = ToTAL1# + VAL(Txt$(32))
    TOTal2# = TOTal2# + VAL(Txt$(33))
    TTax$ = MID$(a$, 406)    'taxed
    TTax$ = QPTrim$(TTax$)
    TTx# = TTx# + VAL(TTax$)
'  ELSE
'    GOTO 1
'  END IF
  
  'FIRST DECODE THE PIN # TO FOR ACCT LOOKUP
  'PIN# IS IN Txt$(18)
  map$ = QPTrim$(Txt$(18))
  'L = 1
  'WHILE L <= LEN(map$)
  '  IF MID$(map$, L, 1) = " " THEN
  '    map$ = LEFT$(map$, L - 1) + MID$(map$, L + 1, LEN(map$) - L)
  '    L = 1
  '  END IF
  '  L = L + 1
  'WEND
  
  'FIND THE ACCOUNT NOW
'  FOR Srch& = 1 TO LOF(PropTaxFile) / LEN(RealRec(1))
'    GET PropTaxFile, Srch&, PROPERTYRec(1)
'    IF RTRIM$(PROPERTYRec(1).RealPin) = RTRIM$(map$) THEN
'      'Account Located Update It Now
'      GOSUB UpdateTheAccount
'      GOTO 1
'    END IF
'  NEXT Srch&
  
  'New Account to Add
  
  'Decode Name Here
  nm$ = UCASE$(Txt$(6))
  IF INSTR(nm$, "INC") > 0 THEN
    LN$ = UCASE$(Txt$(6))
    FM$ = ""
    SN$ = LN$
  ELSE
    kk = INSTR(nm$, " ")
    IF kk > 0 THEN
      LN$ = LEFT$(Txt$(6), kk - 1)
      FM$ = RIGHT$(Txt$(6), (LEN(Txt$(6)) - kk))
      SN$ = LN$
    ELSE
      LN$ = UCASE$(Txt$(6))
      FM$ = ""
      SN$ = LN$
    END IF
  END IF
  
  Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1
  TaxCustRec(1).FNAME = LTRIM$(FM$)
  TaxCustRec(1).LName = LTRIM$(LN$)
  TaxCustRec(1).SName = LTRIM$(SN$)
  TaxCustRec(1).Acct = Record!
  IF LEN(RTRIM$(Txt$(7))) = 0 THEN
    TaxCustRec(1).Addr1 = LTRIM$(UCASE$(Txt$(8)))
    TaxCustRec(1).Addr2 = ""
  ELSE
    TaxCustRec(1).Addr1 = LTRIM$(UCASE$(Txt$(7)))
    TaxCustRec(1).Addr2 = LTRIM$(UCASE$(Txt$(8)))
  END IF
  TaxCustRec(1).City = LTRIM$(UCASE$(Txt$(9)))
  TaxCustRec(1).State = LTRIM$(State$)
  TaxCustRec(1).ZIP = Txt$(10)
  TaxCustRec(1).HPHONE = ""
  TaxCustRec(1).WPHONE = ""
  TaxCustRec(1).CSSN = ""
  TaxCustRec(1).SSSN = ""
  TaxCustRec(1).CountyAcctString = Txt$(4)
  TaxCustRec(1).CountyAcct = VAL(Txt$(4))
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
  TaxCustRec(1).OPENDATE = Date2Num%("11-16-2005")
  PUT TaxFile, Record!, TaxCustRec(1)
  PropNumb = 1
  GOSUB Updatereal
  
NEXTONE:
  GOTO 1
  CLOSE
  STOP
  EXIT SUB
  
  
Updatereal:
  RERecord& = LOF(PropTaxFile) / LEN(RealRec(1)) + 1
  PROPERTYRec(1).RealPin = map$
  PROPERTYRec(1).PROPDATE = Date2Num%("11-16-2005")
  PROPERTYRec(1).GISPOS = ""
  PROPERTYRec(1).map = ""
  PROPERTYRec(1).BLOCK = ""
  PROPERTYRec(1).LOTNUMB = ""
  PROPERTYRec(1).LOTACRE = "A"
  PROPERTYRec(1).PROPSIZE = 0
  PROPERTYRec(1).PROPDISC = "N"
  PROPERTYRec(1).LATELIST = "N"
  PROPERTYRec(1).MORTCODE = ""
  PROPERTYRec(1).PROPVALU = VAL(Txt$(32))
  PROPERTYRec(1).EXMPSENI = VAL(Txt$(33))
  PROPERTYRec(1).EXMPOTHR = 0
  
  PROPERTYRec(1).PROPNOT1 = Txt$(15)
  PROPERTYRec(1).PROPNOT2 = Txt$(16)
  PROPERTYRec(1).PROPNOT3 = Txt$(17)
  
  PROPERTYRec(1).CustPin = Record!
  PROPERTYRec(1).NextRec = 0
  PROPERTYRec(1).LastYrPrinted = 0
  PROPERTYRec(1).Deleted = 0
  PROPERTYRec(1).Blank = ""
  PUT PropTaxFile, RERecord&, PROPERTYRec(1)
  
  RETURN
  
InitYear:
  FOR Srch& = 1 TO LOF(PropTaxFile) / LEN(RealRec(1))
    GET PropTaxFile, Srch&, PROPERTYRec(1)
    PROPERTYRec(1).PROPVALU = 0
    PROPERTYRec(1).EXMPSENI = 0
    PUT PropTaxFile, Srch&, PROPERTYRec(1)
  NEXT Srch&
  RETURN
  
  
UpdateTheAccount:
  'First Update Name and Address
  CustRecord! = PROPERTYRec(1).CustPin
  GET TaxFile, CustRecord!, TaxCustRec(1)
  
  nm$ = UCASE$(Txt$(6))
  IF INSTR(nm$, "INC") > 0 THEN
    LN$ = UCASE$(Txt$(6))
    FM$ = ""
    SN$ = LN$
  ELSE
    kk = INSTR(nm$, " ")
    IF kk > 0 THEN
      LN$ = LEFT$(Txt$(6), kk - 1)
      FM$ = RIGHT$(Txt$(6), (LEN(Txt$(6)) - kk))
      SN$ = LN$
    ELSE
      LN$ = UCASE$(Txt$(6))
      FM$ = ""
      SN$ = LN$
    END IF
  END IF
  
  TaxCustRec(1).FNAME = LTRIM$(FM$)
  TaxCustRec(1).LName = LTRIM$(LN$)
  TaxCustRec(1).SName = LTRIM$(SN$)
  
  IF LEN(RTRIM$(Txt$(7))) = 0 THEN
    TaxCustRec(1).Addr1 = LTRIM$(UCASE$(Txt$(8)))
    TaxCustRec(1).Addr2 = ""
  ELSE
    TaxCustRec(1).Addr1 = LTRIM$(UCASE$(Txt$(7)))
    TaxCustRec(1).Addr2 = LTRIM$(UCASE$(Txt$(8)))
  END IF
  TaxCustRec(1).City = LTRIM$(UCASE$(Txt$(9)))
  TaxCustRec(1).State = State$
  TaxCustRec(1).ZIP = Txt$(10)
  TaxCustRec(1).OPENDATE = Date2Num%("11-16-2005")
  PUT TaxFile, CustRecord!, TaxCustRec(1)
  
  PROPERTYRec(1).RealPin = map$
  PROPERTYRec(1).PROPDATE = Date2Num%("11-16-2005")
  IF VAL(Txt$(30)) = 7 THEN
  ELSE
    PROPERTYRec(1).PROPVALU = VAL(Txt$(32))
    PROPERTYRec(1).EXMPSENI = VAL(Txt$(33))
  END IF
  PROPERTYRec(1).PROPNOT1 = Txt$(15)
  PROPERTYRec(1).PROPNOT2 = Txt$(16)
  PROPERTYRec(1).PROPNOT3 = Txt$(17)
  PROPERTYRec(1).LastYrPrinted = 0
  PROPERTYRec(1).Deleted = 0
  PROPERTYRec(1).Blank = ""
  PUT PropTaxFile, Srch&, PROPERTYRec(1)
  
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
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PROPERTYRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PROPERTYRec(1))
END SUB

