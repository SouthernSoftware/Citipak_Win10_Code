DEFINT A-Z
'Real Estate Conv from Text File Information  New Market Information 2005
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
DECLARE FUNCTION uRound# (N#)

  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PROPERTYRec(1) AS PropertyRecType
  DIM SHARED RealRec(1) AS PropertyRecType
  
  
  STACK 5000
  BalanceListing
  CALL KillFile("nmreal05.exe")


  END

SUB BalanceListing
CLS
DIM a$(112)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  CLS

  GOSUB InitYear

  OPEN "NMREAL05.txt" FOR INPUT AS #11

1 IF NOT EOF(11) THEN
    LINE INPUT #11, a$
  ELSE
    'STOP
    CLOSE
    LOCATE 22, 1
    PRINT "Discount: "; TDisc#
    PRINT "Total1: "; TOTAL1#
    PRINT "Total2: "; TOTAL2#

    EXIT SUB
  END IF
  
  c = c + 1
  LOCATE 1, 1
  PRINT a$
  
  a$(6) = MID$(a$, 58, 35)    'Name
  a$(4) = LTRIM$(MID$(a$, 34, 5))      'Acct #
  a$(18) = MID$(a$, 224, 17)  'map #
  a$(7) = MID$(a$, 93, 30)    'Address 1
  a$(8) = MID$(a$, 123, 30)    'Address 2
  a$(9) = MID$(a$, 153, 30)   'Address 3
  a$(9) = RTRIM$(a$(9))
  State$ = RIGHT$(a$(9), 2)
  IF LEN(a$(9)) > 0 THEN a$(9) = LEFT$(a$(9), LEN(a$(9)) - 2)
  a$(10) = MID$(a$, 184, 5)   'Zip

  a$(30) = MID$(a$, 24, 1)   'Class billing half

  a$(15) = MID$(a$, 375, 25)
  a$(16) = MID$(a$, 400, 25)
  a$(17) = MID$(a$, 425, 25)
  a$(32) = MID$(a$, 469, 8) 'land value
  a$(33) = MID$(a$, 450, 9) 'bldg value

  IF VAL(a$(30)) <> 2 THEN
    TOTAL1# = TOTAL1# + VAL(a$(32))
    TOTAL2# = TOTAL2# + VAL(a$(33))
  ELSE
    GOTO 1
  END IF

  Disc# = VAL(MID$(a$, 485, 8))
  IF Disc# > 0 THEN
    TDisc# = uRound#(TDisc# + Disc#)  '(Disc# / 2))
    'STOP
  ELSE
    'STOP
  END IF

 'FIRST DECODE THE PIN # TO FOR ACCT LOOKUP
 'PIN# IS IN A$(18)
   map$ = RTRIM$(a$(18))

  L = 1
  WHILE L <= LEN(map$)
    IF MID$(map$, L, 1) = " " THEN
      map$ = LEFT$(map$, L - 1) + MID$(map$, L + 1, LEN(map$) - L)
      L = 1
    END IF
    L = L + 1
  WEND

 'FIND THE ACCOUNT NOW
   FOR Srch& = 1 TO LOF(PropTaxFile) / LEN(RealRec(1))
    GET PropTaxFile, Srch&, PROPERTYRec(1)
    IF RTRIM$(PROPERTYRec(1).RealPin) = RTRIM$(map$) THEN
    'Account Located Update It Now
     GOSUB UpdateTheAccount
     GOTO 1
    END IF
   NEXT Srch&

  'New Account to Add

  'Decode Name Here
    nm$ = UCASE$(a$(6))
    IF INSTR(nm$, "INC") > 0 THEN
     LN$ = UCASE$(a$(6))
     FM$ = ""
     SN$ = LN$
    ELSE
    kk = INSTR(nm$, " ")
    IF kk > 0 THEN
     LN$ = LEFT$(a$(6), kk - 1)
     FM$ = RIGHT$(a$(6), (LEN(a$(6)) - kk))
     SN$ = LN$
      ELSE
     LN$ = UCASE$(a$(6))
     FM$ = ""
     SN$ = LN$
    END IF
  END IF

   Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1
   TaxCustRec(1).FNAME = LTRIM$(FM$)
   TaxCustRec(1).LName = LTRIM$(LN$)
   TaxCustRec(1).SName = LTRIM$(SN$)
   TaxCustRec(1).Acct = Record!
   IF LEN(RTRIM$(a$(7))) = 0 THEN
     TaxCustRec(1).Addr1 = LTRIM$(UCASE$(a$(8)))
    TaxCustRec(1).Addr2 = ""
   ELSE
    TaxCustRec(1).Addr1 = LTRIM$(UCASE$(a$(7)))
    TaxCustRec(1).Addr2 = LTRIM$(UCASE$(a$(8)))
   END IF
   TaxCustRec(1).CITY = LTRIM$(UCASE$(a$(9)))
   TaxCustRec(1).State = LTRIM$(State$)
   TaxCustRec(1).ZIP = a$(10)
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).CountyAcctString = a$(4)
   TaxCustRec(1).CountyAcct = VAL(a$(4))
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
   TaxCustRec(1).OPENDATE = Date2Num%("04-15-2005")
'unrem
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
    PROPERTYRec(1).PROPDATE = Date2Num%("04-15-2005")
    PROPERTYRec(1).GISPOS = ""
    PROPERTYRec(1).map = ""
    PROPERTYRec(1).BLOCK = ""
    PROPERTYRec(1).LOTNUMB = ""
    PROPERTYRec(1).LOTACRE = "A"
    PROPERTYRec(1).PROPSIZE = 0
    PROPERTYRec(1).PROPDISC = "N"
    PROPERTYRec(1).LATELIST = "N"
    PROPERTYRec(1).MORTCODE = ""
    PROPERTYRec(1).PROPVALU = VAL(a$(32))
    'PROPERTYRec(1).EXMPSENI = Disc#

    PROPERTYRec(1).EXMPSENI = VAL(a$(33))

    PROPERTYRec(1).EXMPOTHR = Disc#

    PROPERTYRec(1).PROPNOT1 = a$(15)
    PROPERTYRec(1).PROPNOT2 = a$(16)
    PROPERTYRec(1).PROPNOT3 = a$(17)
    PROPERTYRec(1).CustPin = Record!
    PROPERTYRec(1).NextRec = 0
    PROPERTYRec(1).LastYrPrinted = 2004
    PROPERTYRec(1).Deleted = 0
    PROPERTYRec(1).Blank = ""
'unrem
    PUT PropTaxFile, RERecord&, PROPERTYRec(1)

    RETURN

InitYear:
    FOR Srch& = 1 TO LOF(PropTaxFile) / LEN(RealRec(1))
     GET PropTaxFile, Srch&, PROPERTYRec(1)
     PROPERTYRec(1).PROPVALU = 0
     PROPERTYRec(1).EXMPSENI = 0
'unrem
    PUT PropTaxFile, Srch&, PROPERTYRec(1)

    NEXT Srch&
    RETURN


UpdateTheAccount:
   'First Update Name and Address
    CustRecord! = PROPERTYRec(1).CustPin
    GET TaxFile, CustRecord!, TaxCustRec(1)

    nm$ = UCASE$(a$(6))
    IF INSTR(nm$, "INC") > 0 THEN
     LN$ = UCASE$(a$(6))
     FM$ = ""
     SN$ = LN$
    ELSE
    kk = INSTR(nm$, " ")
    IF kk > 0 THEN
     LN$ = LEFT$(a$(6), kk - 1)
     FM$ = RIGHT$(a$(6), (LEN(a$(6)) - kk))
     SN$ = LN$
      ELSE
     LN$ = UCASE$(a$(6))
     FM$ = ""
     SN$ = LN$
    END IF
   END IF

   TaxCustRec(1).FNAME = LTRIM$(FM$)
   TaxCustRec(1).LName = LTRIM$(LN$)
   TaxCustRec(1).SName = LTRIM$(SN$)
   
   IF LEN(RTRIM$(a$(7))) = 0 THEN
    TaxCustRec(1).Addr1 = LTRIM$(UCASE$(a$(8)))
    TaxCustRec(1).Addr2 = ""
    ELSE
    TaxCustRec(1).Addr1 = LTRIM$(UCASE$(a$(7)))
    TaxCustRec(1).Addr2 = LTRIM$(UCASE$(a$(8)))
   END IF
   TaxCustRec(1).CITY = LTRIM$(UCASE$(a$(9)))
   TaxCustRec(1).State = State$
   TaxCustRec(1).ZIP = a$(10)
   TaxCustRec(1).OPENDATE = Date2Num%("04-15-2005")

'unrem
   PUT TaxFile, CustRecord!, TaxCustRec(1)


    PROPERTYRec(1).RealPin = map$
    PROPERTYRec(1).PROPDATE = Date2Num%("04-15-2005")
    IF VAL(a$(30)) = 7 THEN
    ELSE
    PROPERTYRec(1).PROPVALU = VAL(a$(32))
    PROPERTYRec(1).EXMPOTHR = Disc#
    PROPERTYRec(1).EXMPSENI = VAL(a$(33))
    END IF
    PROPERTYRec(1).PROPNOT1 = a$(15)
    PROPERTYRec(1).PROPNOT2 = a$(16)
    PROPERTYRec(1).PROPNOT3 = a$(17)
    PROPERTYRec(1).LastYrPrinted = 2004
    PROPERTYRec(1).Deleted = 0
    PROPERTYRec(1).Blank = ""
'unrem
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

