DEFINT A-Z
'Convert Data From County (Maxton : (Robeson Cty Real/Pers 2003))
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
  
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB OpenTaxPropFile (NumOfPropRecs%, PropTaxFile%)
DECLARE SUB OpenTaxPersFile (NumOfPersRecs%, PersTaxFile%)
DECLARE SUB DisplayTaxScrn (ScrnName$)
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
  DIM SHARED PersRec(1) AS PersonalRecType
  
  
  STACK 5000
  BalanceListing
  
  END

DEFDBL A
SUB BalanceListing

  ON ERROR RESUME NEXT
  DIM a$(20)
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
 ' Now Real Estate
  FOR Snt& = 1 TO NumOfPropRecs
    GET PropTaxFile, Snt&, PropertyRec(1)
    PropertyRec(1).PROPVALU = 0
    PropertyRec(1).EXMPSENI = 0
    PropertyRec(1).EXMPOTHR = 0
    'PropertyRec(1).CustPin = 0
    PUT PropTaxFile, Snt&, PropertyRec(1)
  NEXT Snt&
'  FOR x = 1 TO NumOfTaxRecs
'    GET TaxFile, x, TaxCustRec(1)
'    TaxCustRec(1).FirstPropRec = 0
'    PUT TaxFile, x, TaxCustRec(1)
'  NEXT x


  OPEN "maxton.txt" FOR INPUT AS #10
  Record = 0
  CLS

10 LINE INPUT #10, a$
  IF ERR = 62 THEN
    CLOSE
    PRINT
    PRINT "Done."
    END
  END IF
  'PRINT A$
  
  SDPin$ = MID$(a$, 390, 14)
  
  C = C + 1
  RealVal# = VAL(MID$(a$, 199, 8))
  PERSVAL# = VAL(MID$(a$, 468, 8))
  EVal1# = VAL(MID$(a$, 6, 6))
  EVal2# = VAL(MID$(a$, 12, 6))
  CUSTACCT$ = MID$(a$, 374, 12)
  
  PropRec = 0
  FOR x = 1 TO NumOfTaxRecs
    GET TaxFile, x, TaxCustRec(1)
    'IF x = 57 THEN STOP
    ' PRINT x, TAXCUSTREC(1).CountyAcctString
    IF RTRIM$(CUSTACCT$) = RTRIM$(TaxCustRec(1).CountyAcctString) THEN
      Record! = x
      TaxCustRec(1).OpenDate = Date2Num(DATE$)
      PRINT TaxCustRec(1).lnAME + " " + TaxCustRec(1).FNAME, MID$(a$, 30, 30)
      ' SLEEP 1
      IF LEN(RTRIM$(MID$(a$, 60, 30))) = 0 THEN
        TaxCustRec(1).Addr1 = RTRIM$(MID$(a$, 90, 30))
        TaxCustRec(1).Addr2 = ""
      ELSE
        TaxCustRec(1).Addr1 = RTRIM$(MID$(a$, 60, 30))
        TaxCustRec(1).Addr2 = RTRIM$(MID$(a$, 90, 30))
      END IF
      TaxCustRec(1).City = MID$(a$, 120, 21)
      TaxCustRec(1).STATE = MID$(a$, 141, 2)
      TaxCustRec(1).Zip = MID$(a$, 490, 5)
      PUT TaxFile, Record!, TaxCustRec(1)
      PropRec = TaxCustRec(1).FirstPropRec
      GOSUB UpdateReal
      GOSUB UpdatePers
      GOTO 10
    END IF
  NEXT x
  
  Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1
  
  TaxCustRec(1).Acct = Record!
  TaxCustRec(1).OpenDate = Date2Num(DATE$)
  
  CUSTNAME$ = RTRIM$(MID$(a$, 30, 30))
  KK = INSTR(CUSTNAME$, " ")
  IF KK > 0 THEN
    TaxCustRec(1).FNAME = LTRIM$(MID$(CUSTNAME$, KK, (LEN(CUSTNAME$) - KK)))
    TaxCustRec(1).lnAME = LEFT$(CUSTNAME$, KK - 1)
    TaxCustRec(1).SName = LEFT$(CUSTNAME$, KK - 1)
  ELSE
    TaxCustRec(1).FNAME = ""
    TaxCustRec(1).lnAME = RTRIM$(CUSTNAME$)
    TaxCustRec(1).SName = RTRIM$(CUSTNAME$)
  END IF
  TaxCustRec(1).STATE = MID$(a$, 141, 2)
  TaxCustRec(1).Zip = MID$(a$, 490, 5)
  IF LEN(RTRIM$(MID$(a$, 60, 30))) = 0 THEN
    TaxCustRec(1).Addr1 = RTRIM$(MID$(a$, 90, 30))
    TaxCustRec(1).Addr2 = ""
  ELSE
    TaxCustRec(1).Addr1 = RTRIM$(MID$(a$, 60, 30))
    TaxCustRec(1).Addr2 = RTRIM$(MID$(a$, 90, 30))
  END IF
  TaxCustRec(1).City = MID$(a$, 120, 21)
  TaxCustRec(1).STATE = MID$(a$, 141, 2)
  TaxCustRec(1).Zip = MID$(a$, 490, 5)
  TaxCustRec(1).CountyAcct = 0
  TaxCustRec(1).CountyAcctString = MID$(a$, 374, 12)
  TaxCustRec(1).HPHONE = ""
  TaxCustRec(1).WPHONE = ""
  TaxCustRec(1).CSSN = ""
  TaxCustRec(1).SSSN = ""
  TaxCustRec(1).Active = "Y"
  TaxCustRec(1).Interest = "Y"
  TaxCustRec(1).TaxExempt = "N"
  TaxCustRec(1).Penalty = "N"
  TaxCustRec(1).LastTrans = 0
  TaxCustRec(1).FirstPropRec = 0
  TaxCustRec(1).FirstPersRec = 0
  TaxCustRec(1).Pin = Record!
  TaxCustRec(1).Deleted = 0
  TaxCustRec(1).FileVer = 8
  PUT TaxFile, Record!, TaxCustRec(1)
  GOSUB UpdateReal
  GOSUB UpdatePers
  GOTO 10
  
  
UpdateReal:
  
  PropertyRec(1).REALPIN = SDPin$
  PropertyRec(1).PROPDATE = Date2Num%(DATE$)
  PropertyRec(1).GISPOS = ""
  PropertyRec(1).MAP = MID$(a$, 390, 6)
  PropertyRec(1).BLOCK = ""
  PropertyRec(1).LOTNUMB = ""
  PropertyRec(1).LOTACRE = MID$(a$, 192, 1)
  PropertyRec(1).PROPSIZE = VAL(MID$(a$, 184, 8))
  PropertyRec(1).PROPDISC = "N"
  PropertyRec(1).LATELIST = MID$(a$, 144, 1)
  PropertyRec(1).MORTCODE = ""
  PropertyRec(1).PROPVALU = RealVal#
  PropertyRec(1).EXMPSENI = EVal2#
  PropertyRec(1).EXMPOTHR = 0
  PropertyRec(1).PROPNOT1 = MID$(a$, 160, 24)
  PropertyRec(1).PROPNOT2 = ""
  PropertyRec(1).PROPNOT3 = MID$(a$, 390, 24)
  PropertyRec(1).CustPin = Record!
  PropertyRec(1).NextRec = 0
  PropertyRec(1).LastYrPrinted = 0
  PropertyRec(1).Deleted = 0
  PropertyRec(1).Blank = ""
  IF PropRec > 0 THEN
    NextRecord = PropRec
  ELSE
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
  END IF
  PUT PropTaxFile, NextRecord, PropertyRec(1)
  RETURN
  
UpdatePers:
  
  IF PERSVAL# <= 0 THEN RETURN
  
  PersRec(1).PROPPIN = "1"
  PersRec(1).PROPDATE = Date2Num%(DATE$)
  PersRec(1).PERSVAL = PERSVAL#
  PersRec(1).MHVALUE = 0
  PersRec(1).MCVALUE = 0
  PersRec(1).CVALUE = 0
  PersRec(1).MTVALUE = 0
  PersRec(1).EXMPSENI = 0
  PersRec(1).EXMPOTHR = 0
  PersRec(1).DISCOV = "N"
  PersRec(1).LATELIST = "N"
  PersRec(1).Desc1 = ""
  PersRec(1).Desc2 = ""
  PersRec(1).Desc3 = ""
  PersRec(1).DESC4 = ""
  PersRec(1).DESC5 = ""
  
  PersRec(1).CustPin = Record!
  PersRec(1).NextRec = 0
  PersRec(1).LastYrPrinted = 0
  PersRec(1).Deleted = 0
  PersRec(1).Blank = ""
  NextRec = (LOF(PersTaxFile) / LEN(PersRec(1))) + 1
  PUT PersTaxFile, NextRec, PersRec(1)
  RETURN
  
  
  
END SUB

DEFINT A
SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
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

