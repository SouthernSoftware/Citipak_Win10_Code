DEFINT A-Z
'Conversion for Columbia, NC 2003
'DECLARE FUNCTION QPTrim$ (text$)
DECLARE FUNCTION Date2Num% (TheDate$)
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
DECLARE FUNCTION Round# (b#)
'$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (text$)
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
  
  CALL KillFile("taxcon18.exe")
  PRINT
  PRINT "Processing Complete."
  END


 'OPEN "taxtap.tap" FOR INPUT AS #1
 'OPEN "taxtap5.tap" FOR OUTPUT AS #2
 'DO
 '  LINE INPUT #1, a$
 '  IF LEN(QPTrim$(a$)) > 3 THEN
 '    PRINT #2, a$
 '  END IF
 'LOOP UNTIL EOF(1)
 'CLOSE

SUB BalanceListing
  
  DIM a$(120)

  'DIM TaxCustRec(1) AS TaxCustType

  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  OPEN "TAXTAP5.tap" FOR RANDOM ACCESS READ WRITE SHARED AS #8 LEN = 512
  FIELD 8, 11 AS bb$, 6 AS srval$, 6 AS b$, 5 AS zp$, 1 AS b1$, 30 AS n1$, 30 AS n2$, 30 AS add$, 21 AS city$, 1 AS ll$, 15 AS pin$
  FIELD 8, 157 AS b3$, 24 AS desc$, 8 AS size$, 1 AS lot$, 6 AS b4$, 8 AS landval$
  FIELD 8, 204 AS b5$, 8 AS PersVal$, 8 AS invval$, 8 AS Machval$, 8 AS invval1$, 8 AS mvehval$
  FIELD 8, 354 AS b8$, 8 AS deedbk$, 10 AS cacct$

GOTO JUMP2HERE
  
  Record = NumOfTaxRecs

  CustRecLen = LEN(TaxCustRec(1))

  'reset personal property records
  FOR cc = 1 TO NumOfPersRecs
    GET PersTaxFile, cc, PersRec(1)
    PersRec(1).PersVal = 0
    PersRec(1).MHVALUE = 0
    PersRec(1).MCVALUE = 0
    PersRec(1).CVALUE = 0
    PersRec(1).MTVALUE = 0
    PersRec(1).EXMPSENI = 0
    PersRec(1).EXMPOTHR = 0
    PUT PersTaxFile, cc, PersRec(1)
  NEXT cc

 ' Now Real Estate
  FOR Snt& = 1 TO NumOfPropRecs
    GET PropTaxFile, Snt&, PropertyRec(1)
    PropertyRec(1).PROPVALU = 0
    PropertyRec(1).EXMPSENI = 0
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).CustPin = 0
    PUT PropTaxFile, Snt&, PropertyRec(1)
  NEXT Snt&
  
JUMP2HERE:
1     CLS
bbb:

  FOR x = 1 TO LOF(8) / 512
    GET 8, x
    
    TPers# = TPers# + VAL(PersVal$)
    TLand# = TLand# + VAL(landval$)
    
    TMach# = TMach# + VAL(Machval$)
    TVeh# = TVeh + VAL(mvehval$)
    

    GOTO NotTodayBubba

    GOSUB FindCustRec

    'CustRecNumber=

    IF VAL(landval$) > 0 THEN
      GOSUB UpdateReal
    END IF
    
    IF VAL(PersVal$) > 0 OR VAL(Machval$) > 0 OR VAL(mvehval$) > 0 THEN
      GOSUB UpdatePers
    END IF
NotTodayBubba:
  NEXT x
  CLOSE
  STOP
  END

  EXIT SUB
  
UpdateReal:
  PropertyRec(1).REALPIN = QPTrim$(pin$)
  PropertyRec(1).PROPDATE = Date2Num%(DATE$)
  PropertyRec(1).GISPOS = ""
  PropertyRec(1).MAP = QPTrim$(deedbk$)
  PropertyRec(1).BLOCK = ""
  PropertyRec(1).LOTNUMB = ""
  PropertyRec(1).LOTACRE = lot$
  PropertyRec(1).PROPSIZE = VAL(size$)
  PropertyRec(1).PROPDISC = "N"
  PropertyRec(1).LATELIST = "N"
  PropertyRec(1).MORTCODE = ""
  PropertyRec(1).PROPVALU = VAL(landval$)
  PropertyRec(1).EXMPSENI = VAL(srval$)
  PropertyRec(1).EXMPOTHR = 0
  PropertyRec(1).PROPNOT1 = desc$
  PropertyRec(1).PROPNOT2 = ""
  PropertyRec(1).PROPNOT3 = ""
  'PropertyRec(1).Fill1 = ""
  PropertyRec(1).CustPin = CustRecNumber
  PropertyRec(1).NextRec = 0
  PropertyRec(1).LastYrPrinted = 0
  PropertyRec(1).Deleted = 0
  PropertyRec(1).Blank = ""
  NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
  PUT PropTaxFile, NextRecord, PropertyRec(1)
RETURN

UpdatePers:
  PersRec(1).PROPPIN = "1"
  PersRec(1).PROPDATE = Date2Num%(DATE$)
  PersRec(1).PersVal = VAL(PersVal$)
  PersRec(1).MHVALUE = VAL(mvehval$)
  PersRec(1).MCVALUE = VAL(invval$)
  PersRec(1).CVALUE = 0
  
  'IF x = 527 THEN STOP
  
  Machine$ = Machval$
  Machine$ = QPTrim$(Machine$)
  'IF LEN(Machine$) > 4 THEN
  ' STOP
  'END IF
  TMachine# = TMachine# + Machine#
  
  'IF Machine# > 10000 THEN STOP
  
  PersRec(1).MTVALUE = VAL(Machval$)
  
  PersRec(1).EXMPSENI = 0
  PersRec(1).EXMPOTHR = 0
  PersRec(1).DISCOV = "N"
  PersRec(1).LATELIST = "N"
  PersRec(1).DESC1 = ""
  PersRec(1).DESC2 = ""
  PersRec(1).DESC3 = ""
  PersRec(1).DESC4 = ""
  PersRec(1).DESC5 = ""
  PersRec(1).CustPin = CustRecNumber
  PersRec(1).NextRec = 0
  PersRec(1).LastYrPrinted = 0
  PersRec(1).Deleted = 0
  PersRec(1).Blank = ""
  NextRec = (LOF(PersTaxFile) / LEN(PersRec(1))) + 1
  PUT PersTaxFile, NextRec, PersRec(1)
RETURN

FindCustRec:

  GotAcct = 0

  NumOfTaxRecs = LOF(TaxFile) / CustRecLen

  FOR CustRec& = 1 TO NumOfTaxRecs
    GET TaxFile, CustRec&, TaxCustRec(1)
    TCacct$ = QPTrim$(cacct$)
    TTacct$ = QPTrim$(TaxCustRec(1).CountyAcctString)
    IF TCacct$ = TTacct$ THEN
      GotAcct = -1
      EXIT FOR
    END IF
  NEXT
  IF NOT GotAcct THEN
    CustRecNumber = NumOfTaxRecs + 1
    GOSUB AddNewCust
  ELSE
    CustRecNumber = CustRec&
    GOSUB UPDateCust
  END IF

RETURN

UPDateCust:

  GET TaxFile, CustRec&, TaxCustRec(1)

  City1$ = RTRIM$(city$)
  Record = Record + 1

  PRINT City1$
  kk = INSTR(n1$, " ")
  IF kk > 0 THEN
    TaxCustRec(1).FName = MID$(n1$, kk + 1, (LEN(n1$) - kk + 1))
    TaxCustRec(1).LName = LEFT$(n1$, kk - 1)
    TaxCustRec(1).SName = LEFT$(n1$, kk - 1)
  ELSE
    TaxCustRec(1).FName = ""
    TaxCustRec(1).LName = n1$
    TaxCustRec(1).SName = n1$
  END IF

    TaxCustRec(1).Acct = Record
    TaxCustRec(1).OpenDate = Date2Num%(DATE$)
    IF LEN(QPTrim$(n2$)) = 0 THEN
      TaxCustRec(1).Addr1 = add$
      TaxCustRec(1).Addr2 = ""
    ELSE
      TaxCustRec(1).Addr1 = n2$
      TaxCustRec(1).Addr2 = add$
    END IF
    TaxCustRec(1).city = RTRIM$(LEFT$(City1$, LEN(City1$) - 2))
    TaxCustRec(1).State = RIGHT$(RTRIM$(city$), 2)
    TaxCustRec(1).zip = zp$
    TaxCustRec(1).HPhone = ""
    TaxCustRec(1).WPhone = ""
    TaxCustRec(1).CSSN = ""
    TaxCustRec(1).SSSN = ""
    TaxCustRec(1).CountyAcctString = QPTrim$(cacct$)
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "N"
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).pin = CustRecNumber
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
  '  PUT TaxFile, Record, TaxCustRec(1)

    PUT TaxFile, CustRecNumber, TaxCustRec(1)

RETURN

AddNewCust:

  
  City1$ = RTRIM$(city$)
  Record = Record + 1

  PRINT City1$
  kk = INSTR(n1$, " ")
  IF kk > 0 THEN
    TaxCustRec(1).FName = MID$(n1$, kk + 1, (LEN(n1$) - kk + 1))
    TaxCustRec(1).LName = LEFT$(n1$, kk - 1)
    TaxCustRec(1).SName = LEFT$(n1$, kk - 1)
  ELSE
    TaxCustRec(1).FName = ""
    TaxCustRec(1).LName = n1$
    TaxCustRec(1).SName = n1$
  END IF

    TaxCustRec(1).Acct = Record
    TaxCustRec(1).OpenDate = Date2Num%(DATE$)
    IF LEN(QPTrim$(n2$)) = 0 THEN
      TaxCustRec(1).Addr1 = add$
      TaxCustRec(1).Addr2 = ""
    ELSE
      TaxCustRec(1).Addr1 = n2$
      TaxCustRec(1).Addr2 = add$
    END IF
    TaxCustRec(1).city = RTRIM$(LEFT$(City1$, LEN(City1$) - 2))
    TaxCustRec(1).State = RIGHT$(RTRIM$(city$), 2)
    TaxCustRec(1).zip = zp$
    TaxCustRec(1).HPhone = ""
    TaxCustRec(1).WPhone = ""
    TaxCustRec(1).CSSN = ""
    TaxCustRec(1).SSSN = ""
    TaxCustRec(1).CountyAcctString = QPTrim$(cacct$)
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "N"
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).pin = CustRecNumber
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
'    NumOfCust = LOF(TaxFile) / CustRecLen
'    NumOfCust = NumOfCust + 1

    PUT TaxFile, CustRecNumber, TaxCustRec(1)

RETURN

END SUB

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

