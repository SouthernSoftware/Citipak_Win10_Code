'spruce pines
'MAGGIE VALLEY NC FROM EXCEL TO ACCESS CONVERS
DEFINT A-Z
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

SUB BalanceListing

  DIM A$(120)

'  CALL killfile("taxcust.dat")
  CALL killfile("taxpers.dat")
  CALL killfile("taxprop.dat")

  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

'  'reset personal property records
'  FOR cc = 1 TO NumOfPersRecs
'    GET PersTaxFile, cc, PersRec(1)
'    PersRec(1).PersVal = 0
'    PersRec(1).MHVALUE = 0
'    PersRec(1).MCVALUE = 0
'    PersRec(1).CVALUE = 0
'    PersRec(1).MTVALUE = 0
'    PersRec(1).EXMPSENI = 0
'    PersRec(1).EXMPOTHR = 0
'    PUT PersTaxFile, cc, PersRec(1)
'  NEXT cc
'
'  ' Now Real Estate
'  FOR Snt& = 1 TO NumOfPropRecs
'    GET PropTaxFile, Snt&, PropertyRec(1)
'    PropertyRec(1).PROPVALU = 0
'    PropertyRec(1).EXMPSENI = 0
'    PropertyRec(1).EXMPOTHR = 0
'    PropertyRec(1).CustPin = 0
'    PUT PropTaxFile, Snt&, PropertyRec(1)
'  NEXT Snt&
  
  OPEN "spruce.TXT" FOR INPUT AS #11
  Record = NumOfTaxRecs
  CLS

  DidCnt& = -1

  ON ERROR RESUME NEXT
Top:
  'STOP
  'IF ReadCnt = 28 THEN
  '  STOP
  'END IF

  FOR X = 1 TO 28               'maybe
    INPUT #11, A$(X)
    IF ERR = 62 THEN
      CLOSE
      LOCATE 16, 1
      PRINT "Data import complete."
      PRINT "Whewwww!"
      CALL killfile("txspruce.exe")
      END
    ELSEIF ERR > 0 THEN
      CLOSE
      LOCATE 12, 1
      PRINT "UNEXPECTED ERROR OCCURED"
      PRINT "Ahhhhhh!"
      END
    END IF
  NEXT X

 ReadCnt = ReadCnt + 1

'  LOCATE 3, 1
'  PRINT A$(1); "   "; A$(2)

  LOCATE 4, 1
  PRINT SPACE$(78);

  LOCATE 4, 1
  PRINT A$(4);



  IF DidCnt& = -1 THEN   'skip the header
    DidCnt& = DidCnt& + 1
    GOTO Top
  END IF

'  LOCATE 6, 1
'  PRINT A$(1); "   "; A$(2)

  LOCATE 13, 1
  PRINT A$(1); "   "; A$(2)

  IF A$(1) <> A$(2) THEN
    STOP
  END IF


 FoundEm = False

 Cnt& = 0
 NumOfTaxRecs = (LOF(TaxFile) / LEN(TaxCustRec(1)))
 DO WHILE Cnt& <= NumOfTaxRecs
   Cnt& = Cnt& + 1
   GET TaxFile, Cnt&, TaxCustRec(1)
    cssn$ = RTRIM$(TaxCustRec(1).cssn)
    IF QPTrim$(TaxCustRec(1).CountyAcctString) = A$(1) OR QPTrim$(cssn$) = A$(1) THEN
      'Record = Record + 1
      TaxCustRec(1).Acct = Cnt&
      Record = Cnt&
      FoundCnt = FoundCnt + 1
      LOCATE 7, 1
      PRINT " Updated:"; FoundCnt;
      GOSUB UpdateCustInfo:
      LOCATE 9, 1
      PRINT "ZeroPers:"; Skiped;
      LOCATE 11, 1
      PRINT "ZeroReal:"; RSkiped;

      GOTO Top
    END IF
 LOOP

 NoFind = NoFind + 1
 LOCATE 8, 1
 PRINT "   Added:"; NoFind;

 Record = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1

 GOSUB UpdateCustInfo:
 GOTO Top

  'LOCATE 2, 1
  DidCnt& = DidCnt& + 1

  'PRINT "Processing:"; DidCnt&;

UpdateCustInfo:


  'IF INSTR(a$(1), "001041243") > 0 THEN STOP

  kk = INSTR(A$(4), " ")
  IF kk > 0 THEN
    LN$ = UCASE$(LEFT$(A$(4), kk - 1))
    SN$ = UCASE$(LEFT$(A$(4), kk - 1))
    FF1$ = UCASE$(MID$(A$(4), kk + 1, LEN(A$(4)) - (kk)))
  ELSE
    LN$ = UCASE$(RTRIM$(A$(4)))
    SN$ = UCASE$(RTRIM$(A$(4)))
    FF1$ = ""
  END IF

' PRINT a$(4)

  TaxCustRec(1).OpenDate = Date2Num%(DATE$)  'Date2Num%("07-01-2001")
  TaxCustRec(1).FName = LTRIM$(FF1$)
  TaxCustRec(1).LName = LTRIM$(LN$)
  TaxCustRec(1).SName = LTRIM$(LN$)
  TaxCustRec(1).Addr1 = UCASE$(A$(5))
  TaxCustRec(1).Addr2 = UCASE$(A$(6))
  TaxCustRec(1).City = UCASE$(A$(7))
  TaxCustRec(1).State = UCASE$(A$(8))
  TaxCustRec(1).Zip = A$(9)
  TaxCustRec(1).HPhone = ""
  TaxCustRec(1).WPhone = ""
  TaxCustRec(1).CountyAcctString = A$(1)
  TaxCustRec(1).cssn = A$(1)
  TaxCustRec(1).SSSN = A$(2)
  TaxCustRec(1).Active = "Y"
  TaxCustRec(1).Interest = "Y"
  TaxCustRec(1).TaxExempt = "N"
  TaxCustRec(1).Penalty = "N"
  TaxCustRec(1).LastTrans = 0
  TaxCustRec(1).FirstPropRec = 0
  TaxCustRec(1).FirstPersRec = 0
  TaxCustRec(1).PIN = Record
  TaxCustRec(1).Deleted = 0
  TaxCustRec(1).FileVer = 8

  PUT TaxFile, Record, TaxCustRec(1)

  GOSUB UpdateReal
  GOSUB UpdatePers
  
NextRecd:

  RETURN

  CLOSE
  EXIT SUB

UpdateReal:

  IF (VAL(A$(21)) <= 0) AND (VAL(A$(20)) <= 0) THEN
    RSkiped = RSkiped + 1
    GOTO SkipReal
  END IF


  PropertyRec(1).REALPIN = QPTrim$(A$(12))
  PropertyRec(1).PROPDATE = Date2Num%(DATE$)
  PropertyRec(1).GISPOS = ""
  PropertyRec(1).MAP = ""
  PropertyRec(1).BLOCK = ""
  PropertyRec(1).LOTNUMB = ""
  PropertyRec(1).LOTACRE = UCASE$(A$(25))
  PropertyRec(1).PROPSIZE = VAL(A$(24))
  PropertyRec(1).PROPDISC = "N" 'UCASE$(a$(19))
  PropertyRec(1).LATELIST = "N"
  PropertyRec(1).MORTCODE = ""
  PropertyRec(1).PROPVALU = VAL(A$(21))
  PropertyRec(1).EXMPSENI = VAL(A$(20))
  PropertyRec(1).EXMPOTHR = 0
  PropertyRec(1).PROPNOT1 = UCASE$(A$(19))
  PropertyRec(1).PROPNOT2 = ""
  PropertyRec(1).PROPNOT3 = ""

  PropertyRec(1).StormFee = 0

  PropertyRec(1).CustPin = Record
  PropertyRec(1).NextRec = 0
  PropertyRec(1).LastYrPrinted = 0
  PropertyRec(1).Deleted = 0
  PropertyRec(1).Blank = ""
  NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
  PUT PropTaxFile, NextRecord, PropertyRec(1)

SkipReal:
RETURN

UpdatePers:

  t1# = VAL(A$(14))
  t2# = VAL(A$(15))
  t3# = VAL(A$(16))
  t4# = VAL(A$(17))
  t5# = VAL(A$(18))

  tt# = t1# + t2# + t3# + t4# + t5#
  IF tt# <= 0 THEN
    Skiped = Skiped + 1
    GOTO SkipPersonal
  END IF

  PersRec(1).PROPPIN = QPTrim$(A$(12))
  PersRec(1).PROPDATE = Date2Num(DATE$)
  'OthPers,MobHome,Vehicle,FarmEquip,

  a1# = VAL(A$(14))
  a2# = VAL(A$(17))

  PersRec(1).PersVal = a1# + a2#    '+ a3# + a4#

  PersRec(1).MHVALUE = VAL(A$(15))
  PersRec(1).MCVALUE = VAL(A$(16))
  PersRec(1).CVALUE = VAL(A$(18))
  PersRec(1).MTVALUE = 0
  PersRec(1).EXMPSENI = VAL(A$(13))
  PersRec(1).EXMPOTHR = 0
  PersRec(1).DISCOV = "N"
  PersRec(1).LATELIST = "N"
  PersRec(1).DESC1 = UCASE$(A$(19))
  PersRec(1).DESC2 = ""
  PersRec(1).DESC3 = ""
  PersRec(1).DESC4 = ""
  PersRec(1).DESC5 = ""
  PersRec(1).CustPin = Record
  PersRec(1).NextRec = 0
  PersRec(1).LastYrPrinted = 0
  PersRec(1).Deleted = 0
  PersRec(1).Blank = ""
  NextRec = (LOF(PersTaxFile) / LEN(PersRec(1))) + 1

  PUT PersTaxFile, NextRec, PersRec(1)

SkipPersonal:

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

