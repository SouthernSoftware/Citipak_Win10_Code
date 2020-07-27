  DEFINT A-Z
  'Convert Detail PP Records Elkton VA 2003
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
  DIM a$(31)
  
  CALL KillFile("TAXPERS.DAT")
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  OPEN "ELKDPP05.TXT" FOR INPUT AS #11
  CLS
  b = 0
  
1 FOR cc = 1 TO 11
  INPUT #11, a$(cc)
  IF ERR = 62 THEN
    CLOSE
    PRINT "Done"
    END
  ELSEIF ERR > 0 THEN
    PRINT
    PRINT "UNExpected err:"; ERR
    CLOSE
    END
  END IF
NEXT cc

  g = g + 1
  
  VType$ = a$(8)
  IF VType$ = "LF" THEN GOTO 1
  'OR VType$ = "MC" THEN
  
  acct# = VAL(a$(2))
  Tamt# = Tamt# + VAL(a$(7))
  PRINT acct#
  'First Search for Existing Acct
  FOR Snt& = 1 TO NumOfTaxRecs
    GET TaxFile, Snt&, TaxCustRec(1)
    IF VAL(TaxCustRec(1).CountyAcctString) = acct# THEN
      'If Found then Update Property Only
      AcctRecord! = Snt&
      BillNumber$ = QPTrim$(TaxCustRec(1).TownShip)
      GOSUB UpdatePersProperty
      GOTO NextOne
    END IF
  NEXT Snt&
  STOP
  ' GOSUB UpdatePersProperty
  'STOP
  
NextOne:
  GOTO 1
  CLOSE
  EXIT SUB
  
UpdatePersProperty:
  
  IF VAL(a$(10)) > 0 THEN PPTRA$ = "Y" ELSE PPTRA$ = "N"
  Make$ = QPTrim$(a$(4))
  VIN$ = QPTrim$(a$(5))
  Amt# = VAL(a$(7))
  Yr$ = a$(3)
  TVal# = TVal# + Amt#
  VType$ = a$(8)
  PersRec(1).MCVALUE = 0
  PersRec(1).MTVALUE = 0
  PersRec(1).MHVALUE = 0
  PersRec(1).PERSVAL = 0
  PersRec(1).CValue = 0
  
  IF VType$ = "MT" THEN
    PersRec(1).MTVALUE = Amt#
    GOTO upd1
  END IF
  
  IF VType$ = "MC" THEN
    PersRec(1).MCVALUE = Amt#
    GOTO upd1
  END IF
  IF VType$ = "MH" THEN
    PersRec(1).MHVALUE = Amt#
    GOTO upd1
  END IF
  
  'IF VType$ = "RV" THEN
  ' PersRec(1).MHVALUE = Amt#
  'GOTO upd1
  'END IF
  
  
  PersRec(1).PERSVAL = Amt#
upd1:
  PersRec(1).PROPPIN = BillNumber$
  PersRec(1).PROPDATE = Date2Num%(DATE$)
  PersRec(1).CValue = 0
  PersRec(1).EXMPSENI = 0
  PersRec(1).EXMPOTHR = 0
  PersRec(1).DISCOV = "N"
  PersRec(1).LATELIST = "N"
  PersRec(1).DESC1 = VIN$
  PersRec(1).DESC2 = Make$
  PersRec(1).DESC3 = ""
  PersRec(1).DESC4 = Yr$
  IF PPTRA$ = "Y" THEN
    PersRec(1).DESC5 = "Y" + "/" + STRING$(24, " ") + "2005"
  ELSE
    PersRec(1).DESC5 = "N" + "/" + STRING$(24, " ") + "2005"
  END IF
  
  PersRec(1).CustPin = AcctRecord!
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

