'White Lake Discoveries for 2003 detail
DEFINT A-Z
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
DECLARE FUNCTION Date2Num% (TheDate$)
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

'White Lake NC Version (Bills)
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  DIM a$(120)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  CLS
   
   OPEN "WLDETl.DAT" FOR INPUT AS #11
   Record = NumOfTaxRecs

1 IF EOF(11) THEN
'  FOR Cnt& = 1 TO LOF(TaxFile) / LEN(TaxCustRec(1))
 ' GET TaxFile, Cnt&, TaxCustRec(1)
 ' IF TaxCustRec(1).FirstPropRec = 0 AND TaxCustRec(1).FirstPersRec = 0 AND TaxCustRec(1).LastTrans = 0 THEN
 '     TaxCustRec(1).Deleted = -1
 '     PUT TaxFile, Cnt&, TaxCustRec(1)
 ' END IF
 ' NEXT Cnt&
  END
  END IF




LINE INPUT #11, a$
  
found = 0
Cnt& = 0
     RealValue# = 0
     PersValue# = 0
     CRealValue# = 0
     DRealValue# = 0
     ERealValue# = 0

     Acct$ = MID$(a$, 2, 7)

     'IF MID$(a$, 15, 1) = "E" THEN GOTO 1

     IF MID$(a$, 12, 1) = "C" THEN CRealValue# = CRealValue# + VAL(MID$(a$, 162, 9))
     IF MID$(a$, 12, 1) = "D" THEN DRealValue# = DRealValue# + VAL(MID$(a$, 162, 9))
     IF MID$(a$, 12, 1) = "E" THEN ERealValue# = ERealValue# + VAL(MID$(a$, 162, 9))
     IF MID$(a$, 12, 1) = "F" THEN PersValue# = PersValue# + VAL(MID$(a$, 162, 9))
     IF MID$(a$, 12, 1) = "G" THEN PersValue# = PersValue# + VAL(MID$(a$, 162, 9))
     IF MID$(a$, 12, 1) = "H" THEN PersValue# = PersValue# + VAL(MID$(a$, 162, 9))
     IF MID$(a$, 12, 1) = "I" THEN PersValue# = PersValue# + VAL(MID$(a$, 162, 9))
     IF MID$(a$, 12, 1) = "J" THEN PersValue# = PersValue# + VAL(MID$(a$, 162, 9))
     IF MID$(a$, 12, 1) = "K" THEN PersValue# = PersValue# + VAL(MID$(a$, 162, 9))
     IF MID$(a$, 12, 1) = "L" THEN PersValue# = PersValue# + VAL(MID$(a$, 162, 9))
     IF MID$(a$, 12, 1) = "M" THEN PersValue# = PersValue# + VAL(MID$(a$, 162, 9))
      
      
     
     
   CLS
   PRINT USING "########,#"; CRealValue#
   PRINT USING "########,#"; DRealValue#
   PRINT USING "########,#"; ERealValue#
   PRINT USING "########,#"; PersValue#
   
   RealValue# = CRealValue# + DRealValue# + ERealValue#

   TReal# = TReal# + RealValue#
   TPers# = TPers# + PersValue#


   Acct# = VAL(MID$(a$, 2, 7))
   
  WHILE Cnt& <= NumOfTaxRecs

  Cnt& = Cnt& + 1
  GET TaxFile, Cnt&, TaxCustRec(1)
   IF TaxCustRec(1).CountyAcct = Acct# THEN
     found = 1
     IF RealValue# > 0 AND TaxCustRec(1).FirstPropRec > 0 THEN
      CustRec& = Cnt&
     END IF
  
     IF RealValue# > 0 AND TaxCustRec(1).FirstPropRec = 0 THEN
        GOSUB UpdateReal
        GOTO 1
     END IF

     IF PersValue# > 0 THEN
       GOSUB UpdatePers
       GOTO 1
       
     END IF
      
    
   END IF

NextCnt:

    WEND

    
    GET TaxFile, CustRec&, TaxCustRec(1)
    TaxCustRec(1).FirstPropRec = 0
    PUT TaxFile, (LOF(TaxFile) / LEN(TaxCustRec(1)) + 1), TaxCustRec(1)
    Cnt& = LOF(TaxFile) / LEN(TaxCustRec(1))
    NumOfTaxRecs = Cnt&
    GET TaxFile, Cnt&, TaxCustRec(1)
    GOSUB UpdateReal
    GOTO 1
    CLOSE
    EXIT SUB
  
  

UpdateReal:

    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = ""
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).PROPSIZE = 0
    PropertyRec(1).PROPDISC = "N"
    IF MID$(a$, 68, 1) = "Y" THEN
     PropertyRec(1).LATELIST = "Y"
    ELSE
     PropertyRec(1).LATELIST = "N"
    END IF

    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PropNot1 = ""
    PropertyRec(1).PROPNOT2 = ""
    PropertyRec(1).PROPNOT3 = ""
    'PropertyRec(1).Fill1 = ""
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 0
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    PropertyRec(1).REALPIN = MID$(a$, 182, 12)
    PropertyRec(1).PROPDATE = Date2Num%("08-01-2003")
    PropertyRec(1).LOTNUMB = ""

    PropertyRec(1).LOTACRE = ""
    PropertyRec(1).PROPVALU = RealValue#
    PropertyRec(1).EXMPSENI = VAL(MID$(a$, 139, 5))
    PropertyRec(1).CustPin = Cnt&
    PropertyRec(1).PropNot1 = MID$(a$, 16, 30)
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
      GET TaxFile, Cnt&, TaxCustRec(1)
      TaxCustRec(1).FirstPropRec = NextRecord
      PUT TaxFile, Cnt&, TaxCustRec(1)
    RETURN

UpdatePers:
  IF TaxCustRec(1).FirstPersRec > 0 THEN
   GET PersTaxFile, TaxCustRec(1).FirstPersRec, PersRec(1)
    PersRec(1).MHVALUE = 0
    PersRec(1).MCVALUE = 0
    PersRec(1).CVALUE = 0
    PersRec(1).MTVALUE = 0
    PersRec(1).EXMPOTHR = 0
    PersRec(1).DISCOV = "N"
    IF MID$(a$, 68, 1) = "Y" THEN
    PersRec(1).LATELIST = "Y"
    ELSE
      PersRec(1).LATELIST = "N"
     END IF
    PersRec(1).DESC1 = ""
    PersRec(1).DESC2 = ""
    PersRec(1).DESC3 = ""
    PersRec(1).DESC4 = ""
    PersRec(1).DESC5 = ""
    PersRec(1).NextRec = 0
    PersRec(1).LastYrPrinted = 0
    PersRec(1).Deleted = 0
    PersRec(1).Blank = ""
    PersRec(1).PROPPIN = "1"
    PersRec(1).PROPDATE = Date2Num%("08-01-2003")
    PersRec(1).PersVal = PersRec(1).PersVal + PersValue#
    PersRec(1).EXMPSENI = 0
    PersRec(1).CustPin = Cnt&
    PUT PersTaxFile, TaxCustRec(1).FirstPersRec, PersRec(1)
   ELSE
    PersRec(1).MHVALUE = 0
    PersRec(1).MCVALUE = 0
    PersRec(1).CVALUE = 0
    PersRec(1).MTVALUE = 0
    PersRec(1).EXMPOTHR = 0
    PersRec(1).DISCOV = "N"
    IF VAL(a$(10)) > 0 THEN
     PersRec(1).LATELIST = "Y"
    ELSE
     PersRec(1).LATELIST = "N"
    END IF
    PersRec(1).DESC1 = ""
    PersRec(1).DESC2 = ""
    PersRec(1).DESC3 = ""
    PersRec(1).DESC4 = ""
    PersRec(1).DESC5 = ""
    PersRec(1).NextRec = 0
    PersRec(1).LastYrPrinted = 0
    PersRec(1).Deleted = 0
    PersRec(1).Blank = ""
    PersRec(1).PROPPIN = "1"
    PersRec(1).PROPDATE = Date2Num%("08-01-2003")
    PersRec(1).PersVal = PersValue#
    PersRec(1).EXMPSENI = 0
    PersRec(1).CustPin = Cnt&
    NextRec = (LOF(PersTaxFile) / LEN(PersRec(1))) + 1
    PUT PersTaxFile, NextRec, PersRec(1)
      GET TaxFile, Cnt&, TaxCustRec(1)
      TaxCustRec(1).FirstPersRec = NextRecord
      PUT TaxFile, Cnt&, TaxCustRec(1)

    END IF
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

