'MAXTON 2003 FROM COUNTY -2003 IS FIRST YEAR
'STORED THE PIN # AND THE ACCT #
'PIN # IS PRIMARY KEY WITH ACCT # AS SECONDARY FOR THOSE WITH PERS PROP ONLY

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
DECLARE FUNCTION Round# (b#)
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
   
   OPEN "MAXTON.txt" FOR INPUT AS #11
   

1  LINE INPUT #11, a$
                b$ = LEFT$(a$, 768)
                pin$ = RTRIM$(LEFT$(a$, 24))
                acct$ = RTRIM$(MID$(a$, 25, 12))
                Nam$ = RTRIM$(MID$(a$, 57, 30))
                a1$ = RTRIM$(MID$(a$, 87, 30))
                a2$ = RTRIM$(MID$(a$, 117, 30))
                a3$ = RTRIM$(MID$(a$, 147, 30))
                city$ = RTRIM$(MID$(a$, 177, 18))
                state$ = RTRIM$(MID$(a$, 195, 2))
                zip$ = RTRIM$(MID$(a$, 197, 5))
                legal$ = RTRIM$(MID$(a$, 206, 26))
                legal1$ = RTRIM$(MID$(a$, 232, 14))
                acres$ = RTRIM$(MID$(a$, 246, 10))
                Unit$ = RTRIM$(MID$(a$, 256, 2))
                LandValue# = VAL(MID$(a$, 258, 9)) + VAL(MID$(a$, 267, 9)) + VAL(MID$(a$, 276, 9))
                LandValue1# = VAL(MID$(a$, 294, 9))
                ExemptValue# = VAL(MID$(a$, 321, 9))
                TotalValue# = VAL(MID$(a$, 339, 9))
                IF TotalValue# > 0 AND LandValue1# = 0 THEN
                 PersValue# = TotalValue#
                 ELSE
                 PersValue# = 0
                END IF
                CLS
                PRINT b$
                   STOP
'   FOR Cnt& = 1 TO NumOfTaxRecs
'   GET TaxFile, Cnt&, TaxCustRec(1)
'   found = 0
'   IF TaxCustRec(1).CountyAcct = VAL(A$(2)) THEN
'      TaxCustRec(1).Active = "Y"
'      TaxCustRec(1).TaxExempt = "N"
'      PUT TaxFile, Cnt&, TaxCustRec(1)
'      GET TaxFile, Cnt&, TaxCustRec(1)
'
'   IF VAL(A$(4)) > 0 THEN
'    GOSUB UpdateReal
'    END IF
'    IF VAL(A$(6)) > 0 THEN
'    GOSUB UpdatePers
'    END IF
'    found = 1
'    EXIT FOR
'   END IF
'    NEXT Cnt&
'          IF found = 0 THEN STOP
'   GOTO 1



PRINT Nam$
kk3 = INSTR(Nam$, "LLC")
kk4 = INSTR(Nam$, "INC")
kk5 = INSTR(Nam$, "CHURCH")
kk6 = INSTR(Nam$, "CO ")
IF kk3 > 0 OR kk4 > 0 OR kk5 > 0 OR kk6 > 0 THEN
 n1$ = Nam$
 n2$ = ""
ELSE
 kk = INSTR(Nam$, " ")
 n1$ = LEFT$(Nam$, kk - 1)
 n2$ = MID$(Nam$, kk + 1, LEN(Nam$) - kk)

END IF
 IF LEN(a1$) = 0 THEN
  a1$ = a2$: a2$ = a3$
 END IF
' ADD NEW ACCOUNT HERE
   Record = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1
   TaxCustRec(1).acct = Record
   TaxCustRec(1).OpenDate = Date2Num(DATE$)
   TaxCustRec(1).FName = n2$
   TaxCustRec(1).LNAME = n1$
   TaxCustRec(1).SName = n1$
   TaxCustRec(1).Addr1 = a1$
   TaxCustRec(1).Addr2 = a2$
   TaxCustRec(1).city = city$
   TaxCustRec(1).state = state$
   TaxCustRec(1).zip = zip$
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "N"
   TaxCustRec(1).CountyAcctString = acct$
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).pin = Record
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   PUT TaxFile, Record, TaxCustRec(1)
   IF LandValue1# > 0 THEN
    GOSUB UpdateReal
   END IF
   IF PersValue# > 0 THEN
    GOSUB UpdatePers
   END IF
   GOTO 1
   CLOSE
   EXIT SUB
  
  

UpdateReal:

    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = pin$
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).PROPSIZE = VAL(acres$)
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = legal$
    PropertyRec(1).PROPNOT2 = legal1$
    PropertyRec(1).PROPNOT3 = ""
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 0
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    PropertyRec(1).REALPIN = pin$
    PropertyRec(1).PROPDATE = Date2Num%(DATE$)
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = Unit$
    PropertyRec(1).PROPVALU = LandValue1#
    PropertyRec(1).EXMPSENI = ExemptValue#
    PropertyRec(1).CustPin = Record
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN

UpdatePers:
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
   PersRec(1).Desc4 = ""
   PersRec(1).Desc5 = ""
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 0
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""

   PersRec(1).PROPPIN = "1"
   PersRec(1).PROPDATE = Date2Num%(DATE$)
   PersRec(1).PersVal = PersValue#
   PersRec(1).EXMPSENI = 0
   PersRec(1).CustPin = Record
   NextRec = (LOF(PersTaxFile) / LEN(PersRec(1))) + 1
   PUT PersTaxFile, NextRec, PersRec(1)
   RETURN

InitFiles:
    FOR NT% = 1 TO LOF(PropTaxFile) / LEN(PropertyRec(1))
     GET PropTaxFile, NT%, PropertyRec(1)
      PropertyRec(1).PROPVALU = 0
      PropertyRec(1).EXMPSENI = 0
     PUT PropTaxFile, NT%, PropertyRec(1)
    NEXT NT%

    FOR NT% = 1 TO LOF(PersTaxFile) / LEN(PersRec(1))
     GET PersTaxFile, NT%, PersRec(1)
      PersRec(1).PersVal = 0
     PUT PersTaxFile, NT%, PersRec(1)
    NEXT NT%

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

