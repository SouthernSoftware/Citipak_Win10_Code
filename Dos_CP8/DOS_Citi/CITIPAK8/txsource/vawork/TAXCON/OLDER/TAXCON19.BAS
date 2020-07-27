DEFINT A-Z
'Update Personal Prop Master & Detail Records from Txt file (Warsaw Va 1999)
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
  DIM SHARED TBCustRec(1) AS TBCustType

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM a$(25)
   OpenTaxCustFile NumOfTaxRecs, TaxFile
   OpenTaxPersFile NumOfPersRecs, PersTaxFile

   OPEN "Warsawpp.dat" FOR INPUT AS #11
1 FOR CC = 1 TO 13
  INPUT #11, a$(CC)
  PRINT CC, a$(CC)
  NEXT CC
  PRINT STRING$(79, "-")
  nf = 1

   FOR Cnt& = 1 TO NumOfTaxRecs
   GET TaxFile, Cnt&, TaxCustRec(1)
   
   IF TaxCustRec(1).CountyAcct = VAL(a$(1)) THEN
    Record& = Cnt&
    GOSUB UpdatePersonal
    nf = 0
    EXIT FOR
   END IF
   NEXT Cnt&
   IF nf = 1 THEN
    GOSUB UpdateMaster
    GOSUB UpdatePersonal
   END IF

Nextone:
   GOTO 1
   CLOSE

   EXIT SUB

UpdateMaster:
    Record! = NumOfTaxRecs + 1

    IF a$(9) = "B" OR a$(9) = "W" THEN
     FM1$ = ""
     LNM$ = a$(2)
     SNM$ = a$(2)
    ELSE
     kk = INSTR(a$(2), " ")
     FM1$ = MID$(a$(2), kk + 1, LEN(a$(2)) - kk)
     LNM$ = LEFT$(a$(2), kk - 1)
     SNM$ = LEFT$(a$(2), kk - 1)
    END IF


    TaxCustRec(1).FNAME = FM1$
    TaxCustRec(1).LName = LNM$
    TaxCustRec(1).SName = SNM$
    TaxCustRec(1).Acct = Record!
    TaxCustRec(1).Addr1 = a$(3)
    TaxCustRec(1).Addr2 = a$(4)
    TaxCustRec(1).City = a$(5)
    TaxCustRec(1).STATE = a$(6)
    TaxCustRec(1).Zip = a$(7)
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = ""
    TaxCustRec(1).SSSN = ""
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).Employer = ""
    TaxCustRec(1).Bankrupt = "N"
    TaxCustRec(1).TownShip = ""
    TaxCustRec(1).CountyAcct = VAL(a$(1))
    TaxCustRec(1).CountyAcctstring = a$(1)
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("10-01-1999")
    PUT TaxFile, Record!, TaxCustRec(1)
    NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
    PropNumb = 1
    RETURN

UpdatePersonal:
   MachTool = INSTR(a$(10), "TOOLS")
   MerchCap = INSTR(a$(10), "MERCHANTS")

   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0

   IF MerchCap > 0 THEN
    PersRec(1).MCVALUE = VAL(a$(12))
    GOTO upd1
   END IF
   IF MachTool = 13 THEN
    PersRec(1).MTVALUE = VAL(a$(12))
    GOTO upd1
   END IF
   PersRec(1).PERSVAL = VAL(a$(12))
upd1:
   PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
   PersRec(1).PROPDATE = Date2Num%("11-01-1999")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = a$(11)
   PersRec(1).DESC2 = a$(10)
   PersRec(1).DESC3 = ""
   PersRec(1).DESC4 = LEFT$(a$(10), 4)
   PersRec(1).DESC5 = a$(9)
   PersRec(1).CustPin = Record!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 1998
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

