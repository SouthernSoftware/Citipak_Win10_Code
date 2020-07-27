DEFINT A-Z
'Update Personal Prop Detail Records from Txt file  Shenandoah Va  1999
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
  '$INCLUDE: 'TaxConv.bi'
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED Zip(1) AS ZipType
 

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM a$(75)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  Record! = NumOfTaxRecs + 1

  OPEN "DST05PP.TXT" FOR INPUT AS #11
1 FOR cc = 1 TO 75
  INPUT #11, a$(cc)
   a$(cc) = RTRIM$(a$(cc))
  NEXT cc
  PRINT a$(2)
  PRINT a$(8)
  PRINT a$(9)
  PRINT a$(72)
  PRINT a$(73)
  PRINT a$(74)
  PRINT a$(75)
  PRINT STRING$(79, "-")

  NF = 1
   FOR Cnt& = 1 TO NumOfTaxRecs
    GET TaxFile, Cnt&, TaxCustRec(1)
    IF TaxCustRec(1).CountyAcct = VAL(a$(2)) THEN
    CustRecord! = Cnt&
    GOSUB UpdatePersonal
    NF = 0
    EXIT FOR
   END IF
   NEXT Cnt&
   IF NF = 1 THEN
    GOSUB UpdateCust
    GOSUB UpdatePersonal
   END IF

Nextone:
   GOTO 1
   CLOSE
   EXIT SUB

UpdateCust:

    

    IF LEN(a$(74)) < 5 THEN a$(74) = STRING$(5 - LEN(a$(74)), "0") + a$(74)

    IF a$(69) = "I" THEN
    kk = INSTR(a$(70), " ")
    IF kk > 0 THEN
     LN1$ = LEFT$(a$(70), kk - 1)
     FM1$ = MID$(a$(70), (kk + 1), (LEN(a$(70)) - (kk + 1)))
     SN1$ = LN1$
     ELSE
     LN1$ = a$(70)
     SN1$ = a$(70)
     FM1$ = ""
    END IF
    ELSE
     LN1$ = a$(70)
     SN1$ = a$(70)
     FM1$ = ""

    END IF

  'Zip Search Here
    TaxFile1 = FREEFILE
    OPEN "ZIPCODE1.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile1 LEN = LEN(Zip(1))
    NumOfTaxRecs1 = LOF(TaxFile1) / LEN(Zip(1))
     
   FOR Srch! = 1 TO NumOfTaxRecs1
   GET TaxFile1, Srch!, Zip(1)
   IF VAL(Zip(1).ZipCode) = VAL(a$(74)) THEN
    City$ = Zip(1).City
    EXIT FOR
   END IF
   NEXT Srch!
   CLOSE TaxFile1

    TaxCustRec(1).FNAME = FM1$
    TaxCustRec(1).LName = LN1$
    TaxCustRec(1).SName = SN1$

    TaxCustRec(1).Acct = Record!
    TaxCustRec(1).Addr1 = a$(72)
    TaxCustRec(1).Addr2 = a$(73)
    TaxCustRec(1).City = City$
    TaxCustRec(1).STATE = ""
    TaxCustRec(1).Zip = a$(74)
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = a$(67)
    TaxCustRec(1).SSSN = a$(68)
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).CountyAcct = VAL(a$(2))
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = x
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("01-01-2000")
    PUT TaxFile, Record!, TaxCustRec(1)
    CustRecord! = Record!
    Record! = Record! + 1
    NumOfTaxRecs = NumOfTaxRecs + 1
    PropNumb = 1
    
    RETURN
UpdatePersonal:

   PersRec(1).CVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0

   IF a$(5) = "60" THEN
    PersRec(1).CVALUE = VAL(a$(11))
    GOTO upd1
   END IF
   IF a$(5) = "50" THEN
    PersRec(1).MHVALUE = VAL(a$(11))
    GOTO upd1
   END IF
   IF a$(5) = "30" THEN
    PersRec(1).MTVALUE = VAL(a$(11))
    GOTO upd1
   END IF
   PersRec(1).PERSVAL = VAL(a$(11))
upd1:
   a# = a# + VAL(a$(11))
   PersRec(1).PROPPIN = ""
   PersRec(1).PROPDATE = Date2Num%("01-01-2000")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = a$(8)
   PersRec(1).DESC2 = a$(9)
   PersRec(1).DESC3 = ""
   IF a$(5) = "10" THEN
   PersRec(1).DESC4 = LEFT$(a$(9), 4)
   ELSE
   PersRec(1).DESC4 = ""
   END IF
   IF a$(60) = "Y" THEN
    PersRec(1).DESC5 = "Y" + "/" + STRING$(24, " ") + "2000"
   ELSE
    PersRec(1).DESC5 = "N" + "/" + STRING$(24, " ") + "2000"
   END IF
   PersRec(1).CustPin = CustRecord!
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

