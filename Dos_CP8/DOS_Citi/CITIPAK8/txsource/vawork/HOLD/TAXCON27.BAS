DEFINT A-Z
'Update Personal Prop Records For ALBERTA VA
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NUMOFTAXRECS%, TaxFile%)
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
  DIM SHARED ALBERTA(1) AS ALBERTAPPTYPE

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM A$(25)
   OpenTaxCustFile NUMOFTAXRECS, TaxFile
   OpenTaxPersFile NumOfPersRecs, PersTaxFile

   OPEN "ALBTAXPP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #10 LEN = 1224
   FOR MCnt& = 1 TO LOF(10) / 1224
   GET 10, MCnt&, ALBERTA(1)
   NF = 1

   'First Determine the Ticket Type


   FOR Cnt& = 1 TO NUMOFTAXRECS
   GET TaxFile, Cnt&, TaxCustRec(1)
   NF = 1
   IF TaxCustRec(1).CountyAcct = VAL(ALBERTA(1).Acct) THEN
    GOSUB UpdatePersonal
    NF = 0
    EXIT FOR
   END IF
   NEXT Cnt&
   IF NF = 1 THEN
    GOSUB AddAcct
    GOSUB UpdatePersonal
   END IF

NextOne:
   NEXT MCnt&
   CLOSE
   EXIT SUB


UpdatePersonal:
    

   
    PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PersVal = 0
   PersRec(1).CVALUE = 0

   IF ALBERTA(1).CLASS = "MT" THEN
    PersRec(1).MTVALUE = VAL(ALBERTA(1).AMT)
    GOTO upd1
   END IF
   IF ALBERTA(1).CLASS = "MH" THEN
    PersRec(1).MHVALUE = VAL(ALBERTA(1).AMT)
    GOTO upd1
   END IF

   IF ALBERTA(1).CLASS = "FE" THEN
    PersRec(1).CVALUE = VAL(ALBERTA(1).AMT)
    GOTO upd1
   END IF

   
   PersRec(1).PersVal = VAL(ALBERTA(1).AMT)
upd1:
   PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
   PersRec(1).PROPDATE = Date2Num%("11-01-1999")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
    PersRec(1).DESC1 = ALBERTA(1).Vin
    PersRec(1).DESC2 = ALBERTA(1).Make + "/" + ALBERTA(1).Model + "/" + ALBERTA(1).Body
    PersRec(1).DESC3 = ""
    PersRec(1).Desc4 = ALBERTA(1).YR
    IF ALBERTA(1).DMV = "Y" THEN
     PersRec(1).Desc5 = "Y" + "/"
    ELSE
     PersRec(1).Desc5 = "N" + "/"
    END IF
   PersRec(1).CustPin = RECORD!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 1998
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
   PUT PersTaxFile, PRecord&, PersRec(1)
   RETURN

   

AddAcct:
   NUMOFTAXRECS = NUMOFTAXRECS + 1
   RECORD! = NUMOFTAXRECS
   
   TaxCustRec(1).Acct = RECORD!
   IF ALBERTA(1).TYP = "C" THEN
   F$ = ""
   LN$ = ALBERTA(1).NM
   ELSE
   KK = INSTR(ALBERTA(1).NM, " ")
   LN$ = LEFT$(ALBERTA(1).NM, KK - 1)
   F$ = MID$(ALBERTA(1).NM, KK + 1, LEN(ALBERTA(1).NM) - KK)
   END IF


   TaxCustRec(1).FName = F$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SName = LN$
   TaxCustRec(1).Addr1 = ALBERTA(1).ADD1
   TaxCustRec(1).Addr2 = ALBERTA(1).ADD2
   TaxCustRec(1).City = ALBERTA(1).ADD3
   TaxCustRec(1).State = ""
   TaxCustRec(1).ZIP = ALBERTA(1).ZIP
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   SS1$ = RTRIM$(LTRIM$(ALBERTA(1).SS1))
   IF LEN(SS1$) = 8 THEN SS1$ = "0" + SS1$
   SS2$ = RTRIM$(LTRIM$(ALBERTA(1).SS2))
   IF LEN(SS2$) = 8 THEN SS1$ = "0" + SS2$

   TaxCustRec(1).CSSN = SS1$
   TaxCustRec(1).SSSN = SS2$
   TaxCustRec(1).Employer = ""
   TaxCustRec(1).Bankrupt = "N"
   TaxCustRec(1).TownShip = ""
   TaxCustRec(1).CountyAcctString = ALBERTA(1).SS1
   TaxCustRec(1).CountyAcct = VAL(ALBERTA(1).Acct)
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "Y"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = RECORD!
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("10-01-1999")
   PUT TaxFile, RECORD!, TaxCustRec(1)
   RETURN

END SUB

SUB OpenTaxCustFile (NUMOFTAXRECS, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NUMOFTAXRECS = LOF(TaxFile) / LEN(TaxCustRec(1))
  
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

