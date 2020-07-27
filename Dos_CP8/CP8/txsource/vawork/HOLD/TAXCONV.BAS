DEFINT A-Z
'Chilhowie TownMast Conversion PP
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

  OpenTaxCustFile NumOfTaxRecs, TaxFile
  RECORD! = NumOfTaxRecs


   DIM ARec(1) AS ChilRecType

   OPEN "TOWNMAST.DAT" FOR RANDOM AS #11 LEN = 593
   FOR X = 1 TO LOF(11) / 593
   GET 11, X, ARec(1)
   LOCATE 3, 1: COLOR 15: PRINT ARec(1).Nam, ARec(1).Acct
   'Search For Existing Account on System
   FOR Cnt& = 1 TO NumOfTaxRecs
   GET TaxFile, Cnt&, TaxCustRec(1)
   IF TaxCustRec(1).CountyAcct = VAL(ARec(1).Acct) THEN
   'SLEEP 1
    GOTO NextOne
   END IF
   NEXT Cnt&

    
 PRINT ARec(1).Nam, ne, LOF(11) / 593
  ne = ne + 1
   IF ARec(1).bus = "C" THEN
   firname$ = ""
   lasname$ = ARec(1).Nam
   ELSE
   KK = INSTR(ARec(1).Nam, " ")
   IF KK > 0 THEN
    lasname$ = LEFT$(ARec(1).Nam, KK - 1)
    firname$ = RIGHT$(ARec(1).Nam, LEN(ARec(1).Nam) - KK)
         ELSE
    firname$ = ""
    lasname$ = ARec(1).Nam
   END IF
   END IF
IF LEN(RTRIM$(firname$)) = 0 AND LEN(RTRIM$(lasname$)) = 0 THEN GOTO NextOne
    TaxCustRec(1).FName = LTRIM$(firname$)
    TaxCustRec(1).LName = LTRIM$(lasname$)
    TaxCustRec(1).SName = LTRIM$(lasname$)
    Add3$ = RTRIM$(ARec(1).a3)
    IF LEN(Add3$) > 0 THEN

    cY$ = LEFT$(RTRIM$(ARec(1).a3), LEN(RTRIM$(ARec(1).a3)) - 2)
    st$ = RIGHT$(RTRIM$(Add3$), 2)
    ELSE
     IF LEN(RTRIM$(ARec(1).a2)) > 0 THEN
       cY$ = LEFT$(RTRIM$(ARec(1).a2), LEN(RTRIM$(ARec(1).a2)) - 2)
       st$ = RIGHT$(RTRIM$(ARec(1).a2), 2)
    END IF
   END IF
 
  
  
   RECORD! = RECORD! + 1
   ZipCode$ = LTRIM$(ARec(1).z1)
   IF LEN(RTRIM$(ZipCode$)) = 4 THEN ZipCode$ = "0" + ZipCode$
   z2$ = RTRIM$(ARec(1).z2)
   IF LTRIM$(z2$) = "0" THEN z2$ = ""
   IF LEN(LTRIM$(z2$)) < 4 AND LEN(LTRIM$(z2$)) > 1 THEN
    z2$ = STRING$(4 - LEN(LTRIM$(z2$)), "0") + LTRIM$(z2$)
   END IF

   IF LEN(z2$) = 4 THEN ZipCode$ = ZipCode$ + z2$

  
   TaxCustRec(1).Acct = RECORD!
   TaxCustRec(1).CountyAcct = VAL(ARec(1).Acct)
   TaxCustRec(1).Addr1 = ARec(1).a1
   TaxCustRec(1).Addr2 = ARec(1).a2
   TaxCustRec(1).City = cY$
   TaxCustRec(1).State = st$
   TaxCustRec(1).Zip = ZipCode$
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
  

   TaxCustRec(1).CSSN = LEFT$(ARec(1).ss1, 9)
   TaxCustRec(1).SSSN = LEFT$(ARec(1).ss2, 9)
 
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
   TaxCustRec(1).OPENDATE = Date2Num%("01-01-1998")
   PUT TaxFile, RECORD!, TaxCustRec(1)
   
NextOne:
   NEXT X
   CLOSE
   EXIT SUB



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

