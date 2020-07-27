DEFINT A-Z
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
'  OpenTaxPropFile NumOfPropRecs, PropTaxFile
'  OpenTaxPersFile NumOfPersRecs, PersTaxFile

   DIM Arec(1) AS ChilRecType

   OPEN "TOWNMAST" FOR RANDOM AS #11 LEN = 593
   FOR x = 1 TO LOF(11) / 593
   GET 11, x, Arec(1)


 PRINT Arec(1).Nam

   IF Arec(1).Bus = "C" THEN
   firname$ = ""
   lasname$ = Arec(1).Nam
   ELSE
   KK = INSTR(Arec(1).Nam, " ")
   IF KK > 0 THEN
    lasname$ = LEFT$(Arec(1).Nam, KK - 1)
    firname$ = RIGHT$(Arec(1).Nam, LEN(Arec(1).Nam) - KK)
         ELSE
    firname$ = ""
    lasname$ = Arec(1).Nam
   END IF
   END IF
IF LEN(RTRIM$(firname$)) = 0 AND LEN(RTRIM$(lasname$)) = 0 THEN GOTO Nextone
    TaxCustRec(1).FNAME = LTRIM$(firname$)
    TaxCustRec(1).LName = LTRIM$(lasname$)
    TaxCustRec(1).SName = LTRIM$(lasname$)
    Add3$ = RTRIM$(Arec(1).a3)
    IF LEN(Add3$) > 0 THEN

    cY$ = LEFT$(RTRIM$(Arec(1).a3), LEN(RTRIM$(Arec(1).a3)) - 2)
    st$ = RIGHT$(RTRIM$(Add3$), 2)
    ELSE
     IF LEN(RTRIM$(Arec(1).a2)) > 0 THEN
       cY$ = LEFT$(RTRIM$(Arec(1).a2), LEN(RTRIM$(Arec(1).a2)) - 2)
       st$ = RIGHT$(RTRIM$(Arec(1).a2), 2)
    END IF
   END IF
 
  
  
   Record! = Record! + 1
   ZipCode$ = LTRIM$(Arec(1).z1)
   IF LEN(RTRIM$(ZipCode$)) = 4 THEN ZipCode$ = "0" + ZipCode$
   z2$ = RTRIM$(Arec(1).z2)
   IF LTRIM$(z2$) = "0" THEN z2$ = ""
   IF LEN(LTRIM$(z2$)) < 4 AND LEN(LTRIM$(z2$)) > 1 THEN
    z2$ = STRING$(4 - LEN(LTRIM$(z2$)), "0") + LTRIM$(z2$)
   END IF

   IF LEN(z2$) = 4 THEN ZipCode$ = ZipCode$ + z2$

  
   TaxCustRec(1).Acct = Record!
   TaxCustRec(1).CountyAcct = VAL(Arec(1).Acct)
   TaxCustRec(1).Addr1 = Arec(1).A1
   TaxCustRec(1).Addr2 = Arec(1).a2
   TaxCustRec(1).CITY = cY$
   TaxCustRec(1).STATE = st$
   TaxCustRec(1).ZIP = ZipCode$
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
  

   TaxCustRec(1).CSSN = LEFT$(Arec(1).SS1, 9)
   TaxCustRec(1).SSSN = LEFT$(Arec(1).SS2, 9)
 
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "Y"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = x
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("01-01-1998")
   PUT TaxFile, Record!, TaxCustRec(1)
   PropNumb = 1
   'GOSUB updatepersonal

Nextone:
   NEXT x
   CLOSE
   EXIT SUB

UpdateSameAcct:
   PropNumb = PropNumb + 1
   GOSUB updatepersonal
   RETURN

updatepersonal:
   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0

   'IF Arec(1).PClass = "MC" THEN
   ' PersRec(1).MCVALUE = VAL(Arec(1).PValue)
   ' GOTO upd1
   'END IF
   'IF Arec(1).PClass = "MT" THEN
   ' PersRec(1).MTVALUE = VAL(Arec(1).PValue)
   ' GOTO upd1
   'END IF
   'IF Arec(1).PClass = "MH" THEN
   ' PersRec(1).MHVALUE = VAL(Arec(1).PValue)
   ' GOTO upd1
   'END IF
   'PersRec(1).PERSVAL = VAL(Arec(1).PValue)
upd1:
  ' PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
  ' PersRec(1).PROPDATE = Date2Num%("11-01-1998")
  ' PersRec(1).CVALUE = 0
  ' PersRec(1).EXMPSENI = 0
  ' PersRec(1).EXMPOTHR = 0
  ' PersRec(1).DISCOV = "N"
  ' PersRec(1).LATELIST = "N"
  ' PersRec(1).DESC1 = Arec(1).PVehID
  ' PersRec(1).DESC2 = Arec(1).PMake + "/" + Arec(1).PModel + "/" + Arec(1).PBody
  ' PersRec(1).DESC3 = ""
  ' PersRec(1).DESC4 = Arec(1).PYear
  ' PersRec(1).DESC5 = "Y" + "/" + Arec(1).PClass
  ' PersRec(1).CustPin = Record!
  ' PersRec(1).NextRec = 0
  ' PersRec(1).LastYrPrinted = 0
  ' PersRec(1).Deleted = 0
  ' PersRec(1).Blank = ""
  ' PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
  ' PUT PersTaxFile, PRecord&, PersRec(1)
  ' RETURN



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

