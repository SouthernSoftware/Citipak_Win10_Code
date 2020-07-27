DEFINT A-Z
'Convert ELKTON PP Records 2003 Master PP
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
  '$INCLUDE: 'TaxCnvrt.bi'

  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TaxTrans(1) AS TaxTransactionType


  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM a$(26)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  
  REDIM Elk(1) AS ElktonPPType
  OPEN "ELKPP.TXT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 200
  CLS
  
FOR Cnt& = 1 TO LOF(11) / 200
 GET 11, Cnt&, Elk(1)
 Account$ = Elk(1).AcctNumber
 PRINT Elk(1).AcctNumber
 PRINT Elk(1).Name
 PRINT STRING$(79, "-")

'FIRST LOOK FOR EXISTING ACCOUNT TO UPDATE
 Found = 0
 FOR Snt& = 1 TO NumOfTaxRecs
  GET TaxFile, Snt&, TaxCustRec(1)
  IF VAL(TaxCustRec(1).CountyAcctString) = VAL(Elk(1).AcctNumber) THEN
    kk = INSTR(Elk(1).Name, " ")
    IF kk > 0 THEN
      TaxCustRec(1).FName = LTRIM$(MID$(Elk(1).Name, kk + 1, (LEN(Elk(1).Name) - kk)))
      TaxCustRec(1).LName = LTRIM$(LEFT$(Elk(1).Name, kk - 1))
      TaxCustRec(1).SName = LTRIM$(LEFT$(Elk(1).Name, kk - 1))
      ELSE
      TaxCustRec(1).FName = ""
      TaxCustRec(1).LName = LTRIM$(Elk(1).Name)
      TaxCustRec(1).SName = LTRIM$(Elk(1).Name)
    END IF
  'string out zip
   City$ = Elk(1).City
   FOR ll = 1 TO LEN(City$)
   IF ASC(MID$(City$, ll, 1)) >= 48 AND ASC(MID$(City$, ll, 1)) <= 57 THEN
    City$ = LEFT$(City$, ll - 1)
    EXIT FOR
   END IF
   NEXT ll

   City$ = QPTrim$(City$)
   IF LEN(City$) > 2 THEN
   St$ = RIGHT$(City$, 2)
   City$ = RTRIM$(LEFT$(City$, LEN(City$) - 2))
   END IF

   Add1$ = Elk(1).Add1: Add1$ = QPTrim$(Add1$)
   Add2$ = Elk(1).Add2: Add2$ = QPTrim$(Add2$)
   IF LEN(Add1$) = 0 THEN Add1$ = Add2$: Add2$ = ""


   TaxCustRec(1).Addr1 = Add1$
   TaxCustRec(1).Addr2 = Add2$
   TaxCustRec(1).City = City$
   TaxCustRec(1).State = St$
   TaxCustRec(1).Zip = LEFT$(Elk(1).Zip, 5)
   TaxCustRec(1).CountyAcctString = Elk(1).AcctNumber
   'strip out any hypens
   SS1$ = ""
   FOR lll = 1 TO LEN(Elk(1).SS1)
   IF MID$(Elk(1).SS1, lll, 1) <> "-" THEN
   SS1$ = SS1$ + MID$(Elk(1).SS1, lll, 1)
   END IF
   NEXT lll
   SS2$ = ""
   FOR lll = 1 TO LEN(Elk(1).SS2)
   IF MID$(Elk(1).SS2, lll, 1) <> "-" THEN
   SS2$ = SS2$ + MID$(Elk(1).SS2, lll, 1)
   END IF
   NEXT lll

   TaxCustRec(1).CSSN = SS1$
   TaxCustRec(1).SSSN = Elk(1).BillNo
   TaxCustRec(1).TownShip = Elk(1).BillNo
   TaxCustRec(1).OPENDATE = Date2Num%("11-24-2003")
   PUT TaxFile, Snt&, TaxCustRec(1)
   GOTO Nextone
 END IF
 NEXT Snt&
  GOSUB UpdateCust
Nextone:
   NEXT Cnt&
   CLOSE
   EXIT SUB

UpdateCust:
   Record! = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1
    kk = INSTR(Elk(1).Name, " ")
    IF kk > 0 THEN
   TaxCustRec(1).FName = LTRIM$(MID$(Elk(1).Name, kk + 1, (LEN(Elk(1).Name) - kk)))
   TaxCustRec(1).LName = LTRIM$(LEFT$(Elk(1).Name, kk - 1))
   TaxCustRec(1).SName = LTRIM$(LEFT$(Elk(1).Name, kk - 1))

                ELSE
   TaxCustRec(1).FName = ""
   TaxCustRec(1).LName = LTRIM$(Elk(1).Name)
   TaxCustRec(1).SName = LTRIM$(Elk(1).Name)

    END IF
   
  'string out zip
   City$ = Elk(1).City
   FOR ll = 1 TO LEN(City$)
   IF ASC(MID$(City$, ll, 1)) >= 48 AND ASC(MID$(City$, ll, 1)) <= 57 THEN
    City$ = LEFT$(City$, ll - 1)
    EXIT FOR
   END IF
   NEXT ll

   City$ = QPTrim$(City$)
   IF LEN(City$) > 2 THEN
   St$ = RIGHT$(City$, 2)
   City$ = RTRIM$(LEFT$(City$, LEN(City$) - 2))
   END IF

   Add1$ = Elk(1).Add1: Add1$ = QPTrim$(Add1$)
   Add2$ = Elk(1).Add2: Add2$ = QPTrim$(Add2$)
   IF LEN(Add1$) = 0 THEN Add1$ = Add2$: Add2$ = ""

   
   TaxCustRec(1).Acct = Record!
   TaxCustRec(1).CountyAcctString = Elk(1).AcctNumber
   TaxCustRec(1).Addr1 = Add1$
   TaxCustRec(1).Addr2 = Add2$
   TaxCustRec(1).City = City$
   TaxCustRec(1).State = St$
   TaxCustRec(1).Zip = LEFT$(Elk(1).Zip, 5)
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).TownShip = Elk(1).BillNo

   'strip out any hypens
   SS1$ = ""
   FOR lll = 1 TO LEN(Elk(1).SS1)
   IF MID$(Elk(1).SS1, lll, 1) <> "-" THEN
   SS1$ = SS1$ + MID$(Elk(1).SS1, lll, 1)
   END IF
   NEXT lll
   SS2$ = ""
   FOR lll = 1 TO LEN(Elk(1).SS2)
   IF MID$(Elk(1).SS2, lll, 1) <> "-" THEN
   SS2$ = SS2$ + MID$(Elk(1).SS2, lll, 1)
   END IF
   NEXT lll

   TaxCustRec(1).CSSN = SS1$
   TaxCustRec(1).SSSN = Elk(1).BillNo
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "Y"
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = Record!
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   TaxCustRec(1).OPENDATE = Date2Num%("11-24-2003")
   PUT TaxFile, Record!, TaxCustRec(1)
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

