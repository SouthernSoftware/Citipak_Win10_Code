DEFINT A-Z
'Convert Kenpridge PP Records 2005 Master PP
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
  '$INCLUDE: 'TaxCnvrt.bi'
  
  
  TYPE TBKENPPCust
    AcctNumber       AS STRING * 10
    CustName         AS STRING * 25
    Add1             AS STRING * 25
    Add2             AS STRING * 25
    City             AS STRING * 20
    State            AS STRING * 2
    Zip              AS STRING * 10
    SS1              AS STRING * 12
    SS2              AS STRING * 12
    PersProp         AS STRING * 10
    MachTools        AS STRING * 10
    MobileHome       AS STRING * 10
    MerchCap         AS STRING * 10
    FarmEquip        AS STRING * 10
    Business         AS STRING * 1              'y/n
    District         AS STRING * 2
    LateFile         AS STRING * 1              'y/n
    Deleted          AS STRING * 1
    LName2           AS STRING * 35
    FName2           AS STRING * 35
    HeavyEquip       AS STRING * 10
    Boats            AS STRING * 10
    BusFurn          AS STRING * 10
    BillNo           AS STRING * 7
    TBCurYrPersPropTaxAmount AS STRING * 10
    TBCurYrFarmEquipTaxAmount AS STRING * 10
    TBCurYrMachToolsTaxAmount AS STRING * 10
    TBCurYrMerchCapTaxAmount AS STRING * 10
    TBCurYrMobileHomeTaxAmount AS STRING * 10
    TBLateFeeAmount  AS STRING * 10
    TBPPTRAValue     AS STRING * 10
    TBPPTRADiscnt    AS STRING * 10
    FF               AS STRING * 1
  END TYPE
  
  TYPE PPKenVehType
    VehTyp   AS STRING * 1      'D=Deleted Do Not Use
    VehVin   AS STRING * 22
    VehMake  AS STRING * 20
    VehYear  AS STRING * 4
    VehModel AS STRING * 15
    VehWght  AS STRING * 10
    VehValue AS STRING * 10
    VehQ     AS STRING * 1
    VehLic   AS STRING * 10
    VehAcct  AS STRING * 6
    crlf    AS STRING * 2
  END TYPE
  
  
  
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
  
  DIM KenMP AS TBKENPPCust
  
  DIM Ken AS PPKenVehType
  
  KPLen = LEN(Ken)
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  KPLen = LEN(KenMP)
  
'  OPEN "kenbrdve.txt" FOR RANDOM AS #10 LEN = KPLen
  
  OPEN "KENBRDPP.TXT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = KPLen
  CLS
  
  FOR Cnt& = 1 TO LOF(11) / KPLen
    GET 11, Cnt&, KenMP
    Account$ = KenMP.AcctNumber
    PRINT KenMP.AcctNumber
    PRINT KenMP.CustName
    PRINT STRING$(79, "-")
    'FIRST LOOK FOR EXISTING ACCOUNT TO UPDATE
    Found = 0
    FOR Snt& = 1 TO NumOfTaxRecs
      GET TaxFile, Snt&, TaxCustRec(1)
      IF VAL(TaxCustRec(1).CountyAcctString) = VAL(KenMP.AcctNumber) THEN
        kk = INSTR(KenMP.CustName, " ")
        IF kk > 0 THEN
          TaxCustRec(1).FName = LTRIM$(MID$(KenMP.CustName, kk + 1, (LEN(KenMP.CustName) - kk)))
          TaxCustRec(1).LName = LTRIM$(LEFT$(KenMP.CustName, kk - 1))
          TaxCustRec(1).SName = LTRIM$(LEFT$(KenMP.CustName, kk - 1))
        ELSE
          TaxCustRec(1).FName = ""
          TaxCustRec(1).LName = LTRIM$(KenMP.CustName)
          TaxCustRec(1).SName = LTRIM$(KenMP.CustName)
        END IF
        'string out zip
        City$ = KenMP.City
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
        
        Add1$ = KenMP.Add1: Add1$ = QPTrim$(Add1$)
        Add2$ = KenMP.Add2: Add2$ = QPTrim$(Add2$)
        IF LEN(Add1$) = 0 THEN Add1$ = Add2$: Add2$ = ""
        
        TaxCustRec(1).Addr1 = Add1$
        TaxCustRec(1).Addr2 = Add2$
        TaxCustRec(1).City = City$
        TaxCustRec(1).State = St$
        TaxCustRec(1).Zip = LEFT$(KenMP.Zip, 5)
        TaxCustRec(1).CountyAcctString = KenMP.AcctNumber
        TaxCustRec(1).CountyAcct = VAL(KenMP.AcctNumber)
        'strip out any hypens
        SS1$ = ""
        FOR lll = 1 TO LEN(KenMP.SS1)
          IF MID$(KenMP.SS1, lll, 1) <> "-" THEN
            SS1$ = SS1$ + MID$(KenMP.SS1, lll, 1)
          END IF
        NEXT lll
        SS2$ = ""
        FOR lll = 1 TO LEN(KenMP.SS2)
          IF MID$(KenMP.SS2, lll, 1) <> "-" THEN
            SS2$ = SS2$ + MID$(KenMP.SS2, lll, 1)
          END IF
        NEXT lll
        
        TaxCustRec(1).CSSN = SS1$
        TaxCustRec(1).SSSN = KenMP.BillNo
        TaxCustRec(1).TownShip = KenMP.BillNo
        TaxCustRec(1).OPENDATE = Date2Num%("11-15-2005")
        PUT TaxFile, Snt&, TaxCustRec(1)
        'don't know how this will be accomplished next year?????
        IF VAL(KenMP.MobileHome) > 0 THEN
        'IF VAL(KenMP.MobileHome) > 0 OR VAL(KenMP.FarmEquip) > 0 THEN
        'IF VAL(KenMP.MachTools) > 0 OR VAL(KenMP.MobileHome) > 0 OR VAL(KenMP.MerchCap) > 0 OR VAL(KenMP.FarmEquip) > 0 THEN
          GOSUB MakePersRec
        END IF

        GOTO Nextone
      END IF
    NEXT Snt&

    GOSUB UpdateCust
Nextone:
    'PRINT TaxCustRec(1).FName, TaxCustRec(1).LName
    'SLEEP 1
    
  NEXT Cnt&
  CLOSE
  EXIT SUB
  
UpdateCust:
  Record! = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1
  kk = INSTR(KenMP.CustName, " ")
  IF kk > 0 THEN
    TaxCustRec(1).FName = LTRIM$(MID$(KenMP.CustName, kk + 1, (LEN(KenMP.CustName) - kk)))
    TaxCustRec(1).LName = LTRIM$(LEFT$(KenMP.CustName, kk - 1))
    TaxCustRec(1).SName = LTRIM$(LEFT$(KenMP.CustName, kk - 1))
  ELSE
    TaxCustRec(1).FName = ""
    TaxCustRec(1).LName = LTRIM$(KenMP.CustName)
    TaxCustRec(1).SName = LTRIM$(KenMP.CustName)
  END IF
  
  'string out zip
  City$ = KenMP.City
'  FOR ll = 1 TO LEN(City$)
'    IF ASC(MID$(City$, ll, 1)) >= 48 AND ASC(MID$(City$, ll, 1)) <= 57 THEN
'      City$ = LEFT$(City$, ll - 1)
'      EXIT FOR
'    END IF
'  NEXT ll
  
'  City$ = QPTrim$(City$)
'  IF LEN(City$) > 2 THEN
'    St$ = RIGHT$(City$, 2)
'    City$ = RTRIM$(LEFT$(City$, LEN(City$) - 2))
'  END IF
 St$ = KenMP.State

  Add1$ = KenMP.Add1: Add1$ = QPTrim$(Add1$)
  Add2$ = KenMP.Add2: Add2$ = QPTrim$(Add2$)
  IF LEN(Add1$) = 0 THEN Add1$ = Add2$: Add2$ = ""
  
  
  TaxCustRec(1).acct = Record!
  TaxCustRec(1).CountyAcctString = KenMP.AcctNumber
  TaxCustRec(1).CountyAcct = VAL(KenMP.AcctNumber)
  TaxCustRec(1).Addr1 = Add1$
  TaxCustRec(1).Addr2 = Add2$
  TaxCustRec(1).City = City$
  TaxCustRec(1).State = St$
  TaxCustRec(1).Zip = LEFT$(KenMP.Zip, 5)
  TaxCustRec(1).HPHONE = ""
  TaxCustRec(1).WPHONE = ""
  TaxCustRec(1).TownShip = KenMP.BillNo
  
  'strip out any hypens
  SS1$ = ""
  FOR lll = 1 TO LEN(KenMP.SS1)
    IF MID$(KenMP.SS1, lll, 1) <> "-" THEN
      SS1$ = SS1$ + MID$(KenMP.SS1, lll, 1)
    END IF
  NEXT lll
  SS2$ = ""
  FOR lll = 1 TO LEN(KenMP.SS2)
    IF MID$(KenMP.SS2, lll, 1) <> "-" THEN
      SS2$ = SS2$ + MID$(KenMP.SS2, lll, 1)
    END IF
  NEXT lll
  
  TaxCustRec(1).CSSN = SS1$
  TaxCustRec(1).SSSN = KenMP.BillNo
  TaxCustRec(1).Active = "Y"
  TaxCustRec(1).Interest = "Y"
  TaxCustRec(1).TaxExempt = "N"
  TaxCustRec(1).Penalty = "Y"
  TaxCustRec(1).LastTrans = 0
  TaxCustRec(1).FirstPropRec = 0
  TaxCustRec(1).FirstPersRec = 0
  TaxCustRec(1).pin = Record!
  TaxCustRec(1).Deleted = 0
  TaxCustRec(1).FileVer = 8
  TaxCustRec(1).OPENDATE = Date2Num%("11-15-2005")
  PUT TaxFile, Record!, TaxCustRec(1)
  'don't know how this will be accomplished next year?????
  
  IF VAL(KenMP.MobileHome) > 0 THEN 'OR VAL(KenMP.MachTools) > 0 THEN
  'IF VAL(KenMP.MachTools) > 0 OR VAL(KenMP.MobileHome) > 0 OR VAL(KenMP.MerchCap) > 0 OR VAL(KenMP.FarmEquip) > 0 THEN
    GOSUB MakePersRec
  END IF

  RETURN

MakePersRec:

  PersRec(1).CVALUE = 0 'VAL(KenMP.FarmEquip)
  PersRec(1).MCVALUE = 0
  PersRec(1).MTVALUE = 0'VAL(KenMP.MachTools)
  PersRec(1).MHVALUE = VAL(KenMP.MobileHome)
  PersRec(1).PERSVAL = 0

upd1:
  PersRec(1).PROPPIN = ""
  PersRec(1).PROPDATE = Date2Num%("11-19-2004")
  'PersRec(1).CVALUE = 0
  PersRec(1).EXMPSENI = 0
  PersRec(1).EXMPOTHR = 0
  PersRec(1).DISCOV = "N"
  PersRec(1).LATELIST = "N"
  PersRec(1).DESC1 = ""
  PersRec(1).DESC2 = RTRIM$(Make$) + "/" + RTRIM$(Model$) + "/" + RTRIM$(Body$)
  PersRec(1).DESC3 = ""
  PersRec(1).DESC4 = Yr$
  IF q$ = "Y" THEN
    PersRec(1).DESC5 = "Y" + "/" + STRING$(24, " ") + "2005"
  ELSE
    PersRec(1).DESC5 = "N" + "/" + STRING$(24, " ") + "2005"
  END IF
  PersRec(1).CustPin = Record!
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

