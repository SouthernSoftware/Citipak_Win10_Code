DEFINT A-Z
'Convert ELKTON PP Records from TXT File Already Billed
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
  DIM SHARED TBCustRec(1) AS TBCustType
  DIM SHARED TBCustRec1(1) AS TBCust1Type
  DIM SHARED TaxTrans(1) AS TaxTransactionType
  

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM A$(26)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  
  REDIM Elk(1) AS ElktonPPType
  OPEN "ELKPP.TXT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 200
  CLS
  
FOR Cnt& = 1 TO LOF(11) / 200
 GET 11, Cnt&, Elk(1)
 PRINT Elk(1).AcctNumber
 PRINT Elk(1).Name
 PRINT Elk(1).Add1
 PRINT Elk(1).Add2
 PRINT Elk(1).City
 PRINT Elk(1).Zip
 PRINT Elk(1).TaxDue
 PRINT Elk(1).BillNo
 PRINT Elk(1).ss1
 PRINT Elk(1).ss2
 'SLEEP 1
 PRINT STRING$(79, "-")


 kk = INSTR(Elk(1).ss1, "-")
 IF kk = 3 THEN
  s4$ = LEFT$(Elk(1).ss1, 2) + MID$(Elk(1).ss1, 4, LEN(Elk(1).ss1) - 4)
 END IF
 IF kk = 4 THEN
  s4$ = LEFT$(Elk(1).ss1, 3) + MID$(Elk(1).ss1, 5, 2) + MID$(Elk(1).ss1, 8, 4)
 END IF
 PRINT s4$
 


'FIRST LOOK FOR EXISTING ACCOUNT TO UPDATE
 Found = 0
 FOR Snt& = 1 TO NumOfTaxRecs
 GET TaxFile, Snt&, TaxCustRec(1)
 IF RTRIM$(TaxCustRec(1).CSSN) = RTRIM$(s4$) THEN
  Record! = Snt&
  GOSUB UpdateCust
  Found = 1
  EXIT FOR
 END IF
 NEXT Snt&
 IF Found = 0 THEN
  GOSUB AddCust
 END IF

   
NextOne:
   NEXT Cnt&
   CLOSE
   EXIT SUB

UpdateCust:
    GOSUB UpdateTrans
    RETURN

AddCust:            
    Record! = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1
    IF kk = 3 THEN
     TaxCustRec(1).FName = ""
     TaxCustRec(1).LName = LTRIM$(Elk(1).Name)
     TaxCustRec(1).SName = LTRIM$(Elk(1).Name)
    ELSE
     ll = INSTR(Elk(1).Name, " ")
     TaxCustRec(1).FName = LTRIM$(MID$(Elk(1).Name, ll + 1, LEN(Elk(1).Name) - ll))
     TaxCustRec(1).LName = LTRIM$(LEFT$(Elk(1).Name, ll - 1))
     TaxCustRec(1).SName = LTRIM$(LEFT$(Elk(1).Name, ll - 1))
    END IF
   
    TaxCustRec(1).Acct = Record!
    TaxCustRec(1).Addr1 = Elk(1).Add1
    TaxCustRec(1).Addr2 = Elk(1).Add2
    TaxCustRec(1).City = Elk(1).City
    TaxCustRec(1).State = ""
    TaxCustRec(1).Zip = LEFT$(Elk(1).Zip, 5)
    TaxCustRec(1).CountyAcct = VAL(Elk(1).AcctNumber)
    TaxCustRec(1).CountyAcctString = Elk(1).AcctNumber
    TaxCustRec(1).PIN = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0

    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("08-01-1999")
    PUT TaxFile, Record!, TaxCustRec(1)
    PropNumb = 1
    GOSUB UpdateTrans
    RETURN

UpdateTrans:
  TransFile = FREEFILE
  OPEN "TaxTrans.dat" FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
  NextRec& = LOF(TransFile) / LEN(TaxTrans(1)) + 1

      TaxTrans(1).TaxYear = 1999
      TaxTrans(1).TranType = 1  '1=Bill 2=Payment 3=Release 4=Interest 5=Penalty 6=Collection/Ad Cost Billing
      TaxTrans(1).BillType = "P"                'R=Real P=Personal Property C=Combined (NC/GA)
      TaxTrans(1).Amount = VAL(Elk(1).TaxDue) / 100
      TaxTrans(1).Revenue.Principle1 = VAL(Elk(1).TaxDue) / 100
      TaxTrans(1).Revenue.Principle2 = 0
      TaxTrans(1).Revenue.Principle3 = 0
      TaxTrans(1).Revenue.Principle4 = 0
      TaxTrans(1).Revenue.Principle5 = 0
      TaxTrans(1).Revenue.Interest = 0
      TaxTrans(1).Revenue.Penalty = 0
      TaxTrans(1).Revenue.Collection = 0
      TaxTrans(1).Revenue.Future1 = 0
      TaxTrans(1).Revenue.Future2 = 0
      TaxTrans(1).Revenue.Principle1Pd = 0
      TaxTrans(1).Revenue.Principle2Pd = 0
      TaxTrans(1).Revenue.Principle3Pd = 0
      TaxTrans(1).Revenue.Principle4Pd = 0
      TaxTrans(1).Revenue.Principle5Pd = 0
      TaxTrans(1).Revenue.InterestPd = 0
      TaxTrans(1).Revenue.PenaltyPd = 0
      TaxTrans(1).Revenue.CollectionPd = 0
      TaxTrans(1).Revenue.Future1Pd = 0
      TaxTrans(1).Revenue.Future2Pd = 0
      TaxTrans(1).Description = "Tax Bill #" + Elk(1).BillNo
      TaxTrans(1).Posted2GL = "N"
      TaxTrans(1).CustomerRec = Record!
      TaxTrans(1).LastTrans = TaxCustRec(1).LastTrans
      TaxTrans(1).BelongTo = 0
      TaxTrans(1).Padding = ""

      PUT #TransFile, NextRec&, TaxTrans(1)

     GET #TaxFile, Record!, TaxCustRec(1)
      TaxCustRec(1).LastTrans = NextRec&
     PUT #TaxFile, Record!, TaxCustRec(1)

      CLOSE TransFile

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

