DEFINT A-Z
'Convert Master PP Records from TXT File ST PAUL VA
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
  '$INCLUDE: 'TAXCONV.BI'
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED Pers(1) AS PERSTYPE
  

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM A$(31)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

   OPEN "PPTAXFL.TXT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = LEN(Pers(1))
   CLS
   
  FOR CNT& = 1 TO LOF(11) / LEN(Pers(1))
    GET 11, CNT&, Pers(1)

 'First Search for Existing Acct

    Record! = NumOfTaxRecs + 1
    NumOfTaxRecs = Record!
    TaxCustRec(1).FName = RTRIM$(Pers(1).FIRSTNAME) + " " + Pers(1).INITIAL
    TaxCustRec(1).LName = Pers(1).LASTNAME
    TaxCustRec(1).SName = Pers(1).LASTNAME
    TaxCustRec(1).Acct = Record!
    TaxCustRec(1).Addr1 = Pers(1).ADDRESS1
    TaxCustRec(1).Addr2 = Pers(1).ADDRESS2
    TaxCustRec(1).CITY = Pers(1).CITY
    TaxCustRec(1).STATE = Pers(1).STATE
    TaxCustRec(1).ZIP = Pers(1).ZIP
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = Pers(1).SSN
    TaxCustRec(1).SSSN = Pers(1).SSN2
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).CountyAcct = 0
    TaxCustRec(1).CountyAcctString = A$(1)
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("08-01-1999")
    PUT TaxFile, Record!, TaxCustRec(1)
    PropNumb = 1
    AcctRecord! = Record!
    GOSUB UpdatePersProperty
 
NextOne:
   NEXT CNT&
   CLOSE
   EXIT SUB

UpdatePersProperty:


  
   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PersVal = 0
   PersRec(1).CVALUE = 0

   IF Pers(1).PROPTYPE = "MT" THEN
    PersRec(1).MTVALUE = VAL(Pers(1).AVALUE)
    GOTO upd1
   END IF

   PersRec(1).PersVal = VAL(Pers(1).AVALUE)
upd1:

   PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
   PersRec(1).PROPDATE = Date2Num%("11-01-1999")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = Pers(1).MVVIN
   PersRec(1).DESC2 = Pers(1).MVMAKE
   PersRec(1).DESC3 = ""
   PersRec(1).Desc4 = Pers(1).MVYR
   PersRec(1).Desc5 = RTRIM$(Pers(1).MVPPTRA) + "/" + STRING$(23, " ") + "1999"
   PersRec(1).CustPin = AcctRecord!
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

