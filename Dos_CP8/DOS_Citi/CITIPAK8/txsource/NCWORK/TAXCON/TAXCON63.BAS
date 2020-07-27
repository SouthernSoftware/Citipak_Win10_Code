DEFINT A-Z
'Roseboro 2001
DECLARE FUNCTION Date2Num% (TheDate$)
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
DECLARE FUNCTION Round# (B#)
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
 
  'CONVERSION FROM OLD CITIPAK TAX BILLING TO NEW CITIPAK TAX BILLING
  
   DIM SHARED TaxCustRec(1) AS TaxCustType
   DIM SHARED PropertyRec(1) AS PropertyRecType
   DIM SHARED PersRec(1) AS PersonalRecType
   DIM SHARED TBCust(1) AS Oldtbcustdat
   DIM SHARED TBCust1(1) AS OldTBCUST1DAT
   DIM SHARED TBAB(1) AS TBABDATA
   DIM SHARED TBCustEx(1) AS TBCustEx


  
   STACK 5000
   BalanceListing
  
  END

SUB BalanceListing
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  'ROSEBORO INITIALIZE FIRST 1629 RECORDS SO THAT ACCT #'S WILL MATCH
  FOR CNT& = 1 TO 1629
   TaxCustRec(1).Acct = CNT&
   TaxCustRec(1).FName = ""
   TaxCustRec(1).LNAME = ""
   TaxCustRec(1).SName = ""
   TaxCustRec(1).Addr1 = ""
   TaxCustRec(1).Addr2 = ""
   TaxCustRec(1).CITY = ""
   TaxCustRec(1).State = ""
   TaxCustRec(1).Zip = ""
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).Active = "N"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "N"
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).pin = CNT&
    TaxCustRec(1).Deleted = 1
    TaxCustRec(1).FileVer = 8
    PUT TaxFile, CNT&, TaxCustRec(1)
   NEXT CNT&



   OPEN "TBCust.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 256
   OPEN "TBCust1.Dat" FOR RANDOM ACCESS READ WRITE SHARED AS #12 LEN = 64
   OPEN "TBABDATA.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #13 LEN = 2048
   OPEN "TBCUSTEX.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = 256

   Record = 0

1 FOR x = 1 TO LOF(11) / 256
   GET 11, x, TBCust(1)
   GET 12, x, TBCust1(1)
   GET 13, x, TBAB(1)
   GET 14, x, TBCustEx(1)
   
   
   IF VAL(TBCust(1).Acct) <= 0 THEN 10
   Record = VAL(TBCust(1).Acct)
   GET TaxFile, Record, TaxCustRec(1)
   TaxCustRec(1).OpenDate = Date2Num(DATE$)
   TaxCustRec(1).Acct = VAL(TBCust(1).Acct)
   TaxCustRec(1).FName = TBCust(1).FIRSTNAME
   TaxCustRec(1).LNAME = TBCust(1).LNAME
   TaxCustRec(1).SName = TBCust(1).LNAME
   TaxCustRec(1).Addr1 = TBCust(1).Addr1
   TaxCustRec(1).Addr2 = TBCust(1).Addr2
   TaxCustRec(1).CITY = TBCust(1).CITY
   TaxCustRec(1).State = TBCust(1).State
   TaxCustRec(1).Zip = LEFT$(TBCust(1).Zip, 5)
   TaxCustRec(1).HPHONE = TBCust(1).PHONE
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = TBCust1(1).SSN
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "N"
    TaxCustRec(1).pin = Record
    TaxCustRec(1).CountyAcctString = TBCustEx(1).CTYID
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    PUT TaxFile, Record, TaxCustRec(1)
    GOSUB UpdateReal
    GOSUB UpdatePers

10 NEXT x
   
   CLOSE
   EXIT SUB
  
  

UpdateReal:
    PropVal# = TBCustEx(1).RESVAL + TBCustEx(1).COMVAL + TBCustEx(1).INDVAL
    IF PropVal# <= 0 THEN RETURN
    PropertyRec(1).REALPIN = TBCust(1).Prop
    PropertyRec(1).PROPDATE = Date2Num(DATE$)
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = TBCust(1).MAP
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = ""
    PropertyRec(1).PROPSIZE = 0
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LateList = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = PropVal#
    PropertyRec(1).EXMPSENI = TBCustEx(1).SCRP
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = TBCust(1).Prop
    PropertyRec(1).PROPNOT2 = ""
    PropertyRec(1).PROPNOT3 = ""
    PropertyRec(1).Fill1 = ""
    PropertyRec(1).CustPin = Record
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 0
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN

UpdatePers:
   PersValue# = TBCustEx(1).ME + TBCustEx(1).MH + TBCustEx(1).MV + TBCustEx(1).BOAT + TBCustEx(1).PSER
   IF PersValue# <= 0 THEN RETURN
   PersRec(1).PROPPIN = "1"
   PersRec(1).PROPDATE = Date2Num(DATE$)
    
   PersRec(1).PersVal = PersValue#
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = TBCustEx(1).SCPP
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LateList = "N"
   PersRec(1).DESC1 = ""
   PersRec(1).DESC2 = ""
   PersRec(1).DESC3 = ""
   PersRec(1).Desc4 = ""
   PersRec(1).Desc5 = ""
   PersRec(1).CustPin = Record
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 0
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   NextRec = (LOF(PersTaxFile) / LEN(PersRec(1))) + 1
   PUT PersTaxFile, NextRec, PersRec(1)
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

