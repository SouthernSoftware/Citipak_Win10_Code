'SYLVA 2003
'This is Real
'TaxCon41 is Personal
DEFINT A-Z
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
DECLARE FUNCTION Date2Num% (TheDate$)
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
  

  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  DIM a$(120)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  CLS
  OPEN "SYLVAR~1.txt" FOR INPUT AS #11
  
1   FOR x = 1 TO 20
     INPUT #11, a$(x)
    NEXT x
    PRINT a$(1), a$(20)
   
   tr# = tr# + VAL(a$(17))
   
   FOR Cnt& = 1 TO NumOfTaxRecs
   GET TaxFile, Cnt&, TaxCustRec(1)
   found = 0: record = Cnt&
   IF TaxCustRec(1).CountyAcct = VAL(a$(8)) AND TaxCustRec(1).FirstPropRec = 0 THEN
     GOSUB UpdateReal
        GET TaxFile, Cnt&, TaxCustRec(1)
        TaxCustRec(1).OpenDate = Date2Num%(DATE$)
        TaxCustRec(1).FirstPropRec = NextRecord
        PUT TaxFile, Cnt&, TaxCustRec(1)
     GOTO 1
   END IF
   NEXT Cnt&
   
' ADD NEW ACCOUNT HERE
   record = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1
   TaxCustRec(1).Acct = record
   TaxCustRec(1).OpenDate = Date2Num%(DATE$)
   TaxCustRec(1).FName = a$(2)
   TaxCustRec(1).LName = a$(1)
   TaxCustRec(1).SName = a$(1)
   TaxCustRec(1).Addr1 = a$(3)
   TaxCustRec(1).Addr2 = a$(4)
   TaxCustRec(1).City = a$(5)
   TaxCustRec(1).State = a$(6)
   TaxCustRec(1).Zip = a$(7)
   TaxCustRec(1).HPhone = ""
   TaxCustRec(1).WPhone = ""
   TaxCustRec(1).CSSN = ""
   TaxCustRec(1).SSSN = ""
   TaxCustRec(1).Active = "Y"
   TaxCustRec(1).Interest = "Y"
   TaxCustRec(1).TaxExempt = "N"
   TaxCustRec(1).Penalty = "N"
   TaxCustRec(1).CountyAcct = VAL(a$(8))
   TaxCustRec(1).LastTrans = 0
   TaxCustRec(1).FirstPropRec = 0
   TaxCustRec(1).FirstPersRec = 0
   TaxCustRec(1).PIN = record
   TaxCustRec(1).Deleted = 0
   TaxCustRec(1).FileVer = 8
   PUT TaxFile, record, TaxCustRec(1)
   GOSUB UpdateReal
   GOTO 1
   CLOSE
   EXIT SUB
  
  

UpdateReal:

    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = a$(11)
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).PROPSIZE = VAL(a$(13))
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = a$(14)
    PropertyRec(1).PROPNOT2 = ""
    PropertyRec(1).PROPNOT3 = ""
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 0
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    PropertyRec(1).REALPIN = a$(10)
    PropertyRec(1).PROPDATE = Date2Num%(DATE$)
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = a$(12)
    PropertyRec(1).PROPVALU = VAL(a$(17))
    PropertyRec(1).EXMPSENI = VAL(a$(18)) + VAL(a$(19))
    PropertyRec(1).CustPin = record
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
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

