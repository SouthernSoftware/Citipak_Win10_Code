DEFINT A-Z
'Convert Data From County (Canton 2003)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNumber%)

DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NumOfTaxRecs%, TaxFile%)
DECLARE SUB OpenTaxPropFile (NumofPropRecs%, PropTaxFile%)
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
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  


  
  STACK 5000
  BalanceListing
  
  END

DEFSNG A
SUB BalanceListing
DIM a$(20)
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumofPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
 'Initialize the Props First

  FOR x = 1 TO NumofPropRecs
  GET PropTaxFile, x, PropertyRec(1)
    PropertyRec(1).PROPVALU = 0
    PropertyRec(1).EXMPSENI = 0
  PUT PropTaxFile, x, PropertyRec(1)
  NEXT x

 'Initalize the Customer to Set Prop Flag to 0 so that when updated will not put more than 1 property per account
  FOR x = 1 TO NumOfTaxRecs
    GET TaxFile, x, TaxCustRec(1)
    TaxCustRec(1).FirstPropRec = 0
    PUT TaxFile, x, TaxCustRec(1)
  NEXT x

  ON ERROR RESUME NEXT

  OPEN "canreal2.txt" FOR INPUT AS #10
  DidCnt = 0
  CLS
  PRINT "Importing Tax Data. . ."
10 :

  FOR a = 1 TO 20
    INPUT #10, a$(a)
    IF ERR = 62 THEN
      CLOSE
      LOCATE 12, 1
      PRINT "Processing Complete."
      CALL KillFile("cantreal.exe")
      CALL KillFile("canreal2.txt")
      END
    ELSEIF ERR > 0 THEN
      CLOSE
      PRINT "Unexpected Error Code: "; ERR
      PRINT "AAhhhggggg!!"
      END
    END IF
  NEXT a
  DidCnt = DidCnt + 1
  LOCATE 5, 1
  PRINT "Record:"; DidCnt;
  SDPin$ = a$(10)

  RealVal# = VAL(a$(15))
  BldgVal# = VAL(a$(16))

  EVal1# = VAL(a$(18))
  EVal2# = VAL(a$(19))


  'This Update Must Match Both Pin and Account Number to Work
  'Search the Pin # First
  'Also Before Running the Pers Prop Update Delete the TaxPers.Dat File
  'Then Run Conversion #71


  FOR x = 1 TO NumofPropRecs
   GET PropTaxFile, x, PropertyRec(1)
   IF SDPin$ = RTRIM$(PropertyRec(1).RealPin) THEN
     PropertyRecord& = x
    'Get the Customer and Check the Pin
     GET TaxFile, PropertyRec(1).CustPin, TaxCustRec(1)
     IF VAL(a$(8)) = VAL(TaxCustRec(1).CountyAcctString) AND TaxCustRec(1).FirstPropRec = 0 THEN
      'Update Link List Pointer so Property Has a Record
      TaxCustRec(1).FirstPropRec = PropertyRecord&
      PUT TaxFile, PropertyRec(1).CustPin, TaxCustRec(1)
        
      'Both Match So Update the property Value
        GET PropTaxFile, PropertyRecord&, PropertyRec(1)
        PropertyRec(1).RealPin = SDPin$
        PropertyRec(1).PROPDATE = Date2Num%(DATE$)
        PropertyRec(1).MAP = a$(11)
        PropertyRec(1).LOTACRE = a$(12)
        PropertyRec(1).PROPSIZE = VAL(a$(13))
        PropertyRec(1).PROPVALU = RealVal# + BldgVal#
        PropertyRec(1).EXMPSENI = EVal2#
        PropertyRec(1).PROPNOT1 = a$(14)
        PropertyRec(1).PROPNOT3 = a$(9)
        PUT PropTaxFile, PropertyRecord&, PropertyRec(1)
        GOTO 10
      END IF
      END IF

  ' PRINT SDPin$, PropertyRec(1).RealPin
  NEXT x
  'STOP


    Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1
    TaxCustRec(1).Acct = Record
    TaxCustRec(1).OpenDate = Date2Num(DATE$)
    TaxCustRec(1).FName = LTRIM$(a$(2))
    TaxCustRec(1).LName = LTRIM$(a$(1))
    TaxCustRec(1).SName = LTRIM$(a$(1))
    TaxCustRec(1).Addr1 = a$(3)
    TaxCustRec(1).Addr2 = a$(4)
    TaxCustRec(1).City = a$(5)
    TaxCustRec(1).STATE = a$(6)
    TaxCustRec(1).Zip = a$(7)
    TaxCustRec(1).CountyAcct = 0
    TaxCustRec(1).CountyAcctString = a$(8)
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = QPTrim$(a$(8))
    TaxCustRec(1).SSSN = ""
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "N"
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).Pin = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    PUT TaxFile, Record!, TaxCustRec(1)
    GOSUB UpdateReal
    GOTO 10


UpdateReal:
    PropertyRec(1).RealPin = SDPin$
    PropertyRec(1).PROPDATE = Date2Num%(DATE$)
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).MAP = a$(11)
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = a$(12)
    PropertyRec(1).PROPSIZE = VAL(a$(13))
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = RealVal# + BldgVal#
    PropertyRec(1).EXMPSENI = EVal2#
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = a$(14)
    PropertyRec(1).PROPNOT2 = ""
    PropertyRec(1).PROPNOT3 = a$(9)
    
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 0
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN
END SUB

DEFINT A
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

SUB OpenTaxPropFile (NumofPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumofPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

