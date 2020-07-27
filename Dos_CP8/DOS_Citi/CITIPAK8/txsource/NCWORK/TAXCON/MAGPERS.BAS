  DEFINT A-Z
  'Convert Data From County (Maggie Valley Personal property)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
  
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
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  DIM a$(16)
  
  CALL KillFile("taxpers.dat")
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  CLS
  PRINT "Importing Personal property information. . ."
  
  OPEN "magpers.txt" FOR INPUT AS #10
  
  DO
    FOR a = 1 TO 16
      INPUT #10, a$(a)
    NEXT a
    
    PersVal# = VAL(a$(13))
    Eval# = VAL(a$(14))
    LLF$ = a$(16)
    
    FOR Cnt& = 1 TO NumOfTaxRecs
      GET TaxFile, Cnt&, TaxCustRec(1)
      LOCATE 5, 1:
      PRINT TaxCustRec(1).CountyAcctString;
      IF VAL(TaxCustRec(1).CountyAcctString) = VAL(a$(8)) THEN
        Record! = Cnt&
        TaxCustRec(1).OpenDate = Date2Num(DATE$)
        TaxCustRec(1).FName = LTRIM$(a$(2))
        TaxCustRec(1).LName = LTRIM$(a$(1))
        TaxCustRec(1).SName = LTRIM$(a$(1))
        TaxCustRec(1).Addr1 = a$(3)
        TaxCustRec(1).Addr2 = a$(4)
        TaxCustRec(1).City = a$(5)
        TaxCustRec(1).STATE = a$(6)
        TaxCustRec(1).Zip = a$(7)
        PUT TaxFile, Cnt&, TaxCustRec(1)
        
        GOSUB updatepers
        
        GOTO LoopBackJump
      END IF
      
    NEXT Cnt&
    
    
    'Must Add Account
    'Split Name Here
    Record! = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1
    TaxCustRec(1).Acct = Record!
    TaxCustRec(1).OpenDate = Date2Num(DATE$)
    TaxCustRec(1).FName = LTRIM$(a$(2))
    TaxCustRec(1).LName = LTRIM$(a$(1))
    TaxCustRec(1).SName = LTRIM$(a$(1))
    TaxCustRec(1).Addr1 = a$(3)
    TaxCustRec(1).Addr2 = a$(4)
    TaxCustRec(1).City = a$(5)
    TaxCustRec(1).STATE = a$(6)
    TaxCustRec(1).Zip = a$(7)
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = ""
    TaxCustRec(1).SSSN = ""
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "N"
    TaxCustRec(1).CountyAcct = 0
    TaxCustRec(1).CountyAcctString = a$(8)
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).Pin = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    PUT TaxFile, Record!, TaxCustRec(1)
    GOSUB updatepers
    
LoopBackJump:
    
  LOOP UNTIL EOF(10)
  CLOSE
  LOCATE 12, 1
  PRINT "Processing Complete."
  CALL KillFile("magpers.exe")
  END
  
  
updatepers:
  PersRec(1).PROPPIN = "1"
  PersRec(1).PROPDATE = Date2Num%(DATE$)
  PersRec(1).PersVal = PersVal#
  PersRec(1).MHVALUE = 0
  PersRec(1).MCVALUE = 0
  PersRec(1).CVALUE = 0
  PersRec(1).MTVALUE = 0
  PersRec(1).EXMPSENI = 0
  PersRec(1).EXMPOTHR = 0
  PersRec(1).DISCOV = "N"
  
  PersRec(1).LATELIST = LLF$
  PersRec(1).Desc1 = a$(9)
  PersRec(1).Desc2 = ""
  PersRec(1).Desc3 = ""
  PersRec(1).DESC4 = ""
  PersRec(1).DESC5 = ""
  
  PersRec(1).CustPin = Record!
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

