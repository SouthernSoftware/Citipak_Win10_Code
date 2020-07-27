DEFINT A-Z
'Update Personal Prop Records For Brookneal VA
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NUMOFTAXRECS%, TaxFile%)
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
  DIM SHARED bk(1) AS BrookNealType

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM A$(25)
   OpenTaxCustFile NUMOFTAXRECS, TaxFile
   OpenTaxPersFile NumOfPersRecs, PersTaxFile

   OPEN "BEALTKT" FOR RANDOM ACCESS READ WRITE SHARED AS #10 LEN = 678
   FOR MCnt& = 1 TO LOF(10) / 678
   GET 10, MCnt&, bk(1)
  
   'First Determine the Ticket Type
   'If Ticket TYpe=0 then Goto NextOne
   IF LEFT$(bk(1).NU, 1) = "0" THEN GOTO NextOne
    PType = VAL(LEFT$(bk(1).NU, 1))
   NF = 1
   FOR Cnt& = 1 TO NUMOFTAXRECS
   GET TaxFile, Cnt&, TaxCustRec(1)
   NF = 1
   IF TaxCustRec(1).CountyAcct = VAL(bk(1).AcctNo) THEN
    GOSUB UpdatePersonal
    NF = 0
    EXIT FOR
   END IF
   NEXT Cnt&
   IF NF = 1 THEN
    GOSUB AddAcct
    GOSUB UpdatePersonal
   END IF

NextOne:
   NEXT MCnt&
   CLOSE
   EXIT SUB


UpdatePersonal:
    

   
   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0
   PersRec(1).CVALUE = 0

   IF PType = 3 THEN
    PersRec(1).MCVALUE = VAL(bk(1).MCVALUE)
    GOTO upd1
   END IF
   IF PType = 2 THEN
    PersRec(1).MTVALUE = VAL(bk(1).MTVALUE)
    GOTO upd1
   END IF
   IF PType = 4 THEN
    PersRec(1).MHVALUE = VAL(bk(1).VALU)
    GOTO upd1
   END IF
   
   PersRec(1).PERSVAL = VAL(bk(1).VALU)
upd1:
   PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
   PersRec(1).PROPDATE = Date2Num%("11-01-1998")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   IF PType = 1 THEN
    PersRec(1).DESC1 = bk(1).Vin
    PersRec(1).DESC2 = bk(1).Make + "/" + bk(1).Model
    PersRec(1).DESC3 = ""
    PersRec(1).DESC4 = bk(1).Year
    IF bk(1).Qual = "Y" THEN
     PersRec(1).DESC5 = "Y" + "/"
    ELSE
     PersRec(1).DESC5 = "N" + "/"
    END IF
    ELSE
    PersRec(1).DESC1 = bk(1).Vin
    PersRec(1).DESC2 = bk(1).Make + "/" + bk(1).Model
    PersRec(1).DESC3 = ""
    PersRec(1).DESC4 = bk(1).Year
    PersRec(1).DESC5 = "N"
   END IF
   PersRec(1).CustPin = RECORD!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 1998
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
   PUT PersTaxFile, PRecord&, PersRec(1)
   RETURN

   

AddAcct:
   NUMOFTAXRECS = NUMOFTAXRECS + 1
   RECORD! = NUMOFTAXRECS
   
   Add3$ = RTRIM$(bk(1).Add3)
   Add2$ = RTRIM$(bk(1).Add2)
   IF LEN(Add3$) > 0 THEN
     City$ = Add3$
     TaxCustRec(1).Addr2 = Add2$
    ELSE
     City$ = Add2$
     TaxCustRec(1).Addr2 = ""
   END IF

  TaxCustRec(1).acct = RECORD!
     TaxCustRec(1).FName = bk(1).FName + " " + bk(1).MInt
     TaxCustRec(1).LName = bk(1).LName
     TaxCustRec(1).SName = bk(1).LName

   TaxCustRec(1).Addr1 = bk(1).Add1

   TaxCustRec(1).City = City$
   TaxCustRec(1).State = ""
   TaxCustRec(1).Zip = bk(1).Zip
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = bk(1).SSN
   TaxCustRec(1).SSSN = bk(1).SSN1
   TaxCustRec(1).Employer = ""
   TaxCustRec(1).Bankrupt = "N"
   TaxCustRec(1).TownShip = ""
   TaxCustRec(1).CountyAcctString = bk(1).SSN
   TaxCustRec(1).CountyAcct = VAL(bk(1).AcctNo)
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
   TaxCustRec(1).OPENDATE = Date2Num%("10-01-1999")
   PUT TaxFile, RECORD!, TaxCustRec(1)
   RETURN

END SUB

SUB OpenTaxCustFile (NUMOFTAXRECS, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NUMOFTAXRECS = LOF(TaxFile) / LEN(TaxCustRec(1))
  
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

