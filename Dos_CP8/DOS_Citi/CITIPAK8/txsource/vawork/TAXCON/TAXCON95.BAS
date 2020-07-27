DEFINT A-Z
'Update Personal Prop Records For Pennington Gap 2003 From excel .csv file
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
DECLARE SUB PrintHelp (h$)
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
  '$INCLUDE: 'Taxcnvrt.bi'
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  
  DIM SHARED bk(1) AS BrookNealType

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM a$(44)
   OpenTaxCustFile NumOfTaxRecs, TaxFile
   OpenTaxPersFile NumOfPersRecs, PersTaxFile

'First Initialize all pers prop values now (tax year 2003 only)
 FOR C = 1 TO NumOfPersRecs
 GET PersTaxFile, C, PersRec(1)
   PersRec(1).MCValue = 0
   PersRec(1).MTValue = 0
   PersRec(1).MHValue = 0
   PersRec(1).PersVal = 0
   PersRec(1).CVALUE = 0
 PUT PersTaxFile, C, PersRec(1)
 NEXT C


   OPEN "PGPP.CSV" FOR INPUT AS #11

1 FOR C = 1 TO 36
  INPUT #11, a$(C)
  PRINT a$(C), C
  NEXT C
  IF VAL(a$(1)) = 0 THEN END


 FOR Cnt& = 1 TO LOF(TaxFile) / LEN(TaxCustRec(1))
  GET TaxFile, Cnt&, TaxCustRec(1)
 
 IF VAL(TaxCustRec(1).CSSN) = VAL(a$(11)) AND VAL(a$(11)) > 0 THEN
    
    Record! = Cnt&
    GOSUB UpdateMasterRec
    GOSUB UpdatePersonalRec
    GOTO 1
  END IF

 NEXT Cnt&

   GOSUB AddMasterRec
   GOSUB UpdatePersonalRec
GOTO 1
   CLOSE
   EXIT SUB


UpdatePersonalRec:
   PersRec(1).MCValue = 0
   PersRec(1).MTValue = 0
   PersRec(1).MHValue = 0
   PersRec(1).PersVal = 0
   PersRec(1).CVALUE = 0

   IF a$(13) = "MC" THEN
    PersRec(1).MCValue = VAL(a$(32))
    GOTO upd1
   END IF
   IF a$(13) = "MT" THEN
    PersRec(1).MTValue = VAL(a$(32))
    GOTO upd1
   END IF
   IF a$(13) = "SW" THEN
    PersRec(1).MHValue = VAL(a$(32))
    GOTO upd1
   END IF
   
   PersRec(1).PersVal = VAL(a$(32))
upd1:
   PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
   PersRec(1).PROPDATE = Date2Num%("10-23-2003")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LateList = "N"
   PersRec(1).DESC1 = a$(14)
   PersRec(1).DESC2 = RTRIM$(a$(16)) + "/" + RTRIM$(a$(17)) + "/" + RTRIM$(a$(18))
   PersRec(1).DESC3 = ""
   PersRec(1).Desc4 = a$(15)
    IF a$(34) = "Y" THEN
     PersRec(1).Desc5 = "Y" + "/" + STRING$(24, " ") + "2003"
    ELSE
     PersRec(1).Desc5 = "N" + "/" + STRING$(24, " ") + "2003"
    END IF
   PersRec(1).CustPin = Record!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 2002
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
   PUT PersTaxFile, PRecord&, PersRec(1)
   RETURN

   

UpdateMasterRec:
    IF a$(35) = "B" THEN
     TaxCustRec(1).FName = ""
     TaxCustRec(1).LName = a$(4)
     TaxCustRec(1).SName = a$(4)
    ELSE
     kk = INSTR(a$(4), " ")
     LstName$ = LEFT$(a$(4), kk)
     FstName$ = MID$(a$(4), kk + 1, (LEN(a$(4)) - kk))
     TaxCustRec(1).FName = FstName$
     TaxCustRec(1).LName = LstName$
     TaxCustRec(1).SName = LstName$
    END IF

   TaxCustRec(1).Addr1 = a$(6)
   TaxCustRec(1).Addr2 = a$(7)
   CITY$ = RTRIM$(a$(8))
   State$ = RIGHT$(CITY$, 2)
   CITY$ = LEFT$(CITY$, LEN(CITY$) - 2)

   TaxCustRec(1).CITY = CITY$
   TaxCustRec(1).State = State$
   TaxCustRec(1).ZIP = a$(9)
   TaxCustRec(1).CSSN = a$(11)
   TaxCustRec(1).SSSN = a$(12)
   TaxCustRec(1).CountyAcctString = a$(1)
   TaxCustRec(1).CountyAcct = VAL(a$(1))
   TaxCustRec(1).OPENDATE = Date2Num%("10-23-2003")
   PUT TaxFile, Record!, TaxCustRec(1)
   RETURN

AddMasterRec:
   Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1
     TaxCustRec(1).acct = Record!
    IF a$(35) = "B" THEN
     TaxCustRec(1).FName = ""
     TaxCustRec(1).LName = a$(4)
     TaxCustRec(1).SName = a$(4)
    ELSE
     kk = INSTR(a$(4), " ")
     LstName$ = LEFT$(a$(4), kk)
     FstName$ = MID$(a$(4), kk + 1, (LEN(a$(4)) - kk))
     TaxCustRec(1).FName = FstName$
     TaxCustRec(1).LName = LstName$
     TaxCustRec(1).SName = LstName$
    END IF

   TaxCustRec(1).Addr1 = a$(6)
   TaxCustRec(1).Addr2 = a$(7)
   CITY$ = RTRIM$(a$(8))
   State$ = RIGHT$(CITY$, 2)
   IF LEN(CITY$) > 0 THEN CITY$ = LEFT$(CITY$, LEN(CITY$) - 2)
   
   TaxCustRec(1).CITY = CITY$
   TaxCustRec(1).State = State$
   TaxCustRec(1).ZIP = a$(9)
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = a$(11)
   TaxCustRec(1).SSSN = a$(12)
   TaxCustRec(1).Employer = ""
   TaxCustRec(1).Bankrupt = "N"
   TaxCustRec(1).TownShip = ""
   TaxCustRec(1).CountyAcctString = a$(1)
   TaxCustRec(1).CountyAcct = VAL(a$(1))
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
   TaxCustRec(1).OPENDATE = Date2Num%("10-23-2003")
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

