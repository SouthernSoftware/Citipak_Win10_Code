DEFINT A-Z
'Update Personal Prop Detail Records from Txt file  BrookNeal Va  2003
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NUMOFTAXRECS%, taxFile%)
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
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PersRec(1) AS PersonalRecType
 

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM a$(87)
  OpenTaxCustFile NUMOFTAXRECS, taxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

 'Init Prior Year All but Pers Prop which contains autos

  FOR Cnt& = 1 TO NumOfPersRecs
   GET PersTaxFile, Cnt&, PersRec(1)
    PersRec(1).CVALUE = 0
    PersRec(1).MCVALUE = 0
    PersRec(1).MTVALUE = 0
    PersRec(1).MHVALUE = 0
   PUT PersTaxFile, Cnt&, PersRec(1)
  NEXT Cnt&

  
  OPEN "TBPP2003.txt" FOR INPUT AS #11
1 LINE INPUT #11, a$




  CAcct$ = RTRIM$(MID$(a$, 80, 9))
  CAcct$ = LTRIM$(CAcct$)
  PRINT LEFT$(a$, 50)

   FOR Cnt& = 1 TO LOF(taxFile) / LEN(TaxCustRec(1))
    GET taxFile, Cnt&, TaxCustRec(1)
    nf = 1
   

   IF VAL(TaxCustRec(1).CSSN) = VAL(CAcct$) THEN
    Record! = Cnt&
    GOSUB UpdatePersonal
    nf = 0
    EXIT FOR
   END IF
   NEXT Cnt&

   IF nf = 1 THEN
     GOSUB AddAcct
     GOSUB UpdatePersonal
    END IF

Nextone:
   GOTO 1
   CLOSE
   EXIT SUB


UpdatePersonal:

   PersRec(1).CVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0

   t$ = MID$(a$, 200, 2)

   IF t$ = "MH" THEN
    PersRec(1).MHVALUE = VAL(MID$(a$, 274, 9))
    GOTO upd1
   END IF
   IF t$ = "MT" OR t$ = "4H" THEN
    PersRec(1).MTVALUE = VAL(MID$(a$, 274, 9))
    GOTO upd1
   END IF
   PersRec(1).PERSVAL = VAL(MID$(a$, 274, 9))

upd1:
   a# = a# + VAL(MID$(a$, 274, 9))
   PersRec(1).PROPPIN = ""
   PersRec(1).PROPDATE = Date2Num%("10-14-2003")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = MID$(a$, 253, 18)
   PersRec(1).DESC2 = RTRIM$(MID$(a$, 213, 10)) + "/" + RTRIM$(MID$(a$, 223, 15)) + "/" + RTRIM$(MID$(a$, 238, 15))
   PersRec(1).DESC3 = ""
   PersRec(1).DESC4 = MID$(a$, 209, 4)
   IF MID$(a$, 207, 1) = "Y" THEN
    PersRec(1).Desc5 = "Y" + "/" + STRING$(24, " ") + "2003"
   ELSE
    PersRec(1).Desc5 = "N" + "/" + STRING$(24, " ") + "2003"
   END IF
   PersRec(1).CustPin = Record!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 0
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
   PUT PersTaxFile, PRecord&, PersRec(1)
   RETURN

AddAcct:

   Record! = LOF(taxFile) / LEN(TaxCustRec(1)) + 1
   Add1$ = RTRIM$(MID$(a$, 129, 30))
   Add2$ = RTRIM$(MID$(a$, 99, 30))
   City$ = RTRIM$(MID$(a$, 159, 30))
   State$ = RIGHT$(City$, 2)
   City$ = RTRIM$(LEFT$(City$, LEN(City$) - 2))
   Zip$ = MID$(a$, 190, 5)
   Zip$ = LTRIM$(Zip$)
   IF LEN(Zip$) = 4 THEN Zip$ = "0" + Zip$
   Nme$ = MID$(a$, 9, 35)
   kk = INSTR(Nme$, " ")
   IF kk > 0 THEN
    LName$ = LEFT$(Nme$, kk - 1)
    Fme$ = MID$(Nme$, kk + 1, (LEN(Nme$) - kk))
    ELSE
    LName$ = Nme$
    Fme$ = ""
   END IF

   TaxCustRec(1).acct = Record!
   TaxCustRec(1).FName = Fme$
   TaxCustRec(1).LName = LName$
   TaxCustRec(1).SName = LName$
   TaxCustRec(1).Addr1 = Add1$
   TaxCustRec(1).Addr2 = Add2$
   
   TaxCustRec(1).City = City$
   TaxCustRec(1).State = State$
   TaxCustRec(1).Zip = Zip$
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   TaxCustRec(1).CSSN = MID$(a$, 80, 9)
   TaxCustRec(1).SSSN = MID$(a$, 90, 9)
   TaxCustRec(1).Employer = ""
   TaxCustRec(1).Bankrupt = "N"
   TaxCustRec(1).TownShip = ""
   TaxCustRec(1).CountyAcct = VAL(CAcct$)
   TaxCustRec(1).CountyAcctString = ""
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
   TaxCustRec(1).OPENDATE = Date2Num%(DATE$)
   PUT taxFile, Record!, TaxCustRec(1)
   RETURN

RETURN

END SUB

SUB OpenTaxCustFile (NUMOFTAXRECS, taxFile)
  
  taxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #taxFile LEN = LEN(TaxCustRec(1))
  NUMOFTAXRECS = LOF(taxFile) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, PersTaxFile)
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))
  
END SUB

