DEFINT A-Z
'Update Personal Prop Records For ALBERTA VA 2003
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

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  DIM a$(25)
   OpenTaxCustFile NUMOFTAXRECS, TaxFile
   OpenTaxPersFile NumOfPersRecs, PersTaxFile

   OPEN "ALBERTPP.DAT" FOR INPUT AS #10
1  LINE INPUT #10, a$
   PRINT LEFT$(a$, 8)
   Class$ = MID$(a$, 286, 2)
   IF Class$ = "MC" THEN
     GOTO 1
   END IF
   tmt# = tmt# + VAL(MID$(a$, 309, 9))
   
   

   FOR Cnt& = 1 TO LOF(TaxFile) / LEN(TaxCustRec(1))
   GET TaxFile, Cnt&, TaxCustRec(1)
   IF TaxCustRec(1).CountyAcct = VAL(LEFT$(a$, 8)) THEN
    Record! = Cnt&
    GOSUB UpdatePersonal
    GOTO 1
   END IF
'   IF LEFT$(TaxCustRec(1).CSSN, 9) = MID$(a$, 156, 9) THEN
'    Record! = Cnt&
'    GOSUB UpdatePersonal
'    GOTO 1
'   END IF

   NEXT Cnt&



   
    GOSUB AddAcct
    GOSUB UpdatePersonal
    GOTO 1


UpdatePersonal:
   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PersVal = 0
   PersRec(1).CVALUE = 0

   Class$ = MID$(a$, 286, 2)
   Amt# = VAL(MID$(a$, 309, 9))

   IF Class$ = "MC" THEN
   RETURN
   END IF

   IF Class$ = "MT" THEN
    PersRec(1).MTVALUE = Amt#
    GOTO upd1
   END IF
   IF Class$ = "MH" THEN
    PersRec(1).MHVALUE = Amt#
    GOTO upd1
   END IF
   PersRec(1).PersVal = Amt#
upd1:
   PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
   PersRec(1).PROPDATE = Date2Num%("10-13-2003")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = RTRIM$(MID$(a$, 288, 18))
   PersRec(1).DESC2 = RTRIM$(MID$(a$, 329, 10)) + "/" + RTRIM$(MID$(a$, 339, 15))
   PersRec(1).DESC3 = ""
   PersRec(1).Desc4 = RTRIM$(MID$(a$, 370, 4))
   IF MID$(a$, 527, 1) = "Y" THEN
     PersRec(1).Desc5 = "Y" + "/                       2003"
    ELSE
     PersRec(1).Desc5 = "N" + "/                       2003"
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
  
   Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1
   
   TaxCustRec(1).Acct = Record!

   NM$ = MID$(a$, 19, 35)

   IF MID$(a$, 18, 1) = "C" THEN
    F$ = ""
    LN$ = NM$
    ELSE
    KK = INSTR(NM$, " ")
    LN$ = LEFT$(NM$, KK - 1)
    F$ = MID$(NM$, KK + 1, LEN(NM$) - KK)
   END IF
 
   TaxCustRec(1).FName = F$
   TaxCustRec(1).LName = LN$
   TaxCustRec(1).SName = LN$
   TaxCustRec(1).Addr1 = QPTrim$(MID$(a$, 54, 30))
   TaxCustRec(1).Addr2 = QPTrim$(MID$(a$, 84, 30))
   TaxCustRec(1).City = QPTrim$(MID$(a$, 114, 30))
   TaxCustRec(1).State = ""
   TaxCustRec(1).ZIP = QPTrim$(MID$(a$, 145, 5))
   TaxCustRec(1).HPHONE = ""
   TaxCustRec(1).WPHONE = ""
   SS1$ = QPTrim$(MID$(a$, 156, 9))
   IF LEN(SS1$) = 8 THEN SS1$ = "0" + SS1$
   SS2$ = QPTrim$(MID$(a$, 166, 9))
   IF LEN(SS2$) = 8 THEN SS1$ = "0" + SS2$
 
   TaxCustRec(1).CSSN = SS1$
   TaxCustRec(1).SSSN = SS2$
   TaxCustRec(1).Employer = ""
   TaxCustRec(1).Bankrupt = "N"
   TaxCustRec(1).TownShip = ""
   TaxCustRec(1).CountyAcctString = ""
   TaxCustRec(1).CountyAcct = VAL(LEFT$(a$, 8))
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
   TaxCustRec(1).OPENDATE = Date2Num%("10-13-2003")
   PUT TaxFile, Record!, TaxCustRec(1)
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

