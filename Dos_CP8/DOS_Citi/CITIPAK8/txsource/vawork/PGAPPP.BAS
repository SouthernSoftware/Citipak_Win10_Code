DEFINT A-Z
DECLARE SUB CleanUP20 (bb$)
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
  Filler$ = SPACE$(79)
  DIM a$(25)
  ON ERROR RESUME NEXT

  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  'First Initialize all pers prop values now (tax year 2003 only)
  FOR c = 1 TO NumOfPersRecs
    GET PersTaxFile, c, PersRec(1)
    PersRec(1).MCValue = 0
    PersRec(1).MTValue = 0
    PersRec(1).MHValue = 0
    PersRec(1).PersVal = 0
    PersRec(1).CVALUE = 0
    PUT PersTaxFile, c, PersRec(1)
  NEXT c
  FirstLine = -1
  
  OPEN "PGPP2005.CSV" FOR INPUT AS #11
  
1
  LOCATE 1, 1
  FOR c = 1 TO 25
    INPUT #11, a$(c)
    IF ERR = 62 THEN
      CLOSE
      PRINT "Done"
      CALL KillFile("pgappp.exe")
      END
    ELSEIF ERR > 0 THEN
      CLOSE
      PRINT "UNEXPECTED ERROR CODE IS:"; ERR
      PRINT "CALL SUPPORT. . ."
      END
    END IF
    IF c = 20 THEN
      CleanUP20 a$(c)
    END IF
    IF c < 23 THEN
      LOCATE c, 1
      PRINT Filler$;
      LOCATE c, 1
      PRINT a$(c), c;
    END IF

  NEXT c
  
  IF FirstLine = -1 THEN
    FirstLine = 0
    GOTO 1
  END IF
  ReadCnt = ReadCnt + 1

  'IF VAL(a$(1)) = 0 THEN END
  
  FOR Cnt& = 1 TO LOF(TaxFile) / LEN(TaxCustRec(1))
    GET TaxFile, Cnt&, TaxCustRec(1)
    IF TaxCustRec(1).CountyAcct = VAL(a$(11)) THEN
      'STOP
      'IF VAL(TaxCustRec(1).CSSN) = VAL(a$(9)) AND VAL(a$(9)) > 0 THEN
      Record! = Cnt&
      GOSUB UpdateMasterRec
      GOSUB UpdatePersonalRec
      GOTO 1
    END IF
  NEXT Cnt&
  NewOnes = NewOnes + 1
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

  IF a$(12) = "MC" THEN
    PersRec(1).MCValue = VAL(a$(20))
    GOTO upd1
  END IF
  IF a$(12) = "MT" THEN
    PersRec(1).MTValue = VAL(a$(20))
    GOTO upd1
  END IF
  IF a$(12) = "SW" THEN
    PersRec(1).MHValue = VAL(a$(20))
    GOTO upd1
  END IF
  
  PersRec(1).PersVal = VAL(a$(20))

upd1:
  PersRec(1).PROPPIN = LTRIM$(STR$(PropNumb))
  PersRec(1).PROPDATE = Date2Num%("09-20-2005")
  PersRec(1).CVALUE = 0
  PersRec(1).EXMPSENI = 0
  PersRec(1).EXMPOTHR = 0
  PersRec(1).DISCOV = "N"
  PersRec(1).LateList = "N"
  PersRec(1).DESC1 = a$(13)
  PersRec(1).DESC2 = RTRIM$(a$(14)) + "/" + RTRIM$(a$(15)) + "/" + RTRIM$(a$(16))
  PersRec(1).DESC3 = ""
  PersRec(1).Desc4 = a$(15)
  IF a$(24) = "Y" THEN
    PersRec(1).Desc5 = "Y" + "/" + STRING$(24, " ") + "2005"
  ELSE
    PersRec(1).Desc5 = "N" + "/" + STRING$(24, " ") + "2005"
  END IF
  PersRec(1).CustPin = Record!
  PersRec(1).NextRec = 0
  PersRec(1).LastYrPrinted = 2004
  PersRec(1).Deleted = 0
  PersRec(1).Blank = ""
  PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
  PUT PersTaxFile, PRecord&, PersRec(1)
  RETURN

UpdateMasterRec:
  IF a$(25) = "B" THEN
    TaxCustRec(1).FName = ""
    TaxCustRec(1).LName = a$(2)
    TaxCustRec(1).SName = a$(2)
  ELSE
    kk = INSTR(a$(2), " ")
    LstName$ = LEFT$(a$(2), kk)
    FstName$ = MID$(a$(2), kk + 1, (LEN(a$(2)) - kk))
    TaxCustRec(1).FName = FstName$
    TaxCustRec(1).LName = LstName$
    TaxCustRec(1).SName = LstName$
  END IF
  
  TaxCustRec(1).Addr1 = a$(4)
  TaxCustRec(1).Addr2 = a$(5)
  CITY$ = RTRIM$(a$(6))
  State$ = RIGHT$(CITY$, 2)
  CITY$ = LEFT$(CITY$, LEN(CITY$) - 2)
  
  TaxCustRec(1).CITY = CITY$
  TaxCustRec(1).State = State$
  TaxCustRec(1).ZIP = a$(7)
  TaxCustRec(1).CSSN = a$(9)
  TaxCustRec(1).SSSN = a$(10)
  TaxCustRec(1).CountyAcctString = a$(11)
  TaxCustRec(1).CountyAcct = VAL(a$(11))
  TaxCustRec(1).OPENDATE = Date2Num%("09-20-2005")
  PUT TaxFile, Record!, TaxCustRec(1)
  RETURN
  
AddMasterRec:
  Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1
  TaxCustRec(1).acct = Record!
  IF a$(25) = "B" THEN
    TaxCustRec(1).FName = ""
    TaxCustRec(1).LName = a$(2)
    TaxCustRec(1).SName = a$(2)
  ELSE
    kk = INSTR(a$(2), " ")
    LstName$ = LEFT$(a$(2), kk)
    FstName$ = MID$(a$(2), kk + 1, (LEN(a$(2)) - kk))
    TaxCustRec(1).FName = FstName$
    TaxCustRec(1).LName = LstName$
    TaxCustRec(1).SName = LstName$
  END IF
  
  TaxCustRec(1).Addr1 = a$(4)
  TaxCustRec(1).Addr2 = a$(5)
  CITY$ = RTRIM$(a$(6))
  State$ = RIGHT$(CITY$, 2)
  IF LEN(CITY$) > 0 THEN CITY$ = LEFT$(CITY$, LEN(CITY$) - 2)
  
  TaxCustRec(1).CITY = CITY$
  TaxCustRec(1).State = State$

  TaxCustRec(1).ZIP = a$(7)
  TaxCustRec(1).CSSN = a$(9)
  TaxCustRec(1).SSSN = a$(10)
  TaxCustRec(1).CountyAcctString = a$(11)
  TaxCustRec(1).CountyAcct = VAL(a$(11))
  TaxCustRec(1).Employer = ""
  TaxCustRec(1).Bankrupt = "N"
  TaxCustRec(1).TownShip = ""
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
  TaxCustRec(1).OPENDATE = Date2Num%("09-20-2005")
  PUT TaxFile, Record!, TaxCustRec(1)
  RETURN
  
END SUB

SUB CleanUP20 (bb$)
  bb$ = QPTrim$(bb$)

  IF LEFT$(bb$, 1) = "$" THEN
    bb$ = MID$(bb$, 2)
  END IF
  cpos = INSTR(bb$, ",")
  DO WHILE cpos > 0
    bb$ = LEFT$(bb$, cpos - 1) + MID$(bb$, cpos + 1)
    cpos = INSTR(bb$, ",")
  LOOP
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

