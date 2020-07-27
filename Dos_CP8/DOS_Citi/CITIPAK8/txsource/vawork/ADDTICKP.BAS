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
  
  FirstLine = -1
  
  OPEN "PGPP2005.CSV" FOR INPUT AS #11
  LOCATE 1, 1
  PRINT "Adding Ticket Numbers. . ."
1
  
  FOR c = 1 TO 25
    INPUT #11, a$(c)
    IF ERR = 62 THEN
      CLOSE
      LOCATE 10, 1
      PRINT "Done"
      CALL KillFile("addtickp.exe")
      END
    ELSEIF ERR > 0 THEN
      CLOSE
      PRINT "UNEXPECTED ERROR CODE IS:"; ERR
      PRINT "CALL SUPPORT. . ."
      END
    END IF
    IF c < 5 THEN
      LOCATE c + 3, 1
      PRINT Filler$;
      LOCATE c + 3, 1
      PRINT a$(c);
    END IF

  NEXT c
  
  IF FirstLine = -1 THEN
    FirstLine = 0
    GOTO 1
  END IF
  ReadCnt = ReadCnt + 1
  SSN$ = QPTrim$(a$(1))
  LOCATE 9, 1
  PRINT "Ticket:"; SSN$
  
  FOR Cnt& = 1 TO LOF(TaxFile) / LEN(TaxCustRec(1))
    GET TaxFile, Cnt&, TaxCustRec(1)
    IF TaxCustRec(1).CountyAcct = VAL(a$(11)) THEN
      TaxCustRec(1).CSSN = SSN$
      PUT TaxFile, Cnt&, TaxCustRec(1)
      EXIT FOR
    END IF
  NEXT Cnt&
  GOTO 1

  CLOSE
  EXIT SUB

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

