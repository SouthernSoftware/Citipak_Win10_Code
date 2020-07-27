DEFINT A-Z
'Convert Master PP Records from TXT File (Cedar Bluff 2005)
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
  
  
  'New Market Va Tax PP Conversion of Master (Customers)
  
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
  ON ERROR RESUME NEXT
  DIM A$(55)
  
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  Record! = NumOfTaxRecs
  
  OPEN "CBOTXMST" FOR INPUT AS #11
  CLS
  b = 0
  
1
  LINE INPUT #11, A$
  IF ERR = 62 THEN
    CLOSE
    LOCATE 9, 1
    PRINT "Processed:"; DidCnt
    PRINT "Done."
    CALL KillFile("Taxcon67.exe")
    END
  ELSEIF ERR > 0 THEN
    CLOSE
    PRINT "Unexpected error. Code:"; ERR
    END
  END IF
  DidCnt = DidCnt + 1

  Acct$ = MID$(A$, 6, 8)
  Bus$ = MID$(A$, 18, 1)
  ANAME$ = MID$(A$, 19, 35)
  add1$ = MID$(A$, 54, 30)
  add2$ = MID$(A$, 84, 30)
  add3$ = MID$(A$, 114, 30)
  zip$ = MID$(A$, 145, 5)
  SS$ = MID$(A$, 174, 9)
  SS1$ = MID$(A$, 184, 9)
  LOCATE 1, 1
  PRINT A$
  PRINT STRING$(79, "-")
  
  
  FOR Cnt& = 1 TO NumOfTaxRecs
    GET TaxFile, Cnt&, TaxCustRec(1)
    IF TaxCustRec(1).CountyAcct = VAL(Acct$) AND VAL(Acct$) <> 0 THEN
      'Decode Name Here
      nm$ = UCASE$(ANAME$)
      IF INSTR(nm$, "INC") > 0 THEN
        LN$ = UCASE$(ANAME$)
        FM$ = ""
        SN$ = LN$
      ELSE
        kk = INSTR(nm$, ", ")
        IF kk > 0 THEN
          LN$ = LEFT$(ANAME$, kk - 1)
          FM$ = RIGHT$(ANAME$, (LEN(ANAME$) - kk))
          SN$ = LN$
        ELSE
          LN$ = UCASE$(ANAME$)
          FM$ = ""
          SN$ = LN$
        END IF
      END IF
      
      TaxCustRec(1).FNAME = LTRIM$(FM$)
      TaxCustRec(1).LName = LTRIM$(LN$)
      TaxCustRec(1).SName = LTRIM$(SN$)
      TaxCustRec(1).Acct = Record!
      IF LEN(RTRIM$(add3$)) = 0 THEN
        TaxCustRec(1).Addr1 = add1$
        TaxCustRec(1).Addr2 = ""
        TaxCustRec(1).City = add2$
      ELSE
        TaxCustRec(1).Addr1 = add1$
        TaxCustRec(1).Addr2 = add2$
        TaxCustRec(1).City = add3$
      END IF
      City$ = RTRIM$(TaxCustRec(1).City)
      City$ = LEFT$(City$, LEN(City$) - 2)
      State$ = RIGHT$(RTRIM$(TaxCustRec(1).City), 2)
      TaxCustRec(1).City = City$
      TaxCustRec(1).State = State$
      TaxCustRec(1).zip = zip$
      TaxCustRec(1).CSSN = SS$
      TaxCustRec(1).SSSN = SS1$
      PUT TaxFile, Cnt&, TaxCustRec(1)
      
      GOTO 1
    END IF
  NEXT Cnt&
  
  
  'Decode Name Here
  nm$ = UCASE$(ANAME$)
  IF INSTR(nm$, "INC") > 0 THEN
    LN$ = UCASE$(ANAME$)
    FM$ = ""
    SN$ = LN$
  ELSE
    kk = INSTR(nm$, ", ")
    IF kk > 0 THEN
      LN$ = LEFT$(ANAME$, kk - 1)
      FM$ = RIGHT$(ANAME$, (LEN(ANAME$) - kk))
      SN$ = LN$
    ELSE
      LN$ = UCASE$(ANAME$)
      FM$ = ""
      SN$ = LN$
    END IF
    
  END IF
  
  TaxCustRec(1).FNAME = LTRIM$(FM$)
  TaxCustRec(1).LName = LTRIM$(LN$)
  TaxCustRec(1).SName = LTRIM$(SN$)
  TaxCustRec(1).Acct = Record!
  IF LEN(RTRIM$(add3$)) = 0 THEN
    TaxCustRec(1).Addr1 = add1$
    TaxCustRec(1).Addr2 = ""
    TaxCustRec(1).City = add2$
  ELSE
    TaxCustRec(1).Addr1 = add1$
    TaxCustRec(1).Addr2 = add2$
    TaxCustRec(1).City = add3$
  END IF
  
  TaxCustRec(1).State = ""
  TaxCustRec(1).zip = zip$
  TaxCustRec(1).CSSN = SS$
  TaxCustRec(1).SSSN = SS1$
  TaxCustRec(1).Active = "Y"
  TaxCustRec(1).Interest = "Y"
  TaxCustRec(1).TaxExempt = "N"
  TaxCustRec(1).Penalty = "Y"
  TaxCustRec(1).CountyAcct = VAL(Acct$)
  TaxCustRec(1).LastTrans = 0
  TaxCustRec(1).FirstPropRec = 0
  TaxCustRec(1).FirstPersRec = 0
  TaxCustRec(1).PIN = Record!
  TaxCustRec(1).Deleted = 0
  TaxCustRec(1).FileVer = 8
  TaxCustRec(1).OPENDATE = Date2Num%(DATE$)
  Record! = Record! + 1
  PUT TaxFile, Record!, TaxCustRec(1)

NextOne:
  GOTO 1

  CLOSE
  EXIT SUB
  
  
UpdateRecord:
  TaxCustRec(1).CountyAcct = VAL(Acct$)
  PUT TaxFile, Srch!, TaxCustRec(1)
  RETURN
END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  
END SUB

