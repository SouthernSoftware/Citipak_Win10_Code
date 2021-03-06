DECLARE FUNCTION MakePin$ (TPin$)
DEFINT A-Z
'Real Estate Conversion from Text File Information  CEDAR BLUFF VA 2005
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
  ''$INCLUDE: 'bobtaxes.BI'
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED REALRec(1) AS PropertyRecType
  
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  CLS
  DIM a$(36)
  ON ERROR RESUME NEXT

  CALL KillFile("TAXPROP.DAT")

  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  
  CLS
  OPEN "POTAXMST" FOR INPUT AS #11
  'OPEN "CBTAXMST" FOR INPUT AS #11


1
  LINE INPUT #11, a$
  IF ERR = 62 THEN
    CLOSE
    PRINT "Done"
    CALL KillFile("pocah3.exe")
    END
  ELSEIF ERR > 0 THEN
    CLOSE
    PRINT "Unexpected error. Code:"; ERR
    END
  END IF
  
  IF VAL(MID$(a$, 348, 2)) = 7 THEN GOTO 1      'tax exempt property
  spin$ = RTRIM$(MID$(a$, 239, 8))

  Exempt# = 0
   
  Real# = VAL(MID$(a$, 360, 10))
  Bldg# = VAL(MID$(a$, 370, 10))

  tReal# = tReal# + Real#
  tPers# = tPers# + Bldg#
  LOCATE 5, 1
  PRINT a$
  PRINT "REAL: "; USING "##########,#.##"; tReal#
  PRINT "BLDG: "; USING "##########,#.##"; tPers#
  PRINT " TOT: "; USING "##########,#.##"; tReal# + tPers#
  
  'Find Exiting Acct
  FOR Snt& = 1 TO NumOfTaxRecs
    GET TaxFile, Snt&, TaxCustRec(1)
    
    IF TaxCustRec(1).CountyAcct = VAL(MID$(a$, 6, 12)) THEN
      anam$ = RTRIM$(MID$(a$, 18, 35))
      add1$ = RTRIM$(MID$(a$, 53, 30))
      add2$ = RTRIM$(MID$(a$, 83, 30))
      add3$ = RTRIM$(MID$(a$, 113, 30))
      Zip$ = RTRIM$(MID$(a$, 144, 5))
      kk = INSTR(anam$, ", ")
      IF kk = 0 THEN
        LN$ = anam$
        FM$ = ""
      ELSE
        LN$ = LEFT$(anam$, kk - 1)
        FM$ = LTRIM$(MID$(anam$, kk + 1, LEN(anam$) - kk))
      END IF
      
      SN$ = LN$
      LN$ = UCASE$(LN$)
      FM$ = UCASE$(FM$)
      SN$ = UCASE$(SN$)
      
      TaxCustRec(1).FNAME = FM$
      TaxCustRec(1).LName = LN$
      TaxCustRec(1).SName = SN$
      TaxCustRec(1).Addr1 = add1$
      
      IF LEN(add3$) = 0 THEN
        TaxCustRec(1).Addr2 = ""
        TaxCustRec(1).City = add2$
      ELSE
        TaxCustRec(1).Addr2 = add2$
        TaxCustRec(1).City = add3$
      END IF
      
      'Chop off state in city and put in state
      City$ = RTRIM$(TaxCustRec(1).City)
      State$ = RIGHT$(RTRIM$(City$), 2)
      IF LEN(City$) > 0 THEN City$ = LEFT$(City$, LEN(City$) - 2)
      TaxCustRec(1).City = City$
      TaxCustRec(1).State = State$
      TaxCustRec(1).Zip = Zip$
      TaxCustRec(1).OPENDATE = Date2Num%(DATE$)
      PUT TaxFile, Snt&, TaxCustRec(1)
      Record! = Snt&
      GOSUB UpdateReal
      GOTO 1
      
    END IF
  NEXT Snt&
  
  'Add New Acct Here
  'Decode Name Here
  anam$ = RTRIM$(MID$(a$, 18, 35))
  add1$ = RTRIM$(MID$(a$, 53, 30))
  add2$ = RTRIM$(MID$(a$, 83, 30))
  add3$ = RTRIM$(MID$(a$, 113, 30))
  Zip$ = RTRIM$(MID$(a$, 144, 5))
  
  
  kk = INSTR(anam$, ", ")
  IF kk = 0 THEN
    LN$ = anam$
    FM$ = ""
  ELSE
    LN$ = LEFT$(anam$, kk - 1)
    FM$ = LTRIM$(MID$(anam$, kk + 1, LEN(anam$) - kk))
  END IF
  
  SN$ = LN$
  LN$ = UCASE$(LN$)
  FM$ = UCASE$(FM$)
  SN$ = UCASE$(SN$)
  
  Record! = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1
  
  TaxCustRec(1).FNAME = FM$
  TaxCustRec(1).LName = LN$
  TaxCustRec(1).SName = SN$
  TaxCustRec(1).Addr1 = add1$
  
  IF LEN(add3$) = 0 THEN
    TaxCustRec(1).Addr2 = ""
    TaxCustRec(1).City = add2$
  ELSE
    TaxCustRec(1).Addr2 = add2$
    TaxCustRec(1).City = add3$
  END IF
  'Chop off state in city and put in state
  City$ = RTRIM$(TaxCustRec(1).City)
  State$ = RIGHT$(RTRIM$(City$), 2)
  City$ = LEFT$(City$, LEN(City$) - 2)
  TaxCustRec(1).City = City$
  TaxCustRec(1).State = State$
  TaxCustRec(1).Zip = Zip$
  TaxCustRec(1).CountyAcct = VAL(MID$(a$, 11, 8))
  
  TaxCustRec(1).Acct = Record!
  
  TaxCustRec(1).HPHONE = ""
  TaxCustRec(1).WPHONE = ""
  TaxCustRec(1).CSSN = ""
  TaxCustRec(1).SSSN = ""
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
  PUT TaxFile, Record!, TaxCustRec(1)
  PropNumb = 1
  GOSUB UpdateReal
  
NEXTONE:
  GOTO 1
  CLOSE
  EXIT SUB
  
  
UpdateReal:
  RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1

  PropertyRec(1).RealPin = MakePin$(MID$(a$, 239, 30))

  PropertyRec(1).PROPDATE = Date2Num%(DATE$)
  PropertyRec(1).GISPOS = ""
  PropertyRec(1).Map = ""
  PropertyRec(1).BLOCK = ""
  PropertyRec(1).LOTNUMB = ""
  PropertyRec(1).LOTACRE = "A"
  PropertyRec(1).PropSize = 0
  PropertyRec(1).PROPDISC = "N"
  PropertyRec(1).LateList = "N"
  PropertyRec(1).MORTCODE = ""
  PropertyRec(1).PROPVALU = Real#
  PropertyRec(1).EXMPSENI = Bldg#
  PropertyRec(1).EXMPOTHR = Exempt#
  PropertyRec(1).PROPNOT1 = RTRIM$(MID$(a$, 164, 25))
  PropertyRec(1).PROPNOT2 = RTRIM$(MID$(a$, 189, 25))
  PropertyRec(1).PROPNOT3 = RTRIM$(MID$(a$, 214, 25))
  
  PropertyRec(1).CustPin = Record!
  PropertyRec(1).NextRec = 0
  PropertyRec(1).LastYrPrinted = 2005
  PropertyRec(1).Deleted = 0
  PropertyRec(1).Blank = ""
  PUT PropTaxFile, RERecord&, PropertyRec(1)
  RETURN
  
  
  
END SUB

FUNCTION MakePin$ (TPin$)

  TPin$ = QPTrim$(TPin$)
  DO
    SpacePos = INSTR(TPin$, " ")
    IF SpacePos > 0 THEN
      TPin$ = LEFT$(TPin$, SpacePos - 1) + MID$(TPin$, SpacePos + 1)
    ELSE
      EXIT DO
    END IF
  LOOP
  MakePin$ = TPin$

END FUNCTION

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, REALTaxFile)
  REALTaxFile = FREEFILE
  OPEN "TAXREAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #REALTaxFile LEN = LEN(REALRec(1))
  NumOfPersRecs = LOF(REALTaxFile) / LEN(REALRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

