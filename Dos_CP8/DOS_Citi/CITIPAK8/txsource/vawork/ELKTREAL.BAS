DEFINT A-Z
'Real Estate Conversion from Text File Information ELKTON REAL 2003
'KILL TAXPROP AND RUN THE RELINK BEFORE STARTING CONVERSION
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NumOfTaxRecs%, TAXFILE%)
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
  CONST false = 0, True = NOT false
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  
  DIM SHARED TAXCUSTREC(1) AS TaxCustType
  DIM SHARED TaxTrans(1) AS TaxTransactionType
  DIM SHARED PROPERTYRec(1) AS PropertyRecType
  DIM SHARED REALRec(1) AS PropertyRecType
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  CLS
  ON ERROR RESUME NEXT
  DIM a$(20)

  CALL KillFile("RETXCUST.DAT") 'remove this for 2006
  CALL KillFile("TAXPROP.DAT")

  OpenTaxCustFile NumOfTaxRecs, TAXFILE
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  
  CLS
  DoneSkip = false

  OPEN "ELKTREAL.TXT" FOR INPUT AS #11

1
  FOR Cnt = 1 TO 19
    INPUT #11, a$(Cnt)
    IF ERR = 62 THEN
      CLOSE
      PRINT
      PRINT
      PRINT "Processed:"; AddCnt
      PRINT
      PRINT " Done."
      CALL KillFile("elktreal.exe")
      END
    ELSEIF ERR > 0 THEN
      CLOSE
      PRINT "Unexpected error:"; ERR
      END
    END IF
  NEXT
  IF NOT DoneSkip THEN
    DoneSkip = True
    GOTO 1
  END IF

  t = t + 1
  LNME$ = a$(6)
  'LNME$ = MID$(a$, 33, 5)
  LNME$ = RTRIM$(LNME$)

  'LandVal$ = MID$(a$, 297, 10)
  'BldgVal$ = MID$(a$, 307, 10)

  LandVal$ = a$(15)
  BldgVal$ = a$(16)
  
  TLand# = TLand# + VAL(LandVal$)
  TBldg# = TBldg# + VAL(BldgVal$)
  
  Nm$ = QPTrim$(a$(6))

  ADD1$ = QPTrim$(a$(8))
  'ADD2$ = MID$(a$, 103, 35)
  Zip$ = a$(9)  'MID$(a$, 139, 5)

  City$ = QPTrim$(a$(10))

  City$ = LEFT$(City$, LEN(City$) - 5)
  City$ = QPTrim$(City$)
  St$ = RIGHT$(City$, 2)
  City$ = LEFT$(City$, LEN(City$) - 2)
  City$ = QPTrim$(City$)

  DESC1$ = QPTrim$(a$(11))
  DESC2$ = QPTrim$(a$(12))
  DESC3$ = QPTrim$(a$(13))
  LOCATE 3, 1
  PRINT QPTrim$(a$(19)); TAB(15); LandVal$; TAB(30); BldgVal$; TAB(40); Nm$

  'Search For Record
  FOR Cnt& = 1 TO LOF(TAXFILE) / LEN(TAXCUSTREC(1))
    GET TAXFILE, Cnt&, TAXCUSTREC(1)

   'CACCT& = TAXCUSTREC(1).CountyAcct
    cc$ = QPTrim$(TAXCUSTREC(1).CountyAcctString)
    nc$ = QPTrim$(a$(19))

    IF cc$ = nc$ THEN
      Record! = Cnt&
      UpCnt = UpCnt + 1
      kk = INSTR(Nm$, " ")
      IF kk = 0 THEN
        LN$ = Nm$
        fm$ = ""
      ELSE
        LN$ = LEFT$(Nm$, kk - 1)
        fm$ = LTRIM$(MID$(Nm$, kk + 1, LEN(Nm$) - kk))
      END IF
      SN$ = LN$
      LN$ = UCASE$(LN$)
      fm$ = UCASE$(fm$)
      SN$ = UCASE$(SN$)
      
      'kk1 = INSTR(City$, "  ")
      'IF kk1 > 0 THEN City$ = LEFT$(City$, kk1 - 1)
      
      TAXCUSTREC(1).FName = fm$
      TAXCUSTREC(1).LNAME = LN$
      TAXCUSTREC(1).SName = SN$
      IF LEN(RTRIM$(ADD1$)) = 0 THEN
        TAXCUSTREC(1).ADDR1 = UCASE$(ADD2$)
        TAXCUSTREC(1).ADDR2 = ""
      ELSE
        TAXCUSTREC(1).ADDR1 = UCASE$(ADD1$)
        TAXCUSTREC(1).ADDR2 = UCASE$(ADD2$)
      END IF
      
      TAXCUSTREC(1).City = UCASE$(City$)
      TAXCUSTREC(1).State = St$
      TAXCUSTREC(1).Zip = Zip$
      ssn$ = "000" + a$(2)
      ssn$ = RIGHT$(ssn$, 4)
      TAXCUSTREC(1).CSSn = ssn$
      TAXCUSTREC(1).TownShip = "Y"
      TAXCUSTREC(1).CountyAcctString = QPTrim$(a$(19))
      TAXCUSTREC(1).FirstPropRec = 1
      TAXCUSTREC(1).OPENDATE = Date2Num%(DATE$)
      PUT TAXFILE, Record!, TAXCUSTREC(1)
      GOSUB UpdateReal
      GOTO 1
    END IF
  NEXT Cnt&
  'Add Here CUSTOMER NOT LOCATED
  'Decode Name Here
  AddCnt = AddCnt + 1
  Record! = LOF(TAXFILE) / LEN(TAXCUSTREC(1)) + 1
  
  kk = INSTR(Nm$, " ")
  IF kk = 0 THEN
    LN$ = Nm$
    fm$ = ""
  ELSE
    LN$ = LEFT$(Nm$, kk - 1)
    fm$ = LTRIM$(MID$(Nm$, kk + 1, LEN(Nm$) - kk))
  END IF
  SN$ = LN$
  LN$ = UCASE$(LN$)
  fm$ = UCASE$(fm$)
  SN$ = UCASE$(SN$)
  
  'kk1 = INSTR(City$, "  ")
  'IF kk1 > 0 THEN City$ = LEFT$(City$, kk1 - 1)
  
  Record! = (LOF(TAXFILE) / LEN(TAXCUSTREC(1))) + 1
  TAXCUSTREC(1).FName = fm$
  TAXCUSTREC(1).LNAME = LN$
  TAXCUSTREC(1).SName = SN$
  TAXCUSTREC(1).Acct = Record!
  
  IF LEN(RTRIM$(ADD1$)) = 0 THEN
    TAXCUSTREC(1).ADDR1 = UCASE$(ADD2$)
    TAXCUSTREC(1).ADDR2 = ""
  ELSE
    TAXCUSTREC(1).ADDR1 = UCASE$(ADD1$)
    TAXCUSTREC(1).ADDR2 = UCASE$(ADD2$)
  END IF
  
  TAXCUSTREC(1).City = UCASE$(City$)
  TAXCUSTREC(1).State = St$
  TAXCUSTREC(1).Zip = Zip$
  TAXCUSTREC(1).HPHONE = ""
  TAXCUSTREC(1).WPHONE = ""
  ssn$ = "000" + a$(2)
  ssn$ = RIGHT$(ssn$, 4)
  TAXCUSTREC(1).CSSn = ssn$
  'TaxCustRec(1).SSSN = ""
  TAXCUSTREC(1).CountyAcctString = QPTrim$(a$(19))
  TAXCUSTREC(1).CountyAcct = 0
  TAXCUSTREC(1).TownShip = "Y"
  TAXCUSTREC(1).Active = "Y"
  TAXCUSTREC(1).Interest = "Y"
  TAXCUSTREC(1).TaxExempt = "N"
  TAXCUSTREC(1).Penalty = "Y"
  TAXCUSTREC(1).LastTrans = 0
  TAXCUSTREC(1).FirstPropRec = 0
  TAXCUSTREC(1).FirstPersRec = 0
  TAXCUSTREC(1).Pin = Record!
  TAXCUSTREC(1).Deleted = 0
  TAXCUSTREC(1).FileVer = 8
  TAXCUSTREC(1).OPENDATE = Date2Num%(DATE$)
  PUT TAXFILE, Record!, TAXCUSTREC(1)
  PropNumb = 1
  GOSUB UpdateReal
  GOTO 1
  
NEXTONE:
  
  CLOSE
  EXIT SUB
  
UpdateReal:
  RERecord& = LOF(PropTaxFile) / LEN(REALRec(1)) + 1
  Pin$ = QPTrim$(a$(3))
  Pin$ = Pin$ + QPTrim$(a$(4))
  Pin$ = Pin$ + QPTrim$(a$(5))
  Pin$ = ssn$
  PROPERTYRec(1).REALPIN = "B" + Pin$
  PROPERTYRec(1).PROPDATE = Date2Num%(DATE$)
  PROPERTYRec(1).GISPOS = ""
  PROPERTYRec(1).MAP = QPTrim$(a$(3))
  PROPERTYRec(1).BLOCK = QPTrim$(a$(4))
  PROPERTYRec(1).LOTNUMB = QPTrim$(a$(5))
  PROPERTYRec(1).LOTACRE = "A"
  PROPERTYRec(1).PROPSIZE = VAL(a$(14))
  PROPERTYRec(1).PROPDISC = "N"
  PROPERTYRec(1).LATELIST = "N"
  PROPERTYRec(1).MORTCODE = ""
  PROPERTYRec(1).PROPVALU = VAL(LandVal$)
  PROPERTYRec(1).EXMPSENI = VAL(BldgVal$)
  PROPERTYRec(1).EXMPOTHR = 0
  PROPERTYRec(1).PROPNOT1 = DESC1$
  PROPERTYRec(1).PROPNOT2 = DESC2$
  PROPERTYRec(1).PROPNOT3 = DESC3$
  
  PROPERTYRec(1).CustPin = Record!
  PROPERTYRec(1).NextRec = 0
  PROPERTYRec(1).LastYrPrinted = 0
  PROPERTYRec(1).Deleted = 0
  PROPERTYRec(1).Blank = ""
  PUT PropTaxFile, RERecord&, PROPERTYRec(1)
  RETURN
  
END SUB

SUB keepthis

'  Cnt& = 0
'
'  OPEN "EACCT.TXT" FOR INPUT AS #11
'  DO
'    Cnt& = Cnt& + 1
'    LINE INPUT #11, zz$
'    SSSn$ = MID$(zz$, 10, 4)'
'    Acct$ = RIGHT$(zz$, 9)
'    PRINT Acct$; TAB(12); VAL(Acct$)
    'FOR CustCnt& = 1 TO NumOfTaxRecs
    '  GET TAXFILE, CustCnt&, TAXCUSTREC(1)
    '  CSSn$ = QPTrim$(TAXCUSTREC(1).CSSn)
    '  SSSN$ = QPTrim$(TAXCUSTREC(1).SSSN)
    '  IF LEN(CSSn$) = 0 AND LEN(SSSN$) = 0 THEN
    '    NoSSN = NoSSN + 1
      'STOP
    '  END IF

'      Cssn$ = QPTrim$(TAXCUSTREC(1).Cssn)
'      IF LEN(Cssn$) = 4 THEN
'        IF Cssn$ = SSSn$ THEN
'          TAXCUSTREC(1).CountyAcctString = QPTrim$(STR$(VAL(Acct$)))
'          PUT TAXFILE, CustCnt&, TAXCUSTREC(1)
'          MCnt = MCnt + 1
'          EXIT FOR
'        END IF
'      END IF
'    NEXT
'  LOOP UNTIL EOF(11)
'  CLOSE 11
  'CLOSE
  'END

END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TAXFILE)
  
  TAXFILE = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TAXFILE LEN = LEN(TAXCUSTREC(1))
  NumOfTaxRecs = LOF(TAXFILE) / LEN(TAXCUSTREC(1))
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, REALTaxFile)
  REALTaxFile = FREEFILE
  OPEN "TAXREAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #REALTaxFile LEN = LEN(REALRec(1))
  NumOfPersRecs = LOF(REALTaxFile) / LEN(REALRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PROPERTYRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PROPERTYRec(1))
END SUB

