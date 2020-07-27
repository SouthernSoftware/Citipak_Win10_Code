DEFINT A-Z
'Real Estate Conv from Text File Information  Pen Gap Va from excel .csv file
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
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED RealRec(1) AS PropertyRecType
  
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  ON ERROR RESUME NEXT
  CLS
  DIM a$(112)
  CALL KillFile("TAXPROP.DAT")

  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  CLS
  DoneHeader = False
  Clr$ = SPACE$(79)
  OPEN "real2005.txt" FOR INPUT AS #11
1
  FOR cc = 1 TO 41
    INPUT #11, a$(cc)
    IF ERR = 62 THEN
      LOCATE 7, 1
      PRINT "Updated:"; UPCnt
      PRINT "  Added:"; NewCnt
      PRINT " Skiped:"; Skiped
      PRINT
      PRINT "Done."
      CALL KillFile("indpreal.exe")
      END
    ELSEIF ERR > 0 THEN
      CLOSE
      PRINT "Boom! Error Code:"; ERR
      END
    END IF
  NEXT cc

  IF NOT DoneHeader THEN
    DoneHeader = True
    GOTO 1
  END IF

  DidCnt = DidCnt + 1
  LOCATE 1, 1
  PRINT Clr$;

  LOCATE 1, 1
  PRINT a$(26), a$(11), a$(4)'total1#, total2#

  LOCATE 5, 1
  PRINT "Processing;"; DidCnt
  'GOTO 1

  IF VAL(a$(26)) = 0 THEN
    Skiped = Skiped + 1
    GOTO 1
  END IF
 
  FOR Srch& = 1 TO NumOfTaxRecs   'FIND THE ACCOUNT NOW
    GET TaxFile, Srch&, TaxCustRec(1)
    IF VAL(a$(26)) = TaxCustRec(1).CountyAcct THEN 'Account Located Update It Now
      Record! = Srch&
      'Foundit = True
      Anam$ = RTRIM$(a$(4))
      kk = INSTR(Anam$, ",")
      IF kk = 0 THEN
        kk = INSTR(Anam$, " ")
      END IF
      LN$ = LEFT$(Anam$, kk - 1)
      FM$ = LTRIM$(MID$(Anam$, kk + 1, LEN(Anam$) - kk))
      SN$ = LN$
      LN$ = UCASE$(LN$)
      FM$ = UCASE$(FM$)
      SN$ = UCASE$(SN$)

      TaxCustRec(1).FNAME = QPTrim$(FM$)
      TaxCustRec(1).LName = QPTrim$(LN$)
      TaxCustRec(1).SName = QPTrim$(SN$)
      TaxCustRec(1).Addr1 = QPTrim$(UCASE$(a$(6)))
      TaxCustRec(1).Addr2 = QPTrim$(UCASE$(a$(7)))
      TaxCustRec(1).City = QPTrim$(UCASE$(a$(8)))
      TaxCustRec(1).State = a$(9)
      TaxCustRec(1).ZIP = a$(10)
     'TaxCustRec(1).OPENDATE = Date2Num%("11-11-2005")
      PUT TaxFile, Record!, TaxCustRec(1)
      GOSUB Updatereal
      UPCnt = UPCnt + 1
      GOTO 1
    END IF
  NEXT Srch&

  'IF NOT Foundit THEN
    NewCnt = NewCnt + 1
  'ELSE
  '  UPCnt = UPCnt + 1
  '  GOTO 1
  'END IF

  'New Account to A
  'Decode Name Here
    Anam$ = RTRIM$(a$(4))
    kk = INSTR(Anam$, ",")
    IF kk = 0 THEN
      kk = INSTR(Anam$, " ")
    END IF

    LN$ = LEFT$(Anam$, kk - 1)
    FM$ = LTRIM$(MID$(Anam$, kk + 1, LEN(Anam$) - kk))

    SN$ = LN$
    LN$ = UCASE$(LN$)
    FM$ = UCASE$(FM$)
    SN$ = UCASE$(SN$)

    TaxCustRec(1).FNAME = QPTrim$(FM$)
    TaxCustRec(1).LName = QPTrim$(LN$)
    TaxCustRec(1).SName = QPTrim$(SN$)
    TaxCustRec(1).Addr1 = QPTrim$(UCASE$(a$(6)))
    TaxCustRec(1).Addr2 = QPTrim$(UCASE$(a$(7)))
    TaxCustRec(1).City = QPTrim$(UCASE$(a$(8)))
    TaxCustRec(1).State = a$(9)
    TaxCustRec(1).ZIP = a$(10)
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = ""
    TaxCustRec(1).SSSN = ""
   'TaxCustRec(1).CountyAcctString = a$(1)
    TaxCustRec(1).CountyAcct = VAL(a$(26))
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
    TaxCustRec(1).OPENDATE = Date2Num%("11-10-2005")
    Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1
    TaxCustRec(1).PIN = Record!
    PUT TaxFile, Record!, TaxCustRec(1)
    GOSUB Updatereal
    GOTO 1
   
   'STOP
   EXIT SUB


Updatereal:
    
    PropertyRec(1).RealPin = a$(14)
    PropertyRec(1).PROPDATE = Date2Num%("11-10-2005")
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).map = a$(14)
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = ""
    PropertyRec(1).LOTACRE = "A"
    PropertyRec(1).PROPSIZE = 0
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LATELIST = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = VAL(a$(19))
    PropertyRec(1).EXMPSENI = VAL(a$(20))
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = a$(13) + " " + a$(14)
    PropertyRec(1).PROPNOT2 = a$(15)
    PropertyRec(1).PROPNOT3 = a$(16)
    PropertyRec(1).CustPin = Record!
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 2004
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    RERecord& = LOF(PropTaxFile) / LEN(RealRec(1)) + 1
    PUT PropTaxFile, RERecord&, PropertyRec(1)

    RETURN





END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, REALTaxFile)
  REALTaxFile = FREEFILE
  OPEN "TAXREAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #REALTaxFile LEN = LEN(RealRec(1))
  NumOfPersRecs = LOF(REALTaxFile) / LEN(RealRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

