DEFINT A-Z
'Update Personal Prop Detail Records from Txt file  CedarBluff Va  2005
DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NUMOFTAXRECS%, TAXFILE%)
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
  DIM SHARED PersRec(1) AS PersonalRecType
  
  
  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  CLS
  ON ERROR RESUME NEXT
  DIM a$(88)

  OpenTaxCustFile NUMOFTAXRECS, TAXFILE
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  
  OPEN "CBOTXLNK" FOR INPUT AS #11

1 LINE INPUT #11, a$
  IF ERR = 62 THEN
    CLOSE
    PRINT "Done"
    PRINT "YCnt="; Quil
    PRINT "NCnt="; NonQ
    PRINT "BCnt="; BQui
    PRINT "YAmt="; YAmt#
    PRINT "NAmt="; NAmt#
    PRINT "BAmt="; BAmt#
    PRINT "  MT="; MT#
    PRINT "  MH="; MH#
    PRINT "EAmt="; eamt#
    PRINT "TAmt="; a#

    CALL KillFile("taxcon68.exe")
    END
  ELSEIF ERR > 0 THEN
    CLOSE
    PRINT "Unexpected error. Code:"; ERR
    END
  END IF

  acct$ = MID$(a$, 8, 8)
  Class$ = MID$(a$, 21, 2)
  Vin$ = MID$(a$, 23, 17)
  Amt# = VAL(MID$(a$, 51, 10))
  Make$ = RTRIM$(MID$(a$, 72, 10))
  Model$ = RTRIM$(MID$(a$, 82, 15))
  Body$ = RTRIM$(MID$(a$, 97, 15))
  Yr$ = MID$(a$, 113, 4)
  q$ = MID$(a$, 218, 1)
  PRINT Class$

  IF VAL(MID$(a$, 518, 4)) = 0 THEN eamt# = eamt# + Amt#: GOTO 1
  
  ta# = ta# + Amt#

  PRINT acct$, Vin$
  PRINT STRING$(79, "-")
  
  FOR Cnt& = 1 TO NUMOFTAXRECS
    GET TAXFILE, Cnt&, TaxCustRec(1)
    nf = 1
    IF TaxCustRec(1).CountyAcct = VAL(acct$) THEN
      GOSUB updatepersonal
      nf = 0
      EXIT FOR
    END IF
  NEXT Cnt&
  IF nf = 1 THEN STOP

Nextone:
  GOTO 1
  CLOSE
  EXIT SUB
  
  
updatepersonal:
  
  
  PersRec(1).CVALUE = 0
  PersRec(1).MCVALUE = 0
  PersRec(1).MTVALUE = 0
  PersRec(1).MHVALUE = 0
  PersRec(1).PERSVAL = 0
  
  IF Class$ = "MT" THEN
    PersRec(1).MTVALUE = Amt#
    MT# = MT# + Amt#
    GOTO upd1
  END IF
  IF Class$ = "MH" OR Class$ = "DW" THEN
    PersRec(1).MHVALUE = Amt#
    MH# = MH# + Amt#
    GOTO upd1
  END IF
  
  IF Class$ = "MC" THEN
    Amt# = 0
    PersRec(1).MCVALUE = Amt#
    GOTO upd1
  END IF
  
  PersRec(1).PERSVAL = Amt#
  
upd1:
  a# = a# + Amt#
  PersRec(1).PROPPIN = ""
  PersRec(1).PROPDATE = Date2Num%(DATE$)
  PersRec(1).CVALUE = 0
  PersRec(1).EXMPSENI = 0
  PersRec(1).EXMPOTHR = 0
  PersRec(1).DISCOV = "N"
  PersRec(1).LATELIST = "N"
  PersRec(1).DESC1 = Vin$
  PersRec(1).DESC2 = RTRIM$(Make$) + "/" + RTRIM$(Model$) + "/" + RTRIM$(Body$)
  PersRec(1).DESC3 = ""
  PersRec(1).DESC4 = Yr$
  IF q$ = "Y" THEN
    YAmt# = YAmt# + Amt#
    Quil = Quil + 1
    PersRec(1).DESC5 = "Y" + "/" + STRING$(24, " ") + "2006"
  ELSEIF q$ = "N" THEN
    NAmt# = NAmt# + Amt#
    NonQ = NonQ + 1
    PersRec(1).DESC5 = "N" + "/" + STRING$(24, " ") + "2006"
  ELSE
    BAmt# = BAmt# + Amt#
    BQui = BQui + 1
    PersRec(1).DESC5 = "N" + "/" + STRING$(24, " ") + "2006"
  END IF
  PersRec(1).CustPin = Cnt&
  PersRec(1).NextRec = 0
  PersRec(1).LastYrPrinted = 0
  PersRec(1).Deleted = 0
  PersRec(1).Blank = ""
  PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
  PUT PersTaxFile, PRecord&, PersRec(1)
  RETURN

END SUB

SUB OpenTaxCustFile (NUMOFTAXRECS, TAXFILE)
  
  TAXFILE = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TAXFILE LEN = LEN(TaxCustRec(1))
  NUMOFTAXRECS = LOF(TAXFILE) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, PersTaxFile)
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))
  
END SUB

