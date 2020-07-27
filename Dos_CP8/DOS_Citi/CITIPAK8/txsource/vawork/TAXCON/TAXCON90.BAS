DEFINT A-Z
'Update Personal Prop Detail Records from Txt file  Elkton PP 2000 and create tax bill
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
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TaxTrans(1) AS TaxTransactionType

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  
  DIM Acct!(2000), A$(12), Disc#(2000)

  OpenTaxCustFile NUMOFTAXRECS, TaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

  OPEN "SOSOFT1.TXT" FOR INPUT AS #11
1 FOR cc = 1 TO 10
  IF EOF(11) THEN GOTO 10
  INPUT #11, A$(cc)
  PRINT A$(cc), cc
NEXT cc
  IF C = 0 THEN
   C = 1
   Acct!(C) = VAL(A$(2))
   Disc#(C) = VAL(A$(9))
  ELSE
     FOUNDME = 0
     FOR ZZ = 1 TO C
                 IF VAL(A$(2)) = Acct!(ZZ) THEN
                   Disc#(ZZ) = Disc#(ZZ) + VAL(A$(9))
                   FOUNDME = 1
                   EXIT FOR
                 END IF
    NEXT ZZ
  IF FOUNDME = 0 THEN
  C = C + 1
  Acct!(C) = VAL(A$(2))
  Disc#(C) = VAL(A$(9))
  END IF

  END IF
  GOTO 1
10
   FOR ccc = 1 TO C
   PRINT Acct!(ccc), Disc#(ccc), ccc
   FOR Cnt& = NUMOFTAXRECS TO 1 STEP -1
    GET TaxFile, Cnt&, TaxCustRec(1)
    nf = 1
   IF VAL(TaxCustRec(1).CountyAcctString) = Acct!(ccc) THEN
    PRINT STRING$(79, "*")
    GOSUB updatepersonal
    nf = 0
    EXIT FOR
   END IF
   NEXT Cnt&
   NEXT ccc
Nextone:
   CLOSE
   CLS
   PRINT td#
   STOP
   EXIT SUB


updatepersonal:
  
  TransFile = FREEFILE
  OPEN "TaxTrans.dat" FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
  GET TransFile, TaxCustRec(1).LastTrans, TaxTrans(1)
    IF TaxTrans(1).TaxYear = 2000 AND TaxTrans(1).BillType = "P" AND TaxTrans(1).TranType = 1 THEN
      TaxTrans(1).TransDate = Date2Num("11-01-2000")
      TaxTrans(1).Amount = Round#(TaxTrans(1).Amount - Disc#(ccc))
      TaxTrans(1).Revenue.Principle1 = Round#(TaxTrans(1).Revenue.Principle1 - Disc#(ccc))
      TaxTrans(1).Revenue.Future1 = Round#(Disc#(ccc))
      PUT #TransFile, TaxCustRec(1).LastTrans, TaxTrans(1)
      td# = td# + Disc#(ccc)
    END IF

      CLOSE TransFile

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

