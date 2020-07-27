DEFINT A-Z
'Convert Master PP Records from TXT File (Stephens City 2001)
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
  DIM A$(57)
  OpenTaxCustFile NumOfTaxRecs, TaxFile

   OPEN "SCMSTPP2.csv" FOR INPUT AS #11
   CLS
   b = 0
1  FOR cc = 1 TO 57
    INPUT #11, A$(cc)
   NEXT cc
   PRINT "Name "; A$(7)
   PRINT "Acct Number "; A$(3)
   PRINT "SS# "; A$(18)
   PRINT STRING$(79, "-")

   TaxFile = FREEFILE
   OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
   NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
   FOR Srch! = 1 TO NumOfTaxRecs
   GET TaxFile, Srch!, TaxCustRec(1)

   IF TaxCustRec(1).CountyAcct = VAL(A$(3)) THEN
     GOSUB UpdateRecord
     CLOSE TaxFile
     GOTO 1
    END IF
   
   NEXT Srch!

  'Must Add Record : Could Not Find
  'Decode Name Here
    nm$ = UCASE$(A$(7))
    IF INSTR(nm$, "INC") > 0 THEN
     ln$ = UCASE$(A$(7))
     FM$ = ""
     sn$ = ln$
    ELSE
    kk = INSTR(nm$, " ")
    IF kk > 0 THEN
     ln$ = LEFT$(A$(7), kk - 1)
     FM$ = RIGHT$(A$(7), (LEN(A$(7)) - kk))
     sn$ = ln$
      ELSE
     ln$ = UCASE$(A$(7))
     FM$ = ""
     sn$ = ln$
    END IF

    END IF

    kk3 = INSTR(ln$, ",")
    IF kk3 > 3 THEN
     ln$ = LEFT$(ln$, kk3 - 1)
     sn$ = ln$
    END IF


    
    Record! = NumOfTaxRecs + 1


    
    TaxCustRec(1).FNAME = LTRIM$(FM$)
    TaxCustRec(1).LName = LTRIM$(ln$)
    TaxCustRec(1).SName = LTRIM$(sn$)
    TaxCustRec(1).Acct = Record!
    IF LEN(RTRIM$(A$(10))) = 0 THEN
     TaxCustRec(1).Addr1 = A$(8)
     TaxCustRec(1).Addr2 = ""
     TaxCustRec(1).City = A$(9)
     ELSE
     TaxCustRec(1).Addr1 = A$(8)
     TaxCustRec(1).Addr2 = A$(9)
     TaxCustRec(1).City = A$(10)
    END IF

    State$ = RTRIM$(TaxCustRec(1).City)
    State$ = RIGHT$(State$, 2)
    City$ = RTRIM$(TaxCustRec(1).City)
    City$ = LEFT$(City$, LEN(City$) - 4)
    


     TaxCustRec(1).City$ = City$
     TaxCustRec(1).State = State$


    TaxCustRec(1).Zip = A$(11)
    TaxCustRec(1).HPHONE = A$(25)
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = A$(18)
    TaxCustRec(1).SSSN = A$(19)
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).CountyAcct = VAL(A$(3))
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("01-01-2001")
    PUT TaxFile, Record!, TaxCustRec(1)
    CLOSE TaxFile
NextOne:
   GOTO 1
   CLOSE
   EXIT SUB

UpdateRecord:
     
     IF VAL(TaxCustRec(1).CSSN) = VAL(A$(18)) THEN RETURN

     

     nm$ = UCASE$(A$(7))
   IF INSTR(nm$, "INC") > 0 THEN
     ln$ = UCASE$(A$(7))
     FM$ = ""
     sn$ = ln$
    ELSE
     kk = INSTR(nm$, " ")
     IF kk > 0 THEN
      ln$ = LEFT$(A$(7), kk - 1)
      FM$ = RIGHT$(A$(7), (LEN(A$(7)) - kk))
      sn$ = ln$
      ELSE
      ln$ = UCASE$(A$(7))
      FM$ = ""
      sn$ = ln$
     END IF
    END IF
    
    kk3 = INSTR(ln$, ",")
    IF kk3 > 3 THEN ln$ = LEFT$(ln$, kk3 - 1)

    TaxCustRec(1).FNAME = LTRIM$(FM$)
    TaxCustRec(1).LName = LTRIM$(ln$)
    TaxCustRec(1).SName = LTRIM$(sn$)
    
    IF LEN(RTRIM$(A$(10))) = 0 THEN
     TaxCustRec(1).Addr1 = A$(8)
     TaxCustRec(1).Addr2 = ""
     TaxCustRec(1).City = A$(9)
     ELSE
     TaxCustRec(1).Addr1 = A$(8)
     TaxCustRec(1).Addr2 = A$(9)
     TaxCustRec(1).City = A$(10)
    END IF


    State$ = RTRIM$(TaxCustRec(1).City)
    State$ = RIGHT$(State$, 2)
    City$ = RTRIM$(TaxCustRec(1).City)
    City$ = LEFT$(City$, LEN(City$) - 4)

    TaxCustRec(1).City$ = City$
    TaxCustRec(1).State = State$
    TaxCustRec(1).Zip = A$(11)
    TaxCustRec(1).HPHONE = A$(25)
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = A$(18)
    TaxCustRec(1).SSSN = A$(19)
    TaxCustRec(1).OPENDATE = Date2Num%("01-01-2001")
    PUT TaxFile, Srch!, TaxCustRec(1)
    RETURN



END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, TaxFile)
  
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  
END SUB

