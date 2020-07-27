DEFINT A-Z
'Convert MASTER PP Records from Gordonsville TXT File YEAR=2000
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


  'Stephens City Va Tax PP Conversion of Master (Customers)

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
  DIM a$(38)
  OpenTaxCustFile NumOfTaxRecs, TaxFile


   OPEN "GVOTXMST.DAT" FOR INPUT AS #11
   CLS
   b = 0
1  LINE INPUT #11, a$
   
   IF LEFT$(a$, 1) = "D" THEN GOTO 1

   
   Acct! = VAL(MID$(a$, 6, 8))
   SSN$ = MID$(a$, 174, 9)
   SSN1$ = MID$(a$, 184, 9)
   NM$ = MID$(a$, 19, 34)
   Add1$ = LTRIM$(MID$(a$, 54, 30))
   Add2$ = LTRIM$(MID$(a$, 84, 30))
   Add3$ = LTRIM$(MID$(a$, 114, 30))
   Zip$ = MID$(a$, 145, 5)
   


   PRINT Acct!
   PRINT STRING$(79, "-")
   

   FOR Cnt& = 1 TO NumOfTaxRecs
   GET TaxFile, Cnt&, TaxCustRec(1)
   
   IF RTRIM$(SSN$) = RTRIM$(TaxCustRec(1).CSSN) AND LEN(RTRIM$(SSN$)) > 0 THEN
    Record! = Cnt&
    IF TaxCustRec(1).OPENDATE <> 7611 THEN

    GOSUB UpdateAcct
    GOTO 1
    END IF
   END IF

   NEXT Cnt&
       


    Record! = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1

   kk = INSTR(NM$, ",")
   IF kk = 0 THEN
    TaxCustRec(1).FNAME = ""
    TaxCustRec(1).LName = LTRIM$(NM$)
    TaxCustRec(1).SName = LTRIM$(NM$)
     ELSE
    TaxCustRec(1).FNAME = LTRIM$(MID$(NM$, kk + 1, LEN(NM$) - kk))
    TaxCustRec(1).LName = LTRIM$(LEFT$(NM$, kk - 1))
    TaxCustRec(1).SName = LTRIM$(LEFT$(NM$, kk - 1))
   END IF

    TaxCustRec(1).Acct = Record!
    TaxCustRec(1).Addr1 = Add1$

    IF LEN(RTRIM$(Add3$)) = 0 THEN
      TaxCustRec(1).Addr2 = ""
      TaxCustRec(1).City = Add2$
      TaxCustRec(1).STATE = ""
      TaxCustRec(1).Zip = Zip$
     ELSE
      TaxCustRec(1).Addr2 = Add2$
      TaxCustRec(1).City = Add3$
      TaxCustRec(1).STATE = ""
      TaxCustRec(1).Zip = Zip$
    END IF
   

    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = SSN$
    TaxCustRec(1).SSSN = SSN1$
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).CountyAcct = Acct!
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("11-01-2000")
    PUT TaxFile, Record!, TaxCustRec(1)
    GOTO 1

UpdateAcct:

   kk = INSTR(NM$, ",")
   IF kk = 0 THEN
     TaxCustRec(1).FNAME = ""
     TaxCustRec(1).LName = LTRIM$(NM$)
     TaxCustRec(1).SName = LTRIM$(NM$)
    ELSE
     TaxCustRec(1).FNAME = LTRIM$(MID$(NM$, kk + 1, LEN(NM$) - kk))
     TaxCustRec(1).LName = LTRIM$(LEFT$(NM$, kk - 1))
     TaxCustRec(1).SName = LTRIM$(LEFT$(NM$, kk - 1))
   END IF
    
    TaxCustRec(1).Acct = Record!
    TaxCustRec(1).Addr1 = Add1$

    IF LEN(RTRIM$(Add3$)) = 0 THEN
      TaxCustRec(1).Addr2 = ""
      TaxCustRec(1).City = Add2$
      TaxCustRec(1).STATE = ""
      TaxCustRec(1).Zip = Zip$
     ELSE
      TaxCustRec(1).Addr2 = Add2$
      TaxCustRec(1).City = Add3$
      TaxCustRec(1).STATE = ""
      TaxCustRec(1).Zip = Zip$
    END IF


    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = SSN$
    TaxCustRec(1).SSSN = SSN1$
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).CountyAcct = Acct!
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("11-01-2000")
    PUT TaxFile, Record!, TaxCustRec(1)
    RETURN

   

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

