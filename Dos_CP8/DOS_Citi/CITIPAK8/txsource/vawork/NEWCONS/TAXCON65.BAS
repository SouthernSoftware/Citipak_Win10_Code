DEFINT A-Z
'Update Personal Prop Detail Records from Txt file  Chilhowie Va  2003
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
  ON ERROR RESUME NEXT
  CLS
  DIM a$(88)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile

  OPEN "pp2004c.dat" FOR INPUT AS #11

1

  LINE INPUT #11, a$
  IF ERR = 62 THEN
    CLOSE
    PRINT "Done."
    END
  END IF

  Company$ = MID$(a$, 18, 1)
  acct$ = QPTrim$(MID$(a$, 1, 8))

  Class$ = MID$(a$, 286, 2)
  Vin$ = MID$(a$, 288, 17)
  Amt# = VAL(MID$(a$, 308, 10))
  Make$ = QPTrim$(MID$(a$, 329, 10))
  Model$ = QPTrim$(MID$(a$, 339, 15))
  Body$ = QPTrim$(MID$(a$, 354, 15))
  Yr$ = MID$(a$, 370, 4)
  q$ = MID$(a$, 527, 1)

  
  IF Class$ = "MC" THEN
    GOTO 1
  END IF
  



  PRINT acct$, Vin$
  PRINT STRING$(79, "-")
           
   FOR Cnt& = 1 TO LOF(TaxFile) / LEN(TaxCustRec(1))
    GET TaxFile, Cnt&, TaxCustRec(1)
    nf = 1
   IF TaxCustRec(1).CountyAcct = VAL(acct$) THEN
    Record! = Cnt&
    GOSUB UpdatePersonal
    nf = 0
    EXIT FOR
   END IF
   NEXT Cnt&
    IF nf = 1 THEN
     GOSUB AddNewAcct
     GOSUB UpdatePersonal
    END IF
Nextone:
   GOTO 1
   CLOSE
   EXIT SUB




AddNewAcct:
   acct$ = MID$(a$, 1, 8)
   Bus$ = MID$(a$, 18, 1)
   ANAME$ = QPTrim$(MID$(a$, 19, 35))
   Add1$ = QPTrim$(MID$(a$, 54, 30))
   Add2$ = QPTrim$(MID$(a$, 84, 30))
   add3$ = QPTrim$(MID$(a$, 114, 30))

   IF LEN(Add1$) = 0 THEN
    Add1$ = Add2$: Add2$ = ""
   END IF
   



   zip$ = MID$(a$, 145, 5)
   SS$ = MID$(a$, 156, 9)
   SS1$ = MID$(a$, 166, 9)


  'Decode Name Here
   IF Company$ = "C" THEN
     LN$ = LEFT$(ANAME$, kk - 1)
     FM$ = ""
     SN$ = LN$
      ELSE
      kk = INSTR(ANAME$, " ")
      IF kk > 0 THEN
        LN$ = LEFT$(ANAME$, kk - 1)
        FM$ = MID$(ANAME$, kk + 1, (LEN(ANAME$) - kk))
        SN$ = LN$
     ELSE
     LN$ = ANAME$
     FM$ = ""
     SN$ = LN$

     END IF
    END IF


    Record! = LOF(TaxFile) / LEN(TaxCustRec(1)) + 1



    TaxCustRec(1).FNAME = LTRIM$(FM$)
    TaxCustRec(1).LName = LTRIM$(LN$)
    TaxCustRec(1).SName = LTRIM$(SN$)
    TaxCustRec(1).acct = Record!
    IF LEN(RTRIM$(add3$)) = 0 THEN
     TaxCustRec(1).Addr1 = Add1$
     TaxCustRec(1).Addr2 = ""
     TaxCustRec(1).City = Add2$
     ELSE
     TaxCustRec(1).Addr1 = Add1$
     TaxCustRec(1).Addr2 = Add2$
     TaxCustRec(1).City = add3$
    END IF

    TaxCustRec(1).STATE = ""
    TaxCustRec(1).zip = zip$
    TaxCustRec(1).CSSN = SS$
    TaxCustRec(1).SSSN = SS1$
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).CountyAcct = VAL(acct$)
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = Record!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("09-25-2004")
    PUT TaxFile, Record!, TaxCustRec(1)
    RETURN


UpdatePersonal:
    
      
   PersRec(1).CVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0

   IF Class$ = "MT" THEN
    PersRec(1).MTVALUE = Amt#
    GOTO upd1
   END IF
   IF Class$ = "MH" THEN
    PersRec(1).MHVALUE = Amt#
    GOTO upd1
   END IF

   PersRec(1).PERSVAL = Amt#
upd1:
   a# = a# + Amt#
   PersRec(1).PROPPIN = ""
   PersRec(1).PROPDATE = Date2Num%("09-25-2004")
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
    PersRec(1).DESC5 = "Y" + "/" + STRING$(24, " ") + "2004"
   ELSE
    PersRec(1).DESC5 = "N" + "/" + STRING$(24, " ") + "2004"
   END IF
   PersRec(1).CustPin = Record!
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 0
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   PRecord& = LOF(PersTaxFile) / LEN(PersRec(1)) + 1
   PUT PersTaxFile, PRecord&, PersRec(1)
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

