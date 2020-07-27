DEFINT A-Z
'Update Personal Prop Saltville 2003 from .csv  (from smith county first)
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
  '$INCLUDE: 'TaxConv.bi'
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED Zip(1) AS ZipType
 

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
CLS
  ON ERROR RESUME NEXT
  DIM a$(75)
  OpenTaxCustFile NumOfTaxRecs, TaxFile
  OpenTaxPersFile NumOfPersRecs, PersTaxFile
  Record! = NumOfTaxRecs + 1

  OPEN "saltScpp.txt" FOR INPUT AS #11
1 FOR cc = 1 TO 44
    INPUT #11, a$(cc)
    IF ERR = 62 THEN
      CLOSE
      PRINT "Done"
      CALL KillFile("saltscpp.exe")
      END
    ELSEIF ERR > 0 THEN
      PRINT "Boom!!!"
      CLOSE
      END
    END IF
    a$(cc) = RTRIM$(a$(cc))
  NEXT cc
  DidCnt = DidCnt + 1
  IF DidCnt = 1 THEN
    GOTO 1
  END IF
  LOCATE 4, 1
  PRINT DidCnt; TAB(15); "SC" + RTRIM$(a$(6));
     'STOP
  
   FOR Cnt& = 1 TO LOF(TaxFile) / LEN(TaxCustRec(1))
    GET TaxFile, Cnt&, TaxCustRec(1)
    IF TaxCustRec(1).CountyAcctString = "SC" + RTRIM$(a$(6)) THEN
     CustRecord! = Cnt&
    kk = INSTR(a$(11), " ")
    IF kk > 0 THEN
     LN1$ = LEFT$(a$(11), kk - 1)
     FM1$ = MID$(a$(11), (kk + 1), (LEN(a$(11)) - (kk + 1)))
     SN1$ = LN1$
     ELSE
     LN1$ = a$(11)
     SN1$ = a$(11)
     FM1$ = ""
     END IF
 
     TaxCustRec(1).FNAME = LTRIM$(FM1$)
     TaxCustRec(1).LName = LTRIM$(LN1$)
     TaxCustRec(1).SName = LTRIM$(SN1$)
     TaxCustRec(1).Addr1 = LTRIM$(a$(13))
     TaxCustRec(1).Addr2 = LTRIM$(a$(14))
     TaxCustRec(1).City = a$(15)
      City$ = RTRIM$(TaxCustRec(1).City)
      State$ = RIGHT$(City$, 2)
      City$ = LEFT$(City$, LEN(City$) - 2)
     TaxCustRec(1).City = City$
     TaxCustRec(1).State = State$
     TaxCustRec(1).Zip = a$(16)
     TaxCustRec(1).CSSN = a$(18)
     TaxCustRec(1).SSSN = a$(19)
     TaxCustRec(1).CountyAcct = VAL(a$(6))
     TaxCustRec(1).CountyAcctString = "SC" + a$(6)
     TaxCustRec(1).PIN = CustRecord!
     TaxCustRec(1).OPENDATE = Date2Num%("11-11-2005")
     PUT TaxFile, CustRecord!, TaxCustRec(1)
     GOSUB UpdatePersonal
     GOTO 1
    END IF
    NEXT Cnt&
  
    GOSUB UpdateCust
    GOSUB UpdatePersonal
  

Nextone:
   GOTO 1
   CLOSE
   EXIT SUB

UpdateCust:
    
    kk = INSTR(a$(11), " ")
    IF kk > 0 THEN
     LN1$ = LEFT$(a$(11), kk - 1)
     FM1$ = MID$(a$(11), (kk + 1), (LEN(a$(11)) - (kk + 1)))
     SN1$ = LN1$
     ELSE
     LN1$ = a$(11)
     SN1$ = a$(11)
     FM1$ = ""
 END IF
    CustRecord! = (LOF(TaxFile) / LEN(TaxCustRec(1))) + 1
    TaxCustRec(1).FNAME = LTRIM$(FM1$)
    TaxCustRec(1).LName = LTRIM$(LN1$)
    TaxCustRec(1).SName = LTRIM$(SN1$)
    TaxCustRec(1).Acct = CustRecord!
    IF LEN(a$(15)) = 0 THEN
     TaxCustRec(1).Addr1 = a$(13)
     TaxCustRec(1).Addr2 = ""
     TaxCustRec(1).City = a$(14)
      ELSE
     TaxCustRec(1).Addr1 = a$(13)
     TaxCustRec(1).Addr2 = a$(14)
     TaxCustRec(1).City = a$(15)
    END IF
    City$ = RTRIM$(TaxCustRec(1).City)

    State$ = RIGHT$(City$, 2)
    City$ = LEFT$(City$, LEN(City$) - 2)


    TaxCustRec(1).City = City$
    TaxCustRec(1).State = State$
    TaxCustRec(1).Zip = a$(16)
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    TaxCustRec(1).CSSN = a$(18)
    TaxCustRec(1).SSSN = a$(19)
    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "Y"
    TaxCustRec(1).CountyAcct = VAL(a$(6))
    TaxCustRec(1).CountyAcctString = "SC" + a$(6)
    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = CustRecord!
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    TaxCustRec(1).OPENDATE = Date2Num%("11-11-2005")
    PUT TaxFile, CustRecord!, TaxCustRec(1)
    RETURN


UpdatePersonal:

   PersRec(1).CVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).MHVALUE = 0
   PersRec(1).PERSVAL = 0

   IF VAL(a$(37)) = 2 THEN
    PersRec(1).MHVALUE = VAL(a$(34))
    GOTO upd1
   END IF
   IF VAL(a$(37)) = 3 THEN
    PersRec(1).MTVALUE = VAL(a$(34))
    GOTO upd1
   END IF
   PersRec(1).PERSVAL = VAL(a$(34))
upd1:
   a# = a# + VAL(a$(34))
   PersRec(1).PROPPIN = ""
   PersRec(1).PROPDATE = Date2Num%("11-11-2005")
   PersRec(1).CVALUE = 0
   PersRec(1).EXMPSENI = 0
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   PersRec(1).LATELIST = "N"
   PersRec(1).DESC1 = a$(33)
   PersRec(1).DESC2 = RTRIM$(a$(30)) + "/" + RTRIM$(a$(31)) + "/" + RTRIM$(a$(32))
   PersRec(1).DESC3 = ""
   PersRec(1).DESC4 = a$(27)
   IF a$(43) = "Y" THEN
    PersRec(1).DESC5 = "Y" + "/" + STRING$(24, " ") + "2005"
   ELSE
    PersRec(1).DESC5 = "N" + "/" + STRING$(24, " ") + "2005"
   END IF
   PersRec(1).CustPin = CustRecord!
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

