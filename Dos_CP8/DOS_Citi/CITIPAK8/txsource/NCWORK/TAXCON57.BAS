DEFINT A-Z
'Convert Data From County (Trent Woods 2004)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNumber%)

DECLARE SUB BalanceListing ()
DECLARE SUB OpenTaxCustFile (NumOfTaxRecs%, Taxfile%)
DECLARE SUB OpenTaxPropFile (NumOfPropRecs%, PropTaxFile%)
DECLARE SUB OpenTaxPersFile (NumOfPersRecs%, PersTaxFIle%)
DECLARE SUB DisplayTaxScrn (ScrnName$)
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
  ''$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'bobTaxes.BI'
  '$INCLUDE: 'Taxfiles.BI'
  ''$INCLUDE: 'PROPAbst.BI'
  
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType

  STACK 5000
  BalanceListing
  
  END

SUB BalanceListing
  
  OpenTaxCustFile NumOfTaxRecs, Taxfile
  OpenTaxPropFile NumOfPropRecs, PropTaxFile
  CALL KillFile("TAXPERS.DAT")
  OpenTaxPersFile NumOfPersRecs, PersTaxFIle
  FOR Snt& = 1 TO LOF(Taxfile) / LEN(TaxCustRec(1))
    GET Taxfile, Snt&, TaxCustRec(1)
    TaxCustRec(1).OptSrchDesc = ""
    PUT Taxfile, Snt&, TaxCustRec(1)
  NEXT

InitOldValues:

'  FOR S! = 1 TO NumOfPersRecs
'    GET PersTaxFIle, S!, PersRec(1)
'    PersRec(1).PersVal = 0
'    PersRec(1).MHVALUE = 0
'    PersRec(1).MCVALUE = 0
'    PersRec(1).CVALUE = 0
'    PersRec(1).MTVALUE = 0
'    PersRec(1).EXMPSENI = 0
'    PersRec(1).EXMPOTHR = 0
'    PUT PersTaxFIle, S!, PersRec(1)
'  NEXT
'
  FOR S! = 1 TO NumOfPropRecs
    GET PropTaxFile, S!, PropertyRec(1)
    CLS

'    'reset pin numbers
    PRINT PropertyRec(1).RealPin
    PRINT PropertyRec(1).Map, PropertyRec(1).LOTNUMB
    IF LEN(QPTrim$(PropertyRec(1).RealPin)) <= 9 THEN
      PropertyRec(1).RealPin = PropertyRec(1).Map + PropertyRec(1).LOTNUMB
    END IF
    PRINT PropertyRec(1).RealPin
    PropertyRec(1).PROPVALU = 0
    PropertyRec(1).EXMPSENI = 0
    PUT PropTaxFile, S!, PropertyRec(1)
  NEXT S!

   OPEN "trwbill.txt" FOR INPUT AS #10
   Record = 0
   CLS

   ON ERROR RESUME NEXT

10
    LINE INPUT #10, A$
    IF ERR = 62 THEN
      CLOSE
      PRINT
      PRINT
      PRINT "DATA IMPORT COMPLETE."
      PRINT "New:"; NewOne
      PRINT "PLEASE RELINK ABSTRACTS NEXT!"
      CALL KillFile("txconwod.exe")
      END
    ELSEIF ERR > 0 THEN
      CLOSE
      PRINT
      PRINT "UNEXPECTED ERROR CODE IS:"; ERR
      PRINT "CONVERSION TERMINATED..."
      END
    END IF

    CTYACCT$ = LTRIM$(LEFT$(A$, 8))
    NM$ = RTRIM$(MID$(A$, 13, 35))

    Add1$ = RTRIM$(MID$(A$, 48, 35))
    IF LEN(Add1$) = 0 THEN
      Add1$ = RTRIM$(MID$(A$, 118, 35))
      Add2$ = ""
    ELSE
      Add2$ = RTRIM$(MID$(A$, 118, 35))
    END IF
    City$ = RTRIM$(MID$(A$, 153, 35))
    St$ = MID$(A$, 188, 2)
    Zip$ = MID$(A$, 191, 5) + MID$(A$, 197, 4)
    LateFlg$ = MID$(A$, 281, 1)
    Psc$ = MID$(A$, 263, 8)
    RSC$ = MID$(A$, 431, 8)
    RSC$ = QPTrim$(RSC$)

    'IF VAL(RSC$) > 0 THEN STOP

    RV$ = MID$(A$, 379, 10)
    PV$ = MID$(A$, 253, 10)
    TPv# = TPv# + VAL(PV$)

    Bus$ = MID$(A$, 12, 1)
    BILLNUM$ = QPTrim$(MID$(A$, 501, 6))
   
    DESC1$ = MID$(A$, 308, 20)
    DESC2$ = MID$(A$, 328, 20)
    DESC3$ = MID$(A$, 348, 20)

    PIN$ = MID$(A$, 294, 14)
    PRINT RV$, PV$, Record
    Psc# = Psc# + VAL(Psc$)
    RSC# = RSC# + VAL(RSC$)
    P# = P# + VAL(PV$)
    R# = R# + VAL(RV$)

    IF VAL(RV$) = 0 AND VAL(PV$) > 0 THEN

    'find by acct number and update pers only
    FOR Snt& = 1 TO LOF(Taxfile) / LEN(TaxCustRec(1))
      GET Taxfile, Snt&, TaxCustRec(1)
      IF TaxCustRec(1).CountyAcct = VAL(CTYACCT$) THEN
        'IF Snt& = 417 THEN STOP
        'UpDateCnt = UpDateCnt + 1
        'Split Name Here
        kk = INSTR(NM$, ",")
        IF kk > 0 THEN
          LNM$ = LEFT$(NM$, kk - 1)
          FM$ = MID$(NM$, kk + 1, LEN(NM$) - kk)
        ELSE
          LNM$ = NM$
          FM$ = ""
        END IF
        'TaxCustRec(1).FName = LTRIM$(FM$)
        'TaxCustRec(1).LName = LTRIM$(LNM$)

        'TaxCustRec(1).CustName = QPTrim$(FM$) + " " + QPTrim$(LNM$)
        TaxCustRec(1).CustName = QPTrim$(NM$)
        TaxCustRec(1).SName = LTRIM$(LNM$) + QPTrim$(FM$)
        TaxCustRec(1).Addr1 = Add1$
        TaxCustRec(1).Addr2 = Add2$
        TaxCustRec(1).City = City$
        TaxCustRec(1).State = St$
        TaxCustRec(1).Zip = Zip$
        TaxCustRec(1).CSSN = BILLNUM$
        TaxCustRec(1).CountyAcctString = CTYACCT$
        TaxCustRec(1).CountyAcct = VAL(CTYACCT$)

        'stophere
        IF LEN(QPTrim$(TaxCustRec(1).OptSrchDesc)) = 0 THEN
          TaxCustRec(1).OptSrchDesc = BILLNUM$
        ELSE
          TaxCustRec(1).OptSrchDesc = QPTrim$(TaxCustRec(1).OptSrchDesc) + " " + BILLNUM$
        END IF

        Record = Snt&
        PUT Taxfile, Record, TaxCustRec(1)
        GOSUB UpdatePers
        GOTO 10
      END IF
    NEXT Snt&

END IF
  
  IF LEN(RTRIM$(PIN$)) > 0 THEN
    'Now Find The Account by Pin #
    FOR Cnt! = 1 TO NumOfPropRecs
      GET PropTaxFile, Cnt!, PropertyRec(1)
 
      IF RTRIM$(PropertyRec(1).RealPin) = RTRIM$(PIN$) THEN
        PropertyRec(1).PROPDATE = Date2Num%(DATE$)
        PropertyRec(1).PROPVALU = VAL(RV$)
        PropertyRec(1).EXMPSENI = VAL(RSC$)
        PUT PropTaxFile, Cnt!, PropertyRec(1)
       'Update Pers Value if Needed
        Record = PropertyRec(1).CustPin
        IF VAL(PV$) > 0 THEN GOSUB UpdatePers
       'Update Customer Record
        GET Taxfile, Record, TaxCustRec(1)
        'IF Record = 417 THEN STOP
        kk = INSTR(NM$, ",")
        IF kk > 0 THEN
          LNM$ = LEFT$(NM$, kk - 1)
          FM$ = MID$(NM$, kk + 1, LEN(NM$) - kk)
        ELSE
          LNM$ = NM$
          FM$ = ""
        END IF
        'TaxCustRec(1).FName = LTRIM$(FM$)
        'TaxCustRec(1).LName = LTRIM$(LNM$)
        'TaxCustRec(1).CustName = QPTrim$(FM$) + " " + QPTrim$(LNM$)
        TaxCustRec(1).CustName = QPTrim$(NM$)
        TaxCustRec(1).SName = LTRIM$(LNM$) + QPTrim$(FM$)

        'TaxCustRec(1).SName = LTRIM$(LNM$)
        TaxCustRec(1).Addr1 = Add1$
        TaxCustRec(1).Addr2 = Add2$
        TaxCustRec(1).City = City$
        TaxCustRec(1).State = St$
        TaxCustRec(1).Zip = Zip$
        'TaxCustRec(1).CSSN = BILLNUM$

        IF LEN(QPTrim$(TaxCustRec(1).OptSrchDesc)) = 0 THEN
          TaxCustRec(1).OptSrchDesc = BILLNUM$
        ELSE
          TaxCustRec(1).OptSrchDesc = QPTrim$(TaxCustRec(1).OptSrchDesc) + " " + BILLNUM$
        END IF

        TaxCustRec(1).CountyAcctString = CTYACCT$
        TaxCustRec(1).CountyAcct = VAL(CTYACCT$)
        PUT Taxfile, Record, TaxCustRec(1)
        GOTO 10
      END IF
    NEXT Cnt!
  END IF
  GOSUB UpdateMaster
  IF VAL(RV$) > 0 THEN GOSUB UpdateReal
  IF VAL(PV$) > 0 THEN GOSUB UpdatePers
  GOTO 10

UpdateMaster:
    NewOne = NewOne + 1
    Record = LOF(Taxfile) / LEN(TaxCustRec(1)) + 1
  '  IF Record = 417 THEN STOP
   'Split Name Here
    kk = INSTR(NM$, ",")
    IF kk > 0 THEN
      LNM$ = LEFT$(NM$, kk - 1)
      FM$ = MID$(NM$, kk + 1, LEN(NM$) - kk)
    ELSE
      LNM$ = NM$
      FM$ = ""
    END IF

    TaxCustRec(1).Acct = Record
    TaxCustRec(1).OPENDATE = Date2Num(DATE$)
    'TaxCustRec(1).FName = LTRIM$(FM$)
    'TaxCustRec(1).LName = LTRIM$(LNM$)
    'TaxCustRec(1).CustName = QPTrim$(FM$) + " " + QPTrim$(LNM$)
    TaxCustRec(1).CustName = QPTrim$(NM$)
    TaxCustRec(1).SName = LTRIM$(LNM$) + QPTrim$(FM$)
    'TaxCustRec(1).SName = LTRIM$(LNM$)
    TaxCustRec(1).Addr1 = Add1$
    TaxCustRec(1).Addr2 = Add2$
    TaxCustRec(1).City = City$
    TaxCustRec(1).State = St$
    TaxCustRec(1).Zip = Zip$
    TaxCustRec(1).HPHONE = ""
    TaxCustRec(1).WPHONE = ""
    
    TaxCustRec(1).CSSN = BILLNUM$
    TaxCustRec(1).oSSN = ""

    TaxCustRec(1).Active = "Y"
    TaxCustRec(1).Interest = "Y"
    TaxCustRec(1).TaxExempt = "N"
    TaxCustRec(1).Penalty = "N"
    TaxCustRec(1).CountyAcct = VAL(CTYACCT$)
    TaxCustRec(1).CountyAcctString = CTYACCT$
    IF LEN(QPTrim$(TaxCustRec(1).OptSrchDesc)) = 0 THEN
      TaxCustRec(1).OptSrchDesc = BILLNUM$
    ELSE
      TaxCustRec(1).OptSrchDesc = QPTrim$(TaxCustRec(1).OptSrchDesc) + " " + BILLNUM$
    END IF

    TaxCustRec(1).LastTrans = 0
    TaxCustRec(1).FirstPropRec = 0
    TaxCustRec(1).FirstPersRec = 0
    TaxCustRec(1).PIN = Record
    TaxCustRec(1).Deleted = 0
    TaxCustRec(1).FileVer = 8
    PUT Taxfile, Record, TaxCustRec(1)
    RETURN

UpdateReal:
    PropertyRec(1).RealPin = PIN$
    PropertyRec(1).PROPDATE = Date2Num%(DATE$)
    PropertyRec(1).GISPOS = ""
    PropertyRec(1).Map = LTRIM$(RTRIM$(MID$(A$, 294, 6)))
    PropertyRec(1).BLOCK = ""
    PropertyRec(1).LOTNUMB = LTRIM$(RTRIM$(MID$(A$, 300, 8)))
    PropertyRec(1).LOTACRE = ""
    PropertyRec(1).PropSize = VAL(LTRIM$(RTRIM$(MID$(A$, 369, 96)))) / 1000
    PropertyRec(1).PROPDISC = "N"
    PropertyRec(1).LateList = "N"
    PropertyRec(1).MORTCODE = ""
    PropertyRec(1).PROPVALU = VAL(RV$)
    PropertyRec(1).EXMPSENI = VAL(RSC$)
    PropertyRec(1).EXMPOTHR = 0
    PropertyRec(1).PROPNOT1 = DESC1$
    PropertyRec(1).PROPNOT2 = DESC2$
    PropertyRec(1).PROPNOT3 = DESC3$
    PropertyRec(1).CustPin = Record
    PropertyRec(1).NextRec = 0
    PropertyRec(1).LastYrPrinted = 0
    PropertyRec(1).Deleted = 0
    PropertyRec(1).Blank = ""
    NextRecord = (LOF(PropTaxFile) / LEN(PropertyRec(1))) + 1
    PUT PropTaxFile, NextRecord, PropertyRec(1)
    RETURN

UpdatePers:
   PersRec(1).PropPin = "1"
   PersRec(1).PROPDATE = Date2Num%(DATE$)
   PersRec(1).PersVal = VAL(PV$)
   PersRec(1).MHVALUE = 0
   PersRec(1).MCVALUE = 0
   PersRec(1).CVALUE = 0
   PersRec(1).MTVALUE = 0
   PersRec(1).EXMPSENI = VAL(Psc$)
   PersRec(1).EXMPOTHR = 0
   PersRec(1).DISCOV = "N"
   IF VAL(LateFlg$) > 0 THEN
     PersRec(1).LateList = "Y"
   ELSE
     PersRec(1).LateList = "N"
   END IF
   PersRec(1).DESC1 = ""
   PersRec(1).DESC2 = ""
   PersRec(1).DESC3 = ""
   PersRec(1).Desc4 = ""
   PersRec(1).Desc5 = ""
   PersRec(1).CustPin = Record
   PersRec(1).NextRec = 0
   PersRec(1).LastYrPrinted = 0
   PersRec(1).Deleted = 0
   PersRec(1).Blank = ""
   NextRec = (LOF(PersTaxFIle) / LEN(PersRec(1))) + 1
   PUT PersTaxFIle, NextRec, PersRec(1)
   RETURN

END SUB

SUB OpenTaxCustFile (NumOfTaxRecs, Taxfile)
  
  Taxfile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #Taxfile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(Taxfile) / LEN(TaxCustRec(1))
  
END SUB

SUB OpenTaxPersFile (NumOfPersRecs, PersTaxFIle)
  PersTaxFIle = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFIle LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFIle) / LEN(PersRec(1))
  
END SUB

SUB OpenTaxPropFile (NumOfPropRecs, PropTaxFile)
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
END SUB

