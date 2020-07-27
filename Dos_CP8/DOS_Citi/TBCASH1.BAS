DEFINT A-Z
DECLARE SUB SupTaxPayment (PDATE$, TBPath$, OpNumber%)
DECLARE SUB MiscCodes (PDATE$, TBPath$)
DECLARE SUB SetMessage ()
DECLARE SUB Cleardown ()
DECLARE SUB DecalPayment (PDATE$, TBPath$, OperNumber)
DECLARE SUB SupTaxPayment (PDATE$, TBPath$, OperNumber)
DECLARE SUB PrintReceiptReport (PDATE$, TBPath$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PrintLoading ()

DECLARE FUNCTION num2date$ (Num%)
DECLARE FUNCTION Date2Num% (Dat$)

'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'Qscr.BI'
'$INCLUDE: 'SetCnf.BI'
'$INCLUDE: 'lcTAX.BI'
     

     TBPath$ = "F:\"                    ' change to "F:\" Before compiling

     CLOSE : OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS Year$: GET 1, 1

     CurrentYear$ = Year$: CLOSE
     CurrentYear = VAL(CurrentYear$)

     OPEN "I", 10, "C:\OP.TXT"
     IF EOF(10) THEN
      OpNumber = 0
      ELSE
      INPUT #10, A$
      OpNumber = VAL(A$)
     END IF
     CLOSE 10
     

     DIM SHARED yrdesc(11)
     REDIM TBLog(1) AS TBLog
     DIM SHARED TBRECust(1) AS TBRECust
     DIM SHARED TBRECustBal(1) AS TBREBAL
     DIM SHARED TBPPCust(1) AS TBPPCust
     DIM SHARED TBPPCustBal(1) AS TBPPBal
     DIM SHARED TBTrans(1) AS TBTrans
     DIM SHARED TBRECPT(1) AS DayReceiptsType

     PDATE$ = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 2)
     SupTaxPayment PDATE$, TBPath$, OpNumber

SUB PrintReceiptReport (PDATE$, TBPath$)

    DIM Type1Year(30), Type1Tax#(30), Type1Pen#(30), Type1Int#(30)
    DIM Type2Year(30), Type2Tax#(30), Type2Pen#(30), Type2Int#(30)
    DIM Type3Year(30), Type3Tax#(30), Type3Pen#(30), Type3Int#(30)
    DIM Type4Year(30), Type4Tax#(30), Type4Pen#(30), Type4Int#(30)
    DIM TotalMisc#(200)
    DIM Code(200)


    TotalCash# = 0
    TotalChk# = 0
    TotalOwed# = 0



MainBody:
   Penalty# = 0
   Interest# = 0
   CLS
   scren$ = "8.05.00": CLS : GOSUB Heading
   COLOR 11
   LOCATE 1, 30: PRINT "Cash Receipts Report"
   LOCATE 3, 1
   COLOR 11
   PRINT "Enter the Drawer Number to Print (0=ALL) .. "; : fl = -1
   GOSUB 41000
   Drawer = VAL(in$)
   LOCATE 5, 1
   COLOR 11
   PRINT "Enter the Posting Date of the Receipts to Print (mmddyy) (ESC-Cancel) .. "; : fl = -6
   GOSUB 41000
   IF cf = 1 THEN EXIT SUB
   IF LEN(in$) <> 6 THEN GOTO MainBody
    Month$ = LEFT$(in$, 2)
    Day$ = MID$(in$, 3, 2)
    Year$ = RIGHT$(in$, 2)
   IF VAL(Month$) < 1 OR VAL(Month$) > 12 THEN
    COLOR 15
    PRINT CHR$(7);
    PRINT "Invalid Month!!!"
    SLEEP 3
    GOTO MainBody
   END IF
   IF VAL(Day$) < 1 OR VAL(Day$) > 31 THEN
     COLOR 15
     PRINT CHR$(7);
     PRINT "Invalid Day!!!"
     SLEEP 3
     GOTO MainBody
   END IF
   IF VAL(Year$) > 90 THEN
    Year = VAL(Year$) + 1900
    ELSE
    Year = VAL(Year$) + 2000
   END IF
   Year$ = LTRIM$(STR$(Year))
   ReportDate$ = LEFT$(in$, 2) + "/" + MID$(in$, 3, 2) + "/" + Year$
   ReportDate = Date2Num(ReportDate$)

  'Open the Report File
   ReportFile$ = "TBDAILY.PRN"
   RptHandle% = 10
   OPEN ReportFile$ FOR OUTPUT AS #RptHandle%
   TBRFile = FREEFILE
   OPEN TBPath$ + "TBRECFIL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBRFile LEN = LEN(TBRECPT(1))
   IF LOF(TBRFile) = 0 THEN
    NoReport = 1
    CLOSE TBRFile
    GOTO FinishReport
   END IF
   GOSUB ReportHeading

LinePrint:
   FOR Cnt& = LOF(TBRFile) / LEN(TBRECPT(1)) TO 1 STEP -1
   GET TBRFile, Cnt&, TBRECPT(1)
   
   IF TBRECPT(1).Date < ReportDate THEN
    BadCount = BadCount + 1
    IF BadCount > 5000 THEN GOTO FinishReport
   END IF

  IF TBRECPT(1).Date = ReportDate THEN
   IF Drawer = 0 OR (Drawer > 0 AND TBRECPT(1).Drawer = Drawer) THEN

    'Check for Page Break
     IF lc >= 56 THEN
      PRINT #10, CHR$(12);
      GOSUB ReportHeading
     END IF


    PRINT #10, USING "########"; Cnt&;
    PRINT #10, TAB(10); "";
    IF TBRECPT(1).TranType = 1 THEN PRINT #10, "Reg Real Estate";
    IF TBRECPT(1).TranType = 2 THEN PRINT #10, "Reg Pers Property";
    IF TBRECPT(1).TranType = 3 THEN PRINT #10, "Sup Real Estate";
    IF TBRECPT(1).TranType = 4 THEN PRINT #10, "Sup Pers Property";
    IF TBRECPT(1).TranType = 5 THEN PRINT #10, "Decal";
    IF TBRECPT(1).TranType = 6 THEN PRINT #10, "Misc Payment";

    IF TBRECPT(1).TranType < 1 OR TBRECPT(1).TranType > 6 THEN
     PRINT #10, "Undefined";
    END IF

    IF TBRECPT(1).AmtIn < (TBRECPT(1).CashAmt + TBRECPT(1).ChkAmt) THEN
     TBRECPT(1).AmtIn = (TBRECPT(1).CashAmt + TBRECPT(1).ChkAmt)
     PUT TBRFile, Cnt&, TBRECPT(1)
    END IF

    PRINT #10, TAB(28); USING "######.##"; TBRECPT(1).CashAmt;
    PRINT #10, TAB(38); USING "######.##"; TBRECPT(1).ChkAmt;
    PRINT #10, TAB(48); USING "######.##"; TBRECPT(1).AmtIn;
    PRINT #10, TAB(58); USING "######.##"; TBRECPT(1).AmtOwed;
    PRINT #10, TAB(68); USING "######.##"; TBRECPT(1).AmtIn - TBRECPT(1).AmtOwed

    TotalCash# = TotalCash# + TBRECPT(1).CashAmt
    TotalChk# = TotalChk# + TBRECPT(1).ChkAmt
    TotalIn# = TotalIn# + TBRECPT(1).AmtIn
    TotalOwed# = TotalOwed# + TBRECPT(1).AmtOwed
    lc = lc + 1

    IF TBRECPT(1).TranType = 1 THEN
     PRINT #10, LEFT$(RTRIM$(TBRECPT(1).Desc), 22);
     PRINT #10, "| Tax Pd:"; USING "######.##"; TBRECPT(1).RealTxPd;
     PRINT #10, " Pen Pd:"; USING "####.##"; TBRECPT(1).PenTxPd;
     PRINT #10, " Int Pd:"; USING "####.##"; TBRECPT(1).IntTxPd;
     PRINT #10, " Yr:"; TBRECPT(1).TaxYear
     PRINT #10, STRING$(79, "-")
     lc = lc + 2
    END IF

    IF TBRECPT(1).TranType = 2 THEN
     PRINT #10, LEFT$(RTRIM$(TBRECPT(1).Desc), 22);
     PRINT #10, "| Tax Pd:"; USING "######.##"; TBRECPT(1).PersTxPd;
     PRINT #10, " Pen Pd:"; USING "####.##"; TBRECPT(1).PenTxPd;
     PRINT #10, " Int Pd:"; USING "####.##"; TBRECPT(1).IntTxPd;
     PRINT #10, " Yr:"; TBRECPT(1).TaxYear
     PRINT #10, STRING$(79, "-")
     lc = lc + 2
    END IF

    IF TBRECPT(1).TranType = 3 THEN
     PRINT #10, LEFT$(RTRIM$(TBRECPT(1).Desc), 22);
     PRINT #10, "| Tax Pd:"; USING "######.##"; TBRECPT(1).RealTxPd;
     PRINT #10, " Pen Pd:"; USING "####.##"; TBRECPT(1).PenTxPd;
     PRINT #10, " Int Pd:"; USING "####.##"; TBRECPT(1).IntTxPd;
     PRINT #10, " Yr:"; TBRECPT(1).TaxYear
     PRINT #10, STRING$(79, "-")
     lc = lc + 2
    END IF

    IF TBRECPT(1).TranType = 4 THEN
     PRINT #10, LEFT$(RTRIM$(TBRECPT(1).Desc), 22);
     PRINT #10, "| Tax Pd:"; USING "######.##"; TBRECPT(1).PersTxPd;
     PRINT #10, " Pen Pd:"; USING "####.##"; TBRECPT(1).PenTxPd;
     PRINT #10, " Int Pd:"; USING "####.##"; TBRECPT(1).IntTxPd;
     PRINT #10, " Yr:"; TBRECPT(1).TaxYear
     PRINT #10, STRING$(79, "-")
     lc = lc + 2
    END IF

    IF TBRECPT(1).TranType = 5 THEN
     PRINT #10, RTRIM$(TBRECPT(1).Desc); " || Decal # "; TBRECPT(1).DecalNumb
     PRINT #10, STRING$(79, "-")
     lc = lc + 2
    END IF

    IF TBRECPT(1).TranType = 6 THEN
     PRINT #10, "Name: "; RTRIM$(TBRECPT(1).Desc);
     PRINT #10, " || Misc Code # "; TBRECPT(1).TranCode
     PRINT #10, STRING$(79, "-")
     lc = lc + 2
    END IF
    
  ' Now Add Up Each Type of Payment Per Transaction
    FOUND = 0   'Flag to Indicate Match Set to Null to Begin
    IF TBRECPT(1).TranType = 1 THEN
      IF Type1Cnt& = 0 THEN
       Type1Cnt& = 1
       Type1Year(Type1Cnt&) = TBRECPT(1).TaxYear
       Type1Tax#(Type1Cnt&) = TBRECPT(1).RealTxPd
       Type1Pen#(Type1Cnt&) = TBRECPT(1).PenTxPd
       Type1Int#(Type1Cnt&) = TBRECPT(1).IntTxPd
      ELSE
      FOR C! = 1 TO Type1Cnt&
       IF Type1Year(C!) = TBRECPT(1).TaxYear THEN
        Type1Tax#(C!) = Type1Tax#(C!) + TBRECPT(1).RealTxPd
        Type1Pen#(C!) = Type1Pen#(C!) + TBRECPT(1).PenTxPd
        Type1Int#(C!) = Type1Int#(C!) + TBRECPT(1).IntTxPd
        FOUND = 1
        EXIT FOR
        END IF
        FOUND = 0
       NEXT C!
       IF FOUND = 0 THEN
        Type1Cnt& = Type1Cnt& + 1
        Type1Year(Type1Cnt&) = TBRECPT(1).TaxYear
        Type1Tax#(Type1Cnt&) = TBRECPT(1).RealTxPd
        Type1Pen#(Type1Cnt&) = TBRECPT(1).PenTxPd
        Type1Int#(Type1Cnt&) = TBRECPT(1).IntTxPd

       END IF
     END IF
    END IF

    IF TBRECPT(1).TranType = 2 THEN
     IF Type2Cnt& = 0 THEN
        Type2Cnt& = 1
        Type2Year(Type2Cnt&) = TBRECPT(1).TaxYear
        Type2Tax#(Type2Cnt&) = TBRECPT(1).PersTxPd
        Type2Pen#(Type2Cnt&) = TBRECPT(1).PenTxPd
        Type2Int#(Type2Cnt&) = TBRECPT(1).IntTxPd
      ELSE
      FOR C! = 1 TO Type2Cnt&
       IF Type2Year(C!) = TBRECPT(1).TaxYear THEN
        Type2Tax#(C!) = Type2Tax#(C!) + TBRECPT(1).PersTxPd
        Type2Pen#(C!) = Type2Pen#(C!) + TBRECPT(1).PenTxPd
        Type2Int#(C!) = Type2Int#(C!) + TBRECPT(1).IntTxPd
        FOUND = 1
        EXIT FOR
        END IF
        FOUND = 0
      NEXT C!
       IF FOUND = 0 THEN
        Type2Cnt& = Type2Cnt& + 1
        Type2Year(Type2Cnt&) = TBRECPT(1).TaxYear
        Type2Tax#(Type2Cnt&) = TBRECPT(1).PersTxPd
        Type2Pen#(Type2Cnt&) = TBRECPT(1).PenTxPd
        Type2Int#(Type2Cnt&) = TBRECPT(1).IntTxPd
       END IF
     END IF
    END IF

    IF TBRECPT(1).TranType = 3 THEN
     IF Type3Cnt& = 0 THEN
       Type3Cnt& = 1
       Type3Year(Type3Cnt&) = TBRECPT(1).TaxYear
       Type3Tax#(Type3Cnt&) = TBRECPT(1).RealTxPd
       Type3Pen#(Type3Cnt&) = TBRECPT(1).PenTxPd
       Type3Int#(Type3Cnt&) = TBRECPT(1).IntTxPd
      ELSE
      FOR C! = 1 TO Type3Cnt&
       IF Type3Year(C!) = TBRECPT(1).TaxYear THEN
        Type3Tax#(C!) = Type3Tax#(C!) + TBRECPT(1).RealTxPd
        Type3Pen#(C!) = Type3Pen#(C!) + TBRECPT(1).PenTxPd
        Type3Int#(C!) = Type3Int#(C!) + TBRECPT(1).IntTxPd
        FOUND = 1
        EXIT FOR
        END IF
        FOUND = 0
      NEXT C!
       IF FOUND = 0 THEN
        Type3Cnt& = Type3Cnt& + 1
        Type3Year(Type3Cnt&) = TBRECPT(1).TaxYear
        Type3Tax#(Type3Cnt&) = TBRECPT(1).RealTxPd
        Type3Pen#(Type3Cnt&) = TBRECPT(1).PenTxPd
        Type3Int#(Type3Cnt&) = TBRECPT(1).IntTxPd
       END IF
     END IF
    END IF
    IF TBRECPT(1).TranType = 4 THEN
      IF Type4Cnt& = 0 THEN
       Type4Cnt& = 1
       Type4Year(Type4Cnt&) = TBRECPT(1).TaxYear
       Type4Tax#(Type4Cnt&) = TBRECPT(1).PersTxPd
       Type4Pen#(Type4Cnt&) = TBRECPT(1).PenTxPd
       Type4Int#(Type4Cnt&) = TBRECPT(1).IntTxPd
      ELSE
      FOR C! = 1 TO Type4Cnt&
       IF Type4Year(C!) = TBRECPT(1).TaxYear THEN
        Type4Tax#(C!) = Type4Tax#(C!) + TBRECPT(1).PersTxPd
        Type4Pen#(C!) = Type4Pen#(C!) + TBRECPT(1).PenTxPd
        Type4Int#(C!) = Type4Int#(C!) + TBRECPT(1).IntTxPd
        FOUND = 1
        EXIT FOR
        END IF
        FOUND = 0
      NEXT C!
       IF FOUND = 0 THEN
        Type4Cnt& = Type4Cnt& + 1
        Type4Year(Type4Cnt&) = TBRECPT(1).TaxYear
        Type4Tax#(Type4Cnt&) = TBRECPT(1).PersTxPd
        Type4Pen#(Type4Cnt&) = TBRECPT(1).PenTxPd
        Type4Int#(Type4Cnt&) = TBRECPT(1).IntTxPd
       END IF
     END IF
    END IF

    IF TBRECPT(1).TranType = 5 THEN
     NumberOfDecals! = NumberOfDecals! + 1
     TotalDecals# = TotalDecals# + TBRECPT(1).ChkAmt + TBRECPT(1).CashAmt
    END IF

    IF TBRECPT(1).TranType = 6 THEN
      FOUND = 0
     IF CodeCnt = 0 THEN
      CodeCnt = 1
      TotalMisc#(CodeCnt) = TBRECPT(1).AmtOwed
      Code(CodeCnt) = TBRECPT(1).TranCode
      ELSE
      FOR CCnt = 1 TO CodeCnt
      IF Code(CCnt) = TBRECPT(1).TranCode THEN
       TotalMisc#(CCnt) = TotalMisc#(CCnt) + TBRECPT(1).AmtOwed
       FOUND = 1
       EXIT FOR
      END IF
      NEXT CCnt
      IF FOUND = 0 THEN
       CodeCnt = CodeCnt + 1
       TotalMisc#(CodeCnt) = TBRECPT(1).AmtOwed
       Code(CodeCnt) = TBRECPT(1).TranCode
      END IF
     END IF
    END IF
    END IF
    END IF
  NEXT Cnt&
  
   
FinishReport:
    PRINT #10, "Grand Totals: ";
    PRINT #10, TAB(28); USING "######.##"; TotalCash#;
    PRINT #10, TAB(38); USING "######.##"; TotalChk#;
    PRINT #10, TAB(48); USING "######.##"; TotalIn#;
    PRINT #10, TAB(58); USING "######.##"; TotalOwed#;
    PRINT #10, TAB(68); USING "######.##"; (TotalIn#) - TotalOwed#
    PRINT #10, CHR$(12);
   IF NoReport = 0 THEN
    PRINT #10, "     Lunenburg County : Daily Cash Receipts Report : Date - "; ReportDate$
   IF Drawer = 0 THEN
    PRINT #10, "All Drawers"
     ELSE
    PRINT #10, "Drawer Number "; Drawer
   END IF
    PRINT #10, ""
    PRINT #10, "Totals Summary ..."
    IF Type1Cnt& > 0 THEN
       PRINT #10, ""
       PRINT #10, "Regular Real Estate Taxes"
       PRINT #10, "Tax Year"; TAB(10); "Real Tax Paid"; TAB(25); "Penalty Paid"; TAB(40); "Interest Paid"; TAB(55); " Total Amount"
       PRINT #10, STRING$(79, "-")
       FOR ll = 1 TO Type1Cnt&
        PRINT #10, Type1Year(ll);
         PRINT #10, TAB(10); USING "$$#####,#.##"; Type1Tax#(ll);
         PRINT #10, TAB(25); USING "$$#####,#.##"; Type1Pen#(ll);
         PRINT #10, TAB(40); USING "$$#####,#.##"; Type1Int#(ll);
         PRINT #10, TAB(55); USING "$$#####,#.##"; (Type1Tax#(ll) + Type1Pen#(ll) + Type1Int#(ll))
       NEXT ll
     END IF
    IF Type2Cnt& > 0 THEN
       PRINT #10, ""
       PRINT #10, "Regular Personal Taxes"
       PRINT #10, "Tax Year"; TAB(10); "Pers Tax Paid"; TAB(25); "Penalty Paid"; TAB(40); "Interest Paid"; TAB(55); " Total Amount"
       PRINT #10, STRING$(79, "-")
       FOR ll = 1 TO Type2Cnt&
        PRINT #10, Type2Year(ll);
         PRINT #10, TAB(10); USING "$$#####,#.##"; Type2Tax#(ll);
         PRINT #10, TAB(25); USING "$$#####,#.##"; Type2Pen#(ll);
         PRINT #10, TAB(40); USING "$$#####,#.##"; Type2Int#(ll);
         PRINT #10, TAB(55); USING "$$#####,#.##"; (Type2Tax#(ll) + Type2Pen#(ll) + Type2Int#(ll))
       NEXT ll
     END IF
    IF Type3Cnt& > 0 THEN
       PRINT #10, "Supplemental Real Estate Taxes"
       PRINT #10, "Tax Year"; TAB(10); "Sup Real Tax Pd"; TAB(27); "Penalty Paid"; TAB(42); "Interest Paid"; TAB(57); " Total Amount"
       PRINT #10, STRING$(79, "-")
       FOR ll = 1 TO Type3Cnt&
        PRINT #10, Type3Year(ll);
         PRINT #10, TAB(10); USING "$$#####,#.##"; Type3Tax#(ll);
         PRINT #10, TAB(27); USING "$$#####,#.##"; Type3Pen#(ll);
         PRINT #10, TAB(42); USING "$$#####,#.##"; Type3Int#(ll);
         PRINT #10, TAB(57); USING "$$#####,#.##"; (Type3Tax#(ll) + Type3Pen#(ll) + Type3Int#(ll))
       NEXT ll
     END IF
    IF Type4Cnt& > 0 THEN
       PRINT #10, ""
       PRINT #10, "Supplemental Personal Taxes"
       PRINT #10, "Tax Year"; TAB(10); "Sup Pers Tax Pd"; TAB(27); "Penalty Paid"; TAB(42); "Interest Paid"; TAB(57); " Total Amount"
       PRINT #10, STRING$(79, "-")
       FOR ll = 1 TO Type4Cnt&
        PRINT #10, Type4Year(ll);
         PRINT #10, TAB(10); USING "$$#####,#.##"; Type4Tax#(ll);
         PRINT #10, TAB(27); USING "$$#####,#.##"; Type4Pen#(ll);
         PRINT #10, TAB(42); USING "$$#####,#.##"; Type4Int#(ll);
         PRINT #10, TAB(57); USING "$$#####,#.##"; (Type4Tax#(ll) + Type4Pen#(ll) + Type4Int#(ll))
       NEXT ll
     END IF

    IF NumberOfDecals! > 0 THEN
     PRINT #10, ""
     PRINT #10, "Decals"
     PRINT #10, "Total Decals Issued: "; USING "#####,#"; NumberOfDecals!
     PRINT #10, "Total Amount of Decals: "; USING "######,#.##"; TotalDecals#
    END IF

    IF CodeCnt > 0 THEN
      CLOSE 15
      OPEN TBPath$ + "tbmisc.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #15 LEN = 256
      FIELD 15, 30 AS CodeDesc$, 1 AS Active$, 225 AS nu$
      PRINT #10, CHR$(12);
      PRINT #10, "     Lunenburg County : Daily Cash Receipts Report : Date - "; ReportDate$
      PRINT #10, "Totals Summary ..."
      PRINT #10, "Misc Codes"
      FOR ll = 1 TO CodeCnt
       IF Code(ll) < 1 OR Code(ll) > LOF(15) / 256 THEN
        CodeName$ = "Undefined"
        ELSE
        GET 15, Code(ll)
        CodeName$ = CodeDesc$
      END IF
      PRINT #10, "Code # "; Code(ll); CodeName$; TAB(50); USING "$$######,#.##"; TotalMisc#(ll)
      NEXT ll
      CLOSE 15
    END IF
   END IF
 
  'Laser Printer Ejection for Betsy
   PRINT #10, CHR$(12);
   CLOSE
  'Print the Report Here
   LPTPort% = 1
   EntryPoint% = 1
   Header$ = "Daily Receipts Report"
   PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
   KILL ReportFile$
   EXIT SUB

ReportHeading:
   PRINT #10, "     Lunenburg County : Daily Cash Receipts Report : Date - "; ReportDate$
   PRINT #10, ""
   PRINT #10, "Report Date: "; DATE$
   IF Drawer = 0 THEN
    PRINT #10, "All Drawers"
     ELSE
    PRINT #10, "Drawer Number "; Drawer
   END IF
   PRINT #10, "Receipt#"; TAB(10); "Payment Type"; TAB(28); "Cash Paid"; TAB(38); " Chk Paid"; TAB(48); " Total Pd"; TAB(58); " Amt Owed"; TAB(68); "Change Back"
   PRINT #10, STRING$(79, "=")
   lc = 6
   RETURN
    
41000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
41010 IF INKEY$ <> "" THEN 41010
41020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
41030 z$ = INKEY$: IF z$ = "" THEN 41030
41040 IF z$ <> CHR$(13) THEN 41059 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
41050 GOSUB 41200: RETURN
41059 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(z$) < 2 THEN 41065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN cf = 12: RETURN
41065 IF ASC(z$) = 0 THEN 41030
41070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
41080 IF z$ = " " THEN 41130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 41000
41090 IF z$ <> CHR$(8) THEN 41130 ELSE IF ZL = 0 THEN 41030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 41110 ELSE IF z$ = "," THEN 41120
41100 IF z$ = "." THEN ZD = 0: GOTO 41110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
41110 in$ = LEFT$(in$, LEN(in$) - 1)
41120 ZL = ZL - 1: GOTO 41030
41130 IF ABS(fl) = ZL THEN 41030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 41180
41140 IF z$ = "." AND INSTR(in$, ".") = 0 THEN GOTO 41180
41150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 41190
41160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 41180
41170 IF z$ < "0" OR z$ > "9" THEN 41030
41180 PRINT z$; : in$ = in$ + z$: ZL = ZL + 1
41190 IF ABS(fl) = 1 THEN 41050 ELSE 41030
41200 LOCATE RO, PO: PRINT in$: RETURN
41210 GOSUB 41000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
41220 RETURN
Heading:
     COLOR 10: LOCATE 1, 1
     PRINT "Date: "; : COLOR 15
     PRINT LEFT$(DATE$, 2) + "-" + MID$(DATE$, 4, 2) + "-" + RIGHT$(DATE$, 2);
     COLOR 10
     PRINT TAB(70); scren$
     PRINT STRING$(79, "_")
     RETURN

END SUB

SUB SupTaxPayment (PDATE$, TBPath$, OpNumber)
     PostDate$ = LEFT$(PDATE$, 2) + "-" + MID$(PDATE$, 3, 2) + "-" + RIGHT$(PDATE$, 2)
     CLOSE : OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS Year$: GET 1, 1
     CurrentYear$ = Year$: CLOSE
     CurrentYear = VAL(CurrentYear$)


1026 Cleardown
     Interest# = 0
     Penalty# = 0
     CLOSE
     LOCATE 3, 1: COLOR 11
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; : fl = 1: GOSUB 53210
     IF cf = 1 THEN EXIT SUB
     IF NOT (in$ = "R" OR in$ = "P") THEN 1026
     taxtype$ = UCASE$(in$)



1031 LOCATE 4, 1: COLOR 11
     PRINT "Enter Customer Account Number (TAB-Last Name : ESC-Menu) .. "; : fl = -6
     GOSUB 53000
     IF cf = 1 THEN
      CLOSE : EXIT SUB
     END IF
     IF cf = 2 THEN
      GOSUB NamesearchSup
      IF ggo = 1 THEN 1076
      IF ggo = 0 THEN LOCATE 10, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 53000: GOTO 1026
     END IF
     Record! = VAL(in$)
1076 CLOSE : GOSUB 50000: GOSUB 50025
     IF taxtype$ = "R" THEN
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBRECust(1)) THEN 1031
       LOCK #TBFile, Record!
       GET TBFile, Record!, TBRECust(1)
       GET TBBalFile, Record!, TBRECustBal(1)
       IF TBRECust(1).Deleted = "Y" THEN GOSUB DeletedSup: GOTO 1026
       ELSE
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBPPCust(1)) THEN 1026
       LOCK #TBFile, Record!
       GET TBFile, Record!, TBPPCust(1)
       GET TBBalFile, Record!, TBPPCustBal(1)
       IF TBPPCust(1).Deleted = "Y" THEN GOSUB DeletedSup: GOTO 1026
     END IF

     IF taxtype$ = "R" THEN
       GOSUB ProcessRealSup
        ELSE
       GOSUB ProcessPersonalSup
     END IF
     CLOSE
     EXIT SUB

ProcessRealSup:
    GOSUB FigureBalanceSup
    PayType = 1 'Regular Payment
    PaySource$ = "R"
    CLS
    scren$ = "4.01.00": CLS : GOSUB 53500
    COLOR 11: LOCATE 1, 32: PRINT "Customer Payment Entry - Sup";
    LOCATE 3, 1: PRINT "Real Estate Customer"
    LOCATE 5, 1: COLOR 15
        NME$ = TBRECust(1).CustName
        PRINT TBRECust(1).CustName; TAB(50); "Current Bal: "; USING "$$######,#.##"; CurBal#
        PRINT TBRECust(1).Address1; TAB(50); "   Past Bal: "; USING "$$######,#.##"; PastDue#
        PRINT TBRECust(1).Address2; TAB(50); " Total Owed: "; USING "$$######,#.##"; CurBal# + PastDue#
        PRINT TBRECust(1).Address3
        Owed# = CurBal# + PastDue#
    LOCATE 9, 1: COLOR 11: PRINT STRING$(80, "-")
    LOCATE 10, 1: COLOR 11
    PRINT "Correct Account (Y/N) .. "; : fl = 1: GOSUB 53210
    IF cf = 1 THEN CLOSE : RETURN
    IF in$ = "N" THEN CLOSE : RETURN

    LOCATE 10, 1: PRINT STRING$(79, 32): COLOR 11
    LOCATE 10, 1
    PRINT "Enter Amount to Apply on Account ("; USING "######,#.##"; CurBal# + PastDue#; : PRINT ") .... ";
     PRINT TAB(54); ""; : fl = -9: GOSUB 53000
    IF cf = 1 THEN CLOSE : RETURN
    IF LEN(in$) = 0 THEN
      AmountPaid# = CurBal# + PastDue#
      ELSE
      AmountPaid# = VAL(in$)
    END IF
    LOCATE 10, 54: COLOR 14: PRINT USING "######,#.##"; AmountPaid#

spr1: COLOR 11: LOCATE 11, 1
     PRINT "Enter (1)Cash or (2)Check ... "; : fl = -1: GOSUB 53000
     IF VAL(in$) < 1 OR VAL(in$) > 2 THEN GOTO spr1
     curency = VAL(in$)
spr2: COLOR 11: LOCATE 12, 1
     PRINT "Enter Amount from Customer .. "; : fl = -9: GOSUB 53000
     LOCATE 12, 30: COLOR 14
     PRINT USING "#####,#.##"; VAL(in$)
     AmtIn# = VAL(in$)
     COLOR 11: LOCATE 12, 50: PRINT "Change Due: "; USING "######,#.##"; AmtIn# - AmountPaid#


spr3: LOCATE 13, 1: COLOR 11
     PRINT "Enter Description (if any) .. "; : fl = 15
     GOSUB 53000
     IF cf = 1 THEN CLOSE : RETURN
     Desc$ = UCASE$(in$)
spr4: LOCATE 14, 1: COLOR 11
     PRINT "(A)uto or (M)anually Dist. Payment . "; : fl = 1
     GOSUB 53210
     IF cf = 1 THEN CLOSE : RETURN
     IF NOT (in$ = "A" OR in$ = "M") THEN GOTO spr4
     Dist$ = in$
spr5: LOCATE 15, 1: COLOR 11
     PRINT "Enter Tax Year of Bill Being Paid .. "; : fl = -4
     GOSUB 53000
     IF LEN(in$) <> 4 THEN GOTO spr5
     IF VAL(in$) > VAL(CurrentYear$) THEN
      LOCATE 16, 1: COLOR 11
      PRINT CHR$(7);
      PRINT "Year Entered Exceeds Current Year : Press <ENTER> to Try Again!"; : fl = 1
      GOSUB 53000
      LOCATE 16, 1: PRINT STRING$(79, 32)
      GOTO spr5
     END IF
     ManYear$ = in$
     TaxYear$ = in$

spr6: LOCATE 16, 1: COLOR 11
     PRINT "Payment Entered Correctly? (Y/N) .. "; : fl = 1
     GOSUB 53210
     IF cf = 1 THEN CLOSE : RETURN
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO spr6
     IF in$ = "N" THEN GOTO ProcessRealSup



      GOSUB FilePaymentSup
      CLOSE : RETURN



ProcessPersonalSup:
    GOSUB FigureBalanceSup
    PayType = 1 'Regular Payment
    PaySource$ = "P"
    CLS
    scren$ = "4.01.00": CLS : GOSUB 53500
    COLOR 11: LOCATE 1, 32: PRINT "Customer Payment Entry - Sup";
    LOCATE 3, 1: PRINT "Personal Property Customer"
    LOCATE 5, 1: COLOR 15
        NME$ = TBPPCust(1).CustName
        PRINT TBPPCust(1).CustName; TAB(50); "Current Bal: "; USING "$$######,#.##"; CurBal#
        PRINT TBPPCust(1).Address1; TAB(50); "   Past Bal: "; USING "$$######,#.##"; PastDue#
        PRINT TBPPCust(1).Address2; TAB(50); " Total Owed: "; USING "$$######,#.##"; CurBal# + PastDue#
        PRINT RTRIM$(TBPPCust(1).City); ", "; TBPPCust(1).State; " "; TBPPCust(1).Zip
        Owed# = CurBal# + PastDue#
    LOCATE 9, 1: COLOR 11: PRINT STRING$(80, "-")
    LOCATE 10, 1: COLOR 11
    PRINT "Correct Account (Y/N) .. "; : fl = 1: GOSUB 53210
    IF cf = 1 THEN CLOSE : RETURN
    IF in$ = "N" THEN CLOSE : RETURN


    LOCATE 10, 1: PRINT STRING$(79, 32): COLOR 11
    LOCATE 10, 1
    PRINT "Enter Amount to Apply on Account ("; USING "######,#.##"; CurBal# + PastDue#; : PRINT ") .... ";
     PRINT TAB(54); ""; : fl = -9: GOSUB 53000
    IF cf = 1 THEN CLOSE : RETURN
    IF LEN(in$) = 0 THEN
      AmountPaid# = CurBal# + PastDue#
      ELSE
      AmountPaid# = VAL(in$)
    END IF
    LOCATE 10, 54: COLOR 14: PRINT USING "######,#.##"; AmountPaid#

sppr1: COLOR 11: LOCATE 11, 1
     PRINT "Enter (1)Cash or (2)Check ... "; : fl = -1: GOSUB 53000
     IF VAL(in$) < 1 OR VAL(in$) > 2 THEN GOTO sppr1
     curency = VAL(in$)
sppr2: COLOR 11: LOCATE 12, 1
     PRINT "Enter Amount from Customer .. "; : fl = -9: GOSUB 53000
     LOCATE 12, 30: COLOR 14
     PRINT USING "#####,#.##"; VAL(in$)
     AmtIn# = VAL(in$)
     COLOR 11: LOCATE 12, 50: PRINT "Change Due: "; USING "######,#.##"; AmtIn# - AmountPaid#


sppr3: LOCATE 13, 1: COLOR 11
     PRINT "Enter Description (if any) .. "; : fl = 15
     GOSUB 53000
     IF cf = 1 THEN CLOSE : RETURN
     Desc$ = UCASE$(in$)
sppr4: LOCATE 14, 1: COLOR 11
     PRINT "(A)uto or (M)anually Dist. Payment . "; : fl = 1
     GOSUB 53210
     IF cf = 1 THEN CLOSE : RETURN
     IF NOT (in$ = "A" OR in$ = "M") THEN GOTO sppr4
     Dist$ = in$
sppr5: LOCATE 15, 1: COLOR 11
     PRINT "Enter Tax Year of Bill Being Paid .. "; : fl = -4
     GOSUB 53000
     IF LEN(in$) <> 4 THEN GOTO spr5
     IF VAL(in$) > VAL(CurrentYear$) THEN
      LOCATE 16, 1: COLOR 11
      PRINT CHR$(7);
      PRINT "Year Entered Exceeds Current Year : Press <ENTER> to Try Again!"; : fl = 1
      GOSUB 53000
      LOCATE 16, 1: PRINT STRING$(79, 32)
      GOTO sppr5
     END IF
     ManYear$ = in$
     TaxYear$ = in$

sppr6: LOCATE 16, 1: COLOR 11
     PRINT "Payment Entered Correctly? (Y/N) .. "; : fl = 1
     GOSUB 53210
     IF cf = 1 THEN CLOSE : RETURN
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO sppr6
     IF in$ = "N" THEN GOTO ProcessPersonalSup



      GOSUB FilePaymentSup
      CLOSE : RETURN

FilePaymentSup:
 REDIM PayEdit(1) AS PaymentEdit
 PayEdit(1).PaymentAcct = Record!
 PayEdit(1).PaymentType = PayType
 PayEdit(1).PaymentDate = PDATE$
 PayEdit(1).PaymentAmt = AmountPaid#
 PayEdit(1).PaymentDesc = Desc$
 PayEdit(1).PaymentSource = PaySource$  'R or P
 PayEdit(1).PaymentDist = Dist$         'A=Auto  M=Manual
 PayEdit(1).PaymentManYear = ManYear$



     CLOSE TBFile: CLOSE TBBalFile
     GOSUB 50000: GOSUB 50025


     IF PayEdit(1).PaymentSource = "R" THEN
      IF PayEdit(1).PaymentAcct > 0 THEN
        GET TBFile, PayEdit(1).PaymentAcct, TBRECust(1)
        GET TBBalFile, PayEdit(1).PaymentAcct, TBRECustBal(1)
        GOSUB FigureBalanceSup
        GOSUB PostPaymentSup
        PUT TBFile, PayEdit(1).PaymentAcct, TBRECust(1)
        PUT TBBalFile, PayEdit(1).PaymentAcct, TBRECustBal(1)
        GOSUB FigureBalanceSup
        GOSUB PostRealHistorySup
       END IF
     END IF

     IF PayEdit(1).PaymentSource = "P" THEN
      IF PayEdit(1).PaymentAcct > 0 THEN
       GET TBFile, PayEdit(1).PaymentAcct, TBPPCust(1)
       GET TBBalFile, PayEdit(1).PaymentAcct, TBPPCustBal(1)
       GOSUB FigureBalanceSup
       GOSUB PostPaymentSup
       PUT TBFile, PayEdit(1).PaymentAcct, TBPPCust(1)
       PUT TBBalFile, PayEdit(1).PaymentAcct, TBPPCustBal(1)
       GOSUB FigureBalanceSup
       GOSUB PostPersHistorySup
      END IF
     END IF

   TBRFile = FREEFILE
   OPEN TBPath$ + "TBRECFIL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBRFile LEN = LEN(TBRECPT(1))

   
   IF VAL(RIGHT$(PostDate$, 2)) < 75 THEN
    PostDate$ = LEFT$(PostDate$, 6) + "20" + RIGHT$(PostDate$, 2)
    ELSE
    PostDate$ = LEFT$(PostDate$, 6) + "19" + RIGHT$(PostDate$, 2)
   END IF

   TBRECPT(1).Date = Date2Num%(PostDate$)
   TBRECPT(1).AmtIn = AmtIn#
   TBRECPT(1).AmtOwed = AmountPaid#
   IF curency = 1 THEN
    TBRECPT(1).CashAmt = AmountPaid#
    TBRECPT(1).ChkAmt = 0
   END IF
   IF curency = 2 THEN
    TBRECPT(1).CashAmt = 0
    TBRECPT(1).ChkAmt = AmountPaid#
   END IF

   IF PayEdit(1).PaymentSource = "R" THEN
    TBRECPT(1).TranType = 3
    ELSE
    TBRECPT(1).TranType = 4
   END IF
   TBRECPT(1).TranCode = 0
   TBRECPT(1).TaxYear = VAL(TaxYear$)
   TBRECPT(1).Desc = NME$
   IF PayEdit(1).PaymentSource = "R" THEN
   TBRECPT(1).RealTxPd = PayEdit(1).PaymentAmt - (Interest# + Penalty#)
   ELSE
   TBRECPT(1).PersTxPd = PayEdit(1).PaymentAmt - (Interest# + Penalty#)
   END IF
   TBRECPT(1).MTTxPd = 0
   TBRECPT(1).MCTxPd = 0
   TBRECPT(1).MHTxPd = 0
   TBRECPT(1).FETxPd = 0
   TBRECPT(1).IntTxPd = Interest#
   TBRECPT(1).PenTxPd = Penalty#
   TBRECPT(1).Fudge = ""
   TBRECPT(1).Drawer = OpNumber
   ReceiptNumber& = (LOF(TBRFile) / LEN(TBRECPT(1))) + 1
   PUT TBRFile, ReceiptNumber&, TBRECPT(1)
   CLOSE TBRFile
    IF AmtIn# - AmountPaid# > 0 THEN
    LOCATE 21, 1
    COLOR 15
    PRINT CHR$(7);
    PRINT "CHANGE DUE CUSTOMER .......... "; USING "#####,#.##"; AmtIn# - AmountPaid#
    END IF


RECPT2b:
   MFile = FREEFILE
   OPEN TBPath$ + "TBMESSAG.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #MFile LEN = 256
   FIELD MFile, 30 AS Message$
   IF LOF(MFile) > 0 THEN
    GET MFile, 1
    Mess$ = Message$
   END IF
   CLOSE MFile


     LPRINT "     LUNENBURG COUNTY"
     LPRINT "    RECEIPT FOR PAYMENT"
     IF PayEdit(1).PaymentSource = "R" THEN
      LPRINT "     Real Estate Taxes"
     ELSE
      LPRINT "  Personal Property Taxes"
     END IF
     LPRINT "        "; PostDate$
     LPRINT
     LPRINT "Customer: "; LEFT$(NME$, 18)
     LPRINT
     LPRINT "  Total Taxes Due: "; USING "$#####,#.##"; Owed#
     LPRINT
     LPRINT "Amt of Taxes Paid: "; USING "$#####,#.##"; AmountPaid#
     LPRINT "Rec From Customer: "; USING "$#####,#.##"; AmtIn#
     LPRINT "       Change Due: "; USING "$#####,#.##"; AmtIn# - AmountPaid#
     LPRINT
     LPRINT "Receipt # "; USING "######"; ReceiptNumber&
     LPRINT Mess$
     LPRINT
     LPRINT "    T H A N K   Y O U  !"
     FOR JJ = 1 TO 6
     LPRINT " "
     NEXT JJ
     LPRINT CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
     LOCATE 22, 1: COLOR 11
     PRINT "PRESS <ENTER> TO CONTINUE "; : fl = 1: GOSUB 53000
     CLOSE
     RETURN

PostRealHistorySup:
 GET TBFile, PayEdit(1).PaymentAcct, TBRECust(1)
 GET TBBalFile, PayEdit(1).PaymentAcct, TBRECustBal(1)
 GOSUB FigureBalanceSup
 AcctBalance# = CurBal# + PastDue#
 GOSUB 51000
 PayDate$ = PayEdit(1).PaymentDate
 
 IF VAL(RIGHT$(PayDate$, 2)) < 75 THEN
  PayDate$ = LEFT$(PayDate$, 2) + "/" + MID$(PayDate$, 3, 2) + "/" + "20" + RIGHT$(PayDate$, 2)
  ELSE
  PayDate$ = LEFT$(PayDate$, 2) + "/" + MID$(PayDate$, 3, 2) + "/" + "19" + RIGHT$(PayDate$, 2)
 END IF
 TBTrans(1).TransDate = Date2Num%(PayDate$)
 TBTrans(1).TransAmt = PayEdit(1).PaymentAmt
 TBTrans(1).TrTaxYear = PayEdit(1).PaymentManYear
 TBTrans(1).TrDesc = PayEdit(1).PaymentDesc
 TBTrans(1).TrType = 2                   '1=bill 2=payment 3=INT  4= PEN 5=release 6=BEG BAL
 TBTrans(1).TrTax = PayEdit(1).PaymentAmt - (Interest# + Penalty#)
 TBTrans(1).TrInterest = Interest#
 TBTrans(1).TrPenalty = Penalty#
 TBTrans(1).TrLate = 0
 TBTrans(1).TrMachTools = 0
 TBTrans(1).TrMobileHome = 0
 TBTrans(1).TrMerchCap = 0
 TBTrans(1).TrFarmEquip = 0
 TBTrans(1).TrCustomerRecord = PayEdit(1).PaymentAcct
 TBTrans(1).TrNextTrans = 0
 TBTrans(1).TrRunningBal = AcctBalance#
 TBTrans(1).TrSource = "R"
 TBTrans(1).TrPad = ""
 PUT TBTranFile, LOF(TBTranFile) / LEN(TBTrans(1)) + 1, TBTrans(1)
 TransRecord! = LOF(TBTranFile) / LEN(TBTrans(1))
 IF TBRECust(1).FirstTrans = 0 THEN
    TBRECust(1).FirstTrans = TransRecord!
    TBRECust(1).LastTrans = TransRecord!
    PUT TBFile, PayEdit(1).PaymentAcct, TBRECust(1)
    ELSE
    PrevTrans! = TBRECust(1).LastTrans: TBRECust(1).LastTrans = TransRecord!: PUT TBFile, PayEdit(1).PaymentAcct, TBRECust(1)
    GET TBTranFile, PrevTrans!, TBTrans(1)
    TBTrans(1).TrNextTrans = TransRecord!
    PUT TBTranFile, PrevTrans!, TBTrans(1)
 END IF
 CLOSE TBTranFile
 RETURN


PostPersHistorySup:
 GET TBFile, PayEdit(1).PaymentAcct, TBPPCust(1)
 GET TBBalFile, PayEdit(1).PaymentAcct, TBPPCustBal(1)
 GOSUB FigureBalanceSup
 AcctBalance# = CurBal# + PastDue#
 GOSUB 51000
 IF VAL(RIGHT$(PayDate$, 2)) < 75 THEN
  PayDate$ = LEFT$(PayDate$, 2) + "/" + MID$(PayDate$, 3, 2) + "/" + "20" + RIGHT$(PayDate$, 2)
  ELSE
  PayDate$ = LEFT$(PayDate$, 2) + "/" + MID$(PayDate$, 3, 2) + "/" + "19" + RIGHT$(PayDate$, 2)
 END IF
 TBTrans(1).TransDate = Date2Num%(PayDate$)
 TBTrans(1).TransAmt = PayEdit(1).PaymentAmt
 TBTrans(1).TrTaxYear = PayEdit(1).PaymentManYear
 TBTrans(1).TrDesc = PayEdit(1).PaymentDesc
 TBTrans(1).TrType = 2
 TBTrans(1).TrTax = PayEdit(1).PaymentAmt - (Interest# + Penalty#)
 TBTrans(1).TrInterest = Interest#
 TBTrans(1).TrPenalty = Penalty#
 TBTrans(1).TrLate = 0
 TBTrans(1).TrMachTools = 0
 TBTrans(1).TrMobileHome = 0
 TBTrans(1).TrMerchCap = 0
 TBTrans(1).TrFarmEquip = 0
 TBTrans(1).TrCustomerRecord = PayEdit(1).PaymentAcct
 TBTrans(1).TrNextTrans = 0
 TBTrans(1).TrRunningBal = AcctBalance#
 TBTrans(1).TrSource = "P"
 TBTrans(1).TrPad = ""
 PUT TBTranFile, LOF(TBTranFile) / LEN(TBTrans(1)) + 1, TBTrans(1)
 TransRecord! = LOF(TBTranFile) / LEN(TBTrans(1))
 IF TBPPCust(1).FirstTrans = 0 THEN
    TBPPCust(1).FirstTrans = TransRecord!
    TBPPCust(1).LastTrans = TransRecord!
    PUT TBFile, PayEdit(1).PaymentAcct, TBPPCust(1)
    ELSE
    PrevTrans! = TBPPCust(1).LastTrans: TBPPCust(1).LastTrans = TransRecord!: PUT TBFile, PayEdit(1).PaymentAcct, TBPPCust(1)
    GET TBTranFile, PrevTrans!, TBTrans(1)
    TBTrans(1).TrNextTrans = TransRecord!
    PUT TBTranFile, PrevTrans!, TBTrans(1)
 END IF
 CLOSE TBTranFile
 RETURN

PostPaymentSup:
IF taxtype$ = "R" THEN
  REM process real estate
  IF PayEdit(1).PaymentDist = "M" THEN
   IF VAL(PayEdit(1).PaymentManYear) = CurrentYear THEN
    TotalCredit# = PayEdit(1).PaymentAmt
    IF TBRECust(1).CurYrOthAmount > TotalCredit# THEN
     TBRECust(1).CurYrOthAmount = TBRECust(1).CurYrOthAmount - TotalCredit#
     RETURN
    END IF
    IF TBRECust(1).CurYrOthAmount >= 0 THEN
     TotalCredit# = TotalCredit# - TBRECust(1).CurYrOthAmount
     TBRECust(1).CurYrOthAmount = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBRECust(1).CurYrPenAmount > TotalCredit# THEN
     TBRECust(1).CurYrPenAmount = TBRECust(1).CurYrPenAmount - TotalCredit#
     Penalty# = Penalty# + TotalCredit#
     RETURN
    END IF
    IF TBRECust(1).CurYrPenAmount >= 0 THEN
     TotalCredit# = TotalCredit# - TBRECust(1).CurYrPenAmount
     Penalty# = Penalty# + TBRECust(1).CurYrPenAmount
     TBRECust(1).CurYrPenAmount = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBRECust(1).CurYrIntAmount > TotalCredit# THEN
     TBRECust(1).CurYrIntAmount = TBRECust(1).CurYrIntAmount - TotalCredit#
     Interest# = Interest# + TotalCredit#
     RETURN
    END IF
    IF TBRECust(1).CurYrIntAmount >= 0 THEN
     TotalCredit# = TotalCredit# - TBRECust(1).CurYrIntAmount
     Interest# = Interest# + TBRECust(1).CurYrIntAmount
     TBRECust(1).CurYrIntAmount = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    TBRECust(1).CurYrTaxAmount = TBRECust(1).CurYrTaxAmount - TotalCredit#
    RETURN
   ELSE
   REM process prior year
   Pyr = CurrentYear - VAL(PayEdit(1).PaymentManYear)
    TotalCredit# = PayEdit(1).PaymentAmt
    IF TBRECustBal(1).OthAmt(Pyr) > TotalCredit# THEN
     TBRECustBal(1).OthAmt(Pyr) = TBRECustBal(1).OthAmt(Pyr) - TotalCredit#
     RETURN
    END IF
    IF TBRECustBal(1).OthAmt(Pyr) >= 0 THEN
     TotalCredit# = TotalCredit# - TBRECustBal(1).OthAmt(Pyr)
     TBRECustBal(1).OthAmt(Pyr) = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBRECustBal(1).PenAmt(Pyr) > TotalCredit# THEN
     TBRECustBal(1).PenAmt(Pyr) = TBRECustBal(1).PenAmt(Pyr) - TotalCredit#
     Penalty# = Penalty# + TotalCredit#
     RETURN
    END IF
    IF TBRECustBal(1).PenAmt(Pyr) >= 0 THEN
     TotalCredit# = TotalCredit# - TBRECustBal(1).PenAmt(Pyr)
     Penalty# = Penalty# + TBRECustBal(1).PenAmt(Pyr)
     TBRECustBal(1).PenAmt(Pyr) = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBRECustBal(1).IntAmt(Pyr) > TotalCredit# THEN
     TBRECustBal(1).IntAmt(Pyr) = TBRECustBal(1).IntAmt(Pyr) - TotalCredit#
     Interest# = Interest# + TotalCredit#
     RETURN
    END IF
    IF TBRECustBal(1).IntAmt(Pyr) >= 0 THEN
     TotalCredit# = TotalCredit# - TBRECustBal(1).IntAmt(Pyr)
     Interest# = Interest# + TBRECustBal(1).IntAmt(Pyr)
     TBRECustBal(1).IntAmt(Pyr) = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    TBRECustBal(1).TaxAmt(Pyr) = TBRECustBal(1).TaxAmt(Pyr) - TotalCredit#
    RETURN
   END IF

   ELSE

   REM process auto distribution here
   TotalCredit# = PayEdit(1).PaymentAmt
   IF PastDue# > 0 THEN
   FOR Pyr = 10 TO 1 STEP -1
    IF TBRECustBal(1).OthAmt(Pyr) > TotalCredit# THEN
     TBRECustBal(1).OthAmt(Pyr) = TBRECustBal(1).OthAmt(Pyr) - TotalCredit#
     RETURN
    END IF
    IF TBRECustBal(1).OthAmt(Pyr) >= 0 THEN
     TotalCredit# = TotalCredit# - TBRECustBal(1).OthAmt(Pyr)
     TBRECustBal(1).OthAmt(Pyr) = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBRECustBal(1).PenAmt(Pyr) > TotalCredit# THEN
     TBRECustBal(1).PenAmt(Pyr) = TBRECustBal(1).PenAmt(Pyr) - TotalCredit#
     Penalty# = Penalty# + TotalCredit#
     RETURN
    END IF
    IF TBRECustBal(1).PenAmt(Pyr) >= 0 THEN
     TotalCredit# = TotalCredit# - TBRECustBal(1).PenAmt(Pyr)
     Penalty# = Penalty# + TBRECustBal(1).PenAmt(Pyr)
     TBRECustBal(1).PenAmt(Pyr) = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBRECustBal(1).IntAmt(Pyr) > TotalCredit# THEN
     TBRECustBal(1).IntAmt(Pyr) = TBRECustBal(1).IntAmt(Pyr) - TotalCredit#
     Interest# = Interest# + TotalCredit#
     RETURN
    END IF
    IF TBRECustBal(1).IntAmt(Pyr) >= 0 THEN
     TotalCredit# = TotalCredit# - TBRECustBal(1).IntAmt(Pyr)
     Interest# = Interest# + TBRECustBal(1).IntAmt(Pyr)
     TBRECustBal(1).IntAmt(Pyr) = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBRECustBal(1).TaxAmt(Pyr) > TotalCredit# THEN
     TBRECustBal(1).TaxAmt(Pyr) = TBRECustBal(1).TaxAmt(Pyr) - TotalCredit#
     RETURN
    END IF
    IF TBRECustBal(1).TaxAmt(Pyr) >= 0 THEN
     TotalCredit# = TotalCredit# - TBRECustBal(1).TaxAmt(Pyr)
     TBRECustBal(1).TaxAmt(Pyr) = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    NEXT Pyr
    END IF
    REM put current here
    IF TBRECust(1).CurYrOthAmount > TotalCredit# THEN
     TBRECust(1).CurYrOthAmount = TBRECust(1).CurYrOthAmount - TotalCredit#
     RETURN
    END IF
    IF TBRECust(1).CurYrOthAmount >= 0 THEN
     TotalCredit# = TotalCredit# - TBRECust(1).CurYrOthAmount
     TBRECust(1).CurYrOthAmount = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBRECust(1).CurYrPenAmount >= TotalCredit# THEN
     TBRECust(1).CurYrPenAmount = TBRECust(1).CurYrPenAmount - TotalCredit#
     Penalty# = Penalty# + TotalCredit#
     RETURN
    END IF
    IF TBRECust(1).CurYrPenAmount >= 0 THEN
     TotalCredit# = TotalCredit# - TBRECust(1).CurYrPenAmount
     Penalty# = Penalty# + TBRECust(1).CurYrPenAmount
     TBRECust(1).CurYrPenAmount = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBRECust(1).CurYrIntAmount > TotalCredit# THEN
     TBRECust(1).CurYrIntAmount = TBRECust(1).CurYrIntAmount - TotalCredit#
     Interest# = Interest# + TotalCredit#
     RETURN
    END IF
    IF TBRECust(1).CurYrIntAmount >= 0 THEN
     TotalCredit# = TotalCredit# - TBRECust(1).CurYrIntAmount
     Interest# = Interest# + TBRECust(1).CurYrIntAmount
     TBRECust(1).CurYrIntAmount = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    TBRECust(1).CurYrTaxAmount = TBRECust(1).CurYrTaxAmount - TotalCredit#
    RETURN

  END IF
END IF





IF taxtype$ = "P" THEN
  IF PayEdit(1).PaymentDist = "M" THEN
   IF VAL(PayEdit(1).PaymentManYear) = CurrentYear THEN
    TotalCredit# = PayEdit(1).PaymentAmt
    IF TBPPCust(1).CurYrPenAmount > TotalCredit# THEN
       TBPPCust(1).CurYrPenAmount = TBPPCust(1).CurYrPenAmount - TotalCredit#
       Penalty# = Penalty# + TotalCredit#
       RETURN
    END IF
    IF TBPPCust(1).CurYrPenAmount >= 0 THEN
     TotalCredit# = TotalCredit# - TBPPCust(1).CurYrPenAmount
     Penalty# = Penalty# + TBPPCust(1).CurYrPenAmount
     TBPPCust(1).CurYrPenAmount = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrIntAmount > TotalCredit# THEN
       TBPPCust(1).CurYrIntAmount = TBPPCust(1).CurYrIntAmount - TotalCredit#
       Interest# = Interest# + TotalCredit#
       RETURN
    END IF
    IF TBPPCust(1).CurYrIntAmount >= 0 THEN
     TotalCredit# = TotalCredit# - TBPPCust(1).CurYrIntAmount
     Interest# = Interest# + TBPPCust(1).CurYrIntAmount
     TBPPCust(1).CurYrIntAmount = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrLateFeeTax > TotalCredit# THEN
       TBPPCust(1).CurYrLateFeeTax = TBPPCust(1).CurYrLateFeeTax - TotalCredit#
       RETURN
    END IF
    IF TBPPCust(1).CurYrLateFeeTax >= 0 THEN
     TotalCredit# = TotalCredit# - TBPPCust(1).CurYrLateFeeTax
     TBPPCust(1).CurYrLateFeeTax = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrMobileHomeTax > TotalCredit# THEN
       TBPPCust(1).CurYrMobileHomeTax = TBPPCust(1).CurYrMobileHomeTax - TotalCredit#
       RETURN
    END IF
    IF TBPPCust(1).CurYrMobileHomeTax >= 0 THEN
     TotalCredit# = TotalCredit# - TBPPCust(1).CurYrMobileHomeTax
     TBPPCust(1).CurYrMobileHomeTax = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrMerchCapTax > TotalCredit# THEN
       TBPPCust(1).CurYrMerchCapTax = TBPPCust(1).CurYrMerchCapTax - TotalCredit#
       RETURN
    END IF
    IF TBPPCust(1).CurYrMerchCapTax >= 0 THEN
     TotalCredit# = TotalCredit# - TBPPCust(1).CurYrMerchCapTax
     TBPPCust(1).CurYrMerchCapTax = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrMachToolsTax > TotalCredit# THEN
       TBPPCust(1).CurYrMachToolsTax = TBPPCust(1).CurYrMachToolsTax - TotalCredit#
       RETURN
    END IF
    IF TBPPCust(1).CurYrMachToolsTax >= 0 THEN
     TotalCredit# = TotalCredit# - TBPPCust(1).CurYrMachToolsTax
     TBPPCust(1).CurYrMachToolsTax = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrFarmEquipTax > TotalCredit# THEN
       TBPPCust(1).CurYrFarmEquipTax = TBPPCust(1).CurYrFarmEquipTax - TotalCredit#
       RETURN
    END IF
    IF TBPPCust(1).CurYrFarmEquipTax >= 0 THEN
     TotalCredit# = TotalCredit# - TBPPCust(1).CurYrFarmEquipTax
     TBPPCust(1).CurYrFarmEquipTax = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    TBPPCust(1).CurYrPersPropTax = TBPPCust(1).CurYrPersPropTax - TotalCredit#
    RETURN
    ELSE
    REM process prior year
    Pyr = CurrentYear - VAL(PayEdit(1).PaymentManYear)
     TotalCredit# = PayEdit(1).PaymentAmt
     IF TBPPCustBal(1).PenAmt(Pyr) > TotalCredit# THEN
      TBPPCustBal(1).PenAmt(Pyr) = TBPPCustBal(1).PenAmt(Pyr) - TotalCredit#
      Penalty# = Penalty# + TotalCredit#
      RETURN
     END IF
     IF TBPPCustBal(1).PenAmt(Pyr) >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCustBal(1).PenAmt(Pyr)
      Penalty# = Penalty# + TBPPCustBal(1).PenAmt(Pyr)
      TBPPCustBal(1).PenAmt(Pyr) = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
     END IF
     IF TotalCredit# = 0 THEN RETURN
     IF TBPPCustBal(1).IntAmt(Pyr) > TotalCredit# THEN
      TBPPCustBal(1).IntAmt(Pyr) = TBPPCustBal(1).IntAmt(Pyr) - TotalCredit#
      Interest# = Interest# + TotalCredit#
      RETURN
     END IF
     IF TBPPCustBal(1).IntAmt(Pyr) >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCustBal(1).IntAmt(Pyr)
      Interest# = Interest# + TBPPCustBal(1).IntAmt(Pyr)
      TBPPCustBal(1).IntAmt(Pyr) = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
     END IF
     IF TotalCredit# = 0 THEN RETURN
     IF TBPPCustBal(1).LateFeeAmt(Pyr) > TotalCredit# THEN
      TBPPCustBal(1).LateFeeAmt(Pyr) = TBPPCustBal(1).LateFeeAmt(Pyr) - TotalCredit#
      RETURN
     END IF
     IF TBPPCustBal(1).LateFeeAmt(Pyr) >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCustBal(1).LateFeeAmt(Pyr)
      TBPPCustBal(1).LateFeeAmt(Pyr) = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
     END IF
     IF TotalCredit# = 0 THEN RETURN
     IF TBPPCustBal(1).MobileHomeAmt(Pyr) > TotalCredit# THEN
      TBPPCustBal(1).MobileHomeAmt(Pyr) = TBPPCustBal(1).MobileHomeAmt(Pyr) - TotalCredit#
      RETURN
     END IF
     IF TBPPCustBal(1).MobileHomeAmt(Pyr) >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCustBal(1).MobileHomeAmt(Pyr)
      TBPPCustBal(1).MobileHomeAmt(Pyr) = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
     END IF
     IF TotalCredit# = 0 THEN RETURN
     IF TBPPCustBal(1).MerchCapAmt(Pyr) > TotalCredit# THEN
      TBPPCustBal(1).MerchCapAmt(Pyr) = TBPPCustBal(1).MerchCapAmt(Pyr) - TotalCredit#
      RETURN
     END IF
     IF TBPPCustBal(1).MerchCapAmt(Pyr) >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCustBal(1).MerchCapAmt(Pyr)
      TBPPCustBal(1).MerchCapAmt(Pyr) = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
     END IF
     IF TotalCredit# = 0 THEN RETURN
     IF TBPPCustBal(1).MachToolsAmt(Pyr) > TotalCredit# THEN
      TBPPCustBal(1).MachToolsAmt(Pyr) = TBPPCustBal(1).MachToolsAmt(Pyr) - TotalCredit#
      RETURN
     END IF
     IF TBPPCustBal(1).MachToolsAmt(Pyr) >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCustBal(1).MachToolsAmt(Pyr)
      TBPPCustBal(1).MachToolsAmt(Pyr) = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
     END IF
     IF TotalCredit# = 0 THEN RETURN
     IF TBPPCustBal(1).FarmEquipAmt(Pyr) > TotalCredit# THEN
      TBPPCustBal(1).FarmEquipAmt(Pyr) = TBPPCustBal(1).FarmEquipAmt(Pyr) - TotalCredit#
      RETURN
     END IF
     IF TBPPCustBal(1).FarmEquipAmt(Pyr) >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCustBal(1).FarmEquipAmt(Pyr)
      TBPPCustBal(1).FarmEquipAmt(Pyr) = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
     END IF
     IF TotalCredit# = 0 THEN RETURN
     TBPPCustBal(1).PerPropTaxAmt(Pyr) = TBPPCustBal(1).PerPropTaxAmt(Pyr) - TotalCredit#
     RETURN

    END IF
    ELSE
    REM auto process here
    TotalCredit# = PayEdit(1).PaymentAmt
    IF PastDue# > 0 THEN
    FOR Pyr = 10 TO 1 STEP -1
     IF TBPPCustBal(1).PenAmt(Pyr) > TotalCredit# THEN
      TBPPCustBal(1).PenAmt(Pyr) = TBPPCustBal(1).PenAmt(Pyr) - TotalCredit#
      Penalty# = Penalty# + TotalCredit#
      RETURN
     END IF
     IF TBPPCustBal(1).PenAmt(Pyr) >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCustBal(1).PenAmt(Pyr)
      Penalty# = Penalty# + TBPPCustBal(1).PenAmt(Pyr)
      TBPPCustBal(1).PenAmt(Pyr) = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
     END IF
     IF TotalCredit# = 0 THEN RETURN
     IF TBPPCustBal(1).IntAmt(Pyr) > TotalCredit# THEN
      TBPPCustBal(1).IntAmt(Pyr) = TBPPCustBal(1).IntAmt(Pyr) - TotalCredit#
      Interest# = Interest# + TotalCredit#
      RETURN
     END IF
     IF TBPPCustBal(1).IntAmt(Pyr) >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCustBal(1).IntAmt(Pyr)
      Interest# = Interest# + TBPPCustBal(1).IntAmt(Pyr)
      TBPPCustBal(1).IntAmt(Pyr) = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
     END IF
     IF TotalCredit# = 0 THEN RETURN
     IF TBPPCustBal(1).LateFeeAmt(Pyr) > TotalCredit# THEN
      TBPPCustBal(1).LateFeeAmt(Pyr) = TBPPCustBal(1).LateFeeAmt(Pyr) - TotalCredit#
      RETURN
     END IF
     IF TBPPCustBal(1).LateFeeAmt(Pyr) >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCustBal(1).LateFeeAmt(Pyr)
      TBPPCustBal(1).LateFeeAmt(Pyr) = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
     END IF
     IF TotalCredit# = 0 THEN RETURN
     IF TBPPCustBal(1).MobileHomeAmt(Pyr) > TotalCredit# THEN
      TBPPCustBal(1).MobileHomeAmt(Pyr) = TBPPCustBal(1).MobileHomeAmt(Pyr) - TotalCredit#
      RETURN
     END IF
     IF TBPPCustBal(1).MobileHomeAmt(Pyr) >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCustBal(1).MobileHomeAmt(Pyr)
      TBPPCustBal(1).MobileHomeAmt(Pyr) = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
     END IF
     IF TotalCredit# = 0 THEN RETURN
     IF TBPPCustBal(1).MerchCapAmt(Pyr) > TotalCredit# THEN
      TBPPCustBal(1).MerchCapAmt(Pyr) = TBPPCustBal(1).MerchCapAmt(Pyr) - TotalCredit#
      RETURN
     END IF
     IF TBPPCustBal(1).MerchCapAmt(Pyr) >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCustBal(1).MerchCapAmt(Pyr)
      TBPPCustBal(1).MerchCapAmt(Pyr) = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
     END IF
     IF TotalCredit# = 0 THEN RETURN
     IF TBPPCustBal(1).MachToolsAmt(Pyr) > TotalCredit# THEN
      TBPPCustBal(1).MachToolsAmt(Pyr) = TBPPCustBal(1).MachToolsAmt(Pyr) - TotalCredit#
      RETURN
     END IF
     IF TBPPCustBal(1).MachToolsAmt(Pyr) >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCustBal(1).MachToolsAmt(Pyr)
      TBPPCustBal(1).MachToolsAmt(Pyr) = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
     END IF
     IF TotalCredit# = 0 THEN RETURN
     IF TBPPCustBal(1).FarmEquipAmt(Pyr) > TotalCredit# THEN
      TBPPCustBal(1).FarmEquipAmt(Pyr) = TBPPCustBal(1).FarmEquipAmt(Pyr) - TotalCredit#
      RETURN
     END IF
     IF TBPPCustBal(1).FarmEquipAmt(Pyr) >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCustBal(1).FarmEquipAmt(Pyr)
      TBPPCustBal(1).FarmEquipAmt(Pyr) = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
     END IF
     IF TotalCredit# = 0 THEN RETURN
     IF TBPPCustBal(1).PerPropTaxAmt(Pyr) > TotalCredit# THEN
      TBPPCustBal(1).PerPropTaxAmt(Pyr) = TBPPCustBal(1).PerPropTaxAmt(Pyr) - TotalCredit#
      RETURN
     END IF
     IF TBPPCustBal(1).PerPropTaxAmt(Pyr) >= 0 THEN
      TotalCredit# = TotalCredit# - TBPPCustBal(1).PerPropTaxAmt(Pyr)
      TBPPCustBal(1).PerPropTaxAmt(Pyr) = 0
      TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
     END IF
     IF TotalCredit# = 0 THEN RETURN
    NEXT Pyr

    END IF
    REM current here
    IF TBPPCust(1).CurYrPenAmount > TotalCredit# THEN
       TBPPCust(1).CurYrPenAmount = TBPPCust(1).CurYrPenAmount - TotalCredit#
       Penalty# = Penalty# + TotalCredit#
       RETURN
    END IF
    IF TBPPCust(1).CurYrPenAmount >= 0 THEN
     TotalCredit# = TotalCredit# - TBPPCust(1).CurYrPenAmount
     Penalty# = Penalty# + TBPPCust(1).CurYrPenAmount
     TBPPCust(1).CurYrPenAmount = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrIntAmount > TotalCredit# THEN
       TBPPCust(1).CurYrIntAmount = TBPPCust(1).CurYrIntAmount - TotalCredit#
       Interest# = Interest# + TotalCredit#
       RETURN
    END IF
    IF TBPPCust(1).CurYrIntAmount >= 0 THEN
     TotalCredit# = TotalCredit# - TBPPCust(1).CurYrIntAmount
     Interest# = Interest# + TBPPCust(1).CurYrIntAmount
     TBPPCust(1).CurYrIntAmount = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrLateFeeTax > TotalCredit# THEN
       TBPPCust(1).CurYrLateFeeTax = TBPPCust(1).CurYrLateFeeTax - TotalCredit#
       RETURN
    END IF
    IF TBPPCust(1).CurYrLateFeeTax >= 0 THEN
     TotalCredit# = TotalCredit# - TBPPCust(1).CurYrLateFeeTax
     TBPPCust(1).CurYrLateFeeTax = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrMobileHomeTax > TotalCredit# THEN
       TBPPCust(1).CurYrMobileHomeTax = TBPPCust(1).CurYrMobileHomeTax - TotalCredit#
       RETURN
    END IF
    IF TBPPCust(1).CurYrMobileHomeTax >= 0 THEN
     TotalCredit# = TotalCredit# - TBPPCust(1).CurYrMobileHomeTax
     TBPPCust(1).CurYrMobileHomeTax = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrMerchCapTax > TotalCredit# THEN
       TBPPCust(1).CurYrMerchCapTax = TBPPCust(1).CurYrMerchCapTax - TotalCredit#
       RETURN
    END IF
    IF TBPPCust(1).CurYrMerchCapTax >= 0 THEN
     TotalCredit# = TotalCredit# - TBPPCust(1).CurYrMerchCapTax
     TBPPCust(1).CurYrMerchCapTax = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrMachToolsTax > TotalCredit# THEN
       TBPPCust(1).CurYrMachToolsTax = TBPPCust(1).CurYrMachToolsTax - TotalCredit#
       RETURN
    END IF
    IF TBPPCust(1).CurYrMachToolsTax >= 0 THEN
     TotalCredit# = TotalCredit# - TBPPCust(1).CurYrMachToolsTax
     TBPPCust(1).CurYrMachToolsTax = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    IF TBPPCust(1).CurYrFarmEquipTax > TotalCredit# THEN
       TBPPCust(1).CurYrFarmEquipTax = TBPPCust(1).CurYrFarmEquipTax - TotalCredit#
       RETURN
    END IF
    IF TBPPCust(1).CurYrFarmEquipTax >= 0 THEN
     TotalCredit# = TotalCredit# - TBPPCust(1).CurYrFarmEquipTax
     TBPPCust(1).CurYrFarmEquipTax = 0
     TotalCredit# = INT((TotalCredit# * 100) + .5) / 100
    END IF
    IF TotalCredit# = 0 THEN RETURN
    TBPPCust(1).CurYrPersPropTax = TBPPCust(1).CurYrPersPropTax - TotalCredit#
    RETURN
    END IF
  END IF
RETURN

50000  IF taxtype$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TTRSCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBPath$ + "TTPSCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN

50025 TBBalFile = FREEFILE
      IF taxtype$ = "R" THEN
      OPEN TBPath$ + "TTRSBal.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBPath$ + "TTPSBAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF
50030 RETURN

51000 TBTranFile = FREEFILE
      OPEN TBPath$ + "TTTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTranFile LEN = LEN(TBTrans(1))
51020 RETURN

53000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
53010 IF INKEY$ <> "" THEN 53010
53020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
53030 z$ = INKEY$: IF z$ = "" THEN 53030
53040 IF z$ <> CHR$(13) THEN 53060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
53050 GOSUB 53200: RETURN
53060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(z$) < 2 THEN 53065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN cf = 12: RETURN
53065 IF ASC(z$) = 0 THEN 53030
53070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
53080 IF z$ = " " THEN 53130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 53000
53090 IF z$ <> CHR$(8) THEN 53130 ELSE IF ZL = 0 THEN 53030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 53110 ELSE IF z$ = "," THEN 53120
53100 IF z$ = "." THEN ZD = 0: GOTO 53110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
53110 in$ = LEFT$(in$, LEN(in$) - 1)
53120 ZL = ZL - 1: GOTO 53030
53130 IF ABS(fl) = ZL THEN 53030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 53180
53140 IF z$ = "." AND INSTR(in$, ".") = 0 THEN GOTO 53180
53150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 53190
53160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 53180
53170 IF z$ < "0" OR z$ > "9" THEN 53030
53180 PRINT z$; : in$ = in$ + z$: ZL = ZL + 1
53190 IF ABS(fl) = 1 THEN 53050 ELSE 53030
53200 LOCATE RO, PO: PRINT in$: RETURN
53210 GOSUB 53000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
53220 RETURN
53500 COLOR 10: LOCATE 1, 1
     PRINT "Post Date: "; : COLOR 15
     PRINT LEFT$(PDATE$, 2) + "-" + MID$(PDATE$, 3, 2) + "-" + MID$(PDATE$, 5, 2);
     COLOR 10
     PRINT TAB(70); scren$
     PRINT STRING$(79, "_")
     RETURN

NamesearchSup:
 LOCATE 8, 1:
 COLOR 11
 PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
 GOSUB 53000
 IF cf = 1 THEN
  ggo = 0
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  RETURN
 END IF
 nn$ = in$
 nn = LEN(nn$)
 IF nn > 20 THEN nn = 20
 nn$ = LEFT$(nn$, nn)

 IF taxtype$ = "R" THEN
  TBIDXFile = FREEFILE
  OPEN TBPath$ + "TTRTNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS Mrec$
  GOSUB 50000
  FOR ll = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, ll
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO Supn1
Supn3:
  NEXT ll
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  ggo = 0
  RETURN

Supn1:
GET TBFile, CVS(Mrec$), TBRECust(1)
LOCATE 11, 1: COLOR 11
NME$ = TBRECust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT NME$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(Mrec$)
Supn2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 53210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO Supn2
IF in$ = "Q" THEN
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  ggo = 0
  RETURN
END IF
IF in$ = "Y" THEN
  ggo = 1: Record! = CVS(Mrec$)
  FOR LLL = 8 TO 13:  LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO Supn3
GOTO Supn2
END IF


  TBIDXFile = FREEFILE
  OPEN TBPath$ + "TTPTNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS Mrec$
  GOSUB 50000
  FOR ll = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, ll
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO SupPPn1
SupPPn3:
  NEXT ll
  LOCATE 8, 1: PRINT STRING$(79, 32)
  ggo = 0
  RETURN

SupPPn1:
GET TBFile, CVS(Mrec$), TBPPCust(1)
LOCATE 11, 1: COLOR 11
NME$ = TBPPCust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT NME$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(Mrec$)
SupPPn2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 53210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO SupPPn2
IF in$ = "Q" THEN
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  ggo = 0
  RETURN
END IF
IF in$ = "Y" THEN
  ggo = 1: Record! = CVS(Mrec$)
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO SupPPn3
GOTO SupPPn2

DeletedSup:
    LOCATE 23, 1: COLOR 15
    PRINT "This Account Has Been Deleted! Press <ENTER> to Continue"; :
    fl = 1
    GOSUB 53000
    LOCATE 23, 1: PRINT STRING$(79, 32)
    RETURN

FigureBalanceSup:
     CurBal# = 0: PastDue# = 0
     IF taxtype$ = "R" THEN
     CurBal# = TBRECust(1).CurYrTaxAmount + TBRECust(1).CurYrIntAmount + TBRECust(1).CurYrPenAmount + TBRECust(1).CurYrOthAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR ll = 1 TO 10
     PastDue# = PastDue# + TBRECustBal(1).TaxAmt(ll) + TBRECustBal(1).IntAmt(ll) + TBRECustBal(1).PenAmt(ll) + TBRECustBal(1).OthAmt(ll) + TBRECustBal(1).LateAmt(ll)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT ll
     ELSE
     CurBal# = TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax + TBPPCust(1).CurYrIntAmount + TBPPCust(1). _
CurYrPenAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR ll = 1 TO 10
     PastDue# = PastDue# + TBPPCustBal(1).PerPropTaxAmt(ll) + TBPPCustBal(1).FarmEquipAmt(ll) + TBPPCustBal(1).MachToolsAmt(ll) + TBPPCustBal(1).MerchCapAmt(ll) + TBPPCustBal(1).MobileHomeAmt(ll) + TBPPCustBal(1).LateFeeAmt(ll) + TBPPCustBal(1). _
IntAmt(ll) + TBPPCustBal(1).PenAmt(ll)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT ll
     END IF
     RETURN


END SUB

