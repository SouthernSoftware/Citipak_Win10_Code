DECLARE SUB MiscCodes (PDATE$, TBPath$)
DECLARE SUB SetMessage ()
DECLARE SUB Cleardown ()
DECLARE SUB DecalPayment (PDATE$, TBPath$)
DECLARE SUB SupTaxPayment (PDATE$, TBPath$)
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
    ' TBPath$ = "c:\QB45\LC\"                    ' change to "F:\" Before compiling

     CLOSE : OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS Year$: GET 1, 1
     CurrentYear$ = Year$: CLOSE
     CurrentYear = VAL(CurrentYear$)
     ON ERROR GOTO networkerror

     DIM SHARED yrdesc(11)
     REDIM TBLog(1) AS TBLog
     DIM SHARED TBRECust(1) AS TBRECust
     DIM SHARED TBRECustBal(1) AS TBREBAL
     DIM SHARED TBPPCust(1) AS TBPPCust
     DIM SHARED TBPPCustBal(1) AS TBPPBal
     DIM SHARED TBTrans(1) AS TBTrans
     DIM SHARED TBRECPT(1) AS DayReceiptsType

     PDATE$ = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 2)


10 CLS
110 scren$ = "5.00.00": sysdate$ = DATE$
120 CLS : GOSUB 63500
1000 scren$ = "4.01.00": CLS : GOSUB 63500
      PostDate$ = LEFT$(PDATE$, 2) + "-" + MID$(PDATE$, 3, 2) + "-" + RIGHT$(PDATE$, 2)
1020 COLOR 11: LOCATE 1, 29: PRINT "Cash Receipt Payment Entry";
     LOCATE 4, 1: COLOR 11
     PRINT TAB(29); "   SELECT PAYMENT TYPE"
     PRINT
     PRINT TAB(25); "1-"; : COLOR 15: PRINT " Regular Tax Payment": COLOR 11
     PRINT TAB(25); "2-"; : COLOR 15: PRINT " Supplemental Tax Payment": COLOR 11
     PRINT TAB(25); "3-"; : COLOR 15: PRINT " Decal Payment": COLOR 11
     PRINT TAB(25); "4-"; : COLOR 15: PRINT " Misc Payment": COLOR 11
     PRINT TAB(25); "5-"; : COLOR 15: PRINT " Print Daily Receipts Record": COLOR 11
     PRINT
     PRINT TAB(25); "   Setup Functions"
          PRINT
     PRINT TAB(25); "6-"; : COLOR 3: PRINT " Set Misc Codes": COLOR 11
     PRINT TAB(25); "7-"; : COLOR 3: PRINT " Set Payment Date": COLOR 11
     PRINT TAB(25); "8-"; : COLOR 3: PRINT " Set Receipt Message Line": COLOR 11
     PRINT TAB(25); "   "
     PRINT TAB(25); "   Enter Selection # <ESC>-End: "; : fl = -1
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "TBTMENU"
     IF VAL(in$) = 1 THEN 1025
     IF VAL(in$) = 2 THEN
       SupTaxPayment PDATE$, TBPath$
       GOTO 1000
     END IF
     IF VAL(in$) = 3 THEN
       DecalPayment PDATE$, TBPath$
       GOTO 1000
     END IF
     IF VAL(in$) = 4 THEN
        GOSUB EnterMiscPay
        GOTO 1000
     END IF
     IF VAL(in$) = 5 THEN
        PrintReceiptReport PDATE$, TBPath$
        GOTO 1000
     END IF
     IF VAL(in$) = 6 THEN
        MiscCodes PDATE$, TBPath$
        GOTO 1000
     END IF
     IF VAL(in$) = 7 THEN
      GOSUB SetPostDate
      GOTO 1000
     END IF
     IF VAL(in$) = 8 THEN
      SetMessage
      GOTO 1000
     END IF

     GOTO 1000


1025 Cleardown
     LOCATE 3, 1: COLOR 11
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P") THEN 1025
     taxtype$ = UCASE$(in$)



1030 LOCATE 4, 1: COLOR 11
     PRINT "Enter Customer Account Number (TAB-Last Name : ESC-Menu) .. "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN
      CLOSE : GOTO 10
     END IF
     IF cf = 2 THEN
      GOSUB Namesearch
      IF ggo = 1 THEN 1075
      IF ggo = 0 THEN LOCATE 10, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 62000: GOTO 1000
     END IF
     Record! = VAL(in$)
1075 CLOSE : GOSUB 60000: GOSUB 60025
     IF taxtype$ = "R" THEN
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBRECust(1)) THEN 1030
       LOCK #TBFile, Record!
       GET TBFile, Record!, TBRECust(1)
       GET TBBalFile, Record!, TBRECustBal(1)
       IF TBRECust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 1000
       ELSE
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBPPCust(1)) THEN 1030
       LOCK #TBFile, Record!
       GET TBFile, Record!, TBPPCust(1)
       GET TBBalFile, Record!, TBPPCustBal(1)
       IF TBPPCust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 1000
     END IF

     IF taxtype$ = "R" THEN
       GOSUB ProcessReal
        ELSE
       GOSUB ProcessPersonal
     END IF
     CLOSE
     GOTO 1000

ProcessReal:
    GOSUB FigureBalance
    PayType = 1 'Regular Payment
    PaySource$ = "R"
    CLS
    scren$ = "4.01.00": CLS : GOSUB 63500
    COLOR 11: LOCATE 1, 32: PRINT "Customer Payment Entry";
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
    PRINT "Correct Account (Y/N) .. "; : fl = 1: GOSUB 62210
    IF cf = 1 THEN CLOSE : RETURN
    IF in$ = "N" THEN CLOSE : RETURN

    LOCATE 10, 1: PRINT STRING$(79, 32): COLOR 11
    LOCATE 10, 1
    PRINT "Enter Amount to Apply on Account ("; USING "######,#.##"; CurBal# + PastDue#; : PRINT ") .... ";
     PRINT TAB(54); ""; : fl = -9: GOSUB 62000
    IF cf = 1 THEN CLOSE : RETURN
    IF LEN(in$) = 0 THEN
      AmountPaid# = CurBal# + PastDue#
      ELSE
      AmountPaid# = VAL(in$)
    END IF
    LOCATE 10, 54: COLOR 14: PRINT USING "######,#.##"; AmountPaid#
     
pr1: COLOR 11: LOCATE 11, 1
     PRINT "Enter (1)Cash or (2)Check ... "; : fl = -1: GOSUB 62000
     IF VAL(in$) < 1 OR VAL(in$) > 2 THEN GOTO pr1
     curency = VAL(in$)
pr2: COLOR 11: LOCATE 12, 1
     PRINT "Enter Amount from Customer .. "; : fl = -9: GOSUB 62000
     LOCATE 12, 30: COLOR 14
     PRINT USING "#####,#.##"; VAL(in$)
     AmtIn# = VAL(in$)
     COLOR 11: LOCATE 12, 50: PRINT "Change Due: "; USING "######,#.##"; AmtIn# - AmountPaid#


pr3: LOCATE 13, 1: COLOR 11
     PRINT "Enter Description (if any) .. "; : fl = 15
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RETURN
     Desc$ = UCASE$(in$)
pr4: LOCATE 14, 1: COLOR 11
     PRINT "(A)uto or (M)anually Dist. Payment . "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN CLOSE : RETURN
     IF NOT (in$ = "A" OR in$ = "M") THEN GOTO pr4
     Dist$ = in$
pr5: LOCATE 15, 1: COLOR 11
     PRINT "Enter Tax Year of Bill Being Paid .. "; : fl = -4
     GOSUB 62000
     IF LEN(in$) <> 4 THEN GOTO pr5
     IF VAL(in$) > VAL(CurrentYear$) THEN
      LOCATE 16, 1: COLOR 11
      PRINT CHR$(7);
      PRINT "Year Entered Exceeds Current Year : Press <ENTER> to Try Again!"; : fl = 1
      GOSUB 62000
      LOCATE 16, 1: PRINT STRING$(79, 32)
      GOTO pr5
     END IF
     ManYear$ = in$
     TaxYear$ = in$

pr6: LOCATE 16, 1: COLOR 11
     PRINT "Payment Entered Correctly? (Y/N) .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN CLOSE : RETURN
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO pr6
     IF in$ = "N" THEN GOTO ProcessReal



      GOSUB FilePayment
      CLOSE : RETURN



ProcessPersonal:
    GOSUB FigureBalance
    PayType = 1 'Regular Payment
    PaySource$ = "P"
    CLS
    scren$ = "4.01.00": CLS : GOSUB 63500
    COLOR 11: LOCATE 1, 32: PRINT "Customer Payment Entry";
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
    PRINT "Correct Account (Y/N) .. "; : fl = 1: GOSUB 62210
    IF cf = 1 THEN CLOSE : RETURN
    IF in$ = "N" THEN CLOSE : RETURN


    LOCATE 10, 1: PRINT STRING$(79, 32): COLOR 11
    LOCATE 10, 1
    PRINT "Enter Amount to Apply on Account ("; USING "######,#.##"; CurBal# + PastDue#; : PRINT ") .... ";
     PRINT TAB(54); ""; : fl = -9: GOSUB 62000
    IF cf = 1 THEN CLOSE : RETURN
    IF LEN(in$) = 0 THEN
      AmountPaid# = CurBal# + PastDue#
      ELSE
      AmountPaid# = VAL(in$)
    END IF
    LOCATE 10, 54: COLOR 14: PRINT USING "######,#.##"; AmountPaid#

ppr1: COLOR 11: LOCATE 11, 1
     PRINT "Enter (1)Cash or (2)Check ... "; : fl = -1: GOSUB 62000
     IF VAL(in$) < 1 OR VAL(in$) > 2 THEN GOTO ppr1
     curency = VAL(in$)
ppr2: COLOR 11: LOCATE 12, 1
     PRINT "Enter Amount from Customer .. "; : fl = -9: GOSUB 62000
     LOCATE 12, 30: COLOR 14
     PRINT USING "#####,#.##"; VAL(in$)
     AmtIn# = VAL(in$)
     COLOR 11: LOCATE 12, 50: PRINT "Change Due: "; USING "######,#.##"; AmtIn# - AmountPaid#


ppr3: LOCATE 13, 1: COLOR 11
     PRINT "Enter Description (if any) .. "; : fl = 15
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RETURN
     Desc$ = UCASE$(in$)
ppr4: LOCATE 14, 1: COLOR 11
     PRINT "(A)uto or (M)anually Dist. Payment . "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN CLOSE : RETURN
     IF NOT (in$ = "A" OR in$ = "M") THEN GOTO ppr4
     Dist$ = in$
ppr5: LOCATE 15, 1: COLOR 11
     PRINT "Enter Tax Year of Bill Being Paid .. "; : fl = -4
     GOSUB 62000
     IF LEN(in$) <> 4 THEN GOTO pr5
     IF VAL(in$) > VAL(CurrentYear$) THEN
      LOCATE 16, 1: COLOR 11
      PRINT CHR$(7);
      PRINT "Year Entered Exceeds Current Year : Press <ENTER> to Try Again!"; : fl = 1
      GOSUB 62000
      LOCATE 16, 1: PRINT STRING$(79, 32)
      GOTO ppr5
     END IF
     ManYear$ = in$
     TaxYear$ = in$

ppr6: LOCATE 16, 1: COLOR 11
     PRINT "Payment Entered Correctly? (Y/N) .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN CLOSE : RETURN
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO pr6
     IF in$ = "N" THEN GOTO ProcessPersonal



      GOSUB FilePayment
      CLOSE : RETURN

FilePayment:
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
     GOSUB 60000: GOSUB 60025


     IF PayEdit(1).PaymentSource = "R" THEN
      IF PayEdit(1).PaymentAcct > 0 THEN
        GET TBFile, PayEdit(1).PaymentAcct, TBRECust(1)
        GET TBBalFile, PayEdit(1).PaymentAcct, TBRECustBal(1)
        GOSUB FigureBalance
        GOSUB PostPayment
        PUT TBFile, PayEdit(1).PaymentAcct, TBRECust(1)
        PUT TBBalFile, PayEdit(1).PaymentAcct, TBRECustBal(1)
        GOSUB FigureBalance
        GOSUB PostRealHistory
       END IF
     END IF

     IF PayEdit(1).PaymentSource = "P" THEN
      IF PayEdit(1).PaymentAcct > 0 THEN
       GET TBFile, PayEdit(1).PaymentAcct, TBPPCust(1)
       GET TBBalFile, PayEdit(1).PaymentAcct, TBPPCustBal(1)
       GOSUB FigureBalance
       GOSUB PostPayment
       PUT TBFile, PayEdit(1).PaymentAcct, TBPPCust(1)
       PUT TBBalFile, PayEdit(1).PaymentAcct, TBPPCustBal(1)
       GOSUB FigureBalance
       GOSUB PostPersHistory
      END IF
     END IF

   TBRFile = FREEFILE
   OPEN TBPath$ + "TBRECFIL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBRFile LEN = LEN(TBRECPT(1))
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
    TBRECPT(1).TranType = 1
    ELSE
    TBRECPT(1).TranType = 2
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
   ReceiptNumber& = (LOF(TBRFile) / LEN(TBRECPT(1))) + 1
   PUT TBRFile, ReceiptNumber&, TBRECPT(1)
   CLOSE TBRFile
    IF AmtIn# - AmountPaid# > 0 THEN
    LOCATE 21, 1
    COLOR 15
    PRINT CHR$(7);
    PRINT "CHANGE DUE CUSTOMER .......... "; USING "#####,#.##"; AmtIn# - AmountPaid#
    END IF


RECPT2a:
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
     PRINT "PRESS <ENTER> TO CONTINUE "; : fl = 1: GOSUB 62000
     CLOSE
     RETURN

PostRealHistory:
 GET TBFile, PayEdit(1).PaymentAcct, TBRECust(1)
 GET TBBalFile, PayEdit(1).PaymentAcct, TBRECustBal(1)
 GOSUB FigureBalance
 AcctBalance# = CurBal# + PastDue#
 GOSUB 61000
 TBTrans(1).TransDate = Date2Num%(PayEdit(1).PaymentDate)
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


PostPersHistory:
 GET TBFile, PayEdit(1).PaymentAcct, TBPPCust(1)
 GET TBBalFile, PayEdit(1).PaymentAcct, TBPPCustBal(1)
 GOSUB FigureBalance
 AcctBalance# = CurBal# + PastDue#
 GOSUB 61000
 TBTrans(1).TransDate = Date2Num%(PayEdit(1).PaymentDate)
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

PostPayment:
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

60000  IF taxtype$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TBRTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBPath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN

60025 TBBalFile = FREEFILE
      IF taxtype$ = "R" THEN
      OPEN TBPath$ + "TBRTBal.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBPath$ + "TBPTBAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF
60030 RETURN

61000 TBTranFile = FREEFILE
      OPEN TBPath$ + "TBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTranFile LEN = LEN(TBTrans(1))
61020 RETURN

62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(z$) < 2 THEN 62065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN cf = 12: RETURN
62065 IF ASC(z$) = 0 THEN 62030
62070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
62080 IF z$ = " " THEN 62130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 62000
62090 IF z$ <> CHR$(8) THEN 62130 ELSE IF ZL = 0 THEN 62030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 62110 ELSE IF z$ = "," THEN 62120
62100 IF z$ = "." THEN ZD = 0: GOTO 62110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
62110 in$ = LEFT$(in$, LEN(in$) - 1)
62120 ZL = ZL - 1: GOTO 62030
62130 IF ABS(fl) = ZL THEN 62030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 62180
62140 IF z$ = "." AND INSTR(in$, ".") = 0 THEN GOTO 62180
62150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF z$ < "0" OR z$ > "9" THEN 62030
62180 PRINT z$; : in$ = in$ + z$: ZL = ZL + 1
62190 IF ABS(fl) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT in$: RETURN
62210 GOSUB 62000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
62220 RETURN

63500 COLOR 10: LOCATE 1, 1
     PRINT "Post Date: "; : COLOR 15
     PRINT LEFT$(PDATE$, 2) + "-" + MID$(PDATE$, 3, 2) + "-" + MID$(PDATE$, 5, 2);
     COLOR 10
     PRINT TAB(70); scren$
     PRINT STRING$(79, "_")
     RETURN

Namesearch:
 LOCATE 8, 1:
 COLOR 11
 PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
 GOSUB 62000
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
  OPEN TBPath$ + "TBRTNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS Mrec$
  GOSUB 60000
  FOR LL = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, LL
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO n1
n3:
  NEXT LL
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  ggo = 0
  RETURN

n1:
GET TBFile, CVS(Mrec$), TBRECust(1)
LOCATE 11, 1: COLOR 11
NME$ = TBRECust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT NME$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(Mrec$)
n2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO n2
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
IF in$ = "N" THEN GOTO n3
GOTO n2
END IF


  TBIDXFile = FREEFILE
  OPEN TBPath$ + "TBPTNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS Mrec$
  GOSUB 60000
  FOR LL = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, LL
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO PPn1
PPn3:
  NEXT LL
  LOCATE 8, 1: PRINT STRING$(79, 32)
  ggo = 0
  RETURN

PPn1:
GET TBFile, CVS(Mrec$), TBPPCust(1)
LOCATE 11, 1: COLOR 11
NME$ = TBPPCust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT NME$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(Mrec$)
PPn2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO PPn2
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
IF in$ = "N" THEN GOTO PPn3
GOTO PPn2

Deleted:
    LOCATE 23, 1: COLOR 15
    PRINT "This Account Has Been Deleted! Press <ENTER> to Continue"; :
    fl = 1
    GOSUB 62000
    LOCATE 23, 1: PRINT STRING$(79, 32)
    RETURN

FigureBalance:
     CurBal# = 0: PastDue# = 0
     IF taxtype$ = "R" THEN
     CurBal# = TBRECust(1).CurYrTaxAmount + TBRECust(1).CurYrIntAmount + TBRECust(1).CurYrPenAmount + TBRECust(1).CurYrOthAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR LL = 1 TO 10
     PastDue# = PastDue# + TBRECustBal(1).TaxAmt(LL) + TBRECustBal(1).IntAmt(LL) + TBRECustBal(1).PenAmt(LL) + TBRECustBal(1).OthAmt(LL) + TBRECustBal(1).LateAmt(LL)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT LL
     ELSE
     CurBal# = TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax + TBPPCust(1).CurYrIntAmount + TBPPCust(1). _
CurYrPenAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR LL = 1 TO 10
     PastDue# = PastDue# + TBPPCustBal(1).PerPropTaxAmt(LL) + TBPPCustBal(1).FarmEquipAmt(LL) + TBPPCustBal(1).MachToolsAmt(LL) + TBPPCustBal(1).MerchCapAmt(LL) + TBPPCustBal(1).MobileHomeAmt(LL) + TBPPCustBal(1).LateFeeAmt(LL) + TBPPCustBal(1). _
IntAmt(LL) + TBPPCustBal(1).PenAmt(LL)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT LL
     END IF
     RETURN


SetPostDate:
      Cleardown
      LOCATE 4, 1
      COLOR 11
      PRINT "Enter the Posting Date Now (mmddyy) .. "; : fl = -6
      GOSUB 62000
      IF LEN(in$) <> 6 THEN GOTO SetPostDate
      Month$ = LEFT$(in$, 2)
      Day$ = MID$(in$, 3, 2)
      IF VAL(Month$) < 1 OR VAL(Month$) > 12 THEN
       PRINT CHR$(7);
       PRINT "Invalid Month!!!!"
       SLEEP 3
       GOTO SetPostDate
      END IF
      IF VAL(Day$) < 1 OR VAL(Day$) > 31 THEN
       PRINT CHR$(7);
       PRINT "Invalid Day!!!!"
       SLEEP 3
       GOTO SetPostDate
      END IF
      PDATE$ = in$
      RETURN
misccodefile:
 CLOSE 15
 OPEN TBPath$ + "tbmisc.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #15 LEN = 256
 FIELD 15, 30 AS CODEDESC$, 1 AS Active$, 225 AS nu$
 RETURN

EnterMiscPay:
     
1800 Cleardown
     COLOR 11
     LOCATE 3, 1: PRINT "Enter Payee Name (ESC-MENU).. "; : fl = 30: GOSUB 62000
     IF cf = 1 THEN 1000 ELSE NME$ = in$
     COLOR 11: LOCATE 4, 1: PRINT "Enter Address ............... "; : fl = 27: GOSUB 62000: Desc$ = in$
     COLOR 11: LOCATE 5, 1: PRINT "Enter Amount Owed ........... "; : fl = -10: GOSUB 62000: Owed# = VAL(in$)
     COLOR 11: LOCATE 6, 1: PRINT "Enter Amount Paid ........... "; : fl = -10: GOSUB 62000: amount# = VAL(in$)
     COLOR 11: LOCATE 6, 50: PRINT "Change: "; : COLOR 15: PRINT USING "####.##"; amount# - Owed#
1805 COLOR 11: LOCATE 7, 1: PRINT "Enter (1)Cash or (2)Check ... "; : fl = -1: GOSUB 62000
     IF VAL(in$) < 1 OR VAL(in$) > 2 THEN 1805
     curency = VAL(in$)
1808 COLOR 11: LOCATE 8, 1: PRINT "Enter Misc Source Code ...... "; : fl = -2: GOSUB 62000
     GOSUB misccodefile
     IF VAL(in$) < 1 OR VAL(in$) > LOF(15) / 256 THEN 1808
     mpcode = VAL(in$)
     GET 15, mpcode
     LOCATE 8, 45
      COLOR 15
      PRINT CODEDESC$
      Catagory$ = CODEDESC$
1810 COLOR 11: LOCATE 11, 1: PRINT "Is the Above Correct (Y/N) .. "; : fl = 1: GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN 1810 ELSE LOCATE 11, 31: PRINT in$
     IF in$ = "N" THEN FOR J = 3 TO 11: LOCATE J, 1: PRINT STRING$(79, 32): NEXT J: GOTO 1800

   TBRFile = FREEFILE
   OPEN TBPath$ + "TBRECFIL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBRFile LEN = LEN(TBRECPT(1))
   TBRECPT(1).Date = Date2Num%(PostDate$)
   TBRECPT(1).AmtIn = Paid#
   TBRECPT(1).AmtOwed = Owed#
   IF curency = 1 THEN
    TBRECPT(1).CashAmt = Paid#
    TBRECPT(1).ChkAmt = 0
   END IF
   IF curency = 2 THEN
    TBRECPT(1).CashAmt = 0
    TBRECPT(1).ChkAmt = Paid#
   END IF

   TBRECPT(1).TranType = 6
   TBRECPT(1).TranCode = mpcode
   TBRECPT(1).TaxYear = -32767
   TBRECPT(1).Desc = Desc$
   TBRECPT(1).RealTxPd = 0
   TBRECPT(1).PersTxPd = 0
   TBRECPT(1).MTTxPd = 0
   TBRECPT(1).MCTxPd = 0
   TBRECPT(1).MHTxPd = 0
   TBRECPT(1).FETxPd = 0
   TBRECPT(1).IntTxPd = 0
   TBRECPT(1).PenTxPd = 0
   TBRECPT(1).Fudge = ""
   ReceiptNumber& = (LOF(TBRFile) / LEN(TBRECPT(1))) + 1
   PUT TBRFile, ReceiptNumber&, TBRECPT(1)
   CLOSE TBRFile

FOR LL = 12 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
COLOR 11
LOCATE 15, 1: PRINT "Amount Paid on Account ....... "; USING "#####,#.##"; Owed#
              PRINT "Cash Received from Customer .. "; USING "#####,#.##"; amount#
              PRINT "                               -------------"
              COLOR 15
              PRINT "CHANGE DUE CUSTOMER .......... "; USING "#####,#.##"; amount# - Owed#

RECPT2:
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
     LPRINT "    "; Catagory$
     LPRINT "        "; PostDate$
     LPRINT
     LPRINT "Customer: "; LEFT$(NME$, 18)
     LPRINT ".."; Desc$
     LPRINT
     LPRINT "Bill Amount: "; USING "$####,#.##"; Owed#
     LPRINT "Amount Paid: "; USING "$####,#.##"; amount#
     LPRINT " Change Due: "; USING "$####,#.##"; amount# - Owed#
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
     PRINT "PRESS <ENTER> TO CONTINUE "; : fl = 1: GOSUB 62000
     CLOSE
     RETURN




networkerror:
     REM ERROR CODES
     Number = ERR
     SELECT CASE ERR
          CASE IS = 24, 25
               LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 28
               PRINT "Error Message:"; : COLOR 15: PRINT "Device Time-out! Make Sure Printer Is ON-LINE:Press ENTER to Resume Printing";
               fl = 1: GOSUB 62000: LOCATE 23, 1: PRINT STRING$(79, 32)
               RESUME
          CASE IS = 27
               LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 28
               PRINT "Error Message:"; : COLOR 15: PRINT "Printer Is OUT OF PAPER! Load Paper : Press ENTER to Resume Printing";
               fl = 1: GOSUB 62000: LOCATE 23, 1: PRINT STRING$(79, 32)
               RESUME

          CASE ELSE
               CLS
               LOCATE 15, 1: COLOR 28: PRINT "Error Message: "; : COLOR 15: PRINT "Error "; Number; " Has Occured in Line "; ERL
               PRINT "Write Down Message and Hit <ESC> Key to exit."
               fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
            
     END SELECT

SUB Cleardown
   FOR LL = 3 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
END SUB

SUB DecalPayment (PDATE$, TBPath$)
     PostDate$ = LEFT$(PDATE$, 2) + "-" + MID$(PDATE$, 3, 2) + "-" + RIGHT$(PDATE$, 2)
     CLOSE : OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS Year$: GET 1, 1
     CurrentYear$ = Year$: CLOSE
     CurrentYear = VAL(CurrentYear$)
     
10250 GOSUB 63400: COLOR 11
      LOCATE 1, 30: PRINT "Decal Payment Entry"
      Cleardown
      LOCATE 3, 1: COLOR 11
      taxtype$ = "P"
10300 LOCATE 4, 1: COLOR 11
     PRINT "Enter Customer Account Number (TAB-Last Name : ESC-Menu) .. "; : fl = -6
     GOSUB 58000
     IF cf = 1 THEN
      CLOSE : EXIT SUB
     END IF
     IF cf = 2 THEN
      GOSUB NamesearchDecal
      IF ggo = 1 THEN 10750
      IF ggo = 0 THEN LOCATE 10, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 58000: GOTO 10250
     END IF
     Record! = VAL(in$)
10750 CLOSE
      GOSUB 59000
      GOSUB 59025
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBPPCust(1)) THEN 10300
       LOCK #TBFile, Record!
       GET TBFile, Record!, TBPPCust(1)
       GET TBBalFile, Record!, TBPPCustBal(1)
       IF TBPPCust(1).Deleted = "Y" THEN GOSUB DeletedDecal: GOTO 10250
       GOSUB ProcessPersonalDecal
       CLOSE
       GOTO 10250

ProcessPersonalDecal:
    GOSUB FigureBalanceDecal
    PayType = 5 'Regular Payment
    PaySource$ = "P"
    CLS
    scren$ = "4.01.00": CLS :
    COLOR 11: LOCATE 1, 32: PRINT "Customer Payment Entry";
    LOCATE 3, 1: PRINT "Personal Property Customer"
    LOCATE 5, 1: COLOR 15
        NME$ = TBPPCust(1).CustName
        PRINT TBPPCust(1).CustName; TAB(50); "Current Bal: "; USING "$$######,#.##"; CurBal#
        PRINT TBPPCust(1).Address1; TAB(50); "   Past Bal: "; USING "$$######,#.##"; PastDue#
        PRINT TBPPCust(1).Address2; TAB(50); " Total Owed: "; USING "$$######,#.##"; CurBal# + PastDue#
        PRINT RTRIM$(TBPPCust(1).City); ", "; TBPPCust(1).State; " "; TBPPCust(1).Zip;

        IF RTRIM$(TBPPCust(1).DecalBought) = "Y" THEN
          COLOR 14
         PRINT TAB(50); "Decal "; TBPPCust(1).DecalNumber; " Purchased"
         COLOR 11
         ELSE
         PRINT
        END IF
        Owed# = CurBal# + PastDue#
    LOCATE 9, 1: COLOR 11: PRINT STRING$(80, "-")
    LOCATE 10, 1: COLOR 11
    PRINT "Correct Account (Y/N) .. "; : fl = 1: GOSUB 58210
    IF cf = 1 THEN CLOSE : RETURN
    IF in$ = "N" THEN CLOSE : RETURN


    LOCATE 10, 1: PRINT STRING$(79, 32): COLOR 11
    LOCATE 10, 1
    PRINT "Enter Amount Owed for Decal .... ";
      fl = -9: GOSUB 58000
     IF cf = 1 THEN CLOSE : RETURN
     AmountPaid# = VAL(in$)
     LOCATE 10, 30: COLOR 14: PRINT USING "######,#.##"; AmountPaid#

pppr1: COLOR 11: LOCATE 11, 1
     PRINT "Enter (1)Cash or (2)Check ... "; : fl = -1: GOSUB 58000
     IF VAL(in$) < 1 OR VAL(in$) > 2 THEN GOTO pppr1
     curency = VAL(in$)
     COLOR 11: LOCATE 12, 1
      PRINT "Enter Amount from Customer .. "; : fl = -9: GOSUB 58000
     LOCATE 12, 30: COLOR 14
     PRINT USING "#####,#.##"; VAL(in$)
     AmtIn# = VAL(in$)
     COLOR 11: LOCATE 12, 50: PRINT "Change Due: "; USING "######,#.##"; AmtIn# - AmountPaid#


     LOCATE 13, 1: COLOR 11
     PRINT "Enter Decal Number .......... "; : fl = 8
     GOSUB 58000
     IF cf = 1 THEN CLOSE : RETURN
     Desc$ = UCASE$(in$)
     LOCATE 14, 1: COLOR 11
     PRINT "Enter Check Number (if any).. "; : fl = 15
     GOSUB 58000
     IF cf = 1 THEN CLOSE : RETURN
     Desc1$ = UCASE$(in$)


pppr6: LOCATE 16, 1: COLOR 11
     PRINT "Payment Entered Correctly? (Y/N) .. "; : fl = 1
     GOSUB 58210
     IF cf = 1 THEN CLOSE : RETURN
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO pppr6
     IF in$ = "N" THEN GOTO ProcessPersonalDecal
     GOSUB FilePaymentDecal
     CLOSE
     RETURN

FilePaymentDecal:
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
     GOSUB 59000
     GOSUB 59025

     
      IF PayEdit(1).PaymentAcct > 0 THEN
       GET TBFile, PayEdit(1).PaymentAcct, TBPPCust(1)
       GET TBBalFile, PayEdit(1).PaymentAcct, TBPPCustBal(1)
        TBPPCust(1).DecalBought = "Y"
        TBPPCust(1).DecalNumber = Desc$
       PUT TBFile, PayEdit(1).PaymentAcct, TBPPCust(1)
       PUT TBBalFile, PayEdit(1).PaymentAcct, TBPPCustBal(1)
      END IF
     

   TBRFile = FREEFILE
   OPEN TBPath$ + "TBRECFIL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBRFile LEN = LEN(TBRECPT(1))
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
   TBRECPT(1).TranType = 5
   TBRECPT(1).TranCode = 0
   TBRECPT(1).TaxYear = VAL(TaxYear$)
   TBRECPT(1).Desc = NME$
   TBRECPT(1).RealTxPd = 0
   TBRECPT(1).PersTxPd = 0
   TBRECPT(1).MTTxPd = 0
   TBRECPT(1).MCTxPd = 0
   TBRECPT(1).MHTxPd = 0
   TBRECPT(1).FETxPd = 0
   TBRECPT(1).IntTxPd = 0
   TBRECPT(1).PenTxPd = 0
   TBRECPT(1).DecalNumb = Desc$
   TBRECPT(1).Fudge = ""
   ReceiptNumber& = (LOF(TBRFile) / LEN(TBRECPT(1))) + 1
   PUT TBRFile, ReceiptNumber&, TBRECPT(1)
   CLOSE TBRFile
   IF AmtIn# - AmountPaid# > 0 THEN
   LOCATE 21, 1
   COLOR 15
   PRINT CHR$(7);
   PRINT "CHANGE DUE CUSTOMER .......... "; USING "#####,#.##"; AmtIn# - AmountPaid#
   END IF


RECPT2aDecal:
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
     LPRINT "           DECAL"
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
     PRINT "PRESS <ENTER> TO CONTINUE "; : fl = 1: GOSUB 58000
     CLOSE
     RETURN

58000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
58010 IF INKEY$ <> "" THEN 58010
58020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
58030 z$ = INKEY$: IF z$ = "" THEN 58030
58040 IF z$ <> CHR$(13) THEN 58059 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
58050 GOSUB 58200: RETURN
58059 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(z$) < 2 THEN 58065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN cf = 12: RETURN
58065 IF ASC(z$) = 0 THEN 58030
58070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
58080 IF z$ = " " THEN 58130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 58000
58090 IF z$ <> CHR$(8) THEN 58130 ELSE IF ZL = 0 THEN 58030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 58110 ELSE IF z$ = "," THEN 58120
58100 IF z$ = "." THEN ZD = 0: GOTO 58110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
58110 in$ = LEFT$(in$, LEN(in$) - 1)
58120 ZL = ZL - 1: GOTO 58030
58130 IF ABS(fl) = ZL THEN 58030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 58180
58140 IF z$ = "." AND INSTR(in$, ".") = 0 THEN GOTO 58180
58150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 58190
58160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 58180
58170 IF z$ < "0" OR z$ > "9" THEN 58030
58180 PRINT z$; : in$ = in$ + z$: ZL = ZL + 1
58190 IF ABS(fl) = 1 THEN 58050 ELSE 58030
58200 LOCATE RO, PO: PRINT in$: RETURN
58210 GOSUB 58000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
58220 RETURN


59000  IF taxtype$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TBRTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBPath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN

59025 TBBalFile = FREEFILE
      IF taxtype$ = "R" THEN
      OPEN TBPath$ + "TBRTBal.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBPath$ + "TBPTBAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF
59030 RETURN


NamesearchDecal:
 LOCATE 8, 1:
 COLOR 11
 PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
 GOSUB 58000
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
  OPEN TBPath$ + "TBRTNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS Mrec$
  GOSUB 59000
  FOR LL = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, LL
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO Decaln1
Decaln3:
  NEXT LL
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  ggo = 0
  RETURN

Decaln1:
GET TBFile, CVS(Mrec$), TBRECust(1)
LOCATE 11, 1: COLOR 11
NME$ = TBRECust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT NME$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(Mrec$)
Decaln2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 58210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO Decaln2
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
IF in$ = "N" THEN GOTO Decaln3
GOTO Decaln2
END IF


  TBIDXFile = FREEFILE
  OPEN TBPath$ + "TBPTNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS Mrec$
  GOSUB 59000
  FOR LL = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, LL
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO DecalPPn1
DecalPPn3:
  NEXT LL
  LOCATE 8, 1: PRINT STRING$(79, 32)
  ggo = 0
  RETURN

DecalPPn1:
GET TBFile, CVS(Mrec$), TBPPCust(1)
LOCATE 11, 1: COLOR 11
NME$ = TBPPCust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT NME$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(Mrec$)
DecalPPn2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 58210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO DecalPPn2
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
IF in$ = "N" THEN GOTO DecalPPn3
GOTO DecalPPn2

DeletedDecal:
    LOCATE 23, 1: COLOR 15
    PRINT "This Account Has Been Deleted! Press <ENTER> to Continue"; :
    fl = 1
    GOSUB 58000
    LOCATE 23, 1: PRINT STRING$(79, 32)
    RETURN

FigureBalanceDecal:
     CurBal# = 0: PastDue# = 0
     IF taxtype$ = "R" THEN
     CurBal# = TBRECust(1).CurYrTaxAmount + TBRECust(1).CurYrIntAmount + TBRECust(1).CurYrPenAmount + TBRECust(1).CurYrOthAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR LL = 1 TO 10
     PastDue# = PastDue# + TBRECustBal(1).TaxAmt(LL) + TBRECustBal(1).IntAmt(LL) + TBRECustBal(1).PenAmt(LL) + TBRECustBal(1).OthAmt(LL) + TBRECustBal(1).LateAmt(LL)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT LL
     ELSE
     CurBal# = TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax + TBPPCust(1).CurYrIntAmount + TBPPCust(1). _
CurYrPenAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR LL = 1 TO 10
     PastDue# = PastDue# + TBPPCustBal(1).PerPropTaxAmt(LL) + TBPPCustBal(1).FarmEquipAmt(LL) + TBPPCustBal(1).MachToolsAmt(LL) + TBPPCustBal(1).MerchCapAmt(LL) + TBPPCustBal(1).MobileHomeAmt(LL) + TBPPCustBal(1).LateFeeAmt(LL) + TBPPCustBal(1). _
IntAmt(LL) + TBPPCustBal(1).PenAmt(LL)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT LL
     END IF
     RETURN
63400 COLOR 10: LOCATE 1, 1
     PRINT "Post Date: "; : COLOR 15
     PRINT LEFT$(PDATE$, 2) + "-" + MID$(PDATE$, 3, 2) + "-" + MID$(PDATE$, 5, 2);
     COLOR 10
     PRINT TAB(70); scren$
     PRINT STRING$(79, "_")
     RETURN

END SUB

SUB MiscCodes (PDATE$, TBPath$)
8000 scren$ = "4.06.00": CLS : GOSUB 44500
     LOCATE 1, 30: COLOR 11: PRINT "Misc Code Maintenace"
     COLOR 11
     LOCATE 3, 1: PRINT "(A)dd : (C)hange or (P)rint Codes (ESC-Menu) .. "; : fl = 1: GOSUB 44210
     IF cf = 1 THEN EXIT SUB
     IF NOT (in$ = "A" OR in$ = "C" OR in$ = "P") THEN 8000
     GOSUB misccodefile1
     IF in$ = "A" THEN GOSUB 8001
     IF in$ = "C" THEN GOSUB 8010
     IF in$ = "P" THEN GOSUB 8500
     GOTO 8000

8001 Cleardown
     LOCATE 3, 1: COLOR 11
     NextCodeRecord = (LOF(15) \ 256) + 1
     PRINT "Next Code Number is "; NextCodeRecord
     PRINT "Enter the Description for This Code.. "; : fl = 30
     GOSUB 44000
     PRINT "Code "; NextCodeRecord; " is Set to "; in$
     LSET CODEDESC$ = UCASE$(in$)
     LSET Active$ = "Y"
     PUT 15, NextCodeRecord
     CLOSE
     GOTO 8000
8010 Cleardown
     LOCATE 3, 1: COLOR 11: PRINT STRING$(79, 32): LOCATE 3, 1
     PRINT "Enter Code Number <ESC>-Menu .. "; : fl = -2
     GOSUB 44000: IF cf = 1 THEN 8000
       xx = VAL(in$)
     IF xx < 1 OR xx > LOF(15) / 256 THEN
      PRINT "Invalid Code !!!"
      SLEEP 3
      GOTO 8010
     END IF

     GET 15, xx
     LOCATE 6, 1
     COLOR 11
     PRINT "Enter Code Desc.  .. "; : fl = 30
     GOSUB 44000: LSET CODEDESC$ = in$: LSET Active$ = "Y": PUT 15, xx
     PRINT "Updated"
     PRINT CHR$(7);
     SLEEP 3
     LOCATE 6, 1: PRINT STRING$(79, 32): LOCATE 7, 1: PRINT STRING$(79, 32): LOCATE 8, 1: PRINT STRING$(79, 32)
     GOTO 8010

8500 REM print them
     GOSUB misccodefile1
     IF LOF(15) / 256 > 0 THEN
      LPRINT "Lunenburg County"; TAB(60); DATE$
      LPRINT "Misc Code File Descriptions"
      LPRINT
      ELSE
      CLOSE
      PRINT "No Codes On File !!!"
      SLEEP 3
      GOTO 8000
     END IF
     FOR x = 1 TO LOF(15) / 256
     GET 15, x
     IF Active$ = "Y" THEN
      LPRINT "Code # "; x; " Set To "; CODEDESC$
     END IF
     NEXT x
     LPRINT CHR$(12);
     CLOSE
     RETURN

misccodefile1:
 CLOSE 15
 OPEN TBPath$ + "tbmisc.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #15 LEN = 256
 FIELD 15, 30 AS CODEDESC$, 1 AS Active$, 225 AS nu$
 RETURN

44000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
44010 IF INKEY$ <> "" THEN 44010
44020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
44030 z$ = INKEY$: IF z$ = "" THEN 44030
44040 IF z$ <> CHR$(13) THEN 44060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
44050 GOSUB 44200: RETURN
44060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(z$) < 2 THEN 44065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN cf = 12: RETURN
44065 IF ASC(z$) = 0 THEN 44030
44070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
44080 IF z$ = " " THEN 44130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 44000
44090 IF z$ <> CHR$(8) THEN 44130 ELSE IF ZL = 0 THEN 44030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 44110 ELSE IF z$ = "," THEN 44120
44100 IF z$ = "." THEN ZD = 0: GOTO 44110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
44110 in$ = LEFT$(in$, LEN(in$) - 1)
44120 ZL = ZL - 1: GOTO 44030
44130 IF ABS(fl) = ZL THEN 44030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 44180
44140 IF z$ = "." AND INSTR(in$, ".") = 0 THEN GOTO 44180
44150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 44190
44160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 44180
44170 IF z$ < "0" OR z$ > "9" THEN 44030
44180 PRINT z$; : in$ = in$ + z$: ZL = ZL + 1
44190 IF ABS(fl) = 1 THEN 44050 ELSE 44030
44200 LOCATE RO, PO: PRINT in$: RETURN
44210 GOSUB 44000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
44220 RETURN

44500 COLOR 10: LOCATE 1, 1
     PRINT "Post Date: "; : COLOR 15
     PRINT LEFT$(PDATE$, 2) + "-" + MID$(PDATE$, 3, 2) + "-" + MID$(PDATE$, 5, 2);
     COLOR 10
     PRINT TAB(70); scren$
     PRINT STRING$(79, "_")
     RETURN

END SUB

SUB PrintReceiptReport (PDATE$, TBPath$)


MainBody:
   CLS
   scren$ = "8.05.00": CLS : GOSUB Heading
   COLOR 11
   LOCATE 1, 30: PRINT "Cash Receipts Report"
   LOCATE 3, 1
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
    PRINT #10, TAB(28); USING "######.##"; TBRECPT(1).CashAmt;
    PRINT #10, TAB(38); USING "######.##"; TBRECPT(1).ChkAmt;
    PRINT #10, TAB(48); USING "######.##"; TBRECPT(1).AmtIn;
    PRINT #10, TAB(58); USING "######.##"; TBRECPT(1).AmtOwed;
    PRINT #10, TAB(68); USING "######.##"; TBRECPT(1).AmtIn - TBRECPT(1).AmtOwed
    lc = lc + 1
    IF TBRECPT(1).TranType = 1 THEN
     PRINT #10, RTRIM$(TBRECPT(1).Desc);
     PRINT #10, "| Tax Pd:"; USING "######.##"; TBRECPT(1).RealTxPd;
     PRINT #10, " Pen Pd:"; USING "####.##"; TBRECPT(1).PenTxPd;
     PRINT #10, " Int Pd:"; USING "####.##"; TBRECPT(1).IntTxPd;
     PRINT #10, " Yr:"; TBRECPT(1).TaxYear
     PRINT #10, STRING$(79, "-")
     lc = lc + 2
    END IF
    IF TBRECPT(1).TranType = 2 THEN
     PRINT #10, RTRIM$(TBRECPT(1).Desc);
     PRINT #10, "| Tax Pd:"; USING "######.##"; TBRECPT(1).PersTxPd;
     PRINT #10, " Pen Pd:"; USING "####.##"; TBRECPT(1).PenTxPd;
     PRINT #10, " Int Pd:"; USING "####.##"; TBRECPT(1).IntTxPd;
     PRINT #10, " Yr:"; TBRECPT(1).TaxYear
     PRINT #10, STRING$(79, "-")
     lc = lc + 2
    END IF
    IF TBRECPT(1).TranType = 3 THEN
     PRINT #10, RTRIM$(TBRECPT(1).Desc);
     PRINT #10, "| Tax Pd:"; USING "######.##"; TBRECPT(1).RealTxPd;
     PRINT #10, " Pen Pd:"; USING "####.##"; TBRECPT(1).PenTxPd;
     PRINT #10, " Int Pd:"; USING "####.##"; TBRECPT(1).IntTxPd;
     PRINT #10, " Yr:"; TBRECPT(1).TaxYear
     PRINT #10, STRING$(79, "-")
     lc = lc + 2
    END IF
    IF TBRECPT(1).TranType = 4 THEN
     PRINT #10, RTRIM$(TBRECPT(1).Desc);
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
     PRINT #10, "Misc Code # "; TBRECPT(1).TranCode
     PRINT #10, STRING$(79, "-")
     lc = lc + 2
    END IF

   END IF
   NEXT Cnt&
   CLOSE




FinishReport:

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
   PRINT #10, ""
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

SUB SetMessage
   MFile = FREEFILE
   OPEN TBPath$ + "TBMESSAG.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #MFile LEN = 256
   FIELD MFile, 30 AS Message$
   Cleardown
   COLOR 11
   LOCATE 3, 1
   IF LOF(MFile) > 0 THEN
    GET MFile, 1
    PRINT "Message Currently Set to..."
    COLOR 15
    PRINT Message$
   END IF
   LOCATE 7, 1
   COLOR 11
   PRINT "Enter Message Line to Print at Bottom of Receipt or <ESC>-Cancel"
   PRINT "Enter Message Here: "; : fl = 30
   GOSUB 63000
   IF cf = 1 THEN CLOSE : EXIT SUB
   LSET Message$ = in$
   PUT MFile, 1
   CLOSE
   EXIT SUB

63000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
63010 IF INKEY$ <> "" THEN 63010
63020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
63030 z$ = INKEY$: IF z$ = "" THEN 63030
63040 IF z$ <> CHR$(13) THEN 63060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
63050 GOSUB 63200: RETURN
63060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(z$) < 2 THEN 63065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN cf = 12: RETURN
63065 IF ASC(z$) = 0 THEN 63030
63070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
63080 IF z$ = " " THEN 63130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 63000
63090 IF z$ <> CHR$(8) THEN 63130 ELSE IF ZL = 0 THEN 63030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 63110 ELSE IF z$ = "," THEN 63120
63100 IF z$ = "." THEN ZD = 0: GOTO 63110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
63110 in$ = LEFT$(in$, LEN(in$) - 1)
63120 ZL = ZL - 1: GOTO 63030
63130 IF ABS(fl) = ZL THEN 63030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 63180
63140 IF z$ = "." AND INSTR(in$, ".") = 0 THEN GOTO 63180
63150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 63190
63160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 63180
63170 IF z$ < "0" OR z$ > "9" THEN 63030
63180 PRINT z$; : in$ = in$ + z$: ZL = ZL + 1
63190 IF ABS(fl) = 1 THEN 63050 ELSE 63030
63200 LOCATE RO, PO: PRINT in$: RETURN
63210 GOSUB 63000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
63220 RETURN



END SUB

SUB SupTaxPayment (PDATE$, TBPath$)
     PostDate$ = LEFT$(PDATE$, 2) + "-" + MID$(PDATE$, 3, 2) + "-" + RIGHT$(PDATE$, 2)
     CLOSE : OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS Year$: GET 1, 1
     CurrentYear$ = Year$: CLOSE
     CurrentYear = VAL(CurrentYear$)

1026 Cleardown
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
 TBTrans(1).TransDate = Date2Num%(PayEdit(1).PaymentDate)
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
 TBTrans(1).TransDate = Date2Num%(PayEdit(1).PaymentDate)
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
  FOR LL = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, LL
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO Supn1
Supn3:
  NEXT LL
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
  FOR LL = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, LL
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO SupPPn1
SupPPn3:
  NEXT LL
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
     FOR LL = 1 TO 10
     PastDue# = PastDue# + TBRECustBal(1).TaxAmt(LL) + TBRECustBal(1).IntAmt(LL) + TBRECustBal(1).PenAmt(LL) + TBRECustBal(1).OthAmt(LL) + TBRECustBal(1).LateAmt(LL)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT LL
     ELSE
     CurBal# = TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax + TBPPCust(1).CurYrIntAmount + TBPPCust(1). _
CurYrPenAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR LL = 1 TO 10
     PastDue# = PastDue# + TBPPCustBal(1).PerPropTaxAmt(LL) + TBPPCustBal(1).FarmEquipAmt(LL) + TBPPCustBal(1).MachToolsAmt(LL) + TBPPCustBal(1).MerchCapAmt(LL) + TBPPCustBal(1).MobileHomeAmt(LL) + TBPPCustBal(1).LateFeeAmt(LL) + TBPPCustBal(1). _
IntAmt(LL) + TBPPCustBal(1).PenAmt(LL)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT LL
     END IF
     RETURN


END SUB

