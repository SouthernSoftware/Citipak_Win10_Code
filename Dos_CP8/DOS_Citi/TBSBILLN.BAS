10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbbilling
30 REM Tax Billing Menu
40 REM Last Revision Date:11-29-88
50 REM by Nick Picerno
   DECLARE FUNCTION num2date$ (Num%)
   DECLARE FUNCTION Date2Num% (Dat$)
60 CLS : CLEAR

  '$INCLUDE: 'lcTAX.BI'

    TBPath$ = "F:\"                    ' change to "F:\" Before compiling
    CLOSE : OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
    CurrentYear$ = year$: CLOSE
    CurrentYear = VAL(CurrentYear$)
    ON ERROR GOTO networkerror

    DIM SHARED yrdesc(11)

    REDIM TBRECust(1) AS TBRECust
    REDIM TBRECustBal(1) AS TBREBAL
    REDIM TBPPCust(1) AS TBPPCust
    REDIM TBPPCustBal(1) AS TBPPBal
    REDIM TBTrans(1) AS TBTrans

    WIDTH LPRINT 132

110 scren$ = "3.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
130 LOCATE 4, 1: COLOR 11
200 PRINT TAB(30); "Tax Billing System"
210 PRINT TAB(25); "Supplemental Billing/Interest"
220 PRINT : PRINT
250 PRINT TAB(22); "1. Print Real Estate Tax Bills"; ""
260 PRINT
    PRINT TAB(22); "2. Print Personal Property Tax Bills"
    PRINT
    PRINT TAB(22); "3. Print Billing Proof Report"
280 PRINT
290 PRINT TAB(22); "4. Post Billing to Accounts"
300 PRINT
    PRINT TAB(22); "5. Apply Penalty to Accounts"
    PRINT
    PRINT TAB(22); "6. Apply Interest to Accounts"
    PRINT
310 LOCATE 22, 1: COLOR 11: PRINT TAB(22); "Selection No. (ESC-Master Menu) : "; : fl = -1: GOSUB 62000
320 IF CF = 1 THEN CLOSE : RUN "tbtsmenu"
325 IF LEN(in$) = 0 THEN 310
330 IF VAL(in$) < 1 OR VAL(in$) > 6 THEN 310
340 ON VAL(in$) GOTO 1000, 2000, 3000, 4000, 5000, 6000

1000 RUN "TBSPRTRE"
2000 RUN "TBSPRTPP"
3000 RUN "TBSBILRP"

4000 CLS : scren$ = "3.04.00"
     GOSUB 63500: COLOR 11
     LOCATE 1, 30: PRINT "Post Bills to Accounts"
     totreal# = 0: totpers# = 0: tottax# = 0: totbldg# = 0
     LOCATE 3, 1: COLOR 11
     PRINT "(R)eal or (P)ersonal Supplementals (R,P or <ESC>-End) ... "; : fl = 1: GOSUB 62210
     IF CF = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P") THEN 4000
     TaxType$ = UCASE$(in$)
     
4010 LOCATE 5, 1: COLOR 11
     PRINT "Enter Date of Billing (mmddyyyy) (ESC-End) ... "; : fl = -8: GOSUB 62000
     IF CF = 1 THEN 10
     IF LEN(in$) <> 8 THEN 4010
     BillingDate$ = LEFT$(in$, 2) + "/" + MID$(in$, 3, 2) + "/" + MID$(in$, 5, 4)
4015 LOCATE 6, 1: COLOR 11
     PRINT "Important: Please answer the following correctly or your aging will be incorrect!"
4016 LOCATE 7, 1: COLOR 11
     PRINT "Have you posted";
     IF TaxType$ = "P" THEN
      PRINT " Personal Tax Bill Supplementals "
      ELSE
      PRINT " Real Estate Tax Bill Supplementals "
     END IF
     PRINT "For this tax year prior to this (Y/N): "; : fl = 1
     GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN 4016
     AgeThemBefore$ = in$

4020 LOCATE 10, 1: COLOR 11
     PRINT "Is the Above Information Correct (Y/N) .. "; : fl = 1: GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN 4020
     IF in$ = "N" THEN 4000

    CLOSE : OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
    year$ = RIGHT$(BillingDate$, 4)
    PUT 1, 1
    CLOSE 1

     GOSUB 60000
     GOSUB 60025

     IF TaxType$ = "R" THEN
        EndofFile = LOF(TBFile) / LEN(TBRECust(1))
       ELSE
        EndofFile = LOF(TBFile) / LEN(TBPPCust(1))
     END IF

     FOR x = 1 TO EndofFile
      IF TaxType$ = "R" THEN
       GET TBFile, x, TBRECust(1)
      ELSE
       GET TBFile, x, TBPPCust(1)
      END IF
      AcctRecord = x

   LOCATE 15, 1: COLOR 11: PRINT "Process Record # "; x
   IF TaxType$ = "R" THEN
      IF TBRECust(1).Deleted = "Y" THEN 4230
      GOSUB PostRealBill
      ELSE
      IF TBPPCust(1).Deleted = "Y" THEN 4230
      GOSUB PostPersBill
   END IF

4230 NEXT x
     CLOSE
     LOCATE 20, 1
     PRINT "All Done: Press <ENTER> to Continue (ESC-No Report) .. "; : fl = 1
     GOSUB 62000
     GOTO 10


PostRealBill:
 IF TaxType$ = "R" THEN

  IF TBRECust(1).LBCurYrTaxAmount < .01 THEN

  IF AgeThemBefore$ = "N" THEN
  'Age Them Here
   GET TBBalFile, x, TBRECustBal(1)
   TBRECustBal(1).TaxAmt(10) = TBRECustBal(1).TaxAmt(10) + TBRECustBal(1).TaxAmt(9)
   TBRECustBal(1).IntAmt(10) = TBRECustBal(1).IntAmt(10) + TBRECustBal(1).IntAmt(9)
   TBRECustBal(1).PenAmt(10) = TBRECustBal(1).PenAmt(10) + TBRECustBal(1).PenAmt(9)
   TBRECustBal(1).OthAmt(10) = TBRECustBal(1).OthAmt(10) + TBRECustBal(1).OthAmt(9)
   TBRECustBal(1).LateAmt(10) = TBRECustBal(1).LateAmt(10) + TBRECustBal(1).LateAmt(9)
   TBRECustBal(1).TaxNotice(10) = TBRECustBal(1).TaxNotice(9)
   FOR Cnt! = 9 TO 2 STEP -1
        TBRECustBal(1).TaxAmt(Cnt!) = TBRECustBal(1).TaxAmt(Cnt! - 1)
        TBRECustBal(1).IntAmt(Cnt!) = TBRECustBal(1).IntAmt(Cnt! - 1)
        TBRECustBal(1).PenAmt(Cnt!) = TBRECustBal(1).PenAmt(Cnt! - 1)
        TBRECustBal(1).OthAmt(Cnt!) = TBRECustBal(1).OthAmt(Cnt! - 1)
        TBRECustBal(1).LateAmt(Cnt!) = TBRECustBal(1).LateAmt(Cnt! - 1)
        TBRECustBal(1).TaxNotice(Cnt!) = TBRECustBal(1).TaxNotice(Cnt! - 1)
   NEXT Cnt!
  TBRECustBal(1).TaxAmt(1) = TBRECust(1).CurYrTaxAmount
  TBRECustBal(1).IntAmt(1) = TBRECust(1).CurYrIntAmount
  TBRECustBal(1).PenAmt(1) = TBRECust(1).CurYrPenAmount
  TBRECustBal(1).OthAmt(1) = TBRECust(1).CurYrOthAmount
  TBRECustBal(1).LateAmt(1) = 0
  TBRECustBal(1).TaxNotice(1) = TBRECust(1).CurYrBillNumber
  PUT TBBalFile, x, TBRECustBal(1)
  TBRECust(1).CurYrTaxAmount = 0
  TBRECust(1).CurYrIntAmount = 0
  TBRECust(1).CurYrPenAmount = 0
  TBRECust(1).CurYrOthAmount = 0
  TBRECust(1).CurYrBillNumber = ""
  PUT TBFile, x, TBRECust(1)
  END IF
  RETURN' do not allow if no bill this year
 END IF
 END IF

 GET TBBalFile, AcctRecord, TBRECustBal(1)

 IF TaxType$ = "R" THEN

 FOR Cnt! = 10 TO 1 STEP -1
 TBRECustBal(1).TaxAmt(Cnt!) = 0
 TBRECustBal(1).IntAmt(Cnt!) = 0
 TBRECustBal(1).PenAmt(Cnt!) = 0
 TBRECustBal(1).OthAmt(Cnt!) = 0
 TBRECustBal(1).LateAmt(Cnt!) = 0
 TBRECustBal(1).TaxNotice(Cnt!) = ""
 NEXT Cnt!
 PUT TBBalFile, AcctRecord, TBRECustBal(1)

 END IF

'Now to Set Current Stuff
 TBRECust(1).CurYrTaxAmount = TBRECust(1).LBCurYrTaxAmount
 TBRECust(1).CurYrIntAmount = 0
 TBRECust(1).CurYrPenAmount = 0
 TBRECust(1).CurYrOthAmount = 0
 TBRECust(1).CurYrBillNumber = TBRECust(1).LBCurYrBillNumber

 'Store Information for Posting in Detail
 TransactionAmount# = TBRECust(1).LBCurYrTaxAmount
 TransactionDetail$ = "SUP " + TBRECust(1).LBCurYrBillNumber


 ' Now Set Billing Information to Null for Next Yr Printing
 TBRECust(1).LBCurYrTaxAmount = 0
 TBRECust(1).LBCurYrBillNumber = ""
 PUT TBFile, AcctRecord, TBRECust(1)
 GOSUB PostRealHistoryBill
 RETURN


PostPersBill:
   IF VAL(TBPPCust(1).LBCurYrBillNumber) <= 0 THEN
   IF AgeThemBefore$ = "N" THEN
        GET TBBalFile, x, TBPPCustBal(1)
        TBPPCustBal(1).PerPropTaxAmt(10) = TBPPCustBal(1).PerPropTaxAmt(10) + TBPPCustBal(1).PerPropTaxAmt(9)
        TBPPCustBal(1).FarmEquipAmt(10) = TBPPCustBal(1).FarmEquipAmt(10) + TBPPCustBal(1).FarmEquipAmt(9)
        TBPPCustBal(1).MachToolsAmt(10) = TBPPCustBal(1).MachToolsAmt(10) + TBPPCustBal(1).MachToolsAmt(9)
        TBPPCustBal(1).MerchCapAmt(10) = TBPPCustBal(1).MerchCapAmt(10) + TBPPCustBal(1).MerchCapAmt(9)
        TBPPCustBal(1).MobileHomeAmt(10) = TBPPCustBal(1).MobileHomeAmt(10) + TBPPCustBal(1).MobileHomeAmt(9)
        TBPPCustBal(1).LateFeeAmt(10) = TBPPCustBal(1).LateFeeAmt(10) + TBPPCustBal(1).LateFeeAmt(9)
        TBPPCustBal(1).IntAmt(10) = TBPPCustBal(1).IntAmt(10) + TBPPCustBal(1).IntAmt(9)
        TBPPCustBal(1).PenAmt(10) = TBPPCustBal(1).PenAmt(10) + TBPPCustBal(1).PenAmt(9)
        TBPPCustBal(1).TaxNotice(10) = TBPPCustBal(1).TaxNotice(9)
       FOR Cnt! = 9 TO 2 STEP -1
        TBPPCustBal(1).PerPropTaxAmt(Cnt!) = TBPPCustBal(1).PerPropTaxAmt(Cnt! - 1)
        TBPPCustBal(1).FarmEquipAmt(Cnt!) = TBPPCustBal(1).FarmEquipAmt(Cnt! - 1)
        TBPPCustBal(1).MachToolsAmt(Cnt!) = TBPPCustBal(1).MachToolsAmt(Cnt! - 1)
        TBPPCustBal(1).MerchCapAmt(Cnt!) = TBPPCustBal(1).MerchCapAmt(Cnt! - 1)
        TBPPCustBal(1).MobileHomeAmt(Cnt!) = TBPPCustBal(1).MobileHomeAmt(Cnt! - 1)
        TBPPCustBal(1).LateFeeAmt(Cnt!) = TBPPCustBal(1).LateFeeAmt(Cnt! - 1)
        TBPPCustBal(1).IntAmt(Cnt!) = TBPPCustBal(1).IntAmt(Cnt! - 1)
        TBPPCustBal(1).PenAmt(Cnt!) = TBPPCustBal(1).PenAmt(Cnt! - 1)
        TBPPCustBal(1).TaxNotice(Cnt!) = TBPPCustBal(1).TaxNotice(Cnt! - 1)
       NEXT Cnt!
        TBPPCustBal(1).PerPropTaxAmt(1) = TBPPCust(1).CurYrPersPropTax
        TBPPCustBal(1).FarmEquipAmt(1) = TBPPCust(1).CurYrFarmEquipTax
        TBPPCustBal(1).MachToolsAmt(1) = TBPPCust(1).CurYrMachToolsTax
        TBPPCustBal(1).MerchCapAmt(1) = TBPPCust(1).CurYrMerchCapTax
        TBPPCustBal(1).MobileHomeAmt(1) = TBPPCust(1).CurYrMobileHomeTax
        TBPPCustBal(1).LateFeeAmt(1) = TBPPCust(1).CurYrLateFeeTax
        TBPPCustBal(1).IntAmt(1) = TBPPCust(1).CurYrIntAmount
        TBPPCustBal(1).PenAmt(1) = TBPPCust(1).CurYrPenAmount
        TBPPCustBal(1).TaxNotice(1) = TBPPCust(1).CurYrBillNumber
        PUT TBBalFile, x, TBPPCustBal(1)

        TBPPCust(1).CurYrBillNumber = ""
        TBPPCust(1).CurYrPersPropTax = 0
        TBPPCust(1).CurYrFarmEquipTax = 0
        TBPPCust(1).CurYrMachToolsTax = 0
        TBPPCust(1).CurYrMerchCapTax = 0
        TBPPCust(1).CurYrMobileHomeTax = 0
        TBPPCust(1).CurYrLateFeeTax = 0
        TBPPCust(1).CurYrIntAmount = 0
        TBPPCust(1).CurYrPenAmount = 0
        PUT TBFile, x, TBPPCust(1)
      END IF
   RETURN' do not allow
   END IF
   
GET TBBalFile, AcctRecord, TBPPCustBal(1)
 FOR Cnt! = 10 TO 1
  TBPPCustBal(1).PerPropTaxAmt(Cnt!) = 0
  TBPPCustBal(1).FarmEquipAmt(Cnt!) = 0
  TBPPCustBal(1).MachToolsAmt(Cnt!) = 0
  TBPPCustBal(1).MerchCapAmt(Cnt!) = 0
  TBPPCustBal(1).MobileHomeAmt(Cnt!) = 0
  TBPPCustBal(1).LateFeeAmt(Cnt!) = 0
  TBPPCustBal(1).IntAmt(Cnt!) = 0
  TBPPCustBal(1).PenAmt(Cnt!) = 0
  TBPPCustBal(1).TaxNotice(Cnt!) = ""
 NEXT Cnt!
PUT TBBalFile, AcctRecord, TBPPCustBal(1)

 'Now to Set Current Stuff

        TBPPCust(1).CurYrBillNumber = TBPPCust(1).LBCurYrBillNumber
        TBPPCust(1).CurYrPersPropTax = TBPPCust(1).LBCurYrPersPropTaxAmount - TBPPCust(1).LBPPTRADiscnt
        TBPPCust(1).CurYrFarmEquipTax = TBPPCust(1).LBCurYrFarmEquipTaxAmount
        TBPPCust(1).CurYrMachToolsTax = TBPPCust(1).LBCurYrMachToolsTaxAmount
        TBPPCust(1).CurYrMerchCapTax = TBPPCust(1).LBCurYrMerchCapTaxAmount
        TBPPCust(1).CurYrMobileHomeTax = TBPPCust(1).LBCurYrMobileHomeTaxAmount
        TBPPCust(1).CurYrLateFeeTax = TBPPCust(1).LBLateFeeAmount
        TBPPCust(1).CurYrIntAmount = 0
        TBPPCust(1).CurYrPenAmount = 0
        

 'Store Information for Posting in Detail
 TransactionAmount# = TBPPCust(1).LBCurYrPersPropTaxAmount + TBPPCust(1).LBCurYrFarmEquipTaxAmount + TBPPCust(1).LBCurYrMachToolsTaxAmount + TBPPCust(1).LBCurYrMerchCapTaxAmount + TBPPCust(1).LBCurYrMobileHomeTaxAmount + TBPPCust(1).LBLateFeeAmount  _
- TBPPCust(1).LBPPTRADiscnt
 TransactionDetail$ = TBPPCust(1).LBCurYrBillNumber

        PropTax = TBPPCust(1).LBCurYrPersPropTaxAmount - TBPPCust(1).LBPPTRADiscnt
        FarmEquipTax = TBPPCust(1).LBCurYrFarmEquipTaxAmount
        MachToolsTax = TBPPCust(1).LBCurYrMachToolsTaxAmount
        MerchCapTax = TBPPCust(1).LBCurYrMerchCapTaxAmount
        MobileHomeTax = TBPPCust(1).LBCurYrMobileHomeTaxAmount
        LateFeeTax = TBPPCust(1).LBLateFeeAmount
        PPTRA# = TBPPCust(1).LBPPTRADiscnt

 ' Now Set Billing Information to Null for Next Yr Printing
        TBPPCust(1).LBCurYrBillNumber = ""
        TBPPCust(1).LBCurYrPersPropTaxAmount = 0
        TBPPCust(1).LBCurYrFarmEquipTaxAmount = 0
        TBPPCust(1).LBCurYrMachToolsTaxAmount = 0
        TBPPCust(1).LBCurYrMerchCapTaxAmount = 0
        TBPPCust(1).LBCurYrMobileHomeTaxAmount = 0
        TBPPCust(1).LBLateFeeAmount = 0
        TBPPCust(1).LBPPTRADiscnt = 0
        PUT TBFile, AcctRecord, TBPPCust(1)
        GOSUB PostPersHistoryBill
        RETURN


PostRealHistoryBill:
 GET TBFile, AcctRecord, TBRECust(1)
 GET TBBalFile, AcctRecord, TBRECustBal(1)
 GOSUB FigureBalance
 AcctBalance# = CurBal# + PastDue#
 GOSUB 61000
 TBTrans(1).TransDate = Date2Num%(BillingDate$)
 TBTrans(1).TransAmt = TransactionAmount#
 TBTrans(1).TrTaxYear = RIGHT$(BillingDate$, 4)
 TBTrans(1).TrDesc = TransactionDetail$
 TBTrans(1).TrType = 1                   '1=bill 2=payment 3=INT  4= PEN 5=release 6=BEG BAL
 TBTrans(1).TrTax = TransactionAmount#
 TBTrans(1).TrInterest = 0
 TBTrans(1).TrPenalty = 0
 TBTrans(1).TrLate = 0
 TBTrans(1).TrPersProp = 0
 TBTrans(1).TrMachTools = 0
 TBTrans(1).TrMobileHome = 0
 TBTrans(1).TrMerchCap = 0
 TBTrans(1).TrFarmEquip = 0
 TBTrans(1).TrCustomerRecord = AcctRecord
 TBTrans(1).TrNextTrans = 0
 TBTrans(1).TrRunningBal = AcctBalance#
 TBTrans(1).TrSource = "R"
 TBTrans(1).TrPad = ""

 PUT TBTranFile, LOF(TBTranFile) / LEN(TBTrans(1)) + 1, TBTrans(1)
 TransRecord! = LOF(TBTranFile) / LEN(TBTrans(1))
 IF TBRECust(1).FirstTrans = 0 THEN
    TBRECust(1).FirstTrans = TransRecord!
    TBRECust(1).LastTrans = TransRecord!
    PUT TBFile, AcctRecord, TBRECust(1)
   ELSE
    PrevTrans! = TBRECust(1).LastTrans
    TBRECust(1).LastTrans = TransRecord!
    PUT TBFile, AcctRecord, TBRECust(1)
    GET TBTranFile, PrevTrans!, TBTrans(1)
    TBTrans(1).TrNextTrans = TransRecord!
    PUT TBTranFile, PrevTrans!, TBTrans(1)
 END IF
 CLOSE TBTranFile
 RETURN

PostPersHistoryBill:
 GET TBFile, AcctRecord, TBPPCust(1)
 GET TBBalFile, AcctRecord, TBPPCustBal(1)
 GOSUB FigureBalance
 AcctBalance# = CurBal# + PastDue#
 GOSUB 61000
 TBTrans(1).TransDate = Date2Num%(BillingDate$)
 TBTrans(1).TransAmt = TransactionAmount#
 TBTrans(1).TrTaxYear = RIGHT$(BillingDate$, 4)
 TBTrans(1).TrDesc = TransactionDetail$
 TBTrans(1).TrType = 1
 TBTrans(1).TrTax = TransactionAmount#
 TBTrans(1).TrInterest = 0
 TBTrans(1).TrPenalty = 0
 TBTrans(1).TrLate = LateFeeTax
 TBTrans(1).TrPersProp = PropTax
 TBTrans(1).TrMachTools = MachToolsTax
 TBTrans(1).TrMobileHome = MobileHomeTax
 TBTrans(1).TrMerchCap = MerchCapTax
 TBTrans(1).TrFarmEquip = FarmEquipTax
 TBTrans(1).TrCustomerRecord = AcctRecord
 TBTrans(1).TrNextTrans = 0
 TBTrans(1).TrRunningBal = AcctBalance#
 TBTrans(1).TrSource = "P"
 TBTrans(1).TrPPTRA = PPTRA#
 TBTrans(1).TrPad = ""
 PUT TBTranFile, LOF(TBTranFile) / LEN(TBTrans(1)) + 1, TBTrans(1)
 TransRecord! = LOF(TBTranFile) / LEN(TBTrans(1))
 IF TBPPCust(1).FirstTrans = 0 THEN
   TBPPCust(1).FirstTrans = TransRecord!
   TBPPCust(1).LastTrans = TransRecord!
   PUT TBFile, AcctRecord, TBPPCust(1)
  ELSE
   PrevTrans! = TBPPCust(1).LastTrans
   TBPPCust(1).LastTrans = TransRecord!
   PUT TBFile, AcctRecord, TBPPCust(1)
   GET TBTranFile, PrevTrans!, TBTrans(1)
   TBTrans(1).TrNextTrans = TransRecord!
   PUT TBTranFile, PrevTrans!, TBTrans(1)
 END IF
 CLOSE TBTranFile
 RETURN



5000 CLS : scren$ = "3.05.00"
     GOSUB 63500: COLOR 11
     LOCATE 1, 30: PRINT "Apply Penalty to Accounts": totreal# = 0: totpers# = 0: tottax# = 0
     LOCATE 3, 1: COLOR 11
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; : fl = 1: GOSUB 62210
     IF CF = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P" OR in$ = "F" OR in$ = "A") THEN 5000
     TaxType$ = UCASE$(in$)
     FixType$ = UCASE$(in$)
     IF FixType$ = "F" THEN TaxType$ = "P"
     

5010 LOCATE 5, 1: COLOR 11
     PRINT "Enter Date of Penalty (mmddyyyy) (ESC-End) ... "; : fl = -8: GOSUB 62000
     IF CF = 1 THEN 10
     IF LEN(in$) <> 8 THEN 5010
     PenaltyDate$ = LEFT$(in$, 2) + "/" + MID$(in$, 3, 2) + "/" + MID$(in$, 5, 4)


     

5020 LOCATE 7, 1: COLOR 11
     PRINT "Is the Above Information Correct (Y/N) .. "; : fl = 1: GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN 5020
     IF in$ = "N" THEN 5000

     GOSUB 60000: GOSUB 60025
     IF TaxType$ = "R" THEN
        EndofFile = LOF(TBFile) / LEN(TBRECust(1))
       ELSE
        EndofFile = LOF(TBFile) / LEN(TBPPCust(1))
     END IF

     FOR x = 1 TO EndofFile
      IF TaxType$ = "R" THEN
       GET 1, x, TBRECust(1)
      ELSE
       GET 1, x, TBPPCust(1)
      END IF
      AcctRecord = x
     
     LOCATE 15, 1: COLOR 11: PRINT "Process Record # "; x
   IF TaxType$ = "R" THEN
     IF TBRECust(1).Deleted = "Y" THEN 5230
     GOSUB ProcessReal
     IF Penalty# >= .01 THEN
      TBRECust(1).CurYrPenAmount = Penalty#
      PUT TBFile, AcctRecord, TBRECust(1)
      GOSUB PostRealHistory
     END IF

     ELSE

     IF TBPPCust(1).Deleted = "Y" THEN 5230
      IF FixType$ = "F" THEN
       GOSUB FixPenalty
      ELSE
       GOSUB ProcessPers
      END IF
     IF Penalty# <> 0 THEN
      IF FixType$ = "F" THEN
      TBPPCust(1).CurYrPenAmount = TBPPCust(1).CurYrPenAmount + Penalty#
      ELSE
      TBPPCust(1).CurYrPenAmount = Penalty#
      END IF
      PUT TBFile, AcctRecord, TBPPCust(1)
      GOSUB PostPersHistory
     END IF

   END IF

5230 NEXT x
     CLOSE
     ' put report here
     LOCATE 20, 1
     PRINT "Press <ENTER> for HardCopy of Penalties Charged (ESC-No Report) .. "; : fl = 1
     GOSUB 62000
     IF CF = 1 THEN CLOSE : GOTO 10
     COLOR 15
     PRINT "** Report Now Printing **"
     GOSUB PenaltyHeading
     GOSUB 60000: GOSUB 60025
     IF TaxType$ = "R" THEN
        EndofFile = LOF(TBFile) / LEN(TBRECust(1))
       ELSE
        EndofFile = LOF(TBFile) / LEN(TBPPCust(1))
     END IF

     FOR x = 1 TO EndofFile
      IF TaxType$ = "R" THEN
       GET 1, x, TBRECust(1)
      ELSE
       GET 1, x, TBPPCust(1)
      END IF
      AcctRecord = x

   IF TaxType$ = "R" THEN
     IF TBRECust(1).Deleted = "Y" THEN 5250
     IF TBRECust(1).CurYrPenAmount > 0 THEN
      LPRINT USING "######"; x;
       LPRINT TAB(15); TBRECust(1).CustName;
        LPRINT TAB(60); USING "$$#####,#.##"; TBRECust(1).CurYrPenAmount
        lc = lc + 1: TotalPenalty# = TotalPenalty# + TBRECust(1).CurYrPenAmount
        TotalPenalty# = INT((TotalPenalty# * 100) + .5) / 100
     END IF
     
     ELSE
     IF TBPPCust(1).Deleted = "Y" THEN 5250
     IF TBPPCust(1).CurYrPenAmount > 0 THEN
      LPRINT USING "######"; x;
       LPRINT TAB(15); TBPPCust(1).CustName;
        LPRINT TAB(60); USING "$$#####,#.##"; TBPPCust(1).CurYrPenAmount
        lc = lc + 1: TotalPenalty# = TotalPenalty# + TBPPCust(1).CurYrPenAmount
        TotalPenalty# = INT((TotalPenalty# * 100) + .5) / 100
      END IF

     END IF
     IF lc >= 56 THEN LPRINT CHR$(12); : GOSUB PenaltyHeading
5250 NEXT x
     LPRINT TAB(60); USING "$$#####,#.##"; TotalPenalty#
     CLOSE
     LPRINT CHR$(12);
     GOTO 10

PenaltyHeading:
Page = Page + 1
LPRINT "Lunenburg County "
LPRINT "Penalty's Charged on "; PenaltyDate$
IF TaxType$ = "R" THEN
 LPRINT "Type: Real Estate ";
ELSE
 LPRINT "Type: Personal Property ";
END IF
LPRINT TAB(65); "Page #"; Page
LPRINT
LPRINT "Acct Number"; TAB(15); "Name"; TAB(60); "Penalty Amount"
LPRINT STRING$(79, "-")
lc = 6
RETURN
  

ProcessReal:
    
  IF TBRECust(1).CurYrPenAmount > 0 THEN RETURN' do not allow double tax posting
  Penalty# = TBRECust(1).CurYrTaxAmount * .1
  Penalty# = INT((Penalty# * 100) + .5) / 100
  IF Penalty# < 0 THEN Penalty# = 0
  IF LEFT$(TBRECust(1).ClassCode, 1) = "7" THEN Penalty# = 0
  RETURN

ProcessPers:
  IF TBPPCust(1).CurYrPenAmount > 0 THEN RETURN
  CurYrTax# = TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax
  Penalty# = CurYrTax# * .1
  Penalty# = INT((Penalty# * 100) + .5) / 100
  IF Penalty# < 0 THEN Penalty# = 0
  RETURN

FixPenalty:
  CurYrTax# = TBPPCust(1).CurYrLateFeeTax
  Penalty# = CurYrTax# * -.1
  Penalty# = INT((Penalty# * 100) + .5) / 100
  RETURN

PostRealHistory:
 GET TBFile, AcctRecord, TBRECust(1)
 GET TBBalFile, AcctRecord, TBRECustBal(1)
 GOSUB FigureBalance
 AcctBalance# = CurBal# + PastDue#
 GOSUB 61000
 TBTrans(1).TransDate = Date2Num%(PenaltyDate$)
 TBTrans(1).TransAmt = Penalty#
 TBTrans(1).TrTaxYear = CurrentYear$
 TBTrans(1).TrDesc = "Penalty"
 TBTrans(1).TrType = 4                   '1=bill 2=payment 3=INT  4= PEN 5=release 6=BEG BAL
 TBTrans(1).TrTax = 0
 TBTrans(1).TrInterest = 0
 TBTrans(1).TrPenalty = Penalty#
 TBTrans(1).TrLate = 0
 TBTrans(1).TrMachTools = 0
 TBTrans(1).TrMobileHome = 0
 TBTrans(1).TrMerchCap = 0
 TBTrans(1).TrFarmEquip = 0
 TBTrans(1).TrCustomerRecord = AcctRecord
 TBTrans(1).TrNextTrans = 0
 TBTrans(1).TrRunningBal = AcctBalance#
 TBTrans(1).TrSource = "R"
 TBTrans(1).TrPad = ""
 PUT TBTranFile, LOF(TBTranFile) / LEN(TBTrans(1)) + 1, TBTrans(1)
 TransRecord! = LOF(TBTranFile) / LEN(TBTrans(1))
 PrevTrans! = TBRECust(1).LastTrans
 IF PrevTrans! <= 0 THEN
  TBRECust(1).FirstTrans = TransRecord!
  TBRECust(1).LastTrans = TransRecord!
  PUT TBFile, AcctRecord, TBRECust(1)
  ELSE
 TBRECust(1).LastTrans = TransRecord!
 PUT TBFile, AcctRecord, TBRECust(1)
 GET TBTranFile, PrevTrans!, TBTrans(1)
 TBTrans(1).TrNextTrans = TransRecord!
 PUT TBTranFile, PrevTrans!, TBTrans(1)
 END IF
 CLOSE TBTranFile
 RETURN

PostPersHistory:
 GET TBFile, AcctRecord, TBPPCust(1)
 GET TBBalFile, AcctRecord, TBPPCustBal(1)
 GOSUB FigureBalance
 AcctBalance# = CurBal# + PastDue#
 GOSUB 61000
 TBTrans(1).TransDate = Date2Num%(PenaltyDate$)
 TBTrans(1).TransAmt = Penalty#
 TBTrans(1).TrTaxYear = CurYear$
 TBTrans(1).TrDesc = "Penalty"
 TBTrans(1).TrType = 4
 TBTrans(1).TrTax = 0
 TBTrans(1).TrInterest = 0
 TBTrans(1).TrPenalty = Penalty#
 TBTrans(1).TrLate = 0
 TBTrans(1).TrMachTools = 0
 TBTrans(1).TrMobileHome = 0
 TBTrans(1).TrMerchCap = 0
 TBTrans(1).TrFarmEquip = 0
 TBTrans(1).TrCustomerRecord = AcctRecord
 TBTrans(1).TrNextTrans = 0
 TBTrans(1).TrRunningBal = AcctBalance#
 TBTrans(1).TrSource = "P"
 TBTrans(1).TrPad = ""
 PUT TBTranFile, LOF(TBTranFile) / LEN(TBTrans(1)) + 1, TBTrans(1)
 TransRecord! = LOF(TBTranFile) / LEN(TBTrans(1))
 PrevTrans! = TBPPCust(1).LastTrans
 TBPPCust(1).LastTrans = TransRecord!
 PUT TBFile, AcctRecord, TBPPCust(1)
 GET TBTranFile, PrevTrans!, TBTrans(1)
 TBTrans(1).TrNextTrans = TransRecord!
 PUT TBTranFile, PrevTrans!, TBTrans(1)
 CLOSE TBTranFile
 RETURN
6000 RUN "TBSCALIN"

7000 GOTO 10

60000  IF TaxType$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TTRSCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBPath$ + "TTPSCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN

60025 TBBalFile = FREEFILE
      IF TaxType$ = "R" THEN
      OPEN TBPath$ + "TTRSBAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBPath$ + "TTPSBAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF
60030 RETURN

61000 TBTranFile = FREEFILE
      OPEN TBPath$ + "TTTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTranFile LEN = LEN(TBTrans(1))
61020 RETURN




62000 COLOR 14: CF = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN CF = 1: RETURN
      IF z$ = CHR$(9) THEN CF = 2: RETURN
      IF LEN(z$) < 2 THEN 62065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN CF = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN CF = 12: RETURN
62065 IF ASC(z$) = 0 THEN 62030
62070 IF LEN(z$) > 1 THEN CF = ASC(RIGHT$(z$, 1)) - 57: IF CF > 0 THEN RETURN ELSE CF = CF + 27: RETURN
62080 IF z$ = " " THEN 62130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 62000
62090 IF z$ <> CHR$(8) THEN 62130 ELSE IF ZL = 0 THEN 62030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 62110 ELSE IF z$ = "," THEN 62120
62100 IF z$ = "." THEN ZD = 0: GOTO 62110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
62110 in$ = LEFT$(in$, LEN(in$) - 1)
62120 ZL = ZL - 1: GOTO 62030
62130 IF ABS(fl) = ZL THEN 62030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 62180
62140 IF z$ = "." AND ZD = 0 THEN ZD = 1: GOTO 62180
62150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF z$ < "0" OR z$ > "9" THEN 62030
62180 PRINT z$; : in$ = in$ + z$: ZL = ZL + 1
62190 IF ABS(fl) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT in$: RETURN
62210 GOSUB 62000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
62220 RETURN
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); scren$: PRINT STRING$(79, "_"): RETURN

Namesearch:
 LOCATE 8, 1:
 COLOR 11
 PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
 GOSUB 62000
 IF CF = 1 THEN
  ggo = 0
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  RETURN
 END IF
 nn$ = in$
 nn = LEN(nn$)
 IF nn > 20 THEN nn = 20
 nn$ = LEFT$(nn$, nn)

 IF TaxType$ = "R" THEN
  TBIDXFile = FREEFILE
  OPEN TBPath$ + "TBRTNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS mrec$
  GOSUB 60000
  FOR ll = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, ll
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO n1
n3:
  NEXT ll
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  ggo = 0
  RETURN

n1:
GET TBFile, CVS(mrec$), TBRECust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBRECust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(mrec$)
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
  ggo = 1: Record! = CVS(mrec$)
  FOR LLL = 8 TO 13:  LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO n3
GOTO n2
END IF


  TBIDXFile = FREEFILE
  OPEN TBPath$ + "TBPTNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS mrec$
  GOSUB 60000
  FOR ll = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, ll
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO PPn1
PPn3:
  NEXT ll
  LOCATE 8, 1: PRINT STRING$(79, 32)
  ggo = 0
  RETURN

PPn1:
GET TBFile, CVS(mrec$), TBPPCust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBPPCust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(mrec$)
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
  ggo = 1: Record! = CVS(mrec$)
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
    IF TaxType$ = "R" OR TaxType$ = "N" THEN
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
          CASE IS = 63
               RESUME NEXT

          CASE ELSE
               CLS
               LOCATE 15, 1: COLOR 28: PRINT "Error Message: "; : COLOR 15: PRINT "Error "; Number; " Has Occured in Line "; ERL
               PRINT "Write Down Message and Hit <ESC> Key to exit."
               fl = 1: GOSUB 62000: IF CF = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
            
     END SELECT
    

