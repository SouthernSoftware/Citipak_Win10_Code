10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM
30 REM
40 REM Last Revision Date:03-22-99
50 REM by Nick Picerno
  
   DEFINT A-Z

   DECLARE FUNCTION num2date$ (Num%)
   DECLARE FUNCTION Date2Num% (Dat$)
   DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
   DECLARE SUB PrintTitle (Title$)

'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'Qscr.BI'
'$INCLUDE: 'SetCnf.BI'
'$INCLUDE: 'lcTAX.BI'



60 CLS
    TBPath$ = "F:\"                    ' change to "F:\" Before compiling
    CLOSE
    OPEN "R", 1, TBPath$ + "tbyear.dat", 4
    FIELD 1, 4 AS Year$
    GET 1, 1
    CurrentYear$ = Year$
    CLOSE
    CurrentYear = VAL(CurrentYear$)
    ON ERROR GOTO networkerror

    DIM SHARED yrdesc(11)

    REDIM TBRECust(1) AS TBRECust
    REDIM TBRECustBal(1) AS TBREBAL
    REDIM TBPPCust(1) AS TBPPCust
    REDIM TBPPCustBal(1) AS TBPPBal
    REDIM TBTrans(1) AS TBTrans
    REDIM Interest(1) AS Interest

    WIDTH LPRINT 132



5000 CLS : scren$ = "3.06.00"
     GOSUB 63500: COLOR 11
     LOCATE 1, 30: PRINT "Apply Interest to Accounts"
       totreal# = 0: totpers# = 0: tottax# = 0: FixIt$ = ""
     LOCATE 3, 1: COLOR 11
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; : fl = 1: GOSUB 62210
     IF CF = 1 THEN RUN "tbsbilln"
     IF NOT (in$ = "R" OR in$ = "P" OR in$ = "F") THEN 5000
     taxtype$ = UCASE$(in$)
     IF taxtype$ = "F" THEN FixIt$ = "Y": taxtype$ = "R": LOCATE 4, 1: COLOR 11: PRINT "Fixing the interest calc will use pen amt only"
5010 LOCATE 5, 1: COLOR 11
     PRINT "Enter Date of Interest (mmddyyyy) (ESC-End) . "; : fl = -8: GOSUB 62000
     IF CF = 1 THEN 10
     IF LEN(in$) <> 8 THEN 5010
     InterestDate$ = LEFT$(in$, 2) + "/" + MID$(in$, 3, 2) + "/" + MID$(in$, 5, 4)
5015 LOCATE 8, 1: COLOR 11
     PRINT "Current Taxes"
     PRINT "Enter the Interest Percentage (Example: Enter 5 for 5%) .. "; : fl = -5
     GOSUB 62000
     CPercent! = VAL(in$)
     CPercent! = CPercent! / 100
5018 LOCATE 10, 1: COLOR 11
     PRINT "Past Year Taxes"
     PRINT "Enter the Interest Percentage (Example: Enter 5 for 5%) .. "; : fl = -5
     GOSUB 62000
     PPercent! = VAL(in$)
     PPercent! = PPercent! / 100
     GOSUB GetLastTime
     LOCATE 14, 1: COLOR 11
     PRINT "The Last Time You Applied Interest Was : "; : COLOR 15: PRINT LastTime$
5020 LOCATE 17, 1: COLOR 11
     PRINT "Check Carefully!  Is the Above Information Correct (Y/N) .. "; : fl = 1: GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN 5020
     IF in$ = "N" THEN 5000
     
     GOSUB 60000 ' Opens Customer File
     GOSUB 60025 ' Opens Past Year Balance File
     GOSUB 61000 ' Opens Transaction File
      FirstTransRecord! = LOF(TBTranFile) / LEN(TBTrans(1)) + 1
      CLOSE TBTranFile

     IF taxtype$ = "R" THEN
        EndofFile = LOF(TBFile) / LEN(TBRECust(1))
       ELSE
        EndofFile = LOF(TBFile) / LEN(TBPPCust(1))
     END IF

     FOR x! = 1 TO EndofFile
      IF taxtype$ = "R" THEN
       GET TBFile, x!, TBRECust(1)
       GET TBBalFile, x!, TBRECustBal(1)
       
      ELSE
       GET TBFile, x!, TBPPCust(1)
       GET TBBalFile, x!, TBPPCustBal(1)
      END IF
      AcctRecord! = x!
      
   LOCATE 20, 1: COLOR 11: PRINT "Processing Record # "; x!

   IF taxtype$ = "R" THEN
     IF TBRECust(1).Deleted = "Y" THEN 5230
     GOSUB ProcessReal
     IF InterestAmt# >= .01 THEN
      TBRECust(1).CurYrIntAmount = TBRECust(1).CurYrIntAmount + InterestAmt#
      PUT TBFile, AcctRecord!, TBRECust(1)
      Amount# = InterestAmt#
      Year$ = CurrentYear$
      GOSUB PostRealHistory
     END IF
     FOR LL1 = 1 TO 10
      IF PInterestAmt#(LL1) > .01 THEN
        TBRECustBal(1).IntAmt(LL1) = TBRECustBal(1).IntAmt(LL1) + PInterestAmt#(LL1)
        PUT TBBalFile, AcctRecord!, TBRECustBal(1)
        Year$ = LTRIM$(STR$(CurrentYear - LL1))
        Amount# = PInterestAmt#(LL1)
        GOSUB PostRealHistory
      END IF
     NEXT LL1

     
     ELSE

     IF TBPPCust(1).Deleted = "Y" THEN 5230
      GOSUB ProcessPers
     IF InterestAmt# > .01 THEN
      TBPPCust(1).CurYrIntAmount = TBPPCust(1).CurYrIntAmount + InterestAmt#
      PUT TBFile, AcctRecord!, TBPPCust(1)
      Amount# = InterestAmt#
      Year$ = CurrentYear$
      GOSUB PostPersHistory
     END IF
     FOR LL1 = 1 TO 10
      IF PInterestAmt#(LL1) > .01 THEN
        TBPPCustBal(1).IntAmt(LL1) = TBPPCustBal(1).IntAmt(LL1) + PInterestAmt#(LL1)
        PUT TBBalFile, AcctRecord!, TBPPCustBal(1)
        Year$ = LTRIM$(STR$(CurrentYear - LL1))
        Amount# = PInterestAmt#(LL1)
        GOSUB PostPersHistory
      END IF
     NEXT LL1
   END IF

     ' Clear Past Interest Array Here
     FOR LL1 = 1 TO 10: PInterestAmt#(LL1) = 0: NEXT LL1
     ' Get Next Record

5230 NEXT x!
     CLOSE
     GOSUB PutLastTime
5300     ' put report here
     LOCATE 20, 1
     PRINT "Press <ENTER> for HardCopy of Interest Charged (ESC-No Report) .. "; : fl = 1
     GOSUB 62000
     IF CF = 1 THEN CLOSE : GOTO 10
     COLOR 15
     OPEN "O", 10, "TBINT.REP": WIDTH #10, 255
     PRINT "** Report Now Printing **"
     GOSUB PenaltyHeading
     GOSUB 60000
     GOSUB 61000

     EndofFile = LOF(TBTranFile) / LEN(TBTrans(1))
     IF FirstTransRecord! > EndofFile THEN
      PRINT #10, "No Interest Charged This Run"
      PRINT #10, CHR$(12);
      CLOSE
      GOTO FinishReport
     END IF

    FOR x! = FirstTransRecord! TO EndofFile
      
     GET TBTranFile, x!, TBTrans(1)
   
    IF taxtype$ = "R" THEN
     GET TBFile, TBTrans(1).TrCustomerRecord, TBRECust(1)
      PRINT #10, USING "######"; TBTrans(1).TrCustomerRecord;
       PRINT #10, TAB(15); TBRECust(1).CustName;
        PRINT #10, TAB(60); USING "$$#####,#.##"; TBTrans(1).TransAmt;
         PRINT #10, TAB(76); TBTrans(1).TrTaxYear
        lc = lc + 1
        TotalInterest# = TotalInterest# + TBTrans(1).TransAmt
        TotalInterest# = INT((TotalInterest# * 100) + .5) / 100
    
     
     ELSE
     GET TBFile, TBTrans(1).TrCustomerRecord, TBPPCust(1)
      PRINT #10, USING "######"; TBTrans(1).TrCustomerRecord;
       PRINT #10, TAB(15); TBPPCust(1).CustName;
        PRINT #10, TAB(60); USING "$$#####,#.##"; TBTrans(1).TransAmt;
         PRINT #10, TAB(76); TBTrans(1).TrTaxYear
        lc = lc + 1
        TotalInterest# = TotalInterest# + TBTrans(1).TransAmt
        TotalInterest# = INT((TotalInterest# * 100) + .5) / 100
     END IF
     IF lc >= 56 THEN PRINT #10, CHR$(12); : GOSUB PenaltyHeading
5250 NEXT x!
     PRINT #10, TAB(60); USING "$$#####,#.##"; TotalInterest#
     CLOSE
FinishReport:
      'Set Report Defaults
     
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBINT.REP"
      Header$ = "Interest Report"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
      GOTO 10

PenaltyHeading:
Page = Page + 1
PRINT #10, "Detailed Journal of Interest Charged on "; InterestDate$
IF taxtype$ = "R" THEN
 PRINT #10, "Type: Real Estate ";
ELSE
 PRINT #10, "Type: Personal Property ";
END IF
PRINT #10, TAB(65); "Page #"; Page
PRINT #10, TAB(77); "Tax"
PRINT #10, "Acct Number"; TAB(15); "Name"; TAB(60); "Interest Amt"; TAB(76); "Year"
PRINT #10, STRING$(79, "-")
lc = 6
RETURN
  

ProcessReal:
  'Do Current First
  IF CPercent! > 0 THEN
  IF FixIt$ = "Y" THEN
   InterestAmt# = (TBRECust(1).CurYrPenAmount) * CPercent!
   ELSE
   InterestAmt# = (TBRECust(1).CurYrTaxAmount + TBRECust(1).CurYrPenAmount) * CPercent!
  END IF
  InterestAmt# = INT((InterestAmt# * 100) + .5) / 100
  IF InterestAmt# < 0 THEN InterestAmt# = 0
  END IF
  'Do Previous Years Now
  IF PPercent! > 0 THEN
  FOR ll2 = 1 TO 10
   IF TBRECustBal(1).TaxAmt(ll2) > 0 THEN
    IF FixIt$ = "Y" THEN
     PInterestAmt#(ll2) = (TBRECustBal(1).PenAmt(ll2)) * PPercent!
     ELSE
     PInterestAmt#(ll2) = (TBRECustBal(1).TaxAmt(ll2) + TBRECustBal(1).PenAmt(ll2)) * PPercent!
    END IF
    PInterestAmt#(ll2) = INT((PInterestAmt#(ll2) * 100) + .5) / 100
   END IF
  NEXT ll2
  END IF
  RETURN

ProcessPers:
 CurYrTax# = TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax + TBPPCust(1).CurYrPenAmount
  'Do Current First
  IF CPercent! > 0 THEN
   InterestAmt# = CurYrTax# * CPercent!
   InterestAmt# = INT((InterestAmt# * 100) + .5) / 100
   IF InterestAmt# < 0 THEN InterestAmt# = 0
  END IF

  'Do Previous Years Now
  IF PPercent! > 0 THEN
  FOR ll2 = 1 TO 10
   PastYrTax# = TBPPCustBal(1).PerPropTaxAmt(ll2) + TBPPCustBal(1).FarmEquipAmt(ll2) + TBPPCustBal(1).MachToolsAmt(ll2) + TBPPCustBal(1).MerchCapAmt(ll2) + TBPPCustBal(1).MobileHomeAmt(ll2) + TBPPCustBal(1).PenAmt(ll2)
   IF PastYrTax# > 0 THEN
    PInterestAmt#(ll2) = PastYrTax# * PPercent!
    PInterestAmt#(ll2) = INT((PInterestAmt#(ll2) * 100) + .5) / 100
   END IF
  NEXT ll2
  END IF
  RETURN


PostRealHistory:
 GET TBFile, AcctRecord!, TBRECust(1)
 GET TBBalFile, AcctRecord!, TBRECustBal(1)
 GOSUB FigureBalance
 AcctBalance# = CurBal# + PastDue#
 GOSUB 61000
 TBTrans(1).TransDate = Date2Num%(InterestDate$)
 TBTrans(1).TransAmt = Amount#
 TBTrans(1).TrTaxYear = Year$
 TBTrans(1).TrDesc = "Interest"
 TBTrans(1).TrType = 3                   '1=bill 2=payment 3=INT  4= PEN 5=release 6=BEG BAL
 TBTrans(1).TrTax = 0
 TBTrans(1).TrInterest = Amount#
 TBTrans(1).TrPenalty = 0
 TBTrans(1).TrLate = 0
 TBTrans(1).TrMachTools = 0
 TBTrans(1).TrMobileHome = 0
 TBTrans(1).TrMerchCap = 0
 TBTrans(1).TrFarmEquip = 0
 TBTrans(1).TrCustomerRecord = AcctRecord!
 TBTrans(1).TrNextTrans = 0
 TBTrans(1).TrRunningBal = AcctBalance#
 TBTrans(1).TrSource = "R"
 TBTrans(1).TrPad = ""
 PUT TBTranFile, LOF(TBTranFile) / LEN(TBTrans(1)) + 1, TBTrans(1)
 TransRecord! = LOF(TBTranFile) / LEN(TBTrans(1))
 IF TBRECust(1).FirstTrans = 0 THEN
  TBRECust(1).FirstTrans = TransRecord!
  TBRECust(1).LastTrans = TransRecord!
  PUT TBFile, AcctRecord!, TBRECust
  ELSE
  PrevTrans! = TBRECust(1).LastTrans
  TBRECust(1).LastTrans = TransRecord!
  PUT TBFile, AcctRecord!, TBRECust(1)
  GET TBTranFile, PrevTrans!, TBTrans(1)
  TBTrans(1).TrNextTrans = TransRecord!
  PUT TBTranFile, PrevTrans!, TBTrans(1)
 END IF
  CLOSE TBTranFile
 RETURN

PostPersHistory:
 GET TBFile, AcctRecord!, TBPPCust(1)
 GET TBBalFile, AcctRecord!, TBPPCustBal(1)
 GOSUB FigureBalance
 AcctBalance# = CurBal# + PastDue#
 GOSUB 61000
 TBTrans(1).TransDate = Date2Num%(InterestDate$)
 TBTrans(1).TransAmt = Amount#
 TBTrans(1).TrTaxYear = Year$
 TBTrans(1).TrDesc = "Interest"
 TBTrans(1).TrType = 3
 TBTrans(1).TrTax = 0
 TBTrans(1).TrInterest = Amount#
 TBTrans(1).TrPenalty = 0
 TBTrans(1).TrLate = 0
 TBTrans(1).TrMachTools = 0
 TBTrans(1).TrMobileHome = 0
 TBTrans(1).TrMerchCap = 0
 TBTrans(1).TrFarmEquip = 0
 TBTrans(1).TrCustomerRecord = AcctRecord!
 TBTrans(1).TrNextTrans = 0
 TBTrans(1).TrRunningBal = AcctBalance#
 TBTrans(1).TrSource = "P"
 TBTrans(1).TrPad = ""
 PUT TBTranFile, LOF(TBTranFile) / LEN(TBTrans(1)) + 1, TBTrans(1)
 TransRecord! = LOF(TBTranFile) / LEN(TBTrans(1))
 IF TBPPCust(1).FirstTrans = 0 THEN
   TBPPCust(1).FirstTrans = TransRecord!
   TBPPCust(1).LastTrans = TransRecord!
   PUT TBFile, AcctRecord!, TBPPCust(1)
  ELSE
   PrevTrans! = TBPPCust(1).LastTrans
   TBPPCust(1).LastTrans = TransRecord!
   PUT TBFile, AcctRecord!, TBPPCust(1)
   GET TBTranFile, PrevTrans!, TBTrans(1)
   TBTrans(1).TrNextTrans = TransRecord!
   PUT TBTranFile, PrevTrans!, TBTrans(1)
 END IF
 CLOSE TBTranFile
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
      OPEN TBPath$ + "TBRTBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBPath$ + "TBPTBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF
60030 RETURN

61000 TBTranFile = FREEFILE
      OPEN TBPath$ + "TBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTranFile LEN = LEN(TBTrans(1))
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
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; DATE$; TAB(70); scren$: PRINT STRING$(79, "_"): RETURN

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

 IF taxtype$ = "R" THEN
  TBIDXFile = FREEFILE
  OPEN TBPath$ + "TBRTNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS mrec$
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
  FOR LL = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, LL
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO PPn1
PPn3:
  NEXT LL
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
GetLastTime:
     IF taxtype$ = "R" THEN
       OPEN "TBRLTIME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 2
       IF LOF(11) = 0 THEN
        LastTime$ = "First Time Interest Applied"
        ELSE
        GET 11, 1, Interest(1)
        LastTime$ = num2date$(Interest(1).LastTime)
       END IF
       ELSE
       OPEN "TBPLTIME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 2
       IF LOF(11) = 0 THEN
        LastTime$ = "First Time Interest Applied"
        ELSE
        GET 11, 1, Interest(1)
        LastTime$ = num2date$(Interest(1).LastTime)
       END IF
     END IF
     CLOSE 11
     RETURN

PutLastTime:
     IF taxtype$ = "R" THEN
         OPEN "TBRLTIME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 2
      ELSE
         OPEN "TBPLTIME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 2
     END IF

     Interest(1).LastTime = Date2Num%(InterestDate$)
     PUT 11, 1, Interest(1)
     CLOSE 11

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
               fl = 1: GOSUB 62000: IF CF = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
            
     END SELECT
    

