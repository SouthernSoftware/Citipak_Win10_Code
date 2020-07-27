10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbpayment
30 REM Payment Entry Menu
40 REM Last Revision Date: 7-14-87
50 REM by Nick Picerno
  DECLARE FUNCTION Num2Date$ (Num%)
  DECLARE FUNCTION Date2Num% (Dat$)
60 CLS : CLEAR
   
   '$INCLUDE: 'lcTAX.BI'

    TBPath$ = "F:\"                    ' change to "F:\" Before compiling
   

     CLOSE : OPEN "R", 1, TBPath$ + "tbppyear.dat", 4
     FIELD 1, 4 AS Year$: GET 1, 1
     CurrentYear$ = Year$: CLOSE
     CurrentYear = VAL(CurrentYear$)
     ON ERROR GOTO networkerror

     DIM SHARED yrdesc(11)
     REDIM TBLog(1) AS TBLog
     REDIM TBRECust(1) AS TBRECust
     REDIM TBRECustBal(1) AS TBREBAL
     REDIM TBPPCust(1) AS TBPPCust
     REDIM TBPPCustBal(1) AS TBPPBal
     REDIM TBTrans(1) AS TBTrans

     PDate$ = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 2)
   
100 GOSUB 63000: REM get town in town$
110 SCREN$ = "5.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
130 LOCATE 5, 1: COLOR 11:
200 PRINT TAB(30); "Tax Billing System"
210 PRINT TAB(30); "     Payments"
220 PRINT : PRINT
230 PRINT TAB(25); "1. Enter Payments"
240 PRINT
250 PRINT TAB(25); "2. Change/Delete Payment Transaction"
260 PRINT
270 PRINT TAB(25); "3. Print Payment Edit Report"
280 PRINT
290 PRINT TAB(25); "4. Post Payments to Accounts"
    PRINT
    PRINT TAB(25); "5. Manually Adjust Account"

310 LOCATE 20, 1: COLOR 11: PRINT TAB(25); "Selection No. (ESC-Master Menu) : "; : fl = -1: GOSUB 62000
320 IF cf = 1 THEN CLOSE : RUN "tbtmenu"
325 IF LEN(in$) = 0 THEN 310
330 IF VAL(in$) < 1 OR VAL(in$) > 5 THEN 310
340 ON VAL(in$) GOTO 1000, 2000, 3000, 4000, 5000
410 RUN "tbcustma"

1000 SCREN$ = "4.01.00": CLS : GOSUB 63500
1020 COLOR 11: LOCATE 1, 32: PRINT "Customer Payment Entry";
1025 LOCATE 3, 1: COLOR 11
     IF taxtype$ = "R" OR taxtype$ = "P" THEN
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; taxtype$
     ELSE
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P" OR in$ = "A") THEN 1025
     taxtype$ = UCASE$(in$)
     END IF
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
'HERE DALE 110605
       'LOCK #TBFile, Record!
       GET TBFile, Record!, TBRECust(1)
       GET TBBalFile, Record!, TBRECustBal(1)
       IF TBRECust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 1000
       ELSE
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBPPCust(1)) THEN 1030
'HERE DALE 111605
       'LOCK #TBFile, Record!
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
    PaySource$ = "R"
    CLS
    SCREN$ = "4.01.00": CLS : GOSUB 63500
    COLOR 11: LOCATE 1, 32: PRINT "Customer Payment Entry";
    LOCATE 3, 1: PRINT "Real Estate Customer"
    LOCATE 5, 1: COLOR 15
        PRINT TBRECust(1).CustName; TAB(50); "Current Bal: "; USING "$$######,#.##"; CurBal#
        PRINT TBRECust(1).Address1; TAB(50); "   Past Bal: "; USING "$$######,#.##"; PastDue#
        PRINT TBRECust(1).Address2; TAB(50); " Total Owed: "; USING "$$######,#.##"; CurBal# + PastDue#
        PRINT TBRECust(1).Address3
    LOCATE 9, 1: COLOR 11: PRINT STRING$(80, "-")
    LOCATE 10, 1: COLOR 11
    PRINT "Correct Account (Y/N) .. "; : fl = 1: GOSUB 62210
    IF cf = 1 THEN CLOSE : RETURN
    IF in$ = "N" THEN CLOSE : RETURN

    LOCATE 10, 1: PRINT STRING$(79, 32): COLOR 11
    LOCATE 10, 1
    PRINT "Enter Amount Paid ("; USING "######,#.##"; CurBal# + PastDue#; : PRINT ") .... ";
     PRINT TAB(38); ""; : fl = -9: GOSUB 62000
     IF cf = 1 THEN CLOSE : RETURN
     IF LEN(in$) = 0 THEN
      AmountPaid# = CurBal# + PastDue#
      ELSE
      AmountPaid# = VAL(in$)
     END IF
     LOCATE 10, 37: PRINT USING "######,#.##"; AmountPaid#
pr1: LOCATE 11, 1: COLOR 11
     PRINT "Enter (1)-Payment (2)-Release ...... "; : fl = 1
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RETURN
     IF VAL(in$) < 1 OR VAL(in$) > 2 THEN GOTO pr1
     PayType = VAL(in$)
pr2: LOCATE 12, 1: COLOR 11
     PRINT "Enter Description (if any) ......... "; : fl = 15
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RETURN
     Desc$ = UCASE$(in$)
pr3: LOCATE 13, 1: COLOR 11
     PRINT "(A)uto or (M)anually Dist. Payment . "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN CLOSE : RETURN
     IF NOT (in$ = "A" OR in$ = "M") THEN GOTO pr3
     Dist$ = in$
pr4: LOCATE 14, 1: COLOR 11
     PRINT "Payment Date ("; PDate$; ") .............. "; : fl = -6
     GOSUB 62000
     IF NOT (LEN(in$) = 0 OR LEN(in$) = 6) THEN GOTO pr4
     IF LEN(in$) = 0 THEN
      LOCATE 14, 38: PRINT PDate$
      ELSE
      PDate$ = in$
     END IF
pr5: LOCATE 16, 1: COLOR 11
     PRINT "Payment Entered Correctly? (Y/N) .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN CLOSE : RETURN
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO pr5
     IF in$ = "N" THEN GOTO ProcessReal

pr6:
     IF Dist$ = "A" THEN
      GOSUB FilePayment
      CLOSE : RETURN
     END IF
pr7: LOCATE 18, 1: COLOR 11
     PRINT "Enter Year to Distribute Manually (YYYY)(ESC-Cancel Entry) .. "; : fl = -4
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RETURN
     IF LEN(in$) <> 4 THEN GOTO pr7
     IF VAL(in$) > VAL(CurrentYear$) THEN
      LOCATE 19, 1: COLOR 11
      PRINT "Year Entered Exceeds Current Year : Press <ENTER> to Try Again!"; : fl = 1
      GOSUB 62000
      GOTO pr7
     END IF
     ManYear$ = in$
     GOSUB FilePayment
     CLOSE
     RETURN


ProcessPersonal:
    GOSUB FigureBalance
    PaySource$ = "P"
    CLS
    SCREN$ = "4.01.00": CLS : GOSUB 63500
    COLOR 11: LOCATE 1, 32: PRINT "Customer Payment Entry";
    LOCATE 3, 1: PRINT "Personal Property Customer"
    LOCATE 5, 1: COLOR 15
        PRINT TBPPCust(1).CustName; TAB(50); "Current Bal: "; USING "$$######,#.##"; CurBal#
        PRINT TBPPCust(1).Address1; TAB(50); "   Past Bal: "; USING "$$######,#.##"; PastDue#
        PRINT TBPPCust(1).Address2; TAB(50); " Total Owed: "; USING "$$######,#.##"; CurBal# + PastDue#
        PRINT RTRIM$(TBPPCust(1).City); ", "; TBPPCust(1).State; " "; TBPPCust(1).Zip
    LOCATE 9, 1: COLOR 11: PRINT STRING$(80, "-")
    LOCATE 10, 1: COLOR 11
    PRINT "Correct Account (Y/N) .. "; : fl = 1: GOSUB 62210
    IF cf = 1 THEN CLOSE : RETURN
    IF in$ = "N" THEN CLOSE : RETURN

    LOCATE 10, 1: PRINT STRING$(79, 32): COLOR 11
    LOCATE 10, 1
    PRINT "Enter Amount Paid ("; USING "######,#.##"; CurBal# + PastDue#; : PRINT ") .... ";
     PRINT TAB(38); ""; : fl = -9: GOSUB 62000
     IF cf = 1 THEN CLOSE : RETURN
     IF LEN(in$) = 0 THEN
      AmountPaid# = CurBal# + PastDue#
      ELSE
      AmountPaid# = VAL(in$)
     END IF
     LOCATE 10, 37: PRINT USING "######,#.##"; AmountPaid#
ppr1: LOCATE 11, 1: COLOR 11
     PRINT "Enter (1)-Payment (2)-Release ...... "; : fl = 1
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RETURN
     IF VAL(in$) < 1 OR VAL(in$) > 2 THEN GOTO ppr1
     PayType = VAL(in$)
ppr2: LOCATE 12, 1: COLOR 11
     PRINT "Enter Description (if any) ......... "; : fl = 15
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RETURN
     Desc$ = UCASE$(in$)
ppr3: LOCATE 13, 1: COLOR 11
     PRINT "(A)uto or (M)anually Dist. Payment . "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN CLOSE : RETURN
     IF NOT (in$ = "A" OR in$ = "M") THEN GOTO ppr3
     Dist$ = in$
ppr4: LOCATE 14, 1: COLOR 11
     PRINT "Payment Date ("; PDate$; ") .............. "; : fl = -6
     GOSUB 62000
     IF NOT (LEN(in$) = 0 OR LEN(in$) = 6) THEN GOTO ppr4
     IF LEN(in$) = 0 THEN
      LOCATE 14, 38: PRINT PDate$
      ELSE
      PDate$ = in$
     END IF
ppr5: LOCATE 16, 1: COLOR 11
     PRINT "Payment Entered Correctly? (Y/N) .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN CLOSE : RETURN
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO ppr5
     IF in$ = "N" THEN GOTO ProcessPersonal

ppr6:
     IF Dist$ = "A" THEN
      GOSUB FilePayment
      CLOSE : RETURN
     END IF
ppr7: LOCATE 18, 1: COLOR 11
     PRINT "Enter Year to Distribute Manually (YYYY)(ESC-Cancel Entry) .. "; : fl = -4
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RETURN
     IF LEN(in$) <> 4 THEN GOTO ppr7
     IF VAL(in$) > VAL(CurrentYear$) THEN
      LOCATE 19, 1: COLOR 11
      PRINT "Year Entered Exceeds Current Year : Press <ENTER> to Try Again!"; : fl = 1
      GOSUB 62000
      GOTO ppr7
     END IF
     ManYear$ = in$
     GOSUB FilePayment
     CLOSE
     RETURN

FilePayment:
 REDIM PayEdit(1) AS PaymentEdit
 TBEdit = FREEFILE
 OPEN "TBEDIT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBEdit LEN = LEN(PayEdit(1))
 PayEdit(1).PaymentAcct = Record!
 PayEdit(1).PaymentType = PayType
 PayEdit(1).PaymentDate = PDate$
 PayEdit(1).PaymentAmt = AmountPaid#
 PayEdit(1).PaymentDesc = Desc$
 PayEdit(1).PaymentSource = PaySource$  'R or P
 PayEdit(1).PaymentDist = Dist$         'A=Auto  M=Manual
 PayEdit(1).PaymentManYear = ManYear$
 PUT TBEdit, LOF(TBEdit) / LEN(PayEdit(1)) + 1, PayEdit(1)
 RETURN

2000 SCREN$ = "4.02.00": CLS : GOSUB 63500
2020 COLOR 11: LOCATE 1, 28: PRINT "Customer Payment Entry Edit";
2025 LOCATE 3, 1: COLOR 11
     PRINT "Enter the Customer's Account Number (ESC-End) .. "; : fl = -5:
     GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 10
     Record! = VAL(in$)
2100 REDIM PayEdit(1) AS PaymentEdit
     TBEdit = FREEFILE
     OPEN "TBEDIT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBEdit LEN = LEN(PayEdit(1))
     FOR Cnt! = 1 TO LOF(TBEdit) / LEN(PayEdit(1))
     GET TBEdit, Cnt!, PayEdit(1)
      IF PayEdit(1).PaymentAcct = Record! THEN GOSUB EditPayment: CLOSE : GOTO 2000
2100.5 NEXT Cnt!
     FOR ll = 3 TO 23: LOCATE ll, 1: PRINT STRING$(79, 32): NEXT ll
     LOCATE 5, 1: COLOR 11
     PRINT "Payment Entry Not On File! Press <ENTER> to Continue "; : fl = 1
     GOSUB 62000: GOTO 2000
EditPayment:
    FOR ll = 3 TO 23: LOCATE ll, 1: PRINT STRING$(79, 32): NEXT ll
    LOCATE 3, 1: COLOR 11
    PRINT "1- Payment Date (mmddyy) ...... "; : COLOR 15: PRINT PayEdit(1).PaymentDate; : COLOR 11: PRINT TAB(50); "Source (Real or Pers): "; : COLOR 15: PRINT PayEdit(1).PaymentSource
    COLOR 11: PRINT "2- Payment Amount ............. "; : COLOR 15: PRINT USING "$$######,#.##"; PayEdit(1).PaymentAmt
    COLOR 11: PRINT "3- Payment Description ........ "; : COLOR 15: PRINT PayEdit(1).PaymentDesc
    COLOR 11: PRINT "4- Distribution Method (A/M) .. "; : COLOR 15: PRINT PayEdit(1).PaymentDist
    COLOR 11: PRINT "5- Manual Year to Apply ....... "; : COLOR 15: PRINT PayEdit(1).PaymentManYear
    COLOR 11: PRINT "6- Pay Type 1=Paymt 2=Release . "; : COLOR 15: PRINT PayEdit(1).PaymentType
EditPay1:
    LOCATE 10, 1: COLOR 11
    PRINT "Selection Number (<F1>-Next Match : <F2>-Delete)(ESC-Done) .. "; : fl = -1
    GOSUB 62000
    IF cf = 1 THEN CLOSE : GOTO 10
    IF cf = 11 THEN
     GOTO 2100.5
    END IF
    IF cf = 12 THEN
     LOCATE 15, 1: COLOR 15
     PRINT "Are You Sure You Wish to Delete This Entry (Y/N)? "; : fl = 1
     GOSUB 62210
     IF in$ = "Y" THEN
      PayEdit(1).PaymentAcct = 0
      PUT TBEdit, Cnt!, PayEdit(1)
      PRINT "Deleted!!!"; CHR$(7);
      CLOSE
      RETURN
     END IF
    END IF
    ON VAL(in$) GOSUB 2101, 2102, 2103, 2104, 2105, 2106
    PUT TBEdit, Cnt!, PayEdit(1)
    GOTO EditPay1

2101 LOCATE 3, 33: fl = -6: GOSUB 62000
     IF LEN(in$) <> 6 THEN 2101
     PayEdit(1).PaymentDate = in$: RETURN
2102 LOCATE 4, 33: fl = -9: GOSUB 62000: PayEdit(1).PaymentAmt = VAL(in$): RETURN
2103 LOCATE 5, 33: fl = 15: GOSUB 62000: PayEdit(1).PaymentDesc = UCASE$(in$): RETURN
2104 LOCATE 6, 33: fl = 1: GOSUB 62210
     IF NOT (in$ = "A" OR in$ = "M") THEN 2104
     PayEdit(1).PaymentDist = in$: RETURN
2105 LOCATE 7, 33: fl = -4: GOSUB 62000
     IF NOT (LEN(in$) = 0 OR LEN(in$) = 4) THEN 2105
     IF PayEdit(1).PaymentDist = "M" THEN
      IF LEN(in$) <> 4 THEN 2105            'force date if manual dist.
     END IF
     PayEdit(1).PaymentManYear = in$: RETURN
2106 LOCATE 8, 33: fl = -1: GOSUB 62000
     IF VAL(in$) < 1 OR VAL(in$) > 2 THEN 2106
     PayEdit(1).PaymentType = VAL(in$)
     RETURN
3000 SCREN$ = "4.03.00": CLS : GOSUB 63500
3020 COLOR 11: LOCATE 1, 28: PRINT "Print Payment Edit List";
3025 LOCATE 3, 1: COLOR 11
     PRINT "Output Specification: Output to (S)creen or (P)rinter (S,P or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
3090 IF cf = 1 THEN 10
3100 IF NOT (in$ = "S" OR in$ = "P") THEN 3025
3105 DevSpec$ = in$: LOCATE 3, 77: PRINT DevSpec$
3110 LOCATE 10, 1: COLOR 11: PRINT STRING$(79, "="): COLOR 10: LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": COLOR 14
3120 IF DevSpec$ = "S" THEN DEV$ = "SCRN:" ELSE DEV$ = "lpt1:"
3130 OPEN "O", 10, DEV$: WIDTH #10, 255
     IF DevSpec$ = "S" THEN CLS : GOSUB 3900 ELSE GOSUB 3900

     REDIM PayEdit(1) AS PaymentEdit
     TBEdit = FREEFILE
     OPEN "TBEDIT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBEdit LEN = LEN(PayEdit(1))
     IF LOF(TBEdit) / LEN(PayEdit(1)) = 0 THEN
      COLOR 11
      LOCATE 10, 1
      PRINT "Nothing in the File to Print : Press <ENTER> to Continue "; : fl = 1
      GOSUB 62000: CLOSE : GOTO 10
     END IF

     FOR MasterLoop = 1 TO 2        '1=Real 2=Pers

     IF MasterLoop = 1 THEN taxtype$ = "R"
     IF MasterLoop = 2 THEN taxtype$ = "P"

     CLOSE TBFile: GOSUB 60000

     FOR Cnt! = 1 TO LOF(TBEdit) / LEN(PayEdit(1))
      GET TBEdit, Cnt!, PayEdit(1)

     IF PayEdit(1).PaymentSource = "R" AND MasterLoop = 1 THEN
      IF PayEdit(1).PaymentAcct > 0 THEN
        GET TBFile, PayEdit(1).PaymentAcct, TBRECust(1)
        GOSUB PrintLine
      END IF
     END IF

     IF PayEdit(1).PaymentSource = "P" AND MasterLoop = 2 THEN
      IF PayEdit(1).PaymentAcct > 0 THEN
       GET TBFile, PayEdit(1).PaymentAcct, TBPPCust(1)
       GOSUB PrintLine
      END IF
     END IF
     NEXT Cnt!
     GOSUB PrintTotal
     NEXT MasterLoop
     PRINT #10, STRING$(80, "-")
     PRINT #10, "Grand Total of Payments ... "; USING "$$######,#.##"; TotalRealPay# + TotalPersPay#
     PRINT #10, "Grand Total of Releases ... "; USING "$$######,#.##"; TotalRealRel# + TotalPersRel#
     IF DevSpec$ = "S" THEN
      COLOR 11
      LOCATE 23, 1: PRINT "End of Report : Press <ENTER> to Continue "; : fl = 1
      GOSUB 62000
      ELSE
      PRINT #10, CHR$(12);
     END IF
     CLOSE
     GOTO 10

PrintTotal:
IF MasterLoop = 1 THEN
PRINT #10, STRING$(80, "-")
PRINT #10, "Total Real Estate Payments .. "; USING "$$######,#.##"; TotalRealPay#
PRINT #10, "Total Real Estate Releases .. "; USING "$$######,#.##"; TotalRealRel#
ELSE
PRINT #10, STRING$(80, "-")
PRINT #10, "Total Pers Property Payments .. "; USING "$$######,#.##"; TotalPersPay#
PRINT #10, "Total Pers Property Releases .. "; USING "$$######,#.##"; TotalPersRel#
END IF
IF DevSpec$ = "S" THEN
 LOCATE 23, 1: PRINT "Press <ENTER> to Continue "; : fl = 1: GOSUB 62000: CLS : GOSUB 3900: RETURN
END IF
PRINT #10, CHR$(12);
GOSUB 3900
RETURN

PrintLine:
     PRINT #10, PayEdit(1).PaymentAcct;
     IF taxtype$ = "R" THEN
      PRINT #10, TAB(10); LEFT$(TBRECust(1).CustName, 28);
      ELSE
      PRINT #10, TAB(10); LEFT$(TBPPCust(1).CustName, 28);
     END IF
      PRINT #10, TAB(40); PayEdit(1).PaymentDesc; TAB(59); PayEdit(1).PaymentSource;
       PRINT #10, TAB(63); PayEdit(1).PaymentDist; " "; PayEdit(1).PaymentManYear;
        PRINT #10, TAB(70); USING "######,#.##"; PayEdit(1).PaymentAmt
        IF taxtype$ = "R" THEN
          IF PayEdit(1).PaymentType = 1 THEN TotalRealPay# = TotalRealPay# + PayEdit(1).PaymentAmt
          IF PayEdit(1).PaymentType = 2 THEN TotalRealRel# = TotalRealRel# + PayEdit(1).PaymentAmt
         ELSE
          IF PayEdit(1).PaymentType = 1 THEN TotalPersPay# = TotalPersPay# + PayEdit(1).PaymentAmt
          IF PayEdit(1).PaymentType = 2 THEN TotalPersRel# = TotalPersRel# + PayEdit(1).PaymentAmt
        END IF
       lc = lc + 1
       RETURN

3900 PRINT #10, "             Tax System Payment Entry Edit Listing For Date: "; DATE$
     PRINT #10, ""
     PRINT #10, "Acct #"; TAB(10); "Customer Name"; TAB(40); "Descripton"; TAB(58); "SRC"; TAB(62); "Dist/Year"; TAB(73); "Amount"
     PRINT #10, STRING$(80, "-"): lc = 5: RETURN

4000 SCREN$ = "4.04.00": CLS : GOSUB 63500
4020 COLOR 11: LOCATE 1, 28: PRINT "Post Payments to Accounts"
4030 LOCATE 3, 1
     COLOR 11
     PRINT "Press <ENTER> to Begin Posting : Press <ESC> to Cancel Posting Process "; : fl = 1
     GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 10
4100 LOCATE 10, 1: COLOR 11: PRINT STRING$(79, "="): COLOR 10: LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": COLOR 14

     REDIM PayEdit(1) AS PaymentEdit
     TBEdit = FREEFILE
     OPEN "TBEDIT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBEdit LEN = LEN(PayEdit(1))
     IF LOF(TBEdit) / LEN(PayEdit(1)) = 0 THEN
      COLOR 11
      LOCATE 15, 1
      PRINT "Nothing in the File to Post : Press <ENTER> to Continue "; : fl = 1
      GOSUB 62000: CLOSE : GOTO 10
     END IF

     FOR MasterLoop = 1 TO 2
     IF MasterLoop = 1 THEN taxtype$ = "R"
     IF MasterLoop = 2 THEN taxtype$ = "P"

     CLOSE TBFile: CLOSE TBBalFile
     GOSUB 60000: GOSUB 60025

     FOR Cnt! = 1 TO LOF(TBEdit) / LEN(PayEdit(1))
      GET TBEdit, Cnt!, PayEdit(1)
      Penalty# = 0: Interest# = 0

     IF PayEdit(1).PaymentSource = "R" AND MasterLoop = 1 THEN
      IF PayEdit(1).PaymentAcct > 0 THEN
        GET TBFile, PayEdit(1).PaymentAcct, TBRECust(1)
        GET TBBalFile, PayEdit(1).PaymentAcct, TBRECustBal(1)
        GOSUB FigureBalance
        GOSUB PostPayment
        PUT TBFile, PayEdit(1).PaymentAcct, TBRECust(1)
        PUT TBBalFile, PayEdit(1).PaymentAcct, TBRECustBal(1)
        GOSUB FigureBalance
        GOSUB PostRealHistory
        PayEdit(1).PaymentAcct = 0
        PUT TBEdit, Cnt!, PayEdit(1)
       END IF
     END IF

     IF PayEdit(1).PaymentSource = "P" AND MasterLoop = 2 THEN
      IF PayEdit(1).PaymentAcct > 0 THEN
       GET TBFile, PayEdit(1).PaymentAcct, TBPPCust(1)
       GET TBBalFile, PayEdit(1).PaymentAcct, TBPPCustBal(1)
       GOSUB FigureBalance
       GOSUB PostPayment
       PUT TBFile, PayEdit(1).PaymentAcct, TBPPCust(1)
       PUT TBBalFile, PayEdit(1).PaymentAcct, TBPPCustBal(1)
       GOSUB FigureBalance
       GOSUB PostPersHistory
       PayEdit(1).PaymentAcct = 0
       PUT TBEdit, Cnt!, PayEdit(1)
      END IF
     END IF
    NEXT Cnt!
    NEXT MasterLoop
    CLOSE
    COLOR 11: LOCATE 15, 1: PRINT "All Items Posted!"; CHR$(7); CHR$(7);
    FileNam$ = Path$ + "tbedit.dat"
    KILL FileNam$
    GOTO 10
PostRealHistory:
 GET TBFile, PayEdit(1).PaymentAcct, TBRECust(1)
 GET TBBalFile, PayEdit(1).PaymentAcct, TBRECustBal(1)
 GOSUB FigureBalance
 AcctBalance# = CurBal# + PastDue#
 GOSUB 61000

 PostDate$ = LEFT$(PayEdit(1).PaymentDate, 2) + "-" + MID$(PayEdit(1).PaymentDate, 3, 2) + "-" + MID$(PayEdit(1).PaymentDate, 5, 2)
 IF VAL(RIGHT$(PostDate$, 2)) < 75 THEN
  PostDate$ = LEFT$(PostDate$, 6) + "20" + RIGHT$(PostDate$, 2)
  ELSE
  PostDate$ = LEFT$(PostDate$, 6) + "19" + RIGHT$(PostDate$, 2)
 END IF
 TBTrans(1).TransDate = Date2Num%(PostDate$)


 TBTrans(1).TransAmt = PayEdit(1).PaymentAmt
 TBTrans(1).TrTaxYear = PayEdit(1).PaymentManYear
 TBTrans(1).TrDesc = PayEdit(1).PaymentDesc
 IF PayEdit(1).PaymentType = 1 THEN
  TBTrans(1).TrType = 2                   '1=bill 2=payment 3=INT  4= PEN 5=release 6=BEG BAL
 ELSE
  TBTrans(1).TrType = 5
 END IF
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
 PostDate$ = LEFT$(PayEdit(1).PaymentDate, 2) + "-" + MID$(PayEdit(1).PaymentDate, 3, 2) + "-" + MID$(PayEdit(1).PaymentDate, 5, 2)
 IF VAL(RIGHT$(PostDate$, 2)) < 75 THEN
  PostDate$ = LEFT$(PostDate$, 6) + "20" + RIGHT$(PostDate$, 2)
  ELSE
  PostDate$ = LEFT$(PostDate$, 6) + "19" + RIGHT$(PostDate$, 2)
 END IF
 TBTrans(1).TransDate = Date2Num%(PostDate$)

 
 TBTrans(1).TransAmt = PayEdit(1).PaymentAmt
 TBTrans(1).TrTaxYear = PayEdit(1).PaymentManYear
 TBTrans(1).TrDesc = PayEdit(1).PaymentDesc
 IF PayEdit(1).PaymentType = 1 THEN
  TBTrans(1).TrType = 2
  ELSE
  TBTrans(1).TrType = 5
 END IF
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
5000 RUN "TBPAY1"
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

62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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
63000 RETURN
63500 COLOR 10: LOCATE 1, 1: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN
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
  FOR ll = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, ll
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO n1
n3:
  NEXT ll
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  ggo = 0
  RETURN

n1:
GET TBFile, CVS(Mrec$), TBRECust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBRECust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
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
  FOR ll = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, ll
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO PPn1
PPn3:
  NEXT ll
  LOCATE 8, 1: PRINT STRING$(79, 32)
  ggo = 0
  RETURN

PPn1:
GET TBFile, CVS(Mrec$), TBPPCust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBPPCust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
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

          CASE ELSE
               CLS
               LOCATE 15, 1: COLOR 28: PRINT "Error Message: "; : COLOR 15: PRINT "Error "; Number; " Has Occured in Line "; ERL
               PRINT "Write Down Message and Hit <ESC> Key to exit."
               fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
            
     END SELECT

