10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbcustrep
30 REM Customer Report Menu
40 REM Last Revision Date: 8-24-87
50 REM by Nick Picerno
   DECLARE FUNCTION num2date$ (Num%)
   DECLARE FUNCTION Date2Num% (Dat$)
60 CLS : CLEAR

    '$INCLUDE: 'lcTAX.BI'

     WIDTH LPRINT 200
    
     ' TBPath$ = "F:\"                    ' change to "F:\" Before compiling
     TBPath$ = "c:\QB45\LC\"                    ' change to "F:\" Before compiling

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

     DEFDBL A-C, I, S
100 REM begin program
110 SCREN$ = "2.07.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
130 LOCATE 3, 1: COLOR 11
    PRINT
200 PRINT TAB(31); "Tax Billing System"
210 PRINT TAB(30); "Delinquent Processing"
220 PRINT : PRINT
    PRINT TAB(25); "1- Print Demand Letters - Individual"
    PRINT
    PRINT TAB(25); "2- Print Demand Letters - All"
    PRINT
    PRINT TAB(25); "3- Create Letter Format"
    PRINT
310 LOCATE 15, 1
    COLOR 11: PRINT TAB(25); "Selection No. (ESC-Master Menu) : "; : fl = -1: GOSUB 62000
320 IF cf = 1 THEN CLOSE : RUN "tbcustre"
325 IF LEN(in$) = 0 THEN 310
330 IF VAL(in$) < 1 OR VAL(in$) > 3 THEN 310
340 ON VAL(in$) GOTO 6000, 7000, 8000

6000 REM notices
     SCREN$ = "2.07.04": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 32
     PRINT "Demand Letters"
6010 LOCATE 3, 1: COLOR 11
     PRINT "Enter Date of Letter to Print (ESC-Menu) ... "; : fl = 18
     GOSUB 62000
     IF cf = 1 THEN 10
     today$ = in$
6050 LOCATE 7, 1: COLOR 11
     PRINT "Enter Letter Format to Use .. "; : fl = -2
     GOSUB 62000
     IF cf = 1 THEN 10
     FORMAT = VAL(in$)
     IF FORMAT < 1 THEN 6000
     OPEN "TBLTRFMT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 1301
     FIELD #11, 1 AS valid$, 65 AS l1$, 65 AS l2$, 65 AS l3$, 65 AS l4$, 65 AS l5$, 65 AS l6$, 65 AS l7$, 65 AS l8$, 65 AS l9$, 65 AS l10$, 65 AS l11$, 65 AS l12$, 65 AS l13$, 65 AS l14$, 65 AS l15$, 65 AS l16$, 65 AS l17$, 65 AS l18$, 65 AS l19$,  _
65 AS l20$
     IF FORMAT > LOF(11) / 1300 THEN
      LOCATE 10, 1: COLOR 11: PRINT "Invalid Format : Please Setup & Try Again!"
      LOCATE 11, 1: PRINT "Press <ENTER> to Continue "; : fl = 1: GOSUB 62000: CLOSE
      LOCATE 10, 1: PRINT STRING$(79, " ")
      LOCATE 11, 1: PRINT STRING$(79, " "): GOTO 6050
     END IF
     GET 11, FORMAT
     IF valid$ <> "Y" THEN
      LOCATE 10, 1: COLOR 11: PRINT "Invalid Format : Please Setup & Try Again!"
      LOCATE 11, 1: PRINT "Press <ENTER> to Continue "; : fl = 1: GOSUB 62000: CLOSE
      LOCATE 10, 1: PRINT STRING$(79, " ")
      LOCATE 11, 1: PRINT STRING$(79, " "): GOTO 6050
     END IF
   
6100 LOCATE 15, 1: PRINT "Ready to Print Using Format #"; FORMAT
     LOCATE 16, 1: PRINT "Press <ENTER> to Begin or <ESC> to Cancel .. "; : fl = 1
     GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 10
6200 COLOR 10
6300 FOR ll = 3 TO 23: LOCATE ll, 1: PRINT STRING$(79, 32): NEXT ll
     LOCATE 3, 1: COLOR 11
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN CLOSE : GOTO 10
     IF NOT (in$ = "R" OR in$ = "P" OR in$ = "A") THEN 6300
     taxtype$ = UCASE$(in$)
6330 LOCATE 4, 1: COLOR 11
     PRINT "Enter Customer Account Number (TAB-Last Name : ESC-Menu) .. "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN
      CLOSE : GOTO 10
     END IF
     IF cf = 2 THEN
      GOSUB Namesearch
      IF ggo = 1 THEN 6375
      IF ggo = 0 THEN LOCATE 10, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 62000: GOTO 6300
     END IF
     Record! = VAL(in$)
6375 CLOSE TBFile: CLOSE TBBalFile
      
      CLOSE 11: OPEN "TBLTRFMT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 1301
       FIELD #11, 1 AS valid$, 65 AS l1$, 65 AS l2$, 65 AS l3$, 65 AS l4$, 65 AS l5$, 65 AS l6$, 65 AS l7$, 65 AS l8$, 65 AS l9$, 65 AS l10$, 65 AS l11$, 65 AS l12$, 65 AS l13$, 65 AS l14$, 65 AS l15$, 65 AS l16$, 65 AS l17$, 65 AS l18$, 65 AS l19$ _
, 65 AS l20$
       GET 11, FORMAT
      CLOSE 10
      OPEN "O", 10, "dem1ltr.rep": WIDTH #10, 255
      GOSUB 60000
      GOSUB 60025
     IF taxtype$ = "R" THEN
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBRECust(1)) THEN 6330
       LOCK #TBFile, Record!
       GET TBFile, Record!, TBRECust(1)
       GET TBBalFile, Record!, TBRECustBal(1)
       IF TBRECust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 6300
       ELSE
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBPPCust(1)) THEN 6330
       LOCK #TBFile, Record!
       GET TBFile, Record!, TBPPCust(1)
       GET TBBalFile, Record!, TBPPCustBal(1)
       IF TBPPCust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 6330
     END IF
     GOSUB FigureBalance
     GOSUB PrintLetter
     CLOSE
     SHELL "Type dem1ltr.rep > lpt1:"
     GOTO 6330
   
PrintLetter:
GOSUB townHEADING
GOSUB address
PRINT #10,
PRINT #10, TAB(5); l1$
PRINT #10,
PRINT #10, TAB(5); l2$
PRINT #10, TAB(5); l3$
PRINT #10, TAB(5); l4$
PRINT #10, TAB(5); l5$
PRINT #10, TAB(5); l6$
PRINT #10, TAB(5); l7$
PRINT #10, TAB(5); l8$
PRINT #10, TAB(5); l9$
PRINT #10, TAB(5); l10$
PRINT #10, TAB(5); l11$
PRINT #10, TAB(5); l12$
PRINT #10, TAB(5); l13$
PRINT #10, TAB(5); l14$
PRINT #10, TAB(5); l15$
PRINT #10,
GOSUB LCHeading
GOSUB LDetail
RETURN
  

LDetail:
     IF CurBal# = 0 THEN GOTO lDETAIL1
      IF taxtype$ = "R" THEN
      PRINT #10, TAB(10); TBRECust(1).CurYrBillNumber;
       PRINT #10, TAB(20); LEFT$(TBRECust(1).PropDesc1, 18);
       PRINT #10, TAB(40); USING "$$#####.##"; TBRECust(1).CurYrTaxAmount;
        PRINT #10, TAB(50); USING "$$#####.##"; (TBRECust(1).CurYrIntAmount + TBRECust(1).CurYrPenAmount);
         PRINT #10, TAB(60); USING "$$#####.##"; CurBal#
         runcur# = runcur# + (TBRECust(1).CurYrTaxAmount)
         runint# = runint# + TBRECust(1).CurYrIntAmount + TBRECust(1).CurYrPenAmount

         ELSE
      PRINT #10, TAB(10); TBPPCust(1).CurYrBillNumber;
       PRINT #10, TAB(20); LEFT$(PDESC$, 18);
       PRINT #10, TAB(40); USING "$$#####.##"; TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax;
        PRINT #10, TAB(50); USING "$$#####.##"; TBPPCust(1).CurYrIntAmount + TBPPCust(1).CurYrPenAmount;
         PRINT #10, TAB(60); USING "$$#####.##"; CurBal#
         runcur# = runcur# + (TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax)
         runint# = runint# + (TBPPCust(1).CurYrIntAmount + TBPPCust(1).CurYrPenAmount)

       END IF
         runtotal# = runtotal# + CurBal#
         printed = 1
        
lDETAIL1:
IF taxtype$ = "R" THEN
    FOR ll = 1 TO 10
    IF (TBRECustBal(1).TaxAmt(ll) + TBRECustBal(1).IntAmt(ll) + TBRECustBal(1).PenAmt(ll)) > 0 THEN
      PRINT #10, TAB(10); TBRECustBal(1).TaxNotice(ll); TAB(20); "Real Property";
      PRINT #10, TAB(40); USING "$$#####.##"; (TBRECustBal(1).TaxAmt(ll));
      PRINT #10, TAB(50); USING "$$#####.##"; (TBRECustBal(1).IntAmt(ll) + TBRECustBal(1).PenAmt(ll));
      PRINT #10, TAB(60); USING "$$#####.##"; (TBRECustBal(1).TaxAmt(ll) + TBRECustBal(1).IntAmt(ll) + TBRECustBal(1).PenAmt(ll))
      runtotal# = runtotal# + (TBRECustBal(1).TaxAmt(ll) + TBRECustBal(1).IntAmt(ll) + TBRECustBal(1).PenAmt(ll))
      runcur# = runcur# + TBRECustBal(1).TaxAmt(ll)
      runint# = runint# + (TBRECustBal(1).IntAmt(ll) + TBRECustBal(1).PenAmt(ll))
      printed = printed + 1
     END IF
      NEXT ll
   ELSE
    FOR ll = 1 TO 10
    IF (TBPPCustBal(1).PerPropTaxAmt(ll) + TBPPCustBal(1).FarmEquipAmt(ll) + TBPPCustBal(1).MachToolsAmt(ll) + TBPPCustBal(1).MerchCapAmt(ll) + TBPPCustBal(1).MobileHomeAmt(ll) + TBPPCustBal(1).LateFeeAmt(ll)) > 0 THEN
      PRINT #10, TAB(10); TBPPCustBal(1).TaxNotice(ll); TAB(20); "Pers Property";
      PRINT #10, TAB(40); USING "$$#####.##"; (TBPPCustBal(1).PerPropTaxAmt(ll) + TBPPCustBal(1).FarmEquipAmt(ll) + TBPPCustBal(1).MachToolsAmt(ll) + TBPPCustBal(1).MerchCapAmt(ll) + TBPPCustBal(1).MobileHomeAmt(ll) + TBPPCustBal(1).LateFeeAmt(ll)); _

      PRINT #10, TAB(50); USING "$$#####.##"; (TBPPCustBal(1).IntAmt(ll) + TBPPCustBal(1).PenAmt(ll));
      PRINT #10, TAB(60); USING "$$#####.##"; (TBPPCustBal(1).PerPropTaxAmt(ll) + TBPPCustBal(1).FarmEquipAmt(ll) + TBPPCustBal(1).MachToolsAmt(ll) + TBPPCustBal(1).MerchCapAmt(ll) + TBPPCustBal(1).MobileHomeAmt(ll) + TBPPCustBal(1).LateFeeAmt(ll))  _
+ TBPPCustBal(1).IntAmt(ll) + TBPPCustBal(1).PenAmt(ll)
      runtotal# = runtotal# + (TBPPCustBal(1).PerPropTaxAmt(ll) + TBPPCustBal(1).FarmEquipAmt(ll) + TBPPCustBal(1).MachToolsAmt(ll) + TBPPCustBal(1).MerchCapAmt(ll) + TBPPCustBal(1).MobileHomeAmt(ll) + TBPPCustBal(1).LateFeeAmt(ll)) + TBPPCustBal(1) _
.IntAmt(ll) + TBPPCustBal(1).PenAmt(ll)
      runcur# = runcur# + (TBPPCustBal(1).PerPropTaxAmt(ll) + TBPPCustBal(1).FarmEquipAmt(ll) + TBPPCustBal(1).MachToolsAmt(ll) + TBPPCustBal(1).MerchCapAmt(ll) + TBPPCustBal(1).MobileHomeAmt(ll) + TBPPCustBal(1).LateFeeAmt(ll))
      runint# = runint# + (TBPPCustBal(1).IntAmt(ll) + TBPPCustBal(1).PenAmt(ll))
      printed = printed + 1
     END IF
      NEXT ll
     END IF
      GOSUB lfinal
6500 runint# = 0: runtotal# = 0: runcur# = 0
     RETURN

LCHeading:
 PRINT #10, TAB(10); "RECEIPT   DESCRIPTION               TAX    PENALTIES   TOTAL"
 PRINT #10, TAB(10); "------------------------------------------------------------"
 RETURN

lfinal:
IF runtotal# = 0 THEN PRINT #10, CHR$(12); : RETURN
PRINT #10,
 PRINT #10, TAB(16); "ACCOUNT TOTAL"; TAB(40); USING "$$#####.##"; runcur#;
  PRINT #10, TAB(50); USING "$$#####.##"; runint#;
   PRINT #10, TAB(60); USING "$$#####.##"; runtotal#
PRINT #10, : PRINT #10,
PRINT #10, TAB(5); l16$
PRINT #10, TAB(5); l17$
PRINT #10, TAB(5); l18$
PRINT #10, TAB(5); l19$
PRINT #10, TAB(5); l20$
PRINT #10, CHR$(12);
RETURN


townHEADING:
PRINT #10,
PRINT #10, TAB(5); today$
PRINT #10,
lc = 3
RETURN
                          
address:
IF taxtype$ = "R" THEN
 PRINT #10,
 PRINT #10, TAB(5); TBRECust(1).CustName
 PRINT #10, TAB(5); TBRECust(1).Address1
 PRINT #10, TAB(5); TBRECust(1).Address2
 PRINT #10, TAB(5); TBRECust(1).Address3
 PRINT #10,
 lc = lc + 6
ELSE
PRINT #10,
PRINT #10, TAB(5); TBPPCust(1).CustName
PRINT #10, TAB(5); TBPPCust(1).Address1
PRINT #10, TAB(5); TBPPCust(1).Address2
PRINT #10, TAB(5); RTRIM$(TBPPCust(1).City); ", "; TBPPCust(1).State; " "; TBPPCust(1).Zip
PRINT #10,
lc = lc + 6
END IF
RETURN


7000 SCREN$ = "2.07.04": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 32
     PRINT "Demand Letters - All"
7010 LOCATE 3, 1: COLOR 11
     PRINT "Enter Date of Letter to Print (ESC-Menu) ... "; : fl = 18
     GOSUB 62000
     IF cf = 1 THEN 10
     today$ = in$
7050 LOCATE 7, 1: COLOR 11
     PRINT "Enter Letter Format to Use .. "; : fl = -2
     GOSUB 62000
     IF cf = 1 THEN 10
     FORMAT = VAL(in$)
     IF FORMAT < 1 THEN 7000
     OPEN "TBLTRFMT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 1301
     FIELD #11, 1 AS valid$, 65 AS l1$, 65 AS l2$, 65 AS l3$, 65 AS l4$, 65 AS l5$, 65 AS l6$, 65 AS l7$, 65 AS l8$, 65 AS l9$, 65 AS l10$, 65 AS l11$, 65 AS l12$, 65 AS l13$, 65 AS l14$, 65 AS l15$, 65 AS l16$, 65 AS l17$, 65 AS l18$, 65 AS l19$,  _
65 AS l20$
     IF FORMAT > LOF(11) / 1301 THEN
      LOCATE 10, 1: COLOR 11: PRINT "Invalid Format : Please Setup & Try Again!"
      LOCATE 11, 1: PRINT "Press <ENTER> to Continue "; : fl = 1: GOSUB 62000: CLOSE
      LOCATE 10, 1: PRINT STRING$(79, " ")
      LOCATE 11, 1: PRINT STRING$(79, " "): GOTO 7050
     END IF
     GET 11, FORMAT
     IF valid$ <> "Y" THEN
      LOCATE 10, 1: COLOR 11: PRINT "Invalid Format : Please Setup & Try Again!"
      LOCATE 11, 1: PRINT "Press <ENTER> to Continue "; : fl = 1: GOSUB 62000: CLOSE
      LOCATE 10, 1: PRINT STRING$(79, " ")
      LOCATE 11, 1: PRINT STRING$(79, " "): GOTO 7050
     END IF

7100 LOCATE 15, 1: PRINT "Ready to Print Using Format #"; FORMAT
     LOCATE 16, 1: PRINT "Press <ENTER> to Begin or <ESC> to Cancel .. "; : fl = 1
     GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 10
7200 COLOR 10
7300 FOR ll = 3 TO 23: LOCATE ll, 1: PRINT STRING$(79, 32): NEXT ll
     LOCATE 3, 1: COLOR 11
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN CLOSE : GOTO 10
     IF NOT (in$ = "R" OR in$ = "P") THEN 6300
     taxtype$ = UCASE$(in$)


     IF taxtype$ = "R" THEN
        OPEN TBPath$ + "TBREname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS mrec$: EndofFile = LOF(4) / 24
        ELSE
        OPEN TBPath$ + "TBPPname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS mrec$: EndofFile = LOF(4) / 24
      END IF

     OPEN "O", 10, "demltr.rep": WIDTH #10, 255

     GOSUB 60000: GOSUB 60025
     FOR Cnt = 1 TO EndofFile
      GET 4, Cnt
      Record! = CVS(mrec$)
      LOCATE 15, 1: COLOR 11
      PRINT "Processing Account # "; : COLOR 15: PRINT Record!
     
     IF taxtype$ = "R" THEN
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBRECust(1)) THEN 7330
       LOCK #TBFile, Record!
       GET TBFile, Record!, TBRECust(1)
       GET TBBalFile, Record!, TBRECustBal(1)
       IF TBRECust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 6300
       ELSE
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBPPCust(1)) THEN 6330
       LOCK #TBFile, Record!
       GET TBFile, Record!, TBPPCust(1)
       GET TBBalFile, Record!, TBPPCustBal(1)
       IF TBPPCust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 6330
     END IF
     GOSUB FigureBalance
     IF CurBal# <> 0 OR PastDue# <> 0 THEN
     GOSUB PrintLetter
     END IF
7330 NEXT Cnt
     CLOSE
     SHELL "Type demltr.rep > lpt1:"
     GOTO 10

8000 REM notices
     SCREN$ = "2.07.06": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 32
     PRINT "Letter Format"
     LOCATE 3, 1: COLOR 11
     PRINT "Enter Format Number (ESC-Menu) .. "; : fl = -2
     GOSUB 62000: IF cf = 1 THEN 10
     FORMAT = VAL(in$)
     OPEN "TBLTRFMT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 1301
     FIELD #1, 1 AS valid$, 65 AS l1$, 65 AS l2$, 65 AS l3$, 65 AS l4$, 65 AS l5$, 65 AS l6$, 65 AS l7$, 65 AS l8$, 65 AS l9$, 65 AS l10$, 65 AS l11$, 65 AS l12$, 65 AS l13$, 65 AS l14$, 65 AS l15$, 65 AS l16$, 65 AS l17$, 65 AS l18$, 65 AS l19$, 65 _
 AS l20$
8100 IF FORMAT > LOF(1) / 1300 THEN
       GOSUB createit
       GOTO 8000
     END IF
     GET 1, FORMAT
     IF valid$ <> "Y" THEN
      COLOR 11: PRINT "Invalid Format Setup "
      COLOR 11: PRINT "Do you wish to Setup Now (Y/N) .. "; : fl = 1
      GOSUB 62210: IF in$ = "Y" THEN GOSUB createit: CLOSE : GOTO 8000
      CLOSE : GOTO 8000
     END IF


GOSUB FORMAT
LOCATE 17, 1
PRINT "E 16-"
PRINT "N 17-"
PRINT "D 18-"
PRINT "' 19-"
PRINT "G 20-"

LOCATE 1, 1: COLOR 11: PRINT STRING$(79, 32): LOCATE 1, 1
PRINT "Format # "; FORMAT
COLOR 15
LOCATE 2, 10: PRINT l1$
LOCATE 3, 10: PRINT l2$
LOCATE 4, 10: PRINT l3$
LOCATE 5, 10: PRINT l4$
LOCATE 6, 10: PRINT l5$
LOCATE 7, 10: PRINT l6$
LOCATE 8, 10: PRINT l7$
LOCATE 9, 10: PRINT l8$
LOCATE 10, 10: PRINT l9$
LOCATE 11, 10: PRINT l10$
LOCATE 12, 10: PRINT l11$
LOCATE 13, 10: PRINT l12$
LOCATE 14, 10: PRINT l13$
LOCATE 15, 10: PRINT l14$
LOCATE 16, 10: PRINT l15$
LOCATE 17, 10: PRINT l16$
LOCATE 18, 10: PRINT l17$
LOCATE 19, 10: PRINT l18$
LOCATE 20, 10: PRINT l19$
LOCATE 21, 10: PRINT l20$
8200 LOCATE 23, 1: COLOR 11: PRINT "Selection # to Change (F1-Remove) (ESC-End) .. "; : fl = -2: GOSUB 62000
IF cf = 1 THEN PUT 1, FORMAT: CLOSE : GOTO 10
IF cf = 11 THEN LSET valid$ = "N": PUT 1, FORMAT: CLOSE : CLS : PRINT "Deleted": PRINT CHR$(7); : GOTO 10
IF VAL(in$) < 1 OR VAL(in$) > 20 THEN 8200
ON VAL(in$) GOSUB 8201, 8202, 8203, 8204, 8205, 8206, 8207, 8208, 8209, 8210, 8211, 8212, 8213, 8214, 8215, 8216, 8217, 8218, 8219, 8220
PUT 1, FORMAT: GOTO 8200

8201 LOCATE 2, 10: fl = 65: GOSUB 62000: LSET l1$ = in$: RETURN
8202 LOCATE 3, 10: fl = 65: GOSUB 62000: LSET l2$ = in$: RETURN
8203 LOCATE 4, 10: fl = 65: GOSUB 62000: LSET l3$ = in$: RETURN
8204 LOCATE 5, 10: fl = 65: GOSUB 62000: LSET l4$ = in$: RETURN
8205 LOCATE 6, 10: fl = 65: GOSUB 62000: LSET l5$ = in$: RETURN
8206 LOCATE 7, 10: fl = 65: GOSUB 62000: LSET l6$ = in$: RETURN
8207 LOCATE 8, 10: fl = 65: GOSUB 62000: LSET l7$ = in$: RETURN
8208 LOCATE 9, 10: fl = 65: GOSUB 62000: LSET l8$ = in$: RETURN
8209 LOCATE 10, 10: fl = 65: GOSUB 62000: LSET l9$ = in$: RETURN
8210 LOCATE 11, 10: fl = 65: GOSUB 62000: LSET l10$ = in$: RETURN
8211 LOCATE 12, 10: fl = 65: GOSUB 62000: LSET l11$ = in$: RETURN
8212 LOCATE 13, 10: fl = 65: GOSUB 62000: LSET l12$ = in$: RETURN
8213 LOCATE 14, 10: fl = 65: GOSUB 62000: LSET l13$ = in$: RETURN
8214 LOCATE 15, 10: fl = 65: GOSUB 62000: LSET l14$ = in$: RETURN
8215 LOCATE 16, 10: fl = 65: GOSUB 62000: LSET l15$ = in$: RETURN
8216 LOCATE 17, 10: fl = 65: GOSUB 62000: LSET l16$ = in$: RETURN
8217 LOCATE 18, 10: fl = 65: GOSUB 62000: LSET l17$ = in$: RETURN
8218 LOCATE 19, 10: fl = 65: GOSUB 62000: LSET l18$ = in$: RETURN
8219 LOCATE 20, 10: fl = 65: GOSUB 62000: LSET l19$ = in$: RETURN
8220 LOCATE 21, 10: fl = 65: GOSUB 62000: LSET l20$ = in$: RETURN



createit:
GOSUB FORMAT
LOCATE 2, 10: fl = 65: GOSUB 62000: LSET l1$ = in$
LOCATE 3, 10: fl = 65: GOSUB 62000: LSET l2$ = in$
LOCATE 4, 10: fl = 65: GOSUB 62000: LSET l3$ = in$
LOCATE 5, 10: fl = 65: GOSUB 62000: LSET l4$ = in$
LOCATE 6, 10: fl = 65: GOSUB 62000: LSET l5$ = in$
LOCATE 7, 10: fl = 65: GOSUB 62000: LSET l6$ = in$
LOCATE 8, 10: fl = 65: GOSUB 62000: LSET l7$ = in$
LOCATE 9, 10: fl = 65: GOSUB 62000: LSET l8$ = in$
LOCATE 10, 10: fl = 65: GOSUB 62000: LSET l9$ = in$
LOCATE 11, 10: fl = 65: GOSUB 62000: LSET l10$ = in$
LOCATE 12, 10: fl = 65: GOSUB 62000: LSET l11$ = in$
LOCATE 13, 10: fl = 65: GOSUB 62000: LSET l12$ = in$
LOCATE 14, 10: fl = 65: GOSUB 62000: LSET l13$ = in$
LOCATE 15, 10: fl = 65: GOSUB 62000: LSET l14$ = in$
LOCATE 16, 10: fl = 65: GOSUB 62000: LSET l15$ = in$
LOCATE 17, 10: fl = 65: GOSUB 62000: LSET l16$ = in$
LOCATE 18, 10: fl = 65: GOSUB 62000: LSET l17$ = in$
LOCATE 19, 10: fl = 65: GOSUB 62000: LSET l18$ = in$
LOCATE 20, 10: fl = 65: GOSUB 62000: LSET l19$ = in$
LOCATE 21, 10: fl = 65: GOSUB 62000: LSET l20$ = in$
LSET valid$ = "Y"
PUT 1, FORMAT
CLOSE 1
RETURN



FORMAT:
CLS
COLOR 11
PRINT "Create the Form Letter :You Have 15 lines of main body & 5 lines of end body"
PRINT "M  1-"
PRINT "A  2-"
PRINT "I  3-"
PRINT "N  4-"
PRINT "   5-"
PRINT "B  6-"
PRINT "O  7-"
PRINT "D  8-"
PRINT "Y  9-"
PRINT "  10-"
PRINT "  11-"
PRINT "  12-"
PRINT "  13-"
PRINT "  14-"
PRINT "  15-"
PRINT "E  1-"
PRINT "N  2-"
PRINT "D  3-"
PRINT "'  4-"
PRINT "G  5-"
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
  OPEN TBPath$ + "TBRENAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
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
  OPEN TBPath$ + "TBPPNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
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




60000  IF taxtype$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBPath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN

60025 TBBalFile = FREEFILE
      IF taxtype$ = "R" THEN
      OPEN TBPath$ + "TBREBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBPath$ + "TBPPBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF
60030 RETURN
62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(z$) <> 2 THEN 62065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN cf = 12: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 61 THEN cf = 13: RETURN

62065 IF ASC(z$) = 0 THEN 62030
62070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
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
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN
networkerror:
     REM ERROR CODES
     number = ERR
     SELECT CASE ERR
          CASE IS = 24, 25
               LOCATE 22, 1: PRINT STRING$(79, 32): LOCATE 22, 1: COLOR 28
               PRINT "Error Message:"; : COLOR 15: PRINT "Make Sure Printer Is ON-LINE:Press ENTER to Resume Printing";
               fl = 1: GOSUB 62000: LOCATE 22, 1: PRINT STRING$(79, 32)
               RESUME
          CASE IS = 27
               LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 28
               PRINT "Error Message:"; : COLOR 15: PRINT "Printer Is OUT OF PAPER! Load Paper : Press ENTER to Resume Printing";
               fl = 1: GOSUB 62000: LOCATE 23, 1: PRINT STRING$(79, 32)
               RESUME

          CASE ELSE
               CLS
               LOCATE 15, 1: COLOR 28: PRINT "Error Message: "; : COLOR 15: PRINT "Error "; number; " Has Occured in Line "; ERL
               PRINT "Write Down Message and Hit <ESC> Key to exit."
               fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
            
     END SELECT


