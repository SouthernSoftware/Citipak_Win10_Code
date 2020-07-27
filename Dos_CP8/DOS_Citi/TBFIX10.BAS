DECLARE SUB Cleardown ()
DECLARE FUNCTION num2date$ (Num%)
DECLARE FUNCTION Date2Num% (Dat$)

'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'Qscr.BI'
'$INCLUDE: 'SetCnf.BI'
'$INCLUDE: 'lcTAX.BI'

     

     TBPath$ = "F:\"                    ' change to "F:\" Before compiling
   
     DIM SHARED yrdesc(11)
     REDIM TBLog(1) AS TBLog
     DIM SHARED TBRECust(1) AS TBRECust
     DIM SHARED TBRECustBal(1) AS TBREBAL
     DIM SHARED TBPPCust(1) AS TBPPCust
     DIM SHARED TBPPCustBal(1) AS TBPPBal
     DIM SHARED TBTrans(1) AS TBTrans
     DIM SHARED TBRECPT(1) AS DayReceiptsType

     PDATE$ = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 2)
     CLS

     PRINT "Working ..."

 'Fix Receipts File First
     TBRFile = FREEFILE
     OPEN TBPath$ + "TBRECFIL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBRFile LEN = LEN(TBRECPT(1))
     FOR Cnt& = 1 TO LOF(TBRFile) / LEN(TBRECPT(1))
     GET TBRFile, Cnt&, TBRECPT(1)

   IF TBRECPT(1).Date = -29219 THEN
    TBRECPT(1).Date = Date2Num%("01-01-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF

   IF TBRECPT(1).Date = -29218 THEN
    TBRECPT(1).Date = Date2Num%("01-02-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF

   IF TBRECPT(1).Date = -29217 THEN
    TBRECPT(1).Date = Date2Num%("01-03-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF

   IF TBRECPT(1).Date = -29216 THEN
    TBRECPT(1).Date = Date2Num%("01-04-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF

   IF TBRECPT(1).Date = -29215 THEN
    TBRECPT(1).Date = Date2Num%("01-05-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29214 THEN
    TBRECPT(1).Date = Date2Num%("01-06-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29213 THEN
    TBRECPT(1).Date = Date2Num%("01-07-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29212 THEN
    TBRECPT(1).Date = Date2Num%("01-08-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29211 THEN
    TBRECPT(1).Date = Date2Num%("01-09-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29210 THEN
    TBRECPT(1).Date = Date2Num%("01-10-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF

   IF TBRECPT(1).Date = -29209 THEN
    TBRECPT(1).Date = Date2Num%("01-11-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29208 THEN
    TBRECPT(1).Date = Date2Num%("01-12-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29207 THEN
    TBRECPT(1).Date = Date2Num%("01-13-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29206 THEN
    TBRECPT(1).Date = Date2Num%("01-14-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29205 THEN
    TBRECPT(1).Date = Date2Num%("01-15-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29204 THEN
    TBRECPT(1).Date = Date2Num%("01-16-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29203 THEN
    TBRECPT(1).Date = Date2Num%("01-17-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29202 THEN
    TBRECPT(1).Date = Date2Num%("01-18-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29201 THEN
    TBRECPT(1).Date = Date2Num%("01-19-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29200 THEN
    TBRECPT(1).Date = Date2Num%("01-20-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29199 THEN
    TBRECPT(1).Date = Date2Num%("01-21-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29198 THEN
    TBRECPT(1).Date = Date2Num%("01-22-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29197 THEN
    TBRECPT(1).Date = Date2Num%("01-23-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29196 THEN
    TBRECPT(1).Date = Date2Num%("01-24-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29195 THEN
    TBRECPT(1).Date = Date2Num%("01-25-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29194 THEN
    TBRECPT(1).Date = Date2Num%("01-26-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29193 THEN
    TBRECPT(1).Date = Date2Num%("01-27-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF
   IF TBRECPT(1).Date = -29192 THEN
    TBRECPT(1).Date = Date2Num%("01-28-2000")
    PUT TBRFile, Cnt&, TBRECPT(1)
   END IF













   NEXT Cnt&
   CLOSE TBRFile

'Now Fix Transaction History
   GOSUB 61000
   Total& = LOF(TBTranFile) / LEN(TBTrans(1))
   FOR Cnt& = LOF(TBTranFile) / LEN(TBTrans(1)) TO (Total& - 1000) STEP -1
   GET TBTranFile, Cnt&, TBTrans(1)
   IF TBTrans(1).TransDate = -29219 THEN
     TBTrans(1).TransDate = Date2Num%("01-01-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29218 THEN
     TBTrans(1).TransDate = Date2Num%("01-02-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29217 THEN
     TBTrans(1).TransDate = Date2Num%("01-03-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29216 THEN
     TBTrans(1).TransDate = Date2Num%("01-04-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29215 THEN
     TBTrans(1).TransDate = Date2Num%("01-05-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29214 THEN
     TBTrans(1).TransDate = Date2Num%("01-06-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29213 THEN
     TBTrans(1).TransDate = Date2Num%("01-07-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29212 THEN
     TBTrans(1).TransDate = Date2Num%("01-08-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29211 THEN
     TBTrans(1).TransDate = Date2Num%("01-09-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29210 THEN
     TBTrans(1).TransDate = Date2Num%("01-10-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29209 THEN
     TBTrans(1).TransDate = Date2Num%("01-11-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29208 THEN
     TBTrans(1).TransDate = Date2Num%("01-12-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29207 THEN
     TBTrans(1).TransDate = Date2Num%("01-13-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29206 THEN
     TBTrans(1).TransDate = Date2Num%("01-14-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29205 THEN
     TBTrans(1).TransDate = Date2Num%("01-15-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29204 THEN
     TBTrans(1).TransDate = Date2Num%("01-16-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29203 THEN
     TBTrans(1).TransDate = Date2Num%("01-17-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29202 THEN
     TBTrans(1).TransDate = Date2Num%("01-18-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29201 THEN
     TBTrans(1).TransDate = Date2Num%("01-19-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29200 THEN
     TBTrans(1).TransDate = Date2Num%("01-20-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF




   NEXT Cnt&
   CLOSE

  'Supplemental History Fix Here

   TBTranFile = FREEFILE
   OPEN TBPath$ + "TTTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTranFile LEN = LEN(TBTrans(1))

   Total& = LOF(TBTranFile) / LEN(TBTrans(1))
   FOR Cnt& = LOF(TBTranFile) / LEN(TBTrans(1)) TO 1 STEP -1
   GET TBTranFile, Cnt&, TBTrans(1)
   IF TBTrans(1).TransDate = -29219 THEN
     TBTrans(1).TransDate = Date2Num%("01-01-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29218 THEN
     TBTrans(1).TransDate = Date2Num%("01-02-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29217 THEN
     TBTrans(1).TransDate = Date2Num%("01-03-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29216 THEN
     TBTrans(1).TransDate = Date2Num%("01-04-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29215 THEN
     TBTrans(1).TransDate = Date2Num%("01-05-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29214 THEN
     TBTrans(1).TransDate = Date2Num%("01-06-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29213 THEN
     TBTrans(1).TransDate = Date2Num%("01-07-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29212 THEN
     TBTrans(1).TransDate = Date2Num%("01-08-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29211 THEN
     TBTrans(1).TransDate = Date2Num%("01-09-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29210 THEN
     TBTrans(1).TransDate = Date2Num%("01-10-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29209 THEN
     TBTrans(1).TransDate = Date2Num%("01-11-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29208 THEN
     TBTrans(1).TransDate = Date2Num%("01-12-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29207 THEN
     TBTrans(1).TransDate = Date2Num%("01-13-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29206 THEN
     TBTrans(1).TransDate = Date2Num%("01-14-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29205 THEN
     TBTrans(1).TransDate = Date2Num%("01-15-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29204 THEN
     TBTrans(1).TransDate = Date2Num%("01-16-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29203 THEN
     TBTrans(1).TransDate = Date2Num%("01-17-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29202 THEN
     TBTrans(1).TransDate = Date2Num%("01-18-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29201 THEN
     TBTrans(1).TransDate = Date2Num%("01-19-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29200 THEN
     TBTrans(1).TransDate = Date2Num%("01-20-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF

   IF TBTrans(1).TransDate = -29199 THEN
     TBTrans(1).TransDate = Date2Num%("01-21-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29198 THEN
     TBTrans(1).TransDate = Date2Num%("01-22-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29197 THEN
     TBTrans(1).TransDate = Date2Num%("01-23-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29196 THEN
     TBTrans(1).TransDate = Date2Num%("01-24-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29195 THEN
     TBTrans(1).TransDate = Date2Num%("01-25-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29194 THEN
     TBTrans(1).TransDate = Date2Num%("01-26-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29193 THEN
     TBTrans(1).TransDate = Date2Num%("01-27-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF
   IF TBTrans(1).TransDate = -29192 THEN
     TBTrans(1).TransDate = Date2Num%("01-28-2000")
     PUT TBTranFile, Cnt&, TBTrans(1)
   END IF










   NEXT Cnt&
   CLOSE









   PRINT "DONE"
   SLEEP 5
   END


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

SUB Cleardown
   FOR ll = 3 TO 23: LOCATE ll, 1: PRINT STRING$(79, 32): NEXT ll
END SUB

SUB DecalPayment (PDATE$, TBPath$)
     PostDate$ = LEFT$(PDATE$, 2) + "-" + MID$(PDATE$, 3, 2) + "-" + RIGHT$(PDATE$, 2)
     CLOSE : OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS CYear$: GET 1, 1
     CurrentYear$ = CYear$: CLOSE
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
     PRINT "Enter Decal Number .......... "; : fl = 10
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
   TBRECPT(1).Drawer = OPNUMBER
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
     LPRINT CHR$(7);
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
  FOR ll = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, ll
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO Decaln1
Decaln3:
  NEXT ll
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
  FOR ll = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, ll
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO DecalPPn1
DecalPPn3:
  NEXT ll
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
     LSET CodeDesc$ = UCASE$(in$)
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
     GOSUB 44000: LSET CodeDesc$ = in$: LSET Active$ = "Y": PUT 15, xx
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
      LPRINT "Code # "; x; " Set To "; CodeDesc$
     END IF
     NEXT x
     LPRINT CHR$(12);
     CLOSE
     RETURN

misccodefile1:
 CLOSE 15
 OPEN TBPath$ + "tbmisc.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #15 LEN = 256
 FIELD 15, 30 AS CodeDesc$, 1 AS Active$, 225 AS nu$
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

