10 DEFINT A-Z
DEFDBL R
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Round# (DblNm#)
CLEAR , , 4000
DIM SHARED BMTAX(10, 1), sdTAX(10, 1), ADVERT(10, 1), COLECT(10, 1), INTEREST(10, 1)
'$INCLUDE: 'lcTAX.BI'

        REDIM TBCust(1) AS TBRECust
        REDIM TBCustBal(1) AS TBREBal


        TBFile = FREEFILE
        OPEN "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBCust(1))

        TBBalFile = FREEFILE
        OPEN "TBREBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBCustBal(1))

        FOR Cnt! = 1 TO LOF(TBFile) / LEN(TBCust(1))
        GET TBBalFile, Cnt!, TBCustBal(1)
        FOR ll = 1 TO 10
        TBCustBal(1).TaxAmt(ll) = 0
        TBCustBal(1).IntAmt(ll) = 0
        TBCustBal(1).PenAmt(ll) = 0
        TBCustBal(1).OthAmt(ll) = 0
        TBCustBal(1).LateAmt(ll) = 0
        TBCustBal(1).TaxNotice(ll) = ""
        NEXT ll
        PUT TBBalFile, Cnt!, TBCustBal(1)
        NEXT Cnt!

        GOSUB 60000
        OPEN "tbacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = 31
        FIELD 14, 27 AS met$, 4 AS mrec$: ENDOFFILE = LOF(14) / 31
       
        FOR xx = 1 TO LOF(14) / 31
         GET 14, xx
         GET 11, CVS(mrec$): GET 13, CVS(mrec$)
         IF VAL(acct$) <= 0 THEN 20
         IF UCASE$(Taxtype$) <> "R" THEN 20
         GOSUB figurebalance
         IF curbal# = 0 AND pastdue# = 0 THEN 20
         
         GET TBFile, VAL(acct$), TBCust(1)
         IF TBCust(1).Deleted = "Y" THEN
         GOSUB UpdateAccount
         END IF
        GET TBBalFile, VAL(acct$), TBCustBal(1)
        
        TBCustBal(1).TaxAmt(1) = Round#(CVS(curtaxamt$) + CVS(curadvamt$) + CVS(curcolamt$))
        TBCustBal(1).LateAmt(1) = Round#(CVS(curlateamt$))
        TBCustBal(1).IntAmt(1) = Round#(CVS(curintamt$))
        TBCustBal(1).PenAmt(1) = 0
        TBCustBal(1).OthAmt(1) = 0
        TBCustBal(1).TaxNotice(1) = STR$(CVS(CTaxNotice$))
         
        FOR Cnt! = 2 TO 9
         TBCustBal(1).TaxAmt(Cnt!) = Round#(CVS(PTAXAMT$(Cnt! - 2)) + CVS(PADVAMT$(Cnt! - 2)) + CVS(PCOLAMT$(Cnt! - 2)))
         TBCustBal(1).LateAmt(Cnt!) = Round#(CVS(PLATEAMT$(Cnt! - 2)))
         TBCustBal(1).IntAmt(Cnt!) = Round#(CVS(PintAMT$(Cnt! - 2)))
         TBCustBal(1).PenAmt(Cnt!) = 0
         TBCustBal(1).OthAmt(Cnt!) = 0
         TBCustBal(1).TaxNotice(Cnt!) = STR$(CVS(ptaxnotice$(Cnt! - 2)))
        NEXT Cnt!

         TBCustBal(1).TaxAmt(10) = Round#(CVS(PTAXAMT$(8)) + CVS(PTAXAMT$(9))) + CVS(PADVAMT$(8)) + CVS(PCOLAMT$(9)) + CVS(PADVAMT$(9)) + CVS(PCOLAMT$(8))
         TBCustBal(1).LateAmt(10) = Round#(CVS(PLATEAMT$(8)) + CVS(PLATEAMT$(9)))
         TBCustBal(1).IntAmt(10) = Round#(CVS(PintAMT$(8)) + CVS(PintAMT$(9)))
         TBCustBal(1).PenAmt(10) = 0
         TBCustBal(1).OthAmt(10) = 0
         TBCustBal(1).TaxNotice(10) = STR$(CVS(ptaxnotice$(8)))
         PUT TBBalFile, VAL(acct$), TBCustBal(1)
20 NEXT xx


        CLOSE
        STOP
UpdateAccount:
         TBCust(1).Deleted = "N"
         TBCust(1).CustName = Lnam$
         TBCust(1).Address1 = A1$
         TBCust(1).Address2 = a2$
         TBCust(1).Address3 = cy$ + " " + St$ + " " + ZP$
         TBCust(1).PropDesc1 = PDesc$
         TBCust(1).MapNumber = Pmap$
         PUT TBFile, VAL(acct$), TBCust(1)
         RETURN



60000 CLOSE 11: OPEN "tbcust.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #11 LEN = 512
60010 FIELD 11, 12 AS acct$, 60 AS Lnam$, 35 AS A1$, 35 AS a2$, 20 AS cy$, 2 AS St$, 10 AS ZP$, 12 AS tp$
      FIELD 11, 186 AS dummy$, 39 AS PDesc$, 22 AS Pmap$, 35 AS psub$, 4 AS plot$, 35 AS pcounty$, 4 AS psize$, 8 AS pland$, 8 AS MV$
      FIELD 11, 341 AS dummy1$, 8 AS MH$, 8 AS FM$, 8 AS LS$, 4 AS OP$, 4 AS ME$, 6 AS HOMEEXP$, 1 AS Taxtype$, 3 AS taxtype1$
      FIELD 11, 381 AS dummy2$, 4 AS curlateamt$, 4 AS curtaxamt$, 4 AS curintamt$, 4 AS curadvamt$, 4 AS curcolamt$, 4 AS CTaxNotice$
      FIELD 11, 405 AS dummy3$, 4 AS ftr$, 4 AS ltr$, 11 AS ss$, 1 AS late$, 15 AS pin$, 1 AS Taxable$, 67 AS blank$
60025 CLOSE 13: OPEN "tbcust1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #13 LEN = 256
      FOR tloop = 0 TO 9
      FIELD 13, (24 * tloop) AS dummy4$, 4 AS PTAXAMT$(tloop), 4 AS PintAMT$(tloop), 4 AS PADVAMT$(tloop), 4 AS PCOLAMT$(tloop), 4 AS PLATEAMT$(tloop), 4 AS ptaxnotice$(tloop): NEXT tloop
60030 RETURN
figurebalance:
     curbal# = 0: pastdue# = 0
     curbal# = (CVS(curtaxamt$) + CVS(curintamt$) + CVS(curadvamt$) + CVS(curcolamt$) + CVS(curlateamt$))
     curbal# = INT((curbal# * 100) + .5) / 100
     FOR ll = 0 TO 9
     pastdue# = pastdue# + (CVS(PTAXAMT$(ll)) + CVS(PintAMT$(ll)) + CVS(PADVAMT$(ll)) + CVS(PCOLAMT$(ll)) + CVS(PLATEAMT$(ll)))
     pastdue# = INT((pastdue# * 100) + .5) / 100
     NEXT ll
     RETURN

62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(xy): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 Z$ = INKEY$: IF Z$ = "" THEN 62030
62040 IF Z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF Z$ = CHR$(27) OR Z$ = CHR$(21) THEN cf = 1: RETURN
62065 IF ASC(Z$) = 0 THEN 62030
62070 IF LEN(Z$) > 1 THEN cf = ASC(RIGHT$(Z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
62080 IF Z$ = " " THEN 62130 ELSE IF Z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 62000
62090 IF Z$ <> CHR$(8) THEN 62130 ELSE IF ZL = 0 THEN 62030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 62110 ELSE IF Z$ = "," THEN 62120
62100 IF Z$ = "." THEN ZD = 0: GOTO 62110 ELSE IF Z$ = "+" OR Z$ = "-" THEN ZS = 0
62110 in$ = LEFT$(in$, LEN(in$) - 1)
62120 ZL = ZL - 1: GOTO 62030
62130 IF ABS(fl) = ZL THEN 62030 ELSE IF fl > 0 THEN IF Z$ >= " " AND Z$ <= "z" THEN 62180
62140 IF Z$ = "." AND INSTR(in$, ".") = 0 THEN GOTO 62180
62150 IF Z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (Z$ = "-" OR Z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF Z$ < "0" OR Z$ > "9" THEN 62030
62180 PRINT Z$; : in$ = in$ + Z$: ZL = ZL + 1
62190 IF ABS(fl) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT in$: RETURN
62210 GOSUB 62000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
62220 RETURN
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); scren$: PRINT STRING$(79, "_"): RETURN
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

DEFINT R
FUNCTION Round# (DblNum#)
Round# = INT((DblNum# * 100) + .5) / 100
END FUNCTION

