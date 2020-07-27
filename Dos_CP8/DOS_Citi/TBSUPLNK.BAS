10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbpayment
30 REM Payment Entry Menu
40 REM Last Revision Date: 7-14-87
50 REM by Nick Picerno
    DECLARE FUNCTION Num2Date$ (Num%)
    DECLARE FUNCTION Date2Num% (Dat$)
60 CLS : CLEAR
   
   '$INCLUDE: 'lcTAX.BI'

    TBPath$ = "c:\qb45\lc\"                    ' change to "F:\" Before compiling

     
     REDIM TBRECust(1) AS TBRECust
     REDIM TBPPCust(1) AS TBPPCust
     REDIM TBTrans(1) AS TBTrans

    
     CLS
     PRINT "Now Resetting the Transaction Pointers : Please Wait!!!"
     taxtype$ = "R"
     GOSUB 60000
     FOR Cnt! = 1 TO LOF(TBFile) / LEN(TBRECust(1))
     GET TBFile, Cnt!, TBRECust(1)
      TBRECust(1).FirstTrans = 0
      TBRECust(1).LastTrans = 0
     PUT TBFile, Cnt!, TBRECust(1)
     NEXT Cnt!
     GOSUB 61000
     FOR Cnt! = 1 TO LOF(TBTranFile) / LEN(TBTrans(1))
     GET TBTranFile, Cnt!, TBTrans(1)
    IF TBTrans(1).TrSource = "R" THEN
      TBTrans(1).TrNextTrans = 0
     PUT TBTranFile, Cnt!, TBTrans(1)
     GET TBTranFile, Cnt!, TBTrans(1)
     Record = TBTrans(1).TrCustomerRecord
     IF Record > 0 THEN
        GET TBFile, Record, TBRECust(1)
     IF TBRECust(1).FirstTrans = 0 THEN
        TBRECust(1).FirstTrans = Cnt!
        TBRECust(1).LastTrans = Cnt!
        PUT TBFile, Record, TBRECust(1)
            ELSE
        Prev! = TBRECust(1).LastTrans
        TBRECust(1).LastTrans = Cnt!
        PUT TBFile, Record, TBRECust(1)
        GET TBTranFile, Prev!, TBTrans(1)
        TBTrans(1).TrNextTrans = Cnt!
        PUT TBTranFile, Prev!, TBTrans(1)
      END IF
     END IF
     END IF
     NEXT Cnt!
     CLOSE
     
' Start Personal Property
     taxtype$ = "P"
     GOSUB speciallunenburg

     GOSUB 60000
     FOR Cnt! = 1 TO LOF(TBFile) / LEN(TBPPCust(1))
      GET TBFile, Cnt!, TBPPCust(1)
      TBPPCust(1).FirstTrans = 0
      TBPPCust(1).LastTrans = 0
      PUT TBFile, Cnt!, TBPPCust(1)
     NEXT Cnt!

     GOSUB 61000
   FOR Cnt! = 1 TO LOF(TBTranFile) / LEN(TBTrans(1))
        GET TBTranFile, Cnt!, TBTrans(1)
   IF TBTrans(1).TrSource = "P" THEN
        Record = TBTrans(1).TrCustomerRecord
     IF Record > 0 THEN
        GET TBFile, Record, TBPPCust(1)
      IF TBPPCust(1).FirstTrans = 0 THEN
        TBPPCust(1).FirstTrans = Cnt!
        TBPPCust(1).LastTrans = Cnt!
        PUT TBFile, Record, TBPPCust(1)
            ELSE
        Prev! = TBPPCust(1).LastTrans
        TBPPCust(1).LastTrans = Cnt!
        PUT TBFile, Record, TBPPCust(1)

        GET TBTranFile, Prev!, TBTrans(1)
        TBTrans(1).TrNextTrans = Cnt!
        PUT TBTranFile, Prev!, TBTrans(1)

       END IF
      END IF
   END IF
     NEXT Cnt!
     CLOSE
     END
   
speciallunenburg:
GOSUB 60000
GOSUB 61000
FOR Cnt! = 1 TO LOF(TBFile) / LEN(TBPPCust(1))
 GET TBFile, Cnt!, TBPPCust(1)
 Record = TBPPCust(1).FirstTrans

WHILE Record > 0
 GET TBTranFile, Record, TBTrans(1)
 IF TBTrans(1).TrCustomerRecord = 0 THEN TBTrans(1).TrCustomerRecord = Cnt!
 PUT TBTranFile, Record, TBTrans(1)
 Record = TBTrans(1).TrNextTrans
WEND
NEXT Cnt!
RETURN


60000  IF taxtype$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TTRSCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBPath$ + "TTPSCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN

60025 TBBalFile = FREEFILE
      IF taxtype$ = "R" THEN
      OPEN TBPath$ + "TTRSBAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBPath$ + "TTPSBAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF
60030 RETURN

61000 TBTranFile = FREEFILE
      OPEN TBPath$ + "TTTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTranFile LEN = LEN(TBTrans(1))
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
63000 RETURN
63500 COLOR 10: LOCATE 1, 1: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN
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

