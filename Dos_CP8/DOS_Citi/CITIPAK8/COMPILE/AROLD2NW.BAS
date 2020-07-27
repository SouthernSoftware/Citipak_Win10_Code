DEFINT A-Z
DECLARE SUB OpenARCustFile (NumOfArRecs%, ARFile%)
DECLARE FUNCTION Date2Num% (Dat$)
  '$INCLUDE: 'AR.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
   DIM SHARED ARCust(1) AS ARCustRecType
   DIM SHARED ARCustRec(1) AS ARCustRecType

CLS
PRINT "Now Converting Old A/R to New A/R File Formats"

' open necessary files for customer conversion
GOSUB 60000

OpenARCustFile NumOfArRecs, ARFile


FOR x = 1 TO LOF(1) / 512
GET 1, x


    ARCustRec(1).Custnumb = LTRIM$(STR$(x))
    IF status$ = "I" THEN
     ARCustRec(1).SortName = "INACTIVE"
    ELSE
     ARCustRec(1).SortName = Nme$
    END IF
    ARCustRec(1).BILLNAME = Nme$
    ARCustRec(1).ADDRESS1 = add1$
    ARCustRec(1).ADDRESS2 = add2$
    ARCustRec(1).city = city$
    ARCustRec(1).STATE = St$
    ARCustRec(1).ZIPCODE = Zp$

    IF LEN(RTRIM$(Contact$)) = 0 THEN
     ARCustRec(1).CustName = Nme$
    ELSE
     ARCustRec(1).CustName = Contact$
    END IF

    ARCustRec(1).BILLCAT = "BL"
    ARCustRec(1).SOSEC = ""
    ARCustRec(1).DRVLIC = ""
    ARCustRec(1).DATEOPED = -32767
    ARCustRec(1).BILLCMT = ""
    ARCustRec(1).PAYCMT = ""
    ARCustRec(1).CASHONLY = "N"
    ARCustRec(1).APPNUMB = 1
    ARCustRec(1).BILLFORM = 1
    ARCustRec(1).HPHONE = Phone$
    ARCustRec(1).WPHONE = ""
    ARCustRec(1).FeeAmt = CVS(stfee$)                                                      'INTEGER
    ARCustRec(1).License = License$
    ARCustRec(1).Valid = Date2Num%(validdate$)       'INTEGER Date Function
    ARCustRec(1).IssueLicense = "N"
    ARCustRec(1).AcctBal = (CVD(curbal$) + CVD(pastdue$))
    ARCustRec(1).OldFirstTrans = 0
    ARCustRec(1).OldLastTrans = 0
    IF status$ = "I" THEN
    ARCustRec(1).Deleted = "Y"
    ELSE
    ARCustRec(1).Deleted = "N"
    END IF
    ARCustRec(1).FirstTrans = CVS(Firsttr$)
    ARCustRec(1).LastTrans = CVS(Lasttr$)
    ARCustRec(1).RoomtoGrow = ""

    PUT ARFile, x, ARCustRec(1)
   NEXT x
   CLOSE


   'transactions next

 GOSUB 61000
 REDIM ARTransRec(1) AS ARTransRecType
 ARTransRecLen = LEN(ARTransRec(1))
 ARTransFile = FREEFILE
 OPEN "ARTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARTransFile LEN = ARTransRecLen
 FOR x = 1 TO LOF(2) / 128
  GET 2, x
  ARTransRec(1).CustomerNumber = LTRIM$(STR$(CVS(Tmrec$)))
  ARTransRec(1).TransDate = Date2Num%(TDate$)
  ARTransRec(1).TransAmount = CVS(Tcuramt$)
  ARTransRec(1).TransType = CVI(ttype$)
  IF CVI(ttype$) > 1 THEN
   ARTransRec(1).TransType = 2
  END IF
  ARTransRec(1).TransDesc = TDesc$
  ARTransRec(1).CashAmount = 0
  ARTransRec(1).ChkAmount = 0
  ARTransRec(1).BalanceAfterTrans = CVS(TRunamt$)
  ARTransRec(1).ExtraRoom = ""
  ARTransRec(1).NextTrans = CVS(Tnexttr$)
  PUT ARTransFile, x, ARTransRec(1)
 NEXT x
 CLOSE
 END


60000 CLOSE 1: OPEN "arcustv1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 512
60010 FIELD 1, 40 AS Nme$, 30 AS add1$, 30 AS add2$, 20 AS city$, 2 AS St$
      FIELD 1, 122 AS du1$, 9 AS Zp$, 12 AS Phone$, 8 AS curbal$, 8 AS pastdue$
      FIELD 1, 159 AS du2$, 1 AS active$, 4 AS Firsttr$, 4 AS Lasttr$
      FIELD 1, 168 AS du3$, 30 AS Contact$, 2 AS Category$, 4 AS stfee$
      FIELD 1, 204 AS du4$, 12 AS License$, 6 AS validdate$, 8 AS gross$, 2 AS billcalc$, 280 AS blank$
60030 RETURN
61000 CLOSE 2: OPEN "artranv1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 128
61010 FIELD 2, 8 AS TDate$, 4 AS Tcuramt$, 4 AS TRunamt$, 2 AS ttype$, 25 AS TDesc$
      FIELD 2, 43 AS dum1$, 4 AS Tmrec$, 4 AS Tnexttr$, 77 AS blank1$
61020 RETURN
62000 CF = 0: COLOR 14: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN CF = 1: RETURN
      IF z$ = CHR$(9) THEN CF = 2: RETURN
      IF LEN(z$) <> 2 THEN 62065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN CF = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN CF = 12: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 61 THEN CF = 13: RETURN

62061 IF z$ = CHR$(9) THEN CF = 2: RETURN
62062 IF z$ = ";" THEN CF = 2: RETURN
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
63000 RETURN
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN

SUB OpenARCustFile (NumOfArRecs, ARFile)
  CLOSE ARFile

  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen

END SUB

