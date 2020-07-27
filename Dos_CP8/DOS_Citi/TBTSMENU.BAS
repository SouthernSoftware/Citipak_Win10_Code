10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbmenu
30 REM main tax billing system menu
40 REM Last Revision Date: 8-26-87
50 REM by Nick Picerno
60      DEFINT A-Z
        TYPE Struct
                who AS STRING * 14
                Recnum AS INTEGER
        END TYPE

   DECLARE SUB SortT (SEG Element AS Struct, NumEls, Dir, SSize, MOff, MSize)
  '$INCLUDE: 'LCTax.bi'
   CLEAR , , 4000          'the sort routines need the extra stack space


        REDIM TBRECust(1) AS TBRECust
        taxtype$ = "R"
        TBPath$ = "F:\"         'Change to F:\ when compiling"


100 GOSUB 63000: REM get town in town$
110 SCREN$ = "Main": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
130 LOCATE 4, 1: COLOR 11
200 PRINT TAB(31); "Tax Billing System"
210 PRINT TAB(23); "Treasurers Supplemental Master Menu"
220 PRINT : PRINT
230 PRINT TAB(28); "1. Supplemental Billing"
240 PRINT
250 PRINT TAB(28); "2. Supplemental Reports"
260 PRINT
270 PRINT TAB(28); "3. DMV Processing"
280 PRINT
    PRINT TAB(28); "4. Re-Index Sup Accts"
310 LOCATE 18, 1: COLOR 11: PRINT TAB(28); "Selection No. (ESC-End) : "; : fl = -1: GOSUB 62000
320 IF CF = 1 THEN CLOSE : CLS : RUN "tbtmenu"
325 IF LEN(IN$) = 0 THEN 310
330 IF VAL(IN$) < 1 OR VAL(IN$) > 4 THEN 310
340 ON VAL(IN$) GOTO 410, 420, 430, 1000
410 RUN "tbsbilln"
420 RUN "tbtscstr"
430 RUN "tbtsdmv"
1000


CLS
Size = 4100
Start = 1               'start at array element 1
Dir = 0                 'sort direction - use anything else for descending
SSize = 16               'total size of each TYPE element
MOff = 0                'offset into the TYPE for the key element
MSize = 14              'size of the key element - coded as follows:
                        '   -1 = integer
                        '   -2 = long integer
                        '   -3 = single precision
                        '   -4 = double precision
                        '   +N = TYPE array/fixed-length string of length N

REDIM Array(1 TO Size) AS Struct

CLS
COLOR 11
PRINT "Tax System : Last Name Sort Utility - Supplementals"
GOSUB 60000
FOR x = 1 TO LOF(TBFile) / LEN(TBRECust(1))
GET 1, x, TBRECust(1)

 IF TBRECust(1).Deleted <> "Y" THEN
    Count = Count + 1
    Array(Count).who = LEFT$(TBRECust(1).CustName, 14)
    Array(Count).Recnum = x
 END IF
NEXT x


SortT Array(Start), Count, Dir, SSize, MOff, MSize


OPEN TBPath$ + "TTSRNAME.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = 24
FIELD 14, 20 AS lastnam$, 4 AS mrec$
CLOSE 14
FileNam$ = TBPath$ + "TTSRName.Dat"
KILL FileNam$
OPEN TBPath$ + "TTSRNAME.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = 24
FIELD 14, 20 AS lastnam$, 4 AS mrec$


FOR x = 1 TO Count
    LSET lastnam$ = Array(x).who
    LSET mrec$ = MKS$(Array(x).Recnum)
    PUT 14, LOF(14) / 24 + 1
NEXT
CLOSE
'Now Do Personal
REDIM TBPPCust(1) AS TBPPCust
taxtype$ = "P"
Size = 4100
Start = 1               'start at array element 1
Dir = 0                 'sort direction - use anything else for descending
SSize = 16               'total size of each TYPE element
MOff = 0                'offset into the TYPE for the key element
MSize = 14              'size of the key element - coded as follows:
                        '   -1 = integer
                        '   -2 = long integer
                        '   -3 = single precision
                        '   -4 = double precision
                        '   +N = TYPE array/fixed-length string of length N

REDIM Array(1 TO Size) AS Struct
Count = 0

GOSUB 60000
FOR x = 1 TO LOF(TBFile) / LEN(TBPPCust(1))
GET 1, x, TBPPCust(1)


 IF TBPPCust(1).Deleted <> "Y" THEN
    Count = Count + 1
    Array(Count).who = LEFT$(TBPPCust(1).CustName, 14)
    Array(Count).Recnum = x
 END IF
NEXT x


SortT Array(Start), Count, Dir, SSize, MOff, MSize

OPEN TBPath$ + "TTSPNAME.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
FIELD 4, 20 AS lastnam$, 4 AS mrec$
CLOSE 4
FileNam$ = TBPath$ + "TTSPName.Dat"
KILL FileNam$
OPEN TBPath$ + "TTSPNAME.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
FIELD 4, 20 AS lastnam$, 4 AS mrec$


FOR x = 1 TO Count

    LSET lastnam$ = Array(x).who
    LSET mrec$ = MKS$(Array(x).Recnum)
    
    
    PUT 4, LOF(4) / 24 + 1
NEXT
CLOSE
SLEEP 1

GOTO 10

60000  IF taxtype$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TTRSCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBPath$ + "TTPSCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN

62000 CF = 0: COLOR 14: IN$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 Z$ = INKEY$: IF Z$ = "" THEN 62030
62040 IF Z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF Z$ = CHR$(27) OR Z$ = CHR$(21) THEN CF = 1: RETURN
62065 IF ASC(Z$) = 0 THEN 62030
62070 IF LEN(Z$) > 1 THEN CF = ASC(RIGHT$(Z$, 1)) - 57: IF CF > 0 THEN RETURN ELSE CF = CF + 27: RETURN
62080 IF Z$ = " " THEN 62130 ELSE IF Z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 62000
62090 IF Z$ <> CHR$(8) THEN 62130 ELSE IF ZL = 0 THEN 62030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 62110 ELSE IF Z$ = "," THEN 62120
62100 IF Z$ = "." THEN ZD = 0: GOTO 62110 ELSE IF Z$ = "+" OR Z$ = "-" THEN ZS = 0
62110 IN$ = LEFT$(IN$, LEN(IN$) - 1)
62120 ZL = ZL - 1: GOTO 62030
62130 IF ABS(fl) = ZL THEN 62030 ELSE IF fl > 0 THEN IF Z$ >= " " AND Z$ <= "z" THEN 62180
62140 IF Z$ = "." AND ZD = 0 THEN ZD = 1: GOTO 62180
62150 IF Z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (Z$ = "-" OR Z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF Z$ < "0" OR Z$ > "9" THEN 62030
62180 PRINT Z$; : IN$ = IN$ + Z$: ZL = ZL + 1
62190 IF ABS(fl) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT IN$: RETURN
62210 GOSUB 62000: IF LEN(IN$) THEN IN$ = CHR$(ASC(IN$) AND 223)
62220 RETURN
63000 OPEN "R", 3, "wbtown.dat": FIELD 3, 40 AS TOWNNAME$: GET 3, 1: TOWN$ = TOWNNAME$: CLOSE 3: RETURN
63500 COLOR 10: LOCATE 1, 1: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN

