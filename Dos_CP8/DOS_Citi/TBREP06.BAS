10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbccustrp
30 REM Customer Report Menu
40 REM Last Revision Date: 8-24-87
50 REM by Nick Picerno
60 '$INCLUDE: 'LCTAX.BI'
    DEFINT A-Z
TYPE Struct
    Map AS STRING * 14
    Rec AS INTEGER
END TYPE

DECLARE SUB SortT (SEG Element AS Struct, NumEls, Dir, SSize, MOff, MSize)
CLEAR , , 4000          'the sort routines need the extra stack space
TBPath$ = "F:\"                     ' Set path to F:\ before compiling
'TBPath$ = "C:\qb45\lc\data\"                     ' Set path to F:\ before compiling
ON ERROR GOTO networkerror
WIDTH LPRINT 136
Size = 6000
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

6000 REM
     SCREN$ = "2.06.00": CLS
     GOSUB 63500: COLOR 11
     LOCATE 1, 30: PRINT "Print Property Record Sheets"
6100 LOCATE 3, 1: COLOR 11
     PRINT "Enter the District Number to Print (1-10) or (ESC-End) .. "; : fl = -2
     GOSUB 62000
     IF cf = 1 THEN RUN "tbccustr"
     IF VAL(in$) < 1 OR VAL(in$) > 10 THEN 6100
     District = VAL(in$)
6105 LOCATE 5, 1: COLOR 11
     PRINT "Enter the Beginning Map Location (ENTER=ALL) (ESC-END) .. "; : fl = 14
     GOSUB 62000
     IF LEN(in$) = 0 THEN Map$ = "ALL" ELSE Map$ = UCASE$(in$)

6106 LOCATE 7, 1: COLOR 11
     PRINT "Enter the Tax Year to Print on Sheet .. "; : fl = 4
     GOSUB 62000
     IF LEN(in$) <> 4 THEN 6106
     TaxYear$ = in$

     GOSUB Alignment
     REDIM TBCust(1) AS TBRECust
     TBFile = FREEFILE
     OPEN TBPath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBCust(1))

     LastRecord = LOF(TBFile) / LEN(TBCust(1))
     LOCATE 10, 1: COLOR 15: PRINT "NOW EXTRACTING THIS DISTRICT : PLEASE WAIT!!!"
     Count = 1
WHILE Count <= LastRecord
     GET TBFile, Count, TBCust(1)
     IF VAL(TBCust(1).District) = District THEN
      IF Map$ = "ALL" THEN
      RCnt = RCnt + 1
      Array(RCnt).Map = TBCust(1).MapNumber
      Array(RCnt).Rec = Count
      ELSE
      IF TBCust(1).MapNumber >= Map$ THEN
      RCnt = RCnt + 1
      Array(RCnt).Map = TBCust(1).MapNumber
      Array(RCnt).Rec = Count
      END IF

      END IF
     END IF
     Count = Count + 1
WEND

   SortT Array(Start), RCnt, Dir, SSize, MOff, MSize
Count = 1
COLOR 15: LOCATE 12, 1: PRINT "Now Printing the Cards : Please Wait"
WHILE Count <= RCnt
     GET TBFile, Array(Count).Rec, TBCust(1)
     GOSUB GetDistrict
     GOSUB PrintCard
     Count = Count + 1
WEND
 PRINT CHR$(12);
 CLOSE
 RUN "tbccustr"

PrintCard:
     IF TBCust(1).Deleted = "Y" THEN RETURN
     LPRINT
     LPRINT TAB(5); DistrictName$; TAB(90); TBCust(1).MapNumber
     LPRINT TAB(64); TBCust(1).ClassCode; TAB(72); TBCust(1).PropDesc1
     LPRINT TAB(5); TBCust(1).CustName; TAB(72); TBCust(1).PropDesc2
     LPRINT TAB(5); TBCust(1).Address1; TAB(72); TBCust(1).PropDesc3
     LPRINT TAB(5); TBCust(1).Address2; TAB(72); TBCust(1).PropDesc4
     LPRINT TAB(5); TBCust(1).Address3; TAB(64); TBCust(1).District; TAB(100); Array(Count).Rec
     LPRINT
     LPRINT
     LPRINT TAB(60); TaxYear$;
      LPRINT TAB(66); USING "########,#"; TBCust(1).LandValue;
       LPRINT TAB(78); USING "########,#"; TBCust(1).BldgValue;
        LPRINT TAB(90); USING "########,#"; (TBCust(1).LandValue + TBCust(1).BldgValue);
        LPRINT TAB(103); LEFT$(LTRIM$(TBCust(1).Acres), 5)
     FOR LCnt = 11 TO 51: LPRINT " ": NEXT LCnt
     RETURN

Alignment:
     LOCATE 7, 1: COLOR 11
     PRINT "Print Alignment (Y/N) or (ESC-ABORT PROCESS) .. "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN CLOSE : RUN "tbccustr"
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO Alignment
     IF in$ = "N" THEN RETURN

     LPRINT
     LPRINT TAB(5); "DISTRICT NAME"; TAB(90); "MAP ###########"
     LPRINT TAB(64); "XX"; TAB(72); STRING$(25, "X")
     LPRINT TAB(5); STRING$(34, "X"); TAB(72); STRING$(25, "X")
     LPRINT TAB(5); STRING$(34, "X"); TAB(72); STRING$(25, "X")
     LPRINT TAB(5); STRING$(34, "X"); TAB(72); STRING$(25, "X")
     LPRINT TAB(5); STRING$(34, "X"); TAB(64); "XX"; TAB(100); "XXXXXXX"
     LPRINT
     LPRINT
     LPRINT TAB(60); TaxYear$; TAB(67); "XXXXXXXX"; TAB(79); "XXXXXXXX"; TAB(91); "XXXXXXXX"; TAB(103); "XXXXX"
     
     FOR Cnt = 11 TO 51: LPRINT " ": NEXT Cnt
     GOTO Alignment
     
62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
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
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; DATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN

GetDistrict:
     DistrictName$ = ""
     IF VAL(TBCust(1).District) = 1 THEN DistrictName$ = "BROWN'S STORE"
     IF VAL(TBCust(1).District) = 2 THEN DistrictName$ = "COLUMBIAN GROVE"
     IF VAL(TBCust(1).District) = 3 THEN DistrictName$ = "LEWISTON"
     IF VAL(TBCust(1).District) = 4 THEN DistrictName$ = "LOCHLEVEN"
     IF VAL(TBCust(1).District) = 5 THEN DistrictName$ = "PLEASANT GROVE"
     IF VAL(TBCust(1).District) = 6 THEN DistrictName$ = "PLYMOUTH"
     IF VAL(TBCust(1).District) = 7 THEN DistrictName$ = "REHOBOTH"
     IF VAL(TBCust(1).District) = 9 THEN DistrictName$ = "TOWN OF KENBRIDGE"
     IF VAL(TBCust(1).District) = 10 THEN DistrictName$ = "TOWN OF VICTORIA"
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

