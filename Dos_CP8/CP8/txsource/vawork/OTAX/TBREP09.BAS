10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbcustrep
30 REM Customer Report Menu
40 REM Last Revision Date: 8-24-87
50 REM by Nick Picerno
    TYPE Struct
        Map AS STRING * 14
        Rec AS INTEGER
    END TYPE
    DEFINT A-Z
    DECLARE SUB SortT (SEG Element AS Struct, NumEls, Dir, SSize, MOff, MSize)

   '$INCLUDE: 'LCTAX.BI'
    CLEAR , , 4000          'the sort routines need the extra stack space

   
    TBPath$ = "F:\"
   'TBPath$ = "C:\qb45\lc\"
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
     LOCATE 1, 30: PRINT "Print Personal Property Sheets"
6100 LOCATE 3, 1: COLOR 11
     PRINT "Enter (M)ain Records or (S)up. Records (M,S or <ESC>-End) .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN RUN "tbcustre"
     IF NOT (in$ = "S" OR in$ = "M") THEN 6100
     WhoseRecords$ = in$
6101 LOCATE 4, 1: COLOR 11
     PRINT "Enter the Alpha Letter to Print (A-Z or <ESC>-End) .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN RUN "tbcustre"
     Alpha$ = UCASE$(in$)
     


     GOSUB Alignment
     REDIM TBCust(1) AS TBPPCust
     tbfile = FREEFILE
     IF WhoseRecords$ = "S" THEN
       OPEN TBPath$ + "TBSPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #tbfile LEN = LEN(TBCust(1))
     ELSE
       OPEN TBPath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #tbfile LEN = LEN(TBCust(1))
     END IF
     LastRecord = LOF(tbfile) / LEN(TBCust(1))

     LOCATE 10, 1: COLOR 15: PRINT "Extracting Records : PLEASE WAIT!!!"

     FOR MasterLoop = 1 TO 2
      Count = 1: RCnt = 0
     WHILE Count <= LastRecord
      GET tbfile, Count, TBCust(1)
      
      IF TBCust(1).Deleted <> "Y" THEN

       IF MasterLoop = 1 AND VAL(TBCust(1).district) <= 8 THEN
        RCnt = RCnt + 1
        Array(RCnt).Map = TBCust(1).district + TBCust(1).CustName
        Array(RCnt).Rec = Count
       END IF

       IF MasterLoop = 2 AND VAL(TBCust(1).district) > 8 THEN
        RCnt = RCnt + 1
        Array(RCnt).Map = TBCust(1).district + TBCust(1).CustName
        Array(RCnt).Rec = Count
       END IF

      END IF
      Count = Count + 1

     WEND

     SortT Array(Start), RCnt, Dir, SSize, MOff, MSize
     Count = 1

COLOR 15: LOCATE 12, 1: PRINT "Now Printing the Cards : Please Wait"
WHILE Count <= RCnt
     GET tbfile, Array(Count).Rec, TBCust(1)
     GOSUB GetDistrict
     GOSUB PrintCard
     Count = Count + 1
WEND
 PRINT CHR$(12);
 NEXT MasterLoop
 
 CLOSE
 RUN "TBCUSTRE"

PrintCard:
     IF VAL(TBCust(1).district) = 0 THEN RETURN
     IF Alpha$ = "A" THEN
      IF LEFT$(TBCust(1).CustName, 1) <= "A" THEN GOSUB PrintTheCard: RETURN
     END IF
     IF LEFT$(TBCust(1).CustName, 1) <> Alpha$ THEN RETURN



PrintTheCard:

'Check for Value Do Not Print if it does not contain a value

  Value# = TBCust(1).PersProp + TBCust(1).MachTools + TBCust(1).MobileHome + TBCust(1).MerchCap + TBCust(1).FarmEquip
  Value# = INT((Value# * 100) + .5) / 100
  IF Value# = 0 AND WhoseRecords$ = "M" THEN RETURN

      
     noline = 0
     FOR Cnt = 1 TO 5
      LPRINT ""
     NEXT Cnt
     IF WhoseRecords$ = "M" THEN                    ' do not print supp acct #
       LPRINT TAB(30); Array(Count).Rec;
     END IF
       LPRINT TAB(47); DistrictName$
      LPRINT TAB(2); TBCust(1).CustName
      IF LEN(RTRIM$(TBCust(1).Address1)) > 0 THEN
       LPRINT TAB(2); TBCust(1).Address1
       ELSE
       noline = noline + 1
      END IF
      IF LEN(RTRIM$(TBCust(1).Address2)) > 0 THEN
       LPRINT TAB(2); TBCust(1).Address2
       ELSE
       noline = noline + 1
      END IF
      LPRINT TAB(2); TBCust(1).City; " "; TBCust(1).State; " "; TBCust(1).Zip
      FOR Ncnt = 1 TO noline: LPRINT " ": NEXT Ncnt
      LPRINT TAB(50); TBCust(1).SocSec;
      IF VAL(TBCust(1).SocSec2) > 0 THEN LPRINT TAB(69); TBCust(1).SocSec2 ELSE LPRINT ""
      LPRINT CHR$(12);
      RETURN

Alignment:
     LOCATE 7, 1: COLOR 11
     PRINT "Print Alignment (Y/N) or (ESC-ABORT PROCESS) .. "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN CLOSE : RUN "tbCustre"
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO Alignment
     IF in$ = "N" THEN RETURN
     FOR Cnt = 1 TO 5
      LPRINT ""
     NEXT Cnt
      LPRINT TAB(47); STRING$(20, "X")
      LPRINT TAB(2); STRING$(40, "X")
      LPRINT TAB(2); STRING$(40, "X")
      LPRINT TAB(2); STRING$(40, "X")
      LPRINT TAB(2); STRING$(40, "X")
      LPRINT TAB(50); "XXX-XX-XXXX"; TAB(69); "XXX-XX-XXXX"
      LPRINT CHR$(12);
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
     DistrictName$ = "Lunenburg County"
     IF VAL(TBCust(1).district) = 9 THEN DistrictName$ = "TOWN OF KENBRIDGE"
     IF VAL(TBCust(1).district) = 10 THEN DistrictName$ = "TOWN OF VICTORIA"
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

