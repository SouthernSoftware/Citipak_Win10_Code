10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbccustrp
30 REM Customer Report Menu
40 REM Last Revision Date: 8-24-87
50 REM by Nick Picerno
   DEFINT A-Z
   TYPE Struct
        Map AS STRING * 14
        rec AS INTEGER
   END TYPE

   DECLARE FUNCTION QPTrim$ (A$)
   DECLARE FUNCTION num2date$ (Num%)
   DECLARE FUNCTION Date2Num% (Dat$)
   DECLARE FUNCTION WaitKey% ()
   DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
   DECLARE SUB PrintTitle (Title$)
   DECLARE SUB PrintLoading ()
   DECLARE SUB ShowCursor ()
   DECLARE SUB HideCursor ()

   DECLARE SUB SortT (SEG Element AS Struct, NumEls, Dir, SSize, MOff, MSize)

'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'Qscr.BI'
'$INCLUDE: 'SetCnf.BI'
'$INCLUDE: 'lcTAX.BI'

  TBPath$ = "F:\"
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
6050 LOCATE 3, 1: COLOR 11
     PRINT "Enter (1-All: 2-Victoria: 3-Kenbridge) (<ESC>-End) .. "; : fl = -1
     GOSUB 62000
     IF cf = 1 THEN RUN "tbccustr"
     WhoToPrint = VAL(in$)
     IF WhoToPrint < 1 OR WhoToPrint > 3 THEN 6050

6100 LOCATE 5, 1: COLOR 11
     PRINT "Enter the Alpha Letter to Print (A-Z or <ESC>-End) .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN RUN "tbccustr"
     Alpha$ = UCASE$(in$)
     LOCATE 6, 1: COLOR 11
     PRINT "Enter New Records Beginning Acct # (<ESC>-End) .. "; : fl = -5
     GOSUB 62000
     IF cf = 1 THEN RUN "tbccustr"
     NewRecords& = VAL(in$)
6110 LOCATE 7, 1
     COLOR 11
     PRINT "Print Only the New Customer Records (Y/N) (<ESC>-End) .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN RUN "tbccustr"
     IF NOT (in$ = "Y" OR in$ = "N") THEN 6110
     PrintNewOnly$ = in$
     GOSUB Alignment
     REDIM TBCust(1) AS TBPPCust
     REDIM tbveh(1) AS PPVehType
     REDIM TBBoat(1) AS PPBoat
     REDIM TBMH(1) AS PPMHome

     tbfile = FREEFILE
     OPEN TBPath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #tbfile LEN = LEN(TBCust(1))
     LastRecord = LOF(tbfile) / LEN(TBCust(1))

     TBVFile = FREEFILE
     OPEN TBPath$ + "TBVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(tbveh(1))
     TBBTFile = FREEFILE
     OPEN TBPath$ + "TBBOAT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBTFile LEN = LEN(TBBoat(1))
     TBMHFile = FREEFILE
     OPEN TBPath$ + "TBMHOME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBMHFile LEN = LEN(TBMH(1))
    
     ReportFile$ = "PERSHT.REP"
     OPEN "O", 10, ReportFile$: WIDTH #10, 255
     
     LOCATE 12, 1: COLOR 15: PRINT "Extracting Records : PLEASE WAIT!!!"

     FOR MasterLoop = 1 TO 2
     count = 1: RCnt = 0
     WHILE count <= LastRecord
      GET tbfile, count, TBCust(1)
      IF PrintNewOnly$ = "Y" AND count < NewRecords& THEN GOTO SkipThisOne

      
      IF TBCust(1).Deleted <> "Y" THEN

    IF MasterLoop = 1 AND VAL(TBCust(1).District) <= 8 THEN
     IF Alpha$ = "A" THEN
      IF LEFT$(TBCust(1).CustName, 1) <= "A" THEN
        RCnt = RCnt + 1
        Array(RCnt).Map = TBCust(1).District + TBCust(1).CustName
        Array(RCnt).rec = count
        GOTO AdvanceOne
      END IF
     END IF
     
     IF LEFT$(TBCust(1).CustName, 1) <> Alpha$ THEN
        ELSE
        RCnt = RCnt + 1
        Array(RCnt).Map = TBCust(1).District + TBCust(1).CustName
        Array(RCnt).rec = count
        GOTO AdvanceOne
     END IF
    END IF

   IF MasterLoop = 2 AND VAL(TBCust(1).District) > 8 THEN
     IF Alpha$ = "A" THEN
       IF LEFT$(TBCust(1).CustName, 1) <= "A" THEN
        RCnt = RCnt + 1
        Array(RCnt).Map = TBCust(1).District + TBCust(1).CustName
        Array(RCnt).rec = count
        GOTO AdvanceOne
      END IF
     END IF

     IF LEFT$(TBCust(1).CustName, 1) <> Alpha$ THEN
        ELSE
        RCnt = RCnt + 1
        Array(RCnt).Map = TBCust(1).District + TBCust(1).CustName
        Array(RCnt).rec = count
        GOTO AdvanceOne
     END IF

   END IF

      END IF
AdvanceOne:
     
SkipThisOne:
      count = count + 1
     WEND

     SortT Array(Start), RCnt, Dir, SSize, MOff, MSize
     count = 1

COLOR 15: LOCATE 12, 1: PRINT "Now Printing the Cards : Please Wait"
WHILE count <= RCnt
        
 GET tbfile, Array(count).rec, TBCust(1)
 IF WhoToPrint = 1 THEN
  GOSUB GetDistrict
  GOSUB PrintCard
 END IF
 IF WhoToPrint = 2 AND VAL(TBCust(1).District) = 10 THEN
  GOSUB GetDistrict
  GOSUB PrintCard
 END IF
 IF WhoToPrint = 3 AND VAL(TBCust(1).District) = 9 THEN
  GOSUB GetDistrict
  GOSUB PrintCard
 END IF
  count = count + 1
WEND

 PRINT #10, CHR$(12);
SkipMaster:
 NEXT MasterLoop
 
 

      CLOSE

      LPTPort% = 1
      EntryPoint% = 1
      Header$ = "Pers Prop Sheets"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
      GOTO 10

PrintCard:
     IF VAL(TBCust(1).District) = 0 THEN RETURN
PrintTheCard:

'Check for Value Do Not Print if it does not contain a value
'  PValue# = 0
'  PValue# = TBCust(1).PersProp + TBCust(1).MachTools + TBCust(1).MobileHome + TBCust(1).MerchCap + TBCust(1).FarmEquip
'  PValue# = PValue# + TBCust(1).HeavyEquip + TBCust(1).Boats + TBCust(1).BusFurn
'  PValue# = INT((PValue# * 100) + .5) / 100
 

      GOSUB SheetHeading
      VP = 0                                    'Set Vehicles Printed to zero
     
     'Now the Vehicles Up to 17 per acct
      VehRecord! = TBCust(1).FirstVeh

      WHILE VehRecord! <> 0
      GET TBVFile, VehRecord!, tbveh(1)
      IF tbveh(1).VehTyp <> "D" THEN
      IF VP >= 17 THEN
       PRINT #10, CHR$(12);
       GOSUB SheetHeading
       VP = 0
      END IF
      VehDesc$ = tbveh(1).VehYear + " " + RTRIM$(tbveh(1).VehMake) + " " + RTRIM$(tbveh(1).VehModel)
      VehDesc$ = LTRIM$(VehDesc$)
      PRINT #10, TAB(2); LEFT$(VehDesc$, 19);
      PRINT #10, TAB(22); tbveh(1).VehVin;
      PRINT #10, TAB(63); tbveh(1).VehLic
      VP = VP + 1
      END IF
        VehRecord! = tbveh(1).VehNext
      WEND

      FOR CNT = VP + 14 TO 39
       PRINT #10, ""
      NEXT CNT



    ' Boats Here
      BoatRecord! = TBCust(1).BoatFirstTrans
      BP = 0

     WHILE BoatRecord! <> 0
      GET TBBTFile, BoatRecord!, TBBoat(1)
     IF TBBoat(1).BoatActive = "Y" THEN
       IF BP > 0 THEN
        PRINT #10, CHR$(12);
        GOSUB SheetHeading
        FOR CNT = 14 TO 39
         PRINT #10, ""
        NEXT CNT
       END IF

      VehDesc$ = TBBoat(1).BoatYear + " " + RTRIM$(TBBoat(1).BoatMake) + " "
      VehDesc$ = LTRIM$(VehDesc$)
      PRINT #10, TAB(12); LEFT$(VehDesc$, 13);
      PRINT #10, TAB(26); TBBoat(1).BoatHull;
      PRINT #10, TAB(40); TBBoat(1).BoatLgth
      PRINT #10, TAB(12); TBBoat(1).BoatMotorMake;
      PRINT #10, TAB(48); TBBoat(1).BoatMotorHP
      BP = 1
     END IF
      BoatRecord! = TBBoat(1).BoatNext
      WEND

    'IF BOAT PRINTED THEN SPACE FOR MOBILE HOME
      IF BP > 0 THEN
       FOR CNT = 42 TO 51:
       PRINT #10, ""
       NEXT CNT
      END IF

     
     ' Mobile Homes Here


      MHRecord! = TBCust(1).MHFirstTrans
      MH = 0

     WHILE MHRecord! <> 0
      GET TBMHFile, MHRecord!, TBMH(1)
     IF TBMH(1).MHActive = "Y" THEN
       IF MH > 2 THEN
        PRINT #10, CHR$(12);
        GOSUB SheetHeading
        FOR CNT = 14 TO 51
         PRINT #10, ""
        NEXT CNT
         MH = 0
       END IF

      MHDesc$ = TBMH(1).MHYear + " " + TBMH(1).MHMake
      PRINT #10, TAB(2); LEFT$(TBMH(1).MHLandOwner, 23);
      PRINT #10, TAB(27); LEFT$(MHDesc$, 13);
      PRINT #10, TAB(41); TBMH(1).MHLgth
      MH = MH + 1
     END IF
      MHRecord! = TBMH(1).MHNext
      WEND

   
      PRINT #10, CHR$(12);
      RETURN


SheetHeading:
     noline = 0
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, TAB(3); Array(count).rec
     PRINT #10, ""
     PRINT #10, ""
  'line 6
     PRINT #10, TAB(2); TBCust(1).CustName;
        PRINT #10, TAB(47); DistrictName$

      IF LEN(RTRIM$(TBCust(1).Address1)) > 0 THEN
       PRINT #10, TAB(2); TBCust(1).Address1
       ELSE
       noline = noline + 1
      END IF

      IF LEN(RTRIM$(TBCust(1).Address2)) > 0 THEN
       PRINT #10, TAB(2); TBCust(1).Address2
       ELSE
       noline = noline + 1
      END IF

      PRINT #10, TAB(2); TBCust(1).City; " "; TBCust(1).State; " "; TBCust(1).Zip
      FOR Ncnt = 1 TO noline: PRINT #10, " ": NEXT Ncnt
      PRINT #10,
      PRINT #10, TAB(50); TBCust(1).SocSec;
      IF VAL(TBCust(1).SocSec2) > 0 THEN PRINT #10, TAB(69); TBCust(1).SocSec2 ELSE PRINT #10, ""
      PRINT #10,
      PRINT #10,
      RETURN

Alignment:
     LOCATE 10, 1: COLOR 11
     PRINT "Print Alignment (Y/N) or (ESC-ABORT PROCESS) .. "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN CLOSE : RUN "tbccustr"
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO Alignment
     IF in$ = "N" THEN RETURN
     FOR CNT = 1 TO 4
      LPRINT ""
     NEXT CNT
      LPRINT
      LPRINT TAB(2); STRING$(40, "X"); TAB(47); STRING$(25, "X")
      LPRINT TAB(2); STRING$(40, "X")
      LPRINT TAB(2); STRING$(40, "X")
      LPRINT TAB(2); STRING$(40, "X")
      LPRINT
      LPRINT TAB(50); "XXX-XX-XXXX"; TAB(69); "XXX-XX-XXXX"
      LPRINT
      LPRINT
      LPRINT TAB(2); STRING$(15, "X"); TAB(22); STRING$(17, "X")
      LPRINT TAB(2); STRING$(15, "X"); TAB(22); STRING$(17, "X")
      LPRINT TAB(2); STRING$(15, "X"); TAB(22); STRING$(17, "X")
     FOR CNT = 17 TO 39
      LPRINT ""
     NEXT CNT
       LPRINT TAB(12); STRING$(12, "X"); TAB(26); STRING$(12, "X")
       LPRINT TAB(12); STRING$(12, "X"); TAB(26); STRING$(12, "X")
     FOR CNT = 42 TO 51
      LPRINT ""
     NEXT CNT
       LPRINT TAB(2); STRING$(20, "X"); TAB(27); STRING$(12, "X")
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


