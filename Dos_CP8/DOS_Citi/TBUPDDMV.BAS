
  DECLARE FUNCTION num2date$ (Num%)
  DECLARE FUNCTION Date2Num% (Dat$)
  DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
  DECLARE SUB PrintTitle (Title$)
  DECLARE SUB PrintLoading ()
'$INCLUDE: 'LCTax.Bi'
'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'Qscr.BI'
'$INCLUDE: 'SetCnf.BI'



10 CLS
   ON ERROR GOTO networkerror
   TBPath$ = "F:\"                    ' change to "F:\" Before compiling
    REDIM TBRECust(1) AS TBRECust
    REDIM TBRECustBal(1) AS TBREBAL
    REDIM TBPPCust(1) AS TBPPCust
    REDIM TBPPCustBal(1) AS TBPPBal
    REDIM TBTrans(1) AS TBTrans
    REDIM TBVeh(1) AS PPVehType
    REDIM TBEX(1) AS TBExceptionType
    REDIM TBBoat(1) AS PPBoat
    DIM VREC!(250)
100 GOSUB 63000: REM get town in town$
110 SCREN$ = "1.00.00": SYSDATE$ = DATE$
     CLOSE
     OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS Year$: GET 1, 1
     CURRENTYEAR$ = Year$: CLOSE
     taxtype$ = ""
120 CLS : GOSUB 63500

10000 SCREN$ = "1.10.00": CLS : GOSUB 63500
10010 LOCATE 1, 32: COLOR 11: PRINT "Update Using DMV Data"; : LOCATE 3, 1
     taxtype$ = "R"
     LOCATE 3, 1
     PRINT "This function requires a Password to Process"
     PRINT
     PRINT "Enter the Password Now (ESC-Cancel) .. "; : fl = 6
     GOSUB 62000: IF cf = 1 THEN RUN "tbccm"

     Password$ = UCASE$(IN$)
     IF Password$ <> "SOSOFT" THEN 10000
   'Clear Exception File
   EFile = FREEFILE
   OPEN TBPath$ + "EXCEP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #EFile LEN = LEN(TBEX(1))
   CLOSE EFile
   df$ = "DEL " + TBPath$ + "EXCEP.DAT"
   SHELL df$

     OPEN "I", 10, TBPath$ + "dmv2003.txt"
     COLOR 11: LOCATE 5, 1
     PRINT "Now Initializing the Existing Vehicle File       "
     'Initialize Vehicle Data File

     taxtype$ = "P"
     GOSUB 60000
     'Boats First
     NumOfTaxRecs = LOF(TBFile) \ LEN(TBPPCust(1))
     FOR Cnt& = 1 TO NumOfTaxRecs
     GET TBFile, Cnt&, TBPPCust(1)
     TBPPCust(1).Boats = 0
     PUT TBFile, Cnt&, TBPPCust(1)
     NEXT Cnt&

     NumOfBRecs = LOF(TBBTFile) \ LEN(TBBoat(1))
     FOR Cnt& = 1 TO NumOfBTRecs
     GET TBBTFile, Cnt&, TBBoat(1)
      TBBoat(1).BoatCost = 0
      TBBoat(1).BoatMotorCost = 0
      PUT TBBTFile, Cnt&, TBBoat(1)
     NEXT Cnt&



     NumOfVRecs = LOF(TBVFile) \ LEN(TBVeh(1))
     FOR Cnt& = 1 TO NumOfVRecs
     GET TBVFile, Cnt&, TBVeh(1)
      TBVeh(1).VehWght = 0
      TBVeh(1).VehValue = 0
      TBVeh(1).VehUpd = "N"
      PUT TBVFile, Cnt&, TBVeh(1)
     NEXT Cnt&
     
GetDMVData:
   LINE INPUT #10, a$
   IF EOF(10) THEN
    CLOSE
    GOSUB PrintExceptions
    RUN "tbccm"
   END IF
   COLOR 11: LOCATE 5, 1
   PRINT "DMV Record                                                            "
   PRINT "SSN 1  : "; MID$(a$, 5, 12); TAB(50); "# Accts Found: "; f
   PRINT "SSN 2  : "; MID$(a$, 137, 12); TAB(50); "# Not Located: "; NF
   PRINT "L Name : "; MID$(a$, 17, 45)
   PRINT "F Name : "; MID$(a$, 62, 33)
   PRINT "M Name : "; MID$(a$, 95, 31)
   PRINT "Type   : "; MID$(a$, 126, 1)
   PRINT "Ownship: "; MID$(a$, 127, 10)
   PRINT "Add1   : "; MID$(a$, 259, 35)
   PRINT "Add2   : "; MID$(a$, 294, 35)
   PRINT "City   : "; MID$(a$, 329, 22)
   PRINT "State  : "; MID$(a$, 351, 2)
   PRINT "Zip    : "; MID$(a$, 353, 9)
   PRINT "VIN    : "; MID$(a$, 1246, 22): VIN$ = RTRIM$(MID$(a$, 1246, 22))
   PRINT "VehMake: "; MID$(a$, 1268, 20)
   PRINT "VehModl: "; MID$(a$, 1292, 15)
   PRINT "Veh Yr : "; MID$(a$, 1288, 4)
   PRINT "Veh Wt : "; MID$(a$, 1676, 6)
   

  
     FOR Cnt& = NumOfVRecs TO 1 STEP -1
     GET TBVFile, Cnt&, TBVeh(1)
      IF TBVeh(1).VehTyp <> "D" THEN
      IF RTRIM$(TBVeh(1).VehVin) = VIN$ THEN
       TBVeh(1).VehYear = MID$(a$, 1288, 4)
       TBVeh(1).VehMake = MID$(a$, 1268, 20)
       TBVeh(1).VehModel = MID$(a$, 1292, 15)
       TBVeh(1).VehWght = VAL(MID$(a$, 1676, 6))
       TBVeh(1).VehValue = VAL(MID$(a$, 1758, 6))
       TBVeh(1).VehQ = "N"
       TBVeh(1).VehLic = MID$(a$, 1596, 10)
       TBVeh(1).VehTyp = MID$(a$, 126, 1)
       TBVeh(1).VehUpd = "Y"
       f = f + 1
       PUT TBVFile, Cnt&, TBVeh(1)
       GOTO GetDMVData
      END IF
      END IF
     NEXT Cnt&

  'IF VIN NOT FOUND THEN SEARCH SS# UPDATED 1-21-02
  'Search Order Changed by Pats letter of Dec 27, 2001
  'Search Vin First  Then SS#

   GOSUB UpdateTaxAcct

   IF Found$ = "Y" THEN GOTO GetDMVData


     GOSUB UpdateNoFind
     GOTO GetDMVData
     
UpdateTaxAcct:
   Found$ = "N"

  'Search Using Social Security Numbers
   FSoc1$ = MID$(a$, 5, 9): FSoc1$ = RTRIM$(FSoc1$)
   FSoc2$ = ""                     'Don't Use SS#2 by Pats Command on 2-18-99
   

   FOR Record! = 1 TO LOF(TBFile) / LEN(TBPPCust(1))  'Main Search Loop Begins
   GET TBFile, Record!, TBPPCust(1)

  'Make Soc Sec Number a Combined String
   Soc1$ = TBPPCust(1).SocSec
   Soc2$ = TBPPCust(1).SocSec2

   IF MID$(TBPPCust(1).SocSec, 4, 1) = " " THEN
    Soc1$ = LEFT$(TBPPCust(1).SocSec, 3) + MID$(TBPPCust(1).SocSec, 5, 2) + MID$(TBPPCust(1).SocSec, 8, 4)
    Soc1$ = RTRIM$(Soc1$)
   END IF
   IF MID$(TBPPCust(1).SocSec2, 4, 1) = " " THEN
    Soc2$ = LEFT$(TBPPCust(1).SocSec2, 3) + MID$(TBPPCust(1).SocSec2, 5, 2) + MID$(TBPPCust(1).SocSec2, 8, 4)
    Soc2$ = RTRIM$(Soc2$)
   END IF
   IF LEN(FSoc1$) > 0 THEN
        IF FSoc1$ = Soc1$ OR FSoc1$ = Soc2$ THEN
         f = f + 1: GOSUB UpdateTheAccount
         RETURN
        END IF
   END IF

   IF MID$(a$, 126, 1) = "B" THEN

   IF MID$(TBPPCust(1).SocSec, 3, 1) = " " THEN
    Soc1$ = LEFT$(TBPPCust(1).SocSec, 2) + MID$(TBPPCust(1).SocSec, 4, 7)
    Soc1$ = RTRIM$(Soc1$)
   END IF
   IF LEN(FSoc1$) > 0 THEN
        IF FSoc1$ = Soc1$ OR FSoc1$ = Soc2$ THEN
         f = f + 1: GOSUB UpdateTheAccount
         RETURN
        END IF
   END IF
  
   END IF

  NEXT Record!          'Main Search Loop Ends
  RETURN



UpdateNoFind:
  'Not Located
   NF = NF + 1          'Count of Not Found

  'Store For Exception
   EFile = FREEFILE
   OPEN TBPath$ + "EXCEP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #EFile LEN = LEN(TBEX(1))
   TBEX(1).SSN1 = MID$(a$, 5, 12)
   TBEX(1).SSN2 = MID$(a$, 137, 12)
   TBEX(1).LName = MID$(a$, 17, 45)
   TBEX(1).FName = MID$(a$, 62, 33)
   TBEX(1).MName = MID$(a$, 95, 31)
   TBEX(1).VType = MID$(a$, 126, 1)
   TBEX(1).Add1 = MID$(a$, 259, 35)
   TBEX(1).Add2 = MID$(a$, 294, 35)
   TBEX(1).City = MID$(a$, 329, 22)
   TBEX(1).State = MID$(a$, 351, 2)
   TBEX(1).Zip = MID$(a$, 353, 9)
   TBEX(1).VIN = MID$(a$, 1246, 22)
   TBEX(1).VehMake = MID$(a$, 1268, 20)
   TBEX(1).VehModl = MID$(a$, 1292, 15)
   TBEX(1).VehYr = MID$(a$, 1288, 4)
   TBEX(1).VehWt = MID$(a$, 1676, 6)
   TBEX(1).TradeVal = MID$(a$, 1758, 6)
   NextRecord! = (LOF(EFile) / LEN(TBEX(1))) + 1
   PUT EFile, NextRecord!, TBEX(1)
   CLOSE EFile
   RETURN

UpdateTheAccount:
    'Check For Existing Vin In This Account and Update it if necessary
     VehRecord = TBPPCust(1).FirstVeh

     WHILE VehRecord <> 0
      GET TBVFile, VehRecord, TBVeh(1)
     IF RTRIM$(TBVeh(1).VehVin) = RTRIM$(VIN$) THEN
      Found$ = "Y"
      TBVeh(1).VehVin = MID$(a$, 1246, 22)
      TBVeh(1).VehYear = MID$(a$, 1288, 4)
      TBVeh(1).VehMake = MID$(a$, 1268, 20)
      TBVeh(1).VehModel = MID$(a$, 1292, 15)
      TBVeh(1).VehWght = VAL(MID$(a$, 1676, 6))
      TBVeh(1).VehValue = VAL(MID$(a$, 1758, 6))
      TBVeh(1).VehQ = "N"
      TBVeh(1).VehLic = MID$(a$, 1596, 10)
      TBVeh(1).VehTyp = MID$(a$, 126, 1)
      TBVeh(1).VehPad = ""
      TBVeh(1).VehAcct = Record!
      TBVeh(1).VehUpd = "Y"
      PUT TBVFile, VehRecord, TBVeh(1)
      RETURN
     END IF
      VehRecord = TBVeh(1).VehNext
     WEND
     
     Found$ = "Y"
     TBVeh(1).VehVin = MID$(a$, 1246, 22)
     TBVeh(1).VehYear = MID$(a$, 1288, 4)
     TBVeh(1).VehMake = MID$(a$, 1268, 20)
     TBVeh(1).VehModel = MID$(a$, 1292, 15)
     TBVeh(1).VehWght = VAL(MID$(a$, 1676, 6))
     TBVeh(1).VehValue = VAL(MID$(a$, 1758, 6))
     TBVeh(1).VehQ = "N"
     TBVeh(1).VehLic = MID$(a$, 1596, 10)
     TBVeh(1).VehTyp = MID$(a$, 126, 1)
     TBVeh(1).VehPad = ""
     TBVeh(1).VehNext = 0
     TBVeh(1).VehAcct = Record!
     TBVeh(1).VehUpd = "Y"
     PUT TBVFile, (LOF(TBVFile) \ LEN(TBVeh(1))) + 1, TBVeh(1)
      'Update Pointer Here
      IF TBPPCust(1).LastVeh = 0 THEN
       TBPPCust(1).FirstVeh = (LOF(TBVFile) \ LEN(TBVeh(1)))
       TBPPCust(1).LastVeh = (LOF(TBVFile) \ LEN(TBVeh(1)))
       PUT TBFile, Record!, TBPPCust(1)
      ELSE
      PrevVeh = TBPPCust(1).LastVeh
       TBPPCust(1).LastVeh = (LOF(TBVFile) \ LEN(TBVeh(1)))
      PUT TBFile, Record!, TBPPCust(1)
      GET TBVFile, PrevVeh, TBVeh(1)
       TBVeh(1).VehNext = (LOF(TBVFile) \ LEN(TBVeh(1)))
      PUT TBVFile, PrevVeh, TBVeh(1)
      END IF
      RETURN

PrintExceptions:

   CLOSE
   EFile = FREEFILE
   OPEN TBPath$ + "EXCEP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #EFile LEN = LEN(TBEX(1))
   IF LOF(EFile) = 0 THEN CLOSE : RETURN

   ReportFile$ = "TBEXC.REP"
   OPEN "O", 12, ReportFile$
   GOSUB PHeading
   FOR RCnt! = 1 TO (LOF(EFile) / LEN(TBEX(1)))
   GET EFile, RCnt!, TBEX(1)
   Nme$ = RTRIM$(TBEX(1).LName) + " " + RTRIM$(TBEX(1).FName) + " " + LEFT$(TBEX(1).MName, 1)

   
   PRINT #12, "Name: "; RTRIM$(Nme$); TAB(50); "SS# "; TBEX(1).SSN1; TAB(80); "SS# "; TBEX(1).SSN2
   PRINT #12, "Addr: "; TBEX(1).Add1; TAB(50); "Veh Type: "; TBEX(1).VType; TAB(90); "Make: "; TBEX(1).VehMake
   PRINT #12, "      "; TBEX(1).Add2; TAB(50); "     VIN: "; TBEX(1).VIN; TAB(90); "Modl: "; TBEX(1).VehModl
   PRINT #12, "      "; RTRIM$(TBEX(1).City) + " " + TBEX(1).State + " " + TBEX(1).Zip;
   PRINT #12, TAB(56); "Yr: "; TBEX(1).VehYr; TAB(80); "Wt: "; TBEX(1).VehWt; TAB(100); "Value: "; USING "######,#.##"; VAL(TBEX(1).TradeVal)
   PRINT #12,
   lc = lc + 6
   IF lc > 58 THEN
    PRINT #12, CHR$(12);
    GOSUB PHeading
   END IF
   NEXT RCnt!
     

     'Now Review Vehicles On File and Not Updated

     taxtype$ = "P"
     GOSUB 60000
     PRINT #12, CHR$(12);
     GOSUB PHeading1

     FOR Cnt& = 1 TO NumOfVRecs
     GET TBVFile, Cnt&, TBVeh(1)
      IF TBVeh(1).VehUpd = "N" AND TBVeh(1).VehTyp <> "D" THEN
       
      IF TBVeh(1).VehAcct = 0 OR TBVeh(1).VehAcct > LOF(TBFile) / LEN(TBPPCust(1)) THEN
       GOTO SKE
       GET TBFile, TBVeh(1).VehAcct, TBPPCust(1)
      END IF
       IF TBPPCust(1).Deleted <> "Y" THEN


       PRINT #12, "VIN # "; TBVeh(1).VehVin; TAB(40); "Acct #"; TBVeh(1).VehAcct; "Not on DMV Tape"
       PRINT #12, "Year:Make:Model: "; TBVeh(1).VehYear; " "; TBVeh(1).VehMake; " "; TBVeh(1).VehModel
       PRINT #12, "Value:Lic#:Type: "; TBVeh(1).VehValue; " "; TBVeh(1).VehLic; " "; TBVeh(1).VehTyp
       PRINT #12, STRING$(79, "-")
       lc = lc + 4
       END IF
            IF lc > 58 THEN
                PRINT #12, CHR$(12);
                GOSUB PHeading1
            END IF
      END IF
SKE:

   NEXT Cnt&


       CLOSE


      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBEXC.REP"
      Header$ = "DMV Exception Listing"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%

      'KILL ReportFile$
      RETURN


PHeading:
    Page = Page + 1
    PRINT #12, "DMV Exception Report"; TAB(65); "Page # "; Page
    PRINT #12, TAB(65); "  Date "; DATE$
    PRINT #12, STRING$(132, "-")
    lc = 3
    RETURN

PHeading1:
    Page = Page + 1
    PRINT #12, "DMV Exception Report : ON FILE BUT NOT FOUND ON DMV TAPE  Page # "; Page
    PRINT #12, TAB(65); "  Date "; DATE$
    PRINT #12, STRING$(132, "-")
    lc = 3
    RETURN

60000  IF taxtype$ = "R" THEN
         CLOSE TBFile
         TBFile = FREEFILE
         OPEN TBPath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSEIF taxtype$ = "P" THEN
         CLOSE TBFile
         TBFile = FREEFILE
         OPEN TBPath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         CLOSE TBVFile
         TBVFile = FREEFILE
         OPEN TBPath$ + "TBVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))
         TBBTFile = FREEFILE
         OPEN TBPath$ + "TBBOAT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBTFile LEN = LEN(TBBoat(1))
        END IF
60025 TBBalFile = FREEFILE
      IF taxtype$ = "R" THEN
      OPEN TBPath$ + "TBREBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBPath$ + "TBPPBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF
60030 RETURN
        

62000 cf = 0: COLOR 14: IN$ = "": RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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
62110 IN$ = LEFT$(IN$, LEN(IN$) - 1)
62120 ZL = ZL - 1: GOTO 62030
62130 IF ABS(fl) = ZL THEN 62030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 62180
62140 IF z$ = "." AND ZD = 0 THEN ZD = 1: GOTO 62180
62150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF z$ < "0" OR z$ > "9" THEN 62030
62180 PRINT z$; : IN$ = IN$ + z$: ZL = ZL + 1
62190 IF ABS(fl) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT IN$: RETURN
62210 GOSUB 62000: IF LEN(IN$) THEN IN$ = CHR$(ASC(IN$) AND 223)
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
          CASE IS = 6
               RESUME NEXT

          CASE ELSE
               CLS
               LOCATE 15, 1: COLOR 28: PRINT "Error Message: "; : COLOR 15: PRINT "Error "; Number; " Has Occured in Line "; ERL
               PRINT "Write Down Message and Hit <ESC> Key to exit."
               fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
            
     END SELECT

