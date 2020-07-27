
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
    DIM VREC!(250)
100 GOSUB 63000: REM get town in town$
110 SCREN$ = "1.00.00": SYSDATE$ = DATE$
     CLOSE
     OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS Year$: GET 1, 1
     CURRENTYEAR$ = Year$: CLOSE
     taxtype$ = ""
120 CLS : GOSUB 63500

5000 CLS : SCREN$ = "1.05.00": CLS : GOSUB 63500
     LOCATE 1, 25: COLOR 11: PRINT "Clear Comment's/High Mileage"; : LOCATE 3, 1
     COLOR 11
     PRINT "Do you wish to CLEAR the COMMENT LINES NOW? (Y/N) .. "; : fl = 1
     GOSUB 62210
     Commentline$ = IN$
     COLOR 11
     PRINT "Do you wish to CLEAR the HIGH MILEAGE FLAG NOW? (Y/N) .. "; : fl = 1
     GOSUB 62210
     HighMile$ = IN$

      IF Commentline$ = "Y" THEN
        taxtype$ = "R"
        GOSUB 60000
       FOR Cnt! = 1 TO LOF(tbfile) / LEN(TBRECust(1))
        COLOR 11
        LOCATE 5, 1: PRINT "Account # "; Cnt!; " Now Clearing Comment Lines"
        GET tbfile, Cnt!, TBRECust(1)
        TBRECust(1).Comment1 = ""
        TBRECust(1).Comment2 = ""
        TBRECust(1).Comment3 = ""
        PUT tbfile, Cnt!, TBRECust(1)
       NEXT Cnt!
       CLOSE tbfile
     END IF

    IF HighMile$ = "Y" THEN
        taxtype$ = "P"
        GOSUB 60000

     NumOfVRecs = LOF(TBVFile) \ LEN(TBVeh(1))
     FOR Cnt& = 1 TO NumOfVRecs
        COLOR 11
        LOCATE 7, 1: PRINT "Vehicle # "; Cnt&; " Now Resetting High Mileage Flag"

        GET TBVFile, Cnt&, TBVeh(1)
         TBVeh(1).VehMiles = ""
        PUT TBVFile, Cnt&, TBVeh(1)

     NEXT Cnt&

     END IF
     CLOSE
     RUN "tbccm"


10000 SCREN$ = "1.10.00": CLS : GOSUB 63500
10010 LOCATE 1, 32: COLOR 11: PRINT "Update Using DMV Data"; : LOCATE 3, 1
     taxtype$ = "R"
     LOCATE 3, 1
     PRINT "This function requires a Password to Process"
     PRINT
     PRINT "Enter the Password Now (ESC-Cancel) .. "; : fl = 6
     GOSUB 62000: IF cf = 1 THEN 10

     Password$ = UCASE$(IN$)
     IF Password$ <> "SOSOFT" THEN 10000
   'Clear Exception File
   EFile = FREEFILE
   OPEN TBPath$ + "EXCEP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #EFile LEN = LEN(TBEX(1))
   CLOSE EFile
   df$ = "DEL " + TBPath$ + "EXCEP.DAT"
   SHELL df$



     OPEN "I", 10, TBPath$ + "dmv2002.txt"
     COLOR 11: LOCATE 5, 1
     PRINT "Now Initializing the Existing Vehicle File       "
     'Initialize Vehicle Data File

     taxtype$ = "P"
     GOSUB 60000

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
    GOTO 10
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
   

   FOR Record! = 1 TO LOF(tbfile) / LEN(TBPPCust(1))  'Main Search Loop Begins
   GET tbfile, Record!, TBPPCust(1)

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
       PUT tbfile, Record!, TBPPCust(1)
      ELSE
      PrevVeh = TBPPCust(1).LastVeh
       TBPPCust(1).LastVeh = (LOF(TBVFile) \ LEN(TBVeh(1)))
      PUT tbfile, Record!, TBPPCust(1)
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
       
      IF TBVeh(1).VehAcct = 0 OR TBVeh(1).VehAcct > LOF(tbfile) / LEN(TBPPCust(1)) THEN
       GOTO SKE
       GET tbfile, TBVeh(1).VehAcct, TBPPCust(1)
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
11000 RUN "TBNADAPP"
12000 RUN "TBNADAUP"

60000  IF taxtype$ = "R" THEN
         CLOSE tbfile
         tbfile = FREEFILE
         OPEN TBPath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #tbfile LEN = LEN(TBRECust(1))
        ELSEIF taxtype$ = "P" THEN
         CLOSE tbfile
         tbfile = FREEFILE
         OPEN TBPath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #tbfile LEN = LEN(TBPPCust(1))
         CLOSE TBVFile
         TBVFile = FREEFILE
         OPEN TBPath$ + "TBVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))
        ELSEIF taxtype$ = "S" THEN
         CLOSE tbfile
         tbfile = FREEFILE
         OPEN TBPath$ + "TBSPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #tbfile LEN = LEN(TBPPCust(1))
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
 nn$ = UCASE$(IN$)
 nn = LEN(nn$)
 IF nn > 20 THEN nn = 20
 nn$ = LEFT$(nn$, nn)

 IF taxtype$ = "R" THEN
  TBIDXFILE = 14
  OPEN TBPath$ + "TBRENAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFILE LEN = 24
  CLOSE 14
  TBIDXFILE = 14
  OPEN TBPath$ + "TBRENAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFILE LEN = 24

  FIELD TBIDXFILE, 20 AS lastnam$, 4 AS Mrec$
  GOSUB 60000
  FOR ll = 1 TO LOF(TBIDXFILE) / 24
  GET TBIDXFILE, ll
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO n1
n3:
  NEXT ll
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  ggo = 0
  RETURN

n1:
  GET tbfile, CVS(Mrec$), TBRECust(1)
  LOCATE 11, 1: COLOR 11
  Nme$ = TBRECust(1).CustName
  COLOR 10
  FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  LOCATE 11, 1: COLOR 11
  PRINT "Full Account Name: "; : COLOR 15: PRINT Nme$
  COLOR 11
  PRINT "   Account Number: "; : COLOR 15: PRINT CVS(Mrec$)
n2:
  LOCATE 13, 1: COLOR 11
  PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
  IF NOT (IN$ = "Y" OR IN$ = "N" OR IN$ = "Q") THEN GOTO n2
  IF IN$ = "Q" THEN
   FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
   CLOSE
   ggo = 0
   RETURN
  END IF
  IF IN$ = "Y" THEN
   ggo = 1: Record! = CVS(Mrec$)
   FOR LLL = 8 TO 13:  LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
   CLOSE
   RETURN
  END IF
  IF IN$ = "N" THEN GOTO n3
  GOTO n2
  END IF

  IF taxtype$ = "P" THEN
  CLOSE
  TBIDXFILE = 14
  OPEN TBPath$ + "TBPPNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFILE LEN = 24
  FIELD TBIDXFILE, 20 AS lastnam$, 4 AS Mrec$
  GOSUB 60000
  FOR ll = 1 TO LOF(TBIDXFILE) / 24
  GET TBIDXFILE, ll
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO PPn1
PPn3:
  NEXT ll
  LOCATE 8, 1: PRINT STRING$(79, 32)
  ggo = 0
  RETURN

PPn1:
GET tbfile, CVS(Mrec$), TBPPCust(1)
LOCATE 11, 1: COLOR 11
Nme$ = TBPPCust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT Nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(Mrec$)
PPn2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (IN$ = "Y" OR IN$ = "N" OR IN$ = "Q") THEN GOTO PPn2
IF IN$ = "Q" THEN
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  ggo = 0
  RETURN
END IF
IF IN$ = "Y" THEN
  ggo = 1: Record! = CVS(Mrec$)
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  RETURN
END IF
IF IN$ = "N" THEN GOTO PPn3
GOTO PPn2
END IF


'************************* M A P  S E A R C H  R O U T I N E ***************

MapSearch:
 LOCATE 8, 1
 COLOR 11
 PRINT "Enter the Map Number (ESC-End) .. "; : fl = 14
 GOSUB 62000
 IF cf = 1 THEN
   ggo = 0
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  RETURN
 END IF

 MapNumber$ = UCASE$(IN$)
 MapLgth = LEN(IN$)
 GOSUB 60000
 FOR Snt& = 1 TO LOF(tbfile) \ (LEN(TBRECust(1)))
 GET tbfile, Snt&, TBRECust(1)

 IF LEFT$(MapNumber$, MapLgth) = LEFT$(RTRIM$(TBRECust(1).MapNumber), MapLgth) THEN
   COLOR 11
   LOCATE 14, 1: PRINT "Full Map Number is "; : COLOR 15
   PRINT LEFT$(TBRECust(1).MapNumber, 3) + "-" + MID$(TBRECust(1).MapNumber, 4, 2) + "-" + MID$(TBRECust(1).MapNumber, 6, 2) + "-" + MID$(TBRECust(1).MapNumber, 8, 2) + "-" + MID$(TBRECust(1).MapNumber, 10, 4) + "-" + MID$(TBRECust(1).MapNumber, 14 _
, 1)
   COLOR 11
   LOCATE 15, 1: PRINT "Customer Name is   "; : COLOR 15: PRINT TBRECust(1).CustName: COLOR 11
   COLOR 11
   LOCATE 16, 1: PRINT "Correct (Y/N) .. "; : fl = 1
   GOSUB 62210
   IF IN$ = "Y" THEN
   Record! = Snt&
   ggo = 1
   RETURN
   END IF
 END IF
 NEXT Snt&
  ggo = 0
  CLOSE
 RETURN

'************************ E N D   M A P   S E A R C H   R O U T I N E  **********

'************************* 911  S E A R C H  R O U T I N E ***************

Search911:
 LOCATE 8, 1
 COLOR 11
 PRINT "Enter the 911 Address (or any part of) (ESC-End) .. "; : fl = 25
 GOSUB 62000
 IF cf = 1 THEN
   ggo = 0
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  RETURN
 END IF

 MapNumber$ = UCASE$(IN$)
 MapLgth = LEN(IN$)
 GOSUB 60000
 FOR Snt& = 1 TO LOF(tbfile) \ (LEN(TBRECust(1)))
 GET tbfile, Snt&, TBRECust(1)

 IF LEFT$(MapNumber$, MapLgth) = LEFT$(RTRIM$(TBRECust(1).NewCom2), MapLgth) THEN
   COLOR 11
   LOCATE 14, 1: PRINT "Full Address is "; : COLOR 15
   PRINT TBRECust(1).NewCom2
   COLOR 11
   LOCATE 15, 1: PRINT "Customer Name is   "; : COLOR 15: PRINT TBRECust(1).CustName: COLOR 11
   COLOR 11
   LOCATE 16, 1: PRINT "Correct (Y/N) .. "; : fl = 1
   GOSUB 62210
   IF IN$ = "Y" THEN
   Record! = Snt&
   ggo = 1
   RETURN
   END IF
 END IF
 NEXT Snt&
  ggo = 0
  CLOSE
 RETURN

'************************ E N D   9 1 1   S E A R C H   R O U T I N E  **********









VINsearch:
 LOCATE 8, 1:
 COLOR 11
 PRINT "Enter Vehicle VIN (Partial Accepted : ESC - End) .. "; : fl = 22
 GOSUB 62000
 IF cf = 1 THEN
  ggo = 0
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  RETURN
 END IF
 VIN$ = UCASE$(IN$)
 VinLgth = LEN(VIN$)
 GOSUB 60000
 FOR Snt& = 1 TO LOF(TBVFile) \ (LEN(TBVeh(1)))
 GET TBVFile, Snt&, TBVeh(1)
  
  IF VIN$ = RTRIM$(LEFT$(TBVeh(1).VehVin, VinLgth)) AND TBVeh(1).VehTyp <> "D" THEN
   Record! = TBVeh(1).VehAcct
   GET tbfile, Record!, TBPPCust(1)
   IF TBPPCust(1).Deleted <> "Y" THEN GOTO VPP1
  END IF
Vpp4:
 NEXT Snt&
 CLOSE TBVFile
 ggo = 0
 RETURN


VPP1:
GET tbfile, Record!, TBPPCust(1)
LOCATE 11, 1: COLOR 11
Nme$ = TBPPCust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT Nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT Record!

VPP2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (IN$ = "Y" OR IN$ = "N" OR IN$ = "Q") THEN GOTO VPP2
IF IN$ = "Q" THEN
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  ggo = 0
  RETURN
END IF
IF IN$ = "Y" THEN
  ggo = 1
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  RETURN
END IF
IF IN$ = "N" THEN GOTO Vpp4
GOTO VPP2

SSsearch:
 LOCATE 8, 1:
 COLOR 11
 PRINT "Enter Social Security # (Partial Accepted : ESC - End) .. "; : fl = 11
 GOSUB 62000
 IF cf = 1 THEN
  ggo = 0
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  RETURN
 END IF
 SSN$ = UCASE$(IN$)
 SSLgth = LEN(SSN$)
 GOSUB 60000
 FOR Record! = 1 TO LOF(tbfile) / LEN(TBPPCust(1))
 GET tbfile, Record!, TBPPCust(1)
 IF SSN$ = LEFT$(TBPPCust(1).SocSec, SSLgth) OR SSN$ = LEFT$(TBPPCust(1).SocSec2, SSLgth) THEN
  IF TBPPCust(1).Deleted <> "Y" THEN GOTO SPP1
 END IF

Spp4:
 NEXT Record!
 CLOSE tbfile
 ggo = 0
 RETURN


SPP1:
LOCATE 11, 1: COLOR 11
Nme$ = TBPPCust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT Nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT Record!

SPP2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (IN$ = "Y" OR IN$ = "N" OR IN$ = "Q") THEN GOTO SPP2
IF IN$ = "Q" THEN
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  ggo = 0
  RETURN
END IF
IF IN$ = "Y" THEN
  ggo = 1
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  RETURN
END IF
IF IN$ = "N" THEN GOTO Spp4
GOTO SPP2



RealNull:
FOR Lnt! = 1 TO 10
 TBRECustBal(1).TaxAmt(Lnt!) = 0
 TBRECustBal(1).IntAmt(Lnt!) = 0
 TBRECustBal(1).PenAmt(Lnt!) = 0
 TBRECustBal(1).OthAmt(Lnt!) = 0
 TBRECustBal(1).LateAmt(Lnt!) = 0
 TBRECustBal(1).TaxNotice(Lnt!) = ""
 NEXT Lnt!
 RETURN
PersNull:
FOR Lnt! = 1 TO 10
 TBPPCustBal(1).PerPropTaxAmt(Lnt!) = 0
 TBPPCustBal(1).FarmEquipAmt(Lnt!) = 0
 TBPPCustBal(1).MachToolsAmt(Lnt!) = 0
 TBPPCustBal(1).MerchCapAmt(Lnt!) = 0
 TBPPCustBal(1).MobileHomeAmt(Lnt!) = 0
 TBPPCustBal(1).LateFeeAmt(Lnt!) = 0
 TBPPCustBal(1).IntAmt(Lnt!) = 0
 TBPPCustBal(1).PenAmt(Lnt!) = 0
 TBPPCustBal(1).TaxNotice(Lnt!) = ""
NEXT Lnt!
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
Deleted:
    PRINT CHR$(7);
    LOCATE 23, 1: COLOR 15
    PRINT "This Account Has Been Deleted! Press <ENTER> to Continue"; :
    fl = 1
    GOSUB 62000
    LOCATE 23, 1: PRINT STRING$(79, 32)
    RETURN
HasBalance:
    PRINT CHR$(7);
    LOCATE 23, 1: COLOR 15
    PRINT "This Account Has a Balance and CANNOT Be Deleted! Press <ENTER> to Continue"; :
    fl = 1
    GOSUB 62000
    LOCATE 23, 1: PRINT STRING$(79, 32)
    RETURN



AddReal1:
        FOR zz = 4 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
        LOCATE 3, 1
        COLOR 10
        PRINT "Real Estate Account Data"; TAB(50); "Account # "
        COLOR 11
PRINT " 1- Map Number ...........    -  -  -  -    -"
PRINT " 2- Name/Address ......... "
PRINT " 3- Address Line 1 ....... "
PRINT " 4- Address Line 2 ....... "
PRINT " 5- Address Line 3 ....... "
PRINT " 6- Zip Code ............. "
PRINT " 7- Prop Descrip Line 1 .."
PRINT " 8- Prop Descrip Line 2 .."
PRINT " 9- Prop Descrip Line 3 .."
PRINT "10- Prop Descrip Line 4 .."
PRINT "11- Zone ................."
PRINT "12- Class Code ..........."
PRINT "13- District ............."
PRINT "14- Size/Acreage ........."
PRINT "15- Comment Line 1 ......."
PRINT "16- Comment Line 2 ......."
PRINT "17- Comment Line 3 ......."
PRINT "18- 911 Address .........."
RETURN

AddReal2:
FOR zz = 4 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
LOCATE 5, 1
PRINT "Property Valuation Data"
PRINT " 1- Land Value ..........."
PRINT " 2- Bldg Value ..........."
PRINT " 3- Land Use Exempt ......"
PRINT " 4- New Construction Y/N ."
PRINT " 5- New Value Y/N ........"
PRINT " 6- New Address .........."
PRINT " 7- Trans Year (yy) ......"
PRINT " 8- New Bldg Value ......."
PRINT " 9- New Bldg Comment 1 ..."
PRINT "10- Mortgage Customer Y .."
RETURN

AddPers1:
FOR zz = 4 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
LOCATE 3, 1
COLOR 10
PRINT "Personal Property Account Data"; TAB(50); "Account # "
COLOR 11
PRINT " 1- Customer Name .........."
PRINT " 2- Address 1 .............."
PRINT " 3- Address 2 .............."
PRINT " 4- City ..................."
PRINT " 5- State .................."
PRINT " 6- Zip Code ..............."
PRINT " 7- Social Sec # ..........."
PRINT " 8- Social Sec # ..........."
PRINT " 9- District ..............."
PRINT "10- Late File (Y/N) ........"
PRINT "11- Business (Y/N) ........."
PRINT "12- Mach Tools Value ......."
PRINT "13- Mobile Home Value ......"
PRINT "14- Merch Cap. Value ......."
PRINT "15- Farm Equip Value ......."
RETURN



AddPers2:
FOR zz = 4 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
LOCATE 3, 1
COLOR 10
PRINT "Personal Property Valuation Detail Screen"
COLOR 11
PRINT " 1- Heavy Equipment ........"
PRINT " 2- Boats .................."
PRINT " 3- Bus Furniture & Equip .."
PRINT "  - Vehicle Total .........."
PRINT "  - Total Personal Property."
PRINT STRING$(79, "-")
PRINT " 4- Vehicle VIN ............"
PRINT " 5- Vehicle Year (YYYY) ...."
PRINT " 6- Vehicle Make ..........."
PRINT " 7- Vehicle Model .........."
PRINT " 8- Vehicle Weight ........."
PRINT " 9- Vehicle Value .........."
PRINT "10- Qualifying Vehicle Y/N ."
PRINT "11- Vehicle License # ......"
PRINT "12- Type/Delete (I/B/D/T/M)."
PRINT "13- Trailor Type (U/B/H/R)..  "
PRINT "14- High Mileage Flag (Y)...  "
PRINT "15- Add New Vehicle"
RETURN

AddBoat:
FOR zz = 4 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
LOCATE 3, 1
COLOR 10
PRINT "Personal Property Boat Valuation Detail Screen"
COLOR 11
PRINT " 1- Boat Year ........"
PRINT " 2- Boat Make ........"
PRINT " 3- Boat Hull ........"
PRINT " 4- Boat Length ......"
PRINT " 5- Boat Date Acq'd..."
PRINT " 6- Boat Cost ........"
PRINT " 7- Motor Make ......."
PRINT " 8- Motor HP ........."
PRINT " 9- Motor Date Acq'd.."
PRINT "10- Motor Cost ......."
PRINT "11- Active (Y) ......."
PRINT "12- Add New Boat/Motor"
RETURN
AddMHome:
FOR zz = 4 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
LOCATE 3, 1
COLOR 10
PRINT "Personal Property Mobile Home Valuation Detail Screen"
COLOR 11
PRINT " 1- MH Land Owner ........"
PRINT " 2- MH Year .............."
PRINT " 3- MH Make  ............."
PRINT " 4- MH Length ............"
PRINT " 5- MH Width ............."
PRINT " 6- MH Date Acq'd  ......."
PRINT " 7- MH Cost .............."
PRINT " 8- Active (Y) ......."
PRINT " 9- Add New Mobile Home"
RETURN

PERSBRKDWN:
FOR zz = 4 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz: COLOR 11
LOCATE 3, 1

COLOR 10

RETURN
printrecord:
taxtype$ = "P"
GOSUB 60000
NumOfVRecs = LOF(TBVFile) \ LEN(TBVeh(1))
     FOR Cnt& = NumOfVRecs TO 1 STEP -1
     GET TBVFile, Cnt&, TBVeh(1)
      
      PRINT Cnt&, TBVeh(1).VehAcct
      IF TBVeh(1).VehAcct = 14604 THEN STOP
      'SLEEP 2
      
      NEXT Cnt&
      CLOSE
      STOP

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
          CASE IS = 6
               RESUME NEXT

          CASE ELSE
               CLS
               LOCATE 15, 1: COLOR 28: PRINT "Error Message: "; : COLOR 15: PRINT "Error "; Number; " Has Occured in Line "; ERL
               PRINT "Write Down Message and Hit <ESC> Key to exit."
               fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
            
     END SELECT

