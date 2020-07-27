10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbccustrp
30 REM Customer Report Menu
40 REM Last Revision Date: 5-1-03
50 REM by Nick Picerno

   ' Must be compiled using the compile directory
   DEFINT A-Z
'$INCLUDE: 'lcTAX.BI'
   
60 CLS :
   TBpath$ = "F:\"                    ' change to "F:\" Before compiling
   CLOSE
   OPEN "R", 1, TBpath$ + "tbyear.dat", 4: FIELD 1, 4 AS YEAR$: GET 1, 1
   CurrentYear$ = YEAR$: CLOSE 1

     REDIM TBPPCust(1) AS TBPPCust
     REDIM tbveh(1) AS PPVehType
     REDIM TBRECust(1) AS TBRECust
     
     WIDTH LPRINT 236
     
100 REM begin program
110 scren$ = "2.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500




1000 REM
1100 scren$ = "2.01.01": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 32: PRINT "Export ALL Data"
     LOCATE 3, 1
     PRINT "Press <ENTER> to Begin or <ESC> to Cancel "; : fl = 1
     GOSUB 62000
     IF cf = 1 THEN RUN "TBCMENU"


     TaxType$ = "P"
     GOSUB 60000
1150 IF LOF(TBFile) / LEN(TBPPCust(1)) = 0 THEN CLOSE : RUN "tbcmenu"
     'OPEN "O", 10, "TBLCVEH.TXT": WIDTH #10, 255
     OPEN "O", 10, "TBLCPP.TXT": WIDTH #10, 255
1170 FOR X = 1 TO LOF(TBFile) / LEN(TBPPCust(1))
     GET TBFile, X, TBPPCust(1)
     LOCATE 3, 1: PRINT STRING$(79, 32)
     LOCATE 3, 1: COLOR 11
     PRINT "Account Being Processed: "; USING "#####,#"; X
     IF TBPPCust(1).Deleted = "Y" THEN
     ELSE
     PRINT #10, CHR$(34); X; CHR$(34);
     PRINT #10, CHR$(34); TBPPCust(1).CustName; CHR$(34);
     PRINT #10, CHR$(34); TBPPCust(1).Address1; CHR$(34);
     PRINT #10, CHR$(34); TBPPCust(1).Address2; CHR$(34);
     PRINT #10, CHR$(34); TBPPCust(1).City; CHR$(34);
     PRINT #10, CHR$(34); TBPPCust(1).State; CHR$(34);
     PRINT #10, CHR$(34); TBPPCust(1).Zip; CHR$(34);
     PRINT #10, CHR$(34); TBPPCust(1).SocSec; CHR$(34);
     PRINT #10, CHR$(34); TBPPCust(1).SocSec2; CHR$(34);
     PRINT #10, CHR$(34); TBPPCust(1).District; CHR$(34);
     PRINT #10, CHR$(34); TBPPCust(1).Business; CHR$(34);
     PRINT #10, CHR$(34); TBPPCust(1).MachTools; CHR$(34);
     PRINT #10, CHR$(34); TBPPCust(1).MerchCap; CHR$(34);
     PRINT #10, CHR$(34); TBPPCust(1).FarmEquip; CHR$(34);
     PRINT #10, CHR$(34); TBPPCust(1).MobileHome; CHR$(34);
     PRINT #10, CHR$(34); TBPPCust(1).PersProp; CHR$(34)
    END IF
    NEXT X
LOCATE 4, 1: COLOR 15: PRINT CHR$(7);
PRINT "Main Customer Pers Property File Converted"

      LPRINT "File Name: TBLCPP.TXT"
      LPRINT "File Desc: Main Pers Property Customer File Info"
      LPRINT "Map of Data Exported"
      LPRINT
      LPRINT "Customer Number"
      LPRINT "Customer Name"
      LPRINT "Customer Name/Address"
      LPRINT "Customer Address"
      LPRINT "City"
      LPRINT "State"
      LPRINT "Zip Code"
      LPRINT "Social Security Number"
      LPRINT "Social Security Number 2"
      LPRINT "District Code #"
      LPRINT "Business Flag YN"
      LPRINT "Machine Tools Value"
      LPRINT "Merchant Capital Value"
      LPRINT "Farm Equip Value"
      LPRINT "Mobile Home Value"
      LPRINT "Pers Prop Value"
      LPRINT CHR$(12);

 LOCATE 5, 1: COLOR 11
 PRINT "Now Processing Vehicles"
   CLOSE 10
    OPEN "O", 10, "TBLCVEH.TXT"
     FOR X = 1 TO LOF(TBFile) / LEN(TBPPCust(1))
     GET TBFile, X, TBPPCust(1)
1191 IF TBPPCust(1).Deleted = "Y" THEN 1230
     
      VehRecord! = TBPPCust(1).FirstVeh
       
   WHILE VehRecord! <> 0
        GET TBVFile, VehRecord!, tbveh(1)
      IF tbveh(1).VehTyp <> "D" THEN
        PRINT #10, CHR$(34); tbveh(1).VehAcct; CHR$(34);
        PRINT #10, CHR$(34); tbveh(1).VehTyp; CHR$(34);
        PRINT #10, CHR$(34); tbveh(1).VehVin; CHR$(34);
        PRINT #10, CHR$(34); tbveh(1).VehMake; CHR$(34);
        PRINT #10, CHR$(34); tbveh(1).VehYear; CHR$(34);
        PRINT #10, CHR$(34); tbveh(1).VehModel; CHR$(34);
        PRINT #10, CHR$(34); tbveh(1).VehWght; CHR$(34);
        PRINT #10, CHR$(34); tbveh(1).VehValue; CHR$(34);
        PRINT #10, CHR$(34); tbveh(1).VehQ; CHR$(34);
        PRINT #10, CHR$(34); tbveh(1).VehLic; CHR$(34);
        PRINT #10, CHR$(34); tbveh(1).VehTrail; CHR$(34);
        PRINT #10, CHR$(34); tbveh(1).VehMiles; CHR$(34)
      END IF
        VehRecord! = tbveh(1).VehNext
     WEND

1230  NEXT X
      LOCATE 6, 1: COLOR 15: PRINT CHR$(7);
      PRINT "Main Vehicle File Converted"

      LPRINT "File Name: TBLCVEH.TXT"
       LPRINT "File Desc: Main Vehicle File Info"
       LPRINT "Map of Data Exported"
       LPRINT
       LPRINT "Customer Account Number"
        LPRINT "Vehicle Type Code"
        LPRINT "Vehicle VIN"
        LPRINT "Vehicle Make"
        LPRINT "Vehicle Model Year"
        LPRINT "Vehicle Model"
        LPRINT "Vehicle Weight"
        LPRINT "Vehicle Value"
        LPRINT "Qualify for PPTRA YN"
        LPRINT "Vehicle License #"
        LPRINT "Trailor Type (if any)"
        LPRINT "High Mileage Flag =Y "
        LPRINT CHR$(12);

1250  CLOSE :


'Now Process the Real Customer Data File"
 CLOSE
 TaxType$ = "R"
 OPEN "O", 10, "TBLCREAL.TXT"
 GOSUB 60000
 COLOR 11
 LOCATE 8, 1: PRINT "Now Converting the Real Estate Customer File"
 FOR X = 1 TO LOF(TBFile) / LEN(TBRECust(1))
 GET TBFile, X, TBRECust(1)
 IF TBRECust(1).Deleted = "Y" THEN
  ELSE
  PRINT #10, CHR$(34); X; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).MapNumber; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).District; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).CustName; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).Address1; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).Address2; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).Address3; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).ZipCode; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).PropDesc1; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).PropDesc2; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).PropDesc3; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).PropDesc4; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).Zone; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).ClassCode; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).Acres; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).Comment1; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).Comment2; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).Comment3; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).LandValue; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).BldgValue; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).LandUseValue; CHR$(34);
  PRINT #10, CHR$(34); TBRECust(1).NewCom2; CHR$(34)
  END IF
  NEXT X



      LOCATE 9, 1: COLOR 15: PRINT CHR$(7);
      PRINT "Real Estate Customer File Converted"

      LPRINT "File Name: TBLCREAL.TXT"
       LPRINT "File Desc: Main Vehicle File Info"
       LPRINT "Map of Data Exported"
       LPRINT
  LPRINT "Customer Acct #"
  LPRINT "Map Number"
  LPRINT "District Code"
  LPRINT "Customer Name"
  LPRINT "Address Line 1"
  LPRINT "Address Line 2"
  LPRINT "Address Line 3"
  LPRINT "Zip Code"
  LPRINT "Prop Desc Line 1"
  LPRINT "Prop Desc Line 2"
  LPRINT "Prop Desc Line 3"
  LPRINT "Prop Desc Line 4"
  LPRINT "Zone Code"
  LPRINT "Class Code"
  LPRINT "Acres"
  LPRINT "Comment 1"
  LPRINT "Comment 2"
  LPRINT "Comment 3"
  LPRINT "Land Value"
  LPRINT "Bldg Value"
  LPRINT "Land Use Exp"
  LPRINT "911 Address"
  
 LPRINT CHR$(12);
 SLEEP 5
 RUN "tbcmenu"

    
60000  IF TaxType$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBpath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBpath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         CLOSE TBVFile
         TBVFile = FREEFILE
         OPEN TBpath$ + "TBVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(tbveh(1))

         END IF
        RETURN



62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(XY): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 61 THEN cf = 13: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 62 THEN cf = 14: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 63 THEN cf = 15: RETURN
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

