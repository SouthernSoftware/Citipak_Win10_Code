10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbccustrp
30 REM Customer Report Menu
40 REM Last Revision Date: 8-24-87
50 REM by Nick Picerno

   ' Must be compiled using the compile directory
   DEFINT A-Z
   DECLARE FUNCTION QPTrim$ (a$)
   DECLARE FUNCTION num2date$ (Num%)
   DECLARE FUNCTION Date2Num% (Dat$)
   DECLARE FUNCTION WaitKey% ()
   DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
   DECLARE SUB PrintTitle (Title$)
   DECLARE SUB PrintLoading ()
   DECLARE SUB ShowCursor ()
   DECLARE SUB HideCursor ()
   
'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'Qscr.BI'
'$INCLUDE: 'SetCnf.BI'
'$INCLUDE: 'lcTAX.BI'
   
60 CLS :
   TBpath$ = "F:\"                    ' change to "F:\" Before compiling
   'TBpath$ = "c:\qb45\lc\"                    ' change to "F:\" Before compiling
   CLOSE
   OPEN "R", 1, TBpath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
   CurrentYear$ = year$: CLOSE 1
   ON ERROR GOTO networkerror

     p = 0

     
     REDIM TBPPCust(1) AS TBPPCust
     REDIM TBRECust(1) AS TBRECust
     
     WIDTH LPRINT 80

     CustCount = 0
     TotalLand# = 0
     TotalBldg# = 0
     

100 REM begin program
110 scren$ = "2.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500

1000 REM
1010 scren$ = "2.01.00": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 29: PRINT "Mortgage Customer Listing"
     TaxType$ = "R"
1011 LOCATE 3, 1: COLOR 11
     PRINT "Press <ENTER> to Get Mortgage Customer Report Listing (ESC-Cancel) .. "; : fl = 1
     GOSUB 62000
     IF cf = 1 THEN RUN "tbccustr"
     
     LOCATE 15, 1: COLOR 11: PRINT "Now Printing the Report : Please Wait!"
1100 GOSUB 60000
1150 IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10
     OPEN "O", 10, "TBREMORT.REP": WIDTH #10, 255
     GOSUB 1900
     OPEN TBpath$ + "TBREname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
     FIELD 4, 20 AS lastnam$, 4 AS MREC$: EndofFile = LOF(4) / 24
2170 FOR x = 1 TO EndofFile
     GET 4, x
     GET 1, CVS(MREC$), TBRECust(1)
     Account! = CVS(MREC$)

2191 IF TBRECust(1).Deleted = "Y" THEN 2230
      IF TBRECust(1).MortgageYN = "Y" THEN 2192 ELSE 2230
      
2192 CustName$ = QPTrim$(TBRECust(1).CustName)
     IF LEN(CustName$) = 0 THEN 2230


     PRINT #10, USING "######"; Account!;
     PRINT #10, TAB(10); TBRECust(1).CustName
     lc = lc + 1
     CustCount = CustCount + 1
2220 IF lc >= 58 THEN PRINT #10, CHR$(12); : GOSUB 1900
2230 NEXT x
     PRINT #10, "Total Customers .. "; USING "######,#"; CustCount
     PRINT CHR$(12);
2240 CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBREMORT.REP"
      Header$ = "Real Estate Mortgage Customers"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
2250 GOTO 10
      


1900 p = p + 1

     PRINT #10, TAB(24); "Mortgage Customer Account Listing"
     PRINT #10, "Date: "; DATE$; TAB(68); "Page #"; USING "#####"; p
1910 PRINT #10, "Acct. #"; TAB(10); "Customer Name "
     PRINT #10, STRING$(75, "=")
     lc = 5
     RETURN
     
60000  IF TaxType$ = "R" THEN
         TBFILE = FREEFILE
         OPEN TBpath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFILE LEN = LEN(TBRECust(1))
        ELSE
         TBFILE = FREEFILE
         OPEN TBpath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFILE LEN = LEN(TBPPCust(1))
         END IF
        RETURN

60025 TBBalFile = FREEFILE
      IF TaxType$ = "R" THEN
      OPEN TBpath$ + "TBREBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBpath$ + "TBPPBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF
60030 RETURN

61000 TBTransFile = FREEFILE
      OPEN TBpath$ + "TBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTransFile LEN = LEN(TBTrans(1))
61020 RETURN

62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(Xy): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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
62140 IF z$ = "." AND ZD = 0 THEN ZD = 1: GOTO 62180
62150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF z$ < "0" OR z$ > "9" THEN 62030
62180 PRINT z$; : in$ = in$ + z$: ZL = ZL + 1
62190 IF ABS(fl) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT in$: RETURN
62210 GOSUB 62000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
62220 RETURN
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); scren$: PRINT STRING$(79, "_"): RETURN

GETDISTRICT:
     DistrictName$ = ""
     DistrictName$ = "Lunenburg County"
     IF District = 9 THEN DistrictName$ = "TOWN OF KENBRIDGE"
     IF District = 10 THEN DistrictName$ = "TOWN OF VICTORIA"
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

