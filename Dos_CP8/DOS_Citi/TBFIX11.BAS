10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbcustrep
30 REM Customer Report Menu
40 REM Last Revision Date: 8-24-87
50 REM by Nick Picerno
 
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

   CLOSE
   p = 0: page = 0
   OPEN "R", 1, TBpath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
   CurrentYear$ = year$
   CLOSE 1
   'ON ERROR GOTO networkerror


     REDIM TransLink!(500)
     REDIM TBRTCust(1) AS TBRECust
     REDIM TBRECust(1) AS TBRECust

MainBody:
     CLS : scren$ = "REAL CERT"
     GOSUB 63500: COLOR 11
     LOCATE 3, 1: COLOR 15
     PRINT "REAL ESTATE CERTIFICATION TO TREASURER"
     PRINT
     PRINT "PLEASE CONTACT TREASURER AND MAKE SURE NO ONE IS IN THE TREASURERS SYSTEM"
     PRINT "BEFORE BEGINNING THIS PROCESS"
     PRINT
MB1: LOCATE 8, 1: COLOR 15
     PRINT "IS EVERYONE OUT OF THE TREASURERS SYSTEM (Y/N) .. "; : fl = 1
     GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO MB1
     IF in$ = "N" THEN RUN "TBCMENU"

MB2: LOCATE 9, 1: COLOR 15
     PRINT
     PRINT CHR$(7); CHR$(7);
     PRINT "ARE YOU SURE : LAST CHANCE BEFORE PROCESSING START (Y/N) .. "; : fl = 1
     GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO MB1
     IF in$ = "N" THEN RUN "TBCMENU"
 
MB3: LOCATE 12, 1: COLOR 11
     PRINT "Now Updating the Treasurer's System : Please Wait!"

    'Open Files
     TBTFile = FREEFILE
     OPEN TBpath$ + "TBRTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTFile LEN = LEN(TBRTCust(1))
     NumOfTRecs& = LOF(TBTFile) / LEN(TBRTCust(1))

     TBFile = FREEFILE
     OPEN TBpath$ + "TBRECUST.FIX" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
     NumOfRecs& = LOF(TBFile) / LEN(TBRECust(1))

MainLoop:
     FOR Cnt& = 1 TO NumOfRecs&



     GET TBFile, Cnt&, TBRECust(1)
     
     LOCATE 15, 1
     COLOR 11
     PRINT "Now Updating Acct # "; USING "######"; Cnt&
     IF Cnt& > NumOfTRecs& THEN GOSUB AddAcct: GOTO NxtOne
     
     GET TBTFile, Cnt&, TBRTCust(1)
      TBRTCust(1).MPCurYrBillNumber = TBRECust(1).MPCurYrBillNumber
      TBRTCust(1).MPCurYrTaxAmount = TBRECust(1).MPCurYrTaxAmount
      TBRTCust(1).LBCurYrBillNumber = TBRECust(1).LBCurYrBillNumber
      TBRTCust(1).LBCurYrTaxAmount = TBRECust(1).LBCurYrTaxAmount
      TBRTCust(1).CustName = TBRECust(1).CustName
      TBRTCust(1).Address1 = TBRECust(1).Address1
      TBRTCust(1).Address2 = TBRECust(1).Address2
      TBRTCust(1).Address3 = TBRECust(1).Address3
      TBRTCust(1).ZipCode = TBRECust(1).ZipCode
      TBRTCust(1).PropDesc1 = TBRECust(1).PropDesc1
      TBRTCust(1).PropDesc2 = TBRECust(1).PropDesc2
      TBRTCust(1).PropDesc3 = TBRECust(1).PropDesc3
      TBRTCust(1).PropDesc4 = TBRECust(1).PropDesc4
      TBRTCust(1).ClassCode = TBRECust(1).ClassCode
      TBRTCust(1).Comment1 = TBRECust(1).Comment1
      TBRTCust(1).Comment2 = TBRECust(1).Comment2
      TBRTCust(1).Comment3 = TBRECust(1).Comment3
      TBRTCust(1).LandValue = TBRECust(1).LandValue
      TBRTCust(1).BldgValue = TBRECust(1).BldgValue
      TBRTCust(1).LandUseValue = TBRECust(1).LandUseValue
     PUT TBTFile, Cnt&, TBRTCust(1)
     

NxtOne:
     NEXT Cnt&
       CLOSE
      LOCATE 20, 1
     COLOR 11
     PRINT CHR$(7);
     PRINT "The Certification Has Completed, Please Notify the Treasurer's Office"
     PRINT "that the SYSTEM is Ready For Printing the Bills."
     PRINT "Press <ENTER> to Continue "; : fl = 1
     GOSUB 62000
     END

AddAcct:
      TBRTCust(1).MPCurYrBillNumber = TBRECust(1).MPCurYrBillNumber
      TBRTCust(1).MPCurYrTaxAmount = TBRECust(1).MPCurYrTaxAmount
      TBRTCust(1).LBCurYrBillNumber = TBRECust(1).LBCurYrBillNumber
      TBRTCust(1).LBCurYrTaxAmount = TBRECust(1).LBCurYrTaxAmount
      TBRTCust(1).CustName = TBRECust(1).CustName
      TBRTCust(1).Address1 = TBRECust(1).Address1
      TBRTCust(1).Address2 = TBRECust(1).Address2
      TBRTCust(1).Address3 = TBRECust(1).Address3
      TBRTCust(1).ZipCode = TBRECust(1).ZipCode
      TBRTCust(1).PropDesc1 = TBRECust(1).PropDesc1
      TBRTCust(1).PropDesc2 = TBRECust(1).PropDesc2
      TBRTCust(1).PropDesc3 = TBRECust(1).PropDesc3
      TBRTCust(1).PropDesc4 = TBRECust(1).PropDesc4
      TBRTCust(1).ClassCode = TBRECust(1).ClassCode
      TBRTCust(1).Comment1 = TBRECust(1).Comment1
      TBRTCust(1).Comment2 = TBRECust(1).Comment2
      TBRTCust(1).Comment3 = TBRECust(1).Comment3
      TBRTCust(1).LandValue = TBRECust(1).LandValue
      TBRTCust(1).BldgValue = TBRECust(1).BldgValue
      TBRTCust(1).LandUseValue = TBRECust(1).LandUseValue
      TBRTCust(1).Zone = TBRECust(1).Zone
      TBRTCust(1).NewConst = TBRECust(1).NewConst
      TBRTCust(1).NewValue = TBRECust(1).NewValue
      TBRTCust(1).NewAddress = TBRECust(1).NewAddress
      TBRTCust(1).LastDateChange = TBRECust(1).LastDateChange
      TBRTCust(1).Deleted = "N"
      TBRTCust(1).FirstTrans = 0
      TBRTCust(1).LastTrans = 0
      TBRTCust(1).RELandValue = TBRECust(1).RELandValue
      TBRTCust(1).REBldgValue = TBRECust(1).REBldgValue
      TBRTCust(1).NewBldgValue = TBRECust(1).NewBldgValue
      TBRTCust(1).NewCom1 = TBRECust(1).NewCom1
      TBRTCust(1).NewCom2 = TBRECust(1).NewCom2
      TBRTCust(1).BBCurYrBillNumber = ""
      TBRTCust(1).BBCurYrTaxAmount = 0
      TBRTCust(1).LastYearBilled = 0
      TBRTCust(1).Pad = ""
      PUT TBTFile, Cnt&, TBRTCust(1)
      RETURN


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
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 61 THEN cf = 13: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 62 THEN cf = 14: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 63 THEN cf = 15: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 64 THEN cf = 16: RETURN

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

