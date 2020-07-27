'**********************************************************************
'Chief,
'   Thank you for helping us with there code, I have no idea why
'   there totals are off by about $30,000 - 130 bills
'
'   Again Thank you,  Dale
'
'**********************************************************************
10
DEFINT A-Z

REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
REM tbcustrep
REM Customer Report Menu
REM Last Revision Date: 8-24-87
REM by Nick Picerno
  
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
  TBPath$ = "F:\"               ' change to "F:\" Before compiling
  CLOSE
  p = 0: page = 0
  OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
  CurrentYear$ = year$
  CLOSE 1
  
  'REDIM TransLink!(500)
  REDIM TBRTCust(1) AS TBRECust
  REDIM TBRTCustBal(1) AS TBReBal
  REDIM TBRTBalCust(1)            AS TBReBal
  REDIM TBRECust(1) AS TBRECust
  'REDIM TBRTCust(1) AS TBRECust
  
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
  OPEN TBPath$ + "TBRTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTFile LEN = LEN(TBRTCust(1))
  NumOfTrecs& = LOF(TBTFile) / LEN(TBRTCust(1))

  TBTBalFile = FREEFILE
  OPEN TBPath$ + "TBRTBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTBalFile LEN = LEN(TBRTCustBal(1))
  OPEN "TREATIK" FOR RANDOM ACCESS READ WRITE SHARED AS #10 LEN = 356
  FIELD 10, 356 AS a$
  
  'OPEN "TREAT.exp" FOR OUTPUT AS #50
  'First determine highest new acct number and setup null fields  np9-23-04

  FOR x = 1 TO LOF(10) / 356
    GET 10, x
   ' PRINT #50, a$
    acct$ = LEFT$(a$, 10)
    jj = VAL(acct$)
    IF jj > H& THEN H& = jj
  NEXT x

  'CLOSE : END

  IF H& > NumOfTrecs& THEN
    FOR Tmp& = (NumOfTrecs&) + 1 TO H&
      GET TBTFile, Tmp&, TBRTCust(1)
      TBRTCust(1).LBCurYrBillNumber = ""
      TBRTCust(1).LBCurYrTaxAmount = 0
      TBRTCust(1).MPCurYrBillNumber = ""
      TBRTCust(1).MPCurYrTaxAmount = 0
      TBRTCust(1).CustName = ""
      TBRTCust(1).Address1 = ""
      TBRTCust(1).Address2 = ""
      TBRTCust(1).Address3 = ""
      TBRTCust(1).ZipCode = ""
      TBRTCust(1).PropDesc1 = ""
      TBRTCust(1).PropDesc2 = ""
      TBRTCust(1).PropDesc3 = ""
      TBRTCust(1).PropDesc4 = ""
      TBRTCust(1).ClassCode = ""
      TBRTCust(1).Landvalue = 0
      TBRTCust(1).BldgValue = 0
      TBRTCust(1).LandUseValue = 0
      TBRTCust(1).MapNumber = ""
      TBRTCust(1).district = ""
      TBRTCust(1).Deleted = "Y"
      TBRTCust(1).FirstTrans = 0
      TBRTCust(1).LastTrans = 0
      TBRTCust(1).BBCurYrTaxAmount = 0
      TBRTCust(1).LastYearBilled = 0
      TBRTCust(1).Pad = ""
      PUT TBTFile, Tmp&, TBRTCust(1)
      GOSUB RealNull
    NEXT Tmp&
  END IF
  
  
  
MainLoop:
  FOR x = 1 TO LOF(10) / 356
    GET 10, x
    acct$ = LEFT$(a$, 10)
    name1$ = MID$(a$, 11, 35)

'IF VAL(acct$) = 78 THEN STOP

    kk = INSTR(name1$, ",")
    IF kk > 0 THEN
      MID$(name1$, kk, 1) = " "
    END IF
    name2$ = MID$(a$, 46, 34)
    address$ = MID$(a$, 80, 26)
    city$ = MID$(a$, 106, 14)
    state$ = MID$(a$, 120, 2)
    zip$ = MID$(a$, 122, 10)
    Desc1$ = MID$(a$, 132, 18)
    Desc2$ = MID$(a$, 150, 18)
    deedbk$ = MID$(a$, 168, 4)
    deedpg$ = MID$(a$, 172, 4)
    willbk$ = MID$(a$, 176, 4)
    willpg$ = MID$(a$, 180, 4)
    plat$ = MID$(a$, 184, 10)
    acres$ = MID$(a$, 194, 6)
    totuseval$ = MID$(a$, 200, 8)
    Landvalue$ = MID$(a$, 208, 8)
    Landimp$ = MID$(a$, 216, 9)
    totval$ = MID$(a$, 225, 10)
    TaxAmt$ = MID$(a$, 235, 11)
    grantor$ = MID$(a$, 246, 35)
    district$ = MID$(a$, 281, 4)
    insert$ = MID$(a$, 285, 6)
    dblcir$ = MID$(a$, 291, 2)
    block$ = MID$(a$, 293, 3)
    parno$ = MID$(a$, 296, 5)
    e911rd$ = MID$(a$, 301, 14)
    e911add$ = MID$(a$, 315, 6)
    class$ = MID$(a$, 321, 2)
    excode$ = MID$(a$, 323, 1)
    ticketnum$ = MID$(a$, 324, 6)
    parcrecnum$ = MID$(a$, 330, 6)
    zoning1$ = MID$(a$, 336, 3)
    zoning2$ = MID$(a$, 339, 3)
    zoning3$ = MID$(a$, 342, 3)
    loans$ = MID$(a$, 345, 1)
    
    Cnt& = VAL(acct$)
    
    T# = T# + VAL(TaxAmt$)
    IF VAL(TaxAmt$) > 0 THEN
      BBCnt = BBCnt + 1
    END IF

    LOCATE 15, 1
    COLOR 11
    PRINT "Now Updating Acct # "; USING "######"; Cnt&
    
    GET TBTFile, Cnt&, TBRTCust(1)
    
    TBRTCust(1).MPCurYrBillNumber = ""
    TBRTCust(1).MPCurYrTaxAmount = 0
    TBRTCust(1).LBCurYrBillNumber = LTRIM$(ticketnum$)
    TBRTCust(1).LBCurYrTaxAmount = VAL(TaxAmt$)

    TBRTCust(1).CustName = name1$
    IF LEN(RTRIM$(name2$)) = 0 THEN
      TBRTCust(1).Address1 = address$
      TBRTCust(1).Address2 = ""
    ELSE
      TBRTCust(1).Address1 = name2$
      TBRTCust(1).Address2 = address$
    END IF
    TBRTCust(1).Address3 = RTRIM$(city$) + " " + state$ + " " + LTRIM$(zip$)
    TBRTCust(1).ZipCode = LTRIM$(zip$)
    TBRTCust(1).PropDesc1 = Desc1$
    TBRTCust(1).PropDesc2 = Desc2$
    TBRTCust(1).PropDesc3 = ""
    TBRTCust(1).PropDesc4 = insert$ + dblcir$ + black$ + parno$
    TBRTCust(1).ClassCode = class$
    TBRTCust(1).Landvalue = VAL(Landvalue$)
    TBRTCust(1).BldgValue = VAL(Landimp$)
    TBRTCust(1).LandUseValue = VAL(totuseval$)
    'Don't redo map numbers in wrong format from Pat now (2004 data)
    '      TBRTCust(1).MapNumber = insert$ + dblcir$ + black$ + parno$
    TBRTCust(1).district = LEFT$(LTRIM$(district$), 2)
    TBRTCust(1).Deleted = "N"
    PUT TBTFile, Cnt&, TBRTCust(1)
    
NxtOne:
  NEXT x
'***********
'Nick, This close statement was causing a crash. I'm unsure if it needs to
'exit here or, continue on to the gosub TOWNSMAINLOOP.
  'CLOSE

  GOSUB TOWNSMAINLOOP

  LOCATE 20, 1
  COLOR 11
  PRINT CHR$(7);
  PRINT "The Certification Has Completed, Please Notify the Treasurer's Office"
  PRINT "that the SYSTEM is Ready For Printing the Bills."
  PRINT "Press <ENTER> to Continue "; : fl = 1
  GOSUB 62000
END
  RUN "TBCMENU"
  
AddAcct:
  TBRTCust(1).LBCurYrBillNumber = LTRIM$(ticketnum$)
  TBRTCust(1).LBCurYrTaxAmount = VAL(TaxAmt$)
  TBRTCust(1).CustName = name1$
  TBRTCust(1).Address1 = name2$
  TBRTCust(1).Address2 = address$
  TBRTCust(1).Address3 = RTRIM$(city$) + " " + state$
  TBRTCust(1).ZipCode = zip$
  TBRTCust(1).PropDesc1 = Desc1$
  TBRTCust(1).PropDesc2 = Desc2$
  TBRTCust(1).PropDesc3 = ""
  TBRTCust(1).PropDesc4 = insert$ + dblcir$ + black$ + parno$
  TBRTCust(1).ClassCode = class$
  TBRTCust(1).Landvalue = VAL(Landval$)
  TBRTCust(1).BldgValue = VAL(Landimp$)
  TBRTCust(1).LandUseValue = VAL(totuseval$)
  TBRTCust(1).MapNumber = insert$ + dblcir$ + black$ + parno$
  TBRTCust(1).district = LEFT$(LTRIM$(district$), 2)
  TBRTCust(1).Deleted = "N"
  TBRTCust(1).FirstTrans = 0
  TBRTCust(1).LastTrans = 0
  TBRTCust(1).BBCurYrTaxAmount = 0
  TBRTCust(1).LastYearBilled = 0
  TBRTCust(1).Pad = ""
  PUT TBTFile, Cnt&, TBRTCust(1)
  RETURN
  
TOWNSMAINLOOP:
  FOR x = 1 TO LOF(10) / 356
    GET 10, x
    acct$ = LEFT$(a$, 10)
    name1$ = MID$(a$, 11, 35)
    kk = INSTR(name1$, ",")
    IF kk > 0 THEN
      MID$(name1$, kk, 1) = " "
    END IF
    name2$ = MID$(a$, 46, 34)
    address$ = MID$(a$, 80, 26)
    city$ = MID$(a$, 106, 14)
    state$ = MID$(a$, 120, 2)
    zip$ = MID$(a$, 122, 10)
    Desc1$ = MID$(a$, 132, 18)
    Desc2$ = MID$(a$, 150, 18)
    deedbk$ = MID$(a$, 168, 4)
    deedpg$ = MID$(a$, 172, 4)
    willbk$ = MID$(a$, 176, 4)
    willpg$ = MID$(a$, 180, 4)
    plat$ = MID$(a$, 184, 10)
    acres$ = MID$(a$, 194, 6)
    totuseval$ = MID$(a$, 200, 8)
    Landvalue$ = MID$(a$, 208, 8)
    Landimp$ = MID$(a$, 216, 9)
    totval$ = MID$(a$, 225, 10)
    TaxAmt1$ = MID$(a$, 347, 11)
    grantor$ = MID$(a$, 246, 35)
    district$ = MID$(a$, 281, 4)
    insert$ = MID$(a$, 285, 6)
    dblcir$ = MID$(a$, 291, 2)
    block$ = MID$(a$, 293, 3)
    parno$ = MID$(a$, 296, 5)
    e911rd$ = MID$(a$, 301, 14)
    e911add$ = MID$(a$, 315, 6)
    class$ = MID$(a$, 321, 2)
    excode$ = MID$(a$, 323, 1)
    ticketnum$ = MID$(a$, 324, 6)
    parcrecnum$ = MID$(a$, 330, 6)
    zoning1$ = MID$(a$, 336, 3)
    zoning2$ = MID$(a$, 339, 3)
    zoning3$ = MID$(a$, 342, 3)
    loans$ = MID$(a$, 345, 1)
    
    Cnt& = VAL(acct$)
    
'    T# = T# + VAL(TaxAmt1$)
    
    LOCATE 15, 1
    COLOR 11
    PRINT "Now Updating Acct # "; USING "######"; Cnt&
    
    GET TBTFile, Cnt&, TBRTCust(1)
    IF VAL(TaxAmt1$) > 0 THEN
      TBRTCust(1).MPCurYrBillNumber = LTRIM$(ticketnum$)
      TBRTCust(1).MPCurYrTaxAmount = VAL(TaxAmt1$)
      TBRTCust(1).CustName = name1$
      IF LEN(RTRIM$(name2$)) = 0 THEN
        TBRTCust(1).Address1 = address$
        TBRTCust(1).Address2 = ""
      ELSE
        TBRTCust(1).Address1 = name2$
        TBRTCust(1).Address2 = address$
      END IF
      TBRTCust(1).Address3 = RTRIM$(city$) + " " + state$ + " " + LTRIM$(zip$)
      TBRTCust(1).ZipCode = LTRIM$(zip$)
      TBRTCust(1).PropDesc1 = Desc1$
      TBRTCust(1).PropDesc2 = Desc2$
      TBRTCust(1).PropDesc3 = ""
      TBRTCust(1).PropDesc4 = insert$ + dblcir$ + black$ + parno$
      TBRTCust(1).ClassCode = class$
      TBRTCust(1).Landvalue = VAL(Landvalue$)
      TBRTCust(1).BldgValue = VAL(Landimp$)
      TBRTCust(1).LandUseValue = VAL(totuseval$)
      'Don't redo map numbers in wrong format from Pat now (2004 data)
      '      TBRTCust(1).MapNumber = insert$ + dblcir$ + black$ + parno$
      TBRTCust(1).district = LEFT$(LTRIM$(district$), 2)
      TBRTCust(1).Deleted = "N"
      
      PUT TBTFile, Cnt&, TBRTCust(1)
      DidCnt = DidCnt + 1
    ELSE
      nobillcnt = nobillcnt + 1
    END IF
    
TOWNSNxtOne:
  NEXT x
  CLOSE
  RETURN
  
RealNull:
  FOR Lnt! = 1 TO 10
    TBRTCustBal(1).TaxAmt(Lnt!) = 0
    TBRTCustBal(1).IntAmt(Lnt!) = 0
    TBRTCustBal(1).PenAmt(Lnt!) = 0
    TBRTCustBal(1).OthAmt(Lnt!) = 0
    TBRTCustBal(1).LateAmt(Lnt!) = 0
    TBRTCustBal(1).TaxNotice(Lnt!) = ""
  NEXT Lnt!
  PUT TBTBalFile, Tmp&, TBRTBalCust(1)
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
  

