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
   'TBpath$ = "F:\"                    ' change to "F:\" Before compiling
   'TBpath$ = "c:\qb45\lc\"                    ' change to "F:\" Before compiling
   CLOSE
   OPEN "R", 1, TBpath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
   CurrentYear$ = year$: CLOSE 1
   ON ERROR GOTO networkerror

     p = 0
     REDIM TransLink!(500)
     REDIM LandValue#(99)
     REDIM BldgValue#(99)
     REDIM TaxAmt#(99)
     REDIM Customer!(99)
    
     REDIM TBRECust(1) AS TBRECust
     REDIM TBRECustBal(1) AS TBREBAL
     REDIM TBPPCust(1) AS TBPPCust
     REDIM TBPPCustBal(1) AS TBPPBal
     REDIM TBTrans(1) AS TBTrans
     REDIM TBPAGE(1) AS TBPAGE
     
     WIDTH LPRINT 236



     CustCount = 0
     TotalLand# = 0
      TotalBldg# = 0
     




     
100 REM begin program
110 scren$ = "2.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500

1000 REM
1010 scren$ = "2.01.00": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 32: PRINT "Alpha Customer Listing"
     LOCATE 3, 1
     PRINT "Enter Tax Type (R=Real : P=Personal : <ESC>-Cancel) . "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN RUN "tbccustr"
     IF NOT (in$ = "R" OR in$ = "P") THEN 1010
     TaxType$ = UCASE$(in$)
1011 LOCATE 4, 1: COLOR 11
     PRINT "Enter District Code to Pull (0=All : <ESC>-Cancel) .. "; : fl = -2
     GOSUB 62000
     District = VAL(in$)
1012 LOCATE 5, 1: COLOR 11
     PRINT "Enter Class Code to Pull (0=All : <ESC>-Cancel) ..... "; : fl = -2
     GOSUB 62000
     Class = VAL(in$)
1013 LOCATE 6, 1: COLOR 11
     PRINT "Enter Minimum Acreage to Pull (0=All : <ESC>-Cancel . "; : fl = -8
     GOSUB 62000
     Acre = VAL(in$)
1014 LOCATE 7, 1: COLOR 11
     PRINT "Minimum Property Valuation (0=All : <ESC>-Cancel .... "; : fl = -12
     GOSUB 62000
     PValue# = VAL(in$)
1016 LOCATE 8, 1: COLOR 11
     PRINT "Print (D)etail Report or (C)ondensed Format (D,C) ... "; : fl = 1
     GOSUB 62210
     IF NOT (in$ = "D" OR in$ = "C") THEN 1016
     Layout$ = in$
     LOCATE 15, 1: COLOR 11: PRINT "Now Printing the Report : Please Wait!"
1100 GOSUB 60000
1150 IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10
     OPEN "O", 10, "TBCUSTAL.REP": WIDTH #10, 255
     GOSUB 1900
1160 IF TaxType$ = "R" THEN
        OPEN TBpath$ + "TBREname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS MREC$: EndofFile = LOF(4) / 24

1170 FOR x = 1 TO EndofFile
     GET 4, x
     GET 1, CVS(MREC$), TBRECust(1)
     Account! = CVS(MREC$)
    
1191 IF TBRECust(1).Deleted = "Y" THEN 1230

   ' Do the tests here
     TValue# = TBRECust(1).LandValue + TBRECust(1).BldgValue

     IF TValue# < PValue# THEN 1230
     IF District > 0 AND District <> VAL(TBRECust(1).District) THEN 1230
     IF Class > 0 AND Class <> VAL(TBRECust(1).ClassCode) THEN 1230
     IF VAL(TBRECust(1).Acres) < Acre THEN 1230
     IF LEN(QPTrim$(TBRECust(1).MapNumber)) = 0 THEN 1230 ' do not allow no map numbers to pass thru
     IF TValue# = 0 THEN 1230

   IF Layout$ = "C" THEN
   PRINT #10, USING "######"; Account!;
     PRINT #10, TAB(10); TBRECust(1).CustName; TAB(45); TBRECust(1).PropDesc1; TAB(80); TBRECust(1).MapNumber;
     PRINT #10, TAB(95); USING "$$#######,#"; TBRECust(1).LandValue;
     PRINT #10, TAB(107); USING "$$#######,#"; TBRECust(1).BldgValue;
     PRINT #10, TAB(120); USING "$$######,#"; TBRECust(1).LandValue + TBRECust(1).BldgValue
    PRINT #10, TAB(10); RTRIM$(TBRECust(1).Address1); "   "; RTRIM$(TBRECust(1).Address2); "   "; RTRIM$(TBRECust(1).Address3);
     PRINT #10, TAB(105); " Size:"; USING "######.##"; VAL(TBRECust(1).Acres);
     PRINT #10, TAB(120); "  Dist# "; TBRECust(1).District
    lc = lc + 2
    'PRINT #10, STRING$(130, "-")
     CustCount = CustCount + 1
     TotalLand# = TotalLand# + TBRECust(1).LandValue
     TotalLand# = INT((TotalLand# * 100) + .5) / 100
     TotalBldg# = TotalBldg# + TBRECust(1).BldgValue
     TotalBldg# = INT((TotalBldg# * 100) + .5) / 100


     ELSE
   PRINT #10, USING "######"; Account!;
     PRINT #10, TAB(10); TBRECust(1).CustName; TAB(45); TBRECust(1).PropDesc1; TAB(80); TBRECust(1).MapNumber;
     PRINT #10, TAB(95); USING "$$#######,#"; TBRECust(1).LandValue;
     PRINT #10, TAB(107); USING "$$#######,#"; TBRECust(1).BldgValue;
     PRINT #10, TAB(120); USING "$$######,#"; TBRECust(1).LandValue + TBRECust(1).BldgValue
    PRINT #10, TAB(10); TBRECust(1).Address1; TAB(45); TBRECust(1).PropDesc2
    PRINT #10, TAB(10); TBRECust(1).Address2; TAB(45); TBRECust(1).PropDesc3; TAB(80); "District #"; TBRECust(1).District
    PRINT #10, TAB(10); RTRIM$(TBRECust(1).Address3); TAB(45); TBRECust(1).PropDesc4
    lc = lc + 5
    PRINT #10, STRING$(130, "-")
     CustCount = CustCount + 1
     TotalLand# = TotalLand# + TBRECust(1).LandValue
     TotalLand# = INT((TotalLand# * 100) + .5) / 100
     TotalBldg# = TotalBldg# + TBRECust(1).BldgValue
     TotalBldg# = INT((TotalBldg# * 100) + .5) / 100
     END IF
      IF lc >= 58 THEN PRINT #10, CHR$(12); : GOSUB 1900
1230 NEXT x
     PRINT #10, "Total Customers .. "; USING "######,#"; CustCount
     PRINT #10, "Total Land Value: "; USING "$$#########,#"; TotalLand#
     PRINT #10, "Total Bldg Value: "; USING "$$#########,#"; TotalBldg#
     PRINT #10, " Total Valuation: "; USING "$$#########,#"; TotalLand# + TotalBldg#
     PRINT CHR$(12);
1240 CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBCUSTAL.REP"
      Header$ = "Land Book Preliminary Report"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
1250 GOTO 10

        ELSE

     'Personal Property Here

        OPEN TBpath$ + "TBPPname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS MREC$: EndofFile = LOF(4) / 24
2170 FOR x = 1 TO EndofFile
     GET 4, x
     GET 1, CVS(MREC$), TBPPCust(1)
     Account! = CVS(MREC$)

2191 IF TBRECust(1).Deleted = "Y" THEN 2230

    PRINT #10, USING "######"; Account!;
    PRINT #10, TAB(10); TBPPCust(1).CustName

    lc = lc + 5
    PRINT #10, STRING$(130, "-")
    CustCount = CustCount + 1

2220 IF lc >= 58 THEN PRINT #10, CHR$(12); : GOSUB 1900
2230 NEXT x
     PRINT #10, "Total Customers .. "; USING "######,#"; CustCount
     PRINT #10, " Total Valuation: "; USING "$$#########,#"; TotalLand# + TotalBldg#

     PRINT CHR$(12);
2240 CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBCUSTAL.REP"
      Header$ = "Personal Property - 'Alpha' Report"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
2250 GOTO 10
      END IF


1900 p = p + 1
     PRINT #10, TAB(27); "Alpha Customer Account Listing"
     PRINT #10, "Date: "; DATE$; TAB(68); "Page #"; USING "#####"; p
     PRINT #10, "District: "; :
       IF District = 0 THEN PRINT #10, "All";  ELSE PRINT #10, District;
     PRINT #10, TAB(30); "Class # ";
       IF Class = 0 THEN PRINT #10, "All";  ELSE PRINT #10, Class;
     PRINT #10, TAB(60); "Acreage: ";
       IF Acre = 0 THEN PRINT #10, "All";  ELSE PRINT #10, USING "#######.###"; Acre
     PRINT #10, "Minimum Property Value: ";
       IF PValue# = 0 THEN PRINT #10, "NONE" ELSE PRINT #10, USING "########,#"; PValue#

       

1910 PRINT #10, "Acct. #"; TAB(10); "Name & Address"; TAB(45); "Legal Description"; TAB(80); "Map Reference"; TAB(102); "Land"; TAB(113); "Bldgs"; TAB(120); "Total Value"
     PRINT #10, STRING$(130, "=")
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

Namesearch:
 LOCATE 8, 1:
 COLOR 11
 PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
 GOSUB 62000
 IF cf = 1 THEN
  ggo = 0
  FOR lll = 8 TO 13: PRINT STRING$(79, 32): NEXT lll
  RETURN
 END IF
 nn$ = in$
 nn = LEN(nn$)
 IF nn > 20 THEN nn = 20
 nn$ = LEFT$(nn$, nn)

 IF TaxType$ = "R" THEN
  TBIDXFile = FREEFILE
  OPEN TBpath$ + "TBRENAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS MREC$
  GOSUB 60000
  FOR LL = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, LL
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO n1
n3:
  NEXT LL
  FOR lll = 8 TO 13: PRINT STRING$(79, 32): NEXT lll
  ggo = 0
  RETURN

n1:
GET TBFILE, CVS(MREC$), TBRECust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBRECust(1).CustName
COLOR 10
FOR lll = 10 TO 12: LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(MREC$)
n2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO n2
IF in$ = "Q" THEN
  FOR lll = 8 TO 13: LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
  CLOSE
  ggo = 0
  RETURN
END IF
IF in$ = "Y" THEN
  ggo = 1: Record! = CVS(MREC$)
  FOR lll = 8 TO 13:  LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO n3
GOTO n2
END IF


  TBIDXFile = FREEFILE
  OPEN TBpath$ + "TBPPNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS MREC$
  GOSUB 60000
  FOR LL = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, LL
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO PPn1
PPn3:
  NEXT LL
  LOCATE 8, 1: PRINT STRING$(79, 32)
  ggo = 0
  RETURN

PPn1:
GET TBFILE, CVS(MREC$), TBPPCust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBPPCust(1).CustName
COLOR 10
FOR lll = 10 TO 12: LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(MREC$)
PPn2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO PPn2
IF in$ = "Q" THEN
  FOR lll = 8 TO 13: LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
  CLOSE
  ggo = 0
  RETURN
END IF
IF in$ = "Y" THEN
  ggo = 1: Record! = CVS(MREC$)
  FOR lll = 8 TO 13: LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO PPn3
GOTO PPn2

GETDISTRICT:
     DistrictName$ = ""
     IF District = 1 THEN DistrictName$ = "BROWN'S STORE"
     IF District = 2 THEN DistrictName$ = "COLUMBIAN GROVE"
     IF District = 3 THEN DistrictName$ = "LEWISTON"
     IF District = 4 THEN DistrictName$ = "LOCHLEVEN"
     IF District = 5 THEN DistrictName$ = "PLEASANT GROVE"
     IF District = 6 THEN DistrictName$ = "PLYMOUTH"
     IF District = 7 THEN DistrictName$ = "REHOBOTH"
     IF District = 9 THEN DistrictName$ = "TOWN OF KENBRIDGE"
     IF District = 10 THEN DistrictName$ = "TOWN OF VICTORIA"
     RETURN


FigureBalance:
     CurBal# = 0: PastDue# = 0
     IF TaxType$ = "R" THEN
     CurBal# = TBRECust(1).CurYrTaxAmount + TBRECust(1).CurYrIntAmount + TBRECust(1).CurYrPenAmount + TBRECust(1).CurYrOthAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR LL = 1 TO 10
     PastDue# = PastDue# + TBRECustBal(1).TaxAmt(LL) + TBRECustBal(1).IntAmt(LL) + TBRECustBal(1).PenAmt(LL) + TBRECustBal(1).OthAmt(LL) + TBRECustBal(1).LateAmt(LL)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT LL
     ELSE
     CurBal# = TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax + TBPPCust(1).CurYrIntAmount + TBPPCust(1). _
CurYrPenAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR LL = 1 TO 10
     PastDue# = PastDue# + TBPPCustBal(1).PerPropTaxAmt(LL) + TBPPCustBal(1).FarmEquipAmt(LL) + TBPPCustBal(1).MachToolsAmt(LL) + TBPPCustBal(1).MerchCapAmt(LL) + TBPPCustBal(1).MobileHomeAmt(LL) + TBPPCustBal(1).LateFeeAmt(LL) + TBPPCustBal(1). _
IntAmt(LL) + TBPPCustBal(1).PenAmt(LL)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT LL


     END IF
     RETURN

Deleted:
    LOCATE 23, 1: COLOR 15
    PRINT "This Account Has Been Deleted! Press <ENTER> to Continue"; :
    fl = 1
    GOSUB 62000
    LOCATE 23, 1: PRINT STRING$(79, 32)
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

