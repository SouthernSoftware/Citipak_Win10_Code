10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbcustrep
30 REM Customer Report Menu
40 REM Last Revision Date: 8-24-87
50 REM by Nick Picerno

   ' Must be compiled using the compile directory
   DEFINT A-Z
   DECLARE FUNCTION QPTrim$ (A$)
   DECLARE FUNCTION Num2Date$ (Num%)
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
 ' TBpath$ = "c:\qb45\lc\"                    ' change to "F:\" Before compiling
   CLOSE
   OPEN "R", 1, TBpath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
   CurrentYear$ = year$: CLOSE 1
   ON ERROR GOTO networkerror


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
     
100 REM begin program
110 SCREN$ = "2.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500

1000 REM
1010 SCREN$ = "2.01.00": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 32: PRINT "Print Reassessment Forms"
1020 LOCATE 3, 1: COLOR 11
        PRINT "Print (I)ndividual or (A)ll Forms (I,A or <ESC>-Menu) .. "; : fl = 1
        GOSUB 62210
        IF CF = 1 THEN RUN "tbcustre"
        IF NOT (in$ = "A" OR in$ = "I") THEN 1020
        IF in$ = "I" THEN 3000

1100 LOCATE 3, 1: COLOR 11
     PRINT "(N)umerical Order or (A)lphabetical Order (N,A or <ESC>-Cancel) .. "; : fl = 1
     GOSUB 62210
     IF CF = 1 THEN RUN "tbcustre"
     IF NOT (in$ = "N" OR in$ = "A") THEN 1100
     SortType$ = in$


1111 LOCATE 4, 1
     COLOR 11
     IF SortType$ = "A" THEN
      PRINT "Enter 1st 3 Characters to Begin Printing (ENTER=ALL)"
      PRINT "Example:  To Begin with CAWLET  enter CAW ........... "; : fl = 3
      GOSUB 62000
      IF LEN(in$) = 0 THEN
       SortStart$ = "ALL"
       ELSE
       SortStart$ = RTRIM$(UCASE$(in$))
      END IF

     END IF

1110 LOCATE 7, 1: COLOR 15
     PRINT "Place your printer in 15 cpi mode now!!!"
     COLOR 11
     PRINT "Print an Alignment Mask (Y/N/Q) .. "; : fl = 1
     GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN 1110
     IF in$ = "Q" THEN RUN "tbcustre"
     IF in$ = "Y" THEN
      GOSUB AlignMent: GOTO 1110
     END IF
     taxtype$ = "R"
     GOSUB 60000
     LOCATE 12, 1: COLOR 11: PRINT "NOW SPOOLING OUT FORMS : PLEASE WAIT!!!"
1150 IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10

     OPEN "O", 10, "TBASSFRM.REP": WIDTH #10, 255
     
     IF SortType$ = "A" THEN
        OPEN TBpath$ + "TBREname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS LastNam$, 4 AS Mrec$: EndofFile = LOF(4) / 24
      ELSE
       OPEN TBpath$ + "TBMAP.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
       FIELD 4, 20 AS LastNam$, 4 AS Mrec$
       EndofFile = LOF(4) / 24
    END IF
     

1170 FOR x = 1 TO EndofFile
     GET 4, x
     Account! = CVS(Mrec$)
     GET TBFile, CVS(Mrec$), TBRECust(1)
     IF SortType$ = "A" THEN
      IF SortStart$ = "ALL" THEN 1191
      IF LEFT$(LastNam$, LEN(SortStart$)) >= SortStart$ THEN 1191
      GOTO 1230 'No Good
     END IF

1191 IF TBRECust(1).Deleted = "Y" THEN 1230
     GOSUB GetDistrict
     PRINT #10, ""
     PRINT #10, TAB(45); "LUNENBURG COUNTY 1998 PROPERTY REASSESSMENT"
     PRINT #10, TAB(45); "       COMMONWEALTH OF VIRGINIA"
     PRINT #10, TAB(15); Account!; TAB(100); TBRECust(1).MapNumber
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, TAB(8); TBRECust(1).PropDesc1
     PRINT #10, TAB(8); TBRECust(1).PropDesc2;
      PRINT #10, TAB(53); USING "$$#######,#"; TBRECust(1).RELandValue;
      PRINT #10, TAB(68); USING "$$#######,#"; TBRECust(1).REBldgValue;
      PRINT #10, TAB(83); USING "$$#######,#"; TBRECust(1).RELandValue + TBRECust(1).REBldgValue
     PRINT #10, TAB(8); TBRECust(1).PropDesc3
     PRINT #10, TAB(8); TBRECust(1).PropDesc4
     PRINT #10, TAB(8); DistrictName$
     PRINT #10,
     PRINT #10,
     PRINT #10, TAB(20); LEFT$(TBRECust(1).CustName, 31); TAB(53); "|Applicants to appear before the Board"
     PRINT #10, TAB(20); LEFT$(TBRECust(1).Address1, 31); TAB(53); "|of Equalization must file from 2/2/98 to 2/27/98."
     PRINT #10, TAB(20); LEFT$(TBRECust(1).Address2, 31); TAB(53); "|The BOE will complete its work by 4/1/98.  For"
     PRINT #10, TAB(20); LEFT$(TBRECust(1).Address3, 31); TAB(53); "|further details refer to your local newspaper."
     PRINT #10,
     PRINT #10, TAB(10); "Hearings will be held at the Lunenburg Human Resources Building, Board Room by appointments only."
     PRINT #10, TAB(10); "If you wish to make an appointment, please call by December 12, 1997,  (804)696-2546."
     PRINT #10, TAB(10); "Monday - Friday 9:00-4:00."
     PRINT #10, ""
1230 NEXT
1240 CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBASSFRM.REP"
      Header$ = "Reassessment Forms"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
1250 RUN "TBCustRe"




AlignMent:
     LPRINT ""
     LPRINT TAB(45); "LUNENBURG COUNTY 1998 PROPERTY REASSESSMENT"
     LPRINT TAB(45); "       COMMONWEALTH OF VIRGINIA"
     LPRINT TAB(15); "ACCT # HERE"; TAB(100); "MAP NUMBER HERE"
     LPRINT
     LPRINT
     LPRINT TAB(8); "Property Description Line 1"
     LPRINT TAB(8); "Property Description Line 2";
      LPRINT TAB(53); "XX,XXX,XXX"; TAB(68); "XX,XXX,XXX"; TAB(83); "XX,XXX,XXX"
     LPRINT TAB(8); "Property Description Line 3"
     LPRINT TAB(8); "Property Description Line 4"
     LPRINT TAB(8); "DISTRICT NAME HERE"
     LPRINT
     LPRINT
     LPRINT TAB(20); "Customer Name"; TAB(53); "|Applicants to appear before the Board"
     LPRINT TAB(20); "Address Line 1"; TAB(53); "|of Equalization must file from "; Date1$; " to "; Date2$; "."
     LPRINT TAB(20); "Address Line 2"; TAB(53); "|The BOE will complete its work by "; Date3$; ".  For"
     LPRINT TAB(20); "Address Line 3"; TAB(53); "|futher details refer to your local newspaper."
     LPRINT TAB(20); STRING$(85, "*")
     LPRINT TAB(20); "Hearings will be held at the Lunenburg County Courthouse by appointments only."
     LPRINT TAB(20); "If you wish to make an appointment please call by December 12, 1997 at"
     LPRINT TAB(20); "(804)696-1404 or (804)696-1801. Mon-Fri. 9:00-4:00."
     LPRINT
     RETURN

3000 REM
3010 SCREN$ = "2.03.00": CLS : GOSUB 63500
3020 COLOR 11: LOCATE 1, 25: PRINT "Print Individual Assessment Notice";
3025 taxtype$ = "R"
     LOCATE 3, 1: COLOR 15
     PRINT "Place your printer in 15 cpi mode now!!!"

3030 LOCATE 5, 1: COLOR 11
     PRINT "Enter Customer Account Number (TAB-Last Name : ESC-Menu) .. "; : fl = -6
     GOSUB 62000
     IF CF = 1 THEN
      CLOSE : RUN "tbcustre"
     END IF
     IF CF = 2 THEN
      GOSUB Namesearch
      IF ggo = 1 THEN 3075
      IF ggo = 0 THEN LOCATE 10, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 62000: GOTO 3000
     END IF
     Record! = VAL(in$)
3075 CLOSE : GOSUB 60000: GOSUB 60025
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBRECust(1)) THEN 3030
       LOCK #TBFile, Record!
       GET TBFile, Record!, TBRECust(1)
      
       IF TBRECust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 3000

     

    
     GOSUB GetDistrict
     LPRINT ""
     LPRINT TAB(45); "LUNENBURG COUNTY 1998 PROPERTY REASSESSMENT"
     LPRINT TAB(45); "       COMMONWEALTH OF VIRGINIA"
     LPRINT TAB(15); Account!; TAB(100); TBRECust(1).MapNumber
     LPRINT ""
     LPRINT ""
     LPRINT TAB(8); TBRECust(1).PropDesc1
     LPRINT TAB(8); TBRECust(1).PropDesc2;
      LPRINT TAB(53); USING "$$#######,#"; TBRECust(1).RELandValue;
      LPRINT TAB(68); USING "$$#######,#"; TBRECust(1).REBldgValue;
      LPRINT TAB(83); USING "$$#######,#"; TBRECust(1).RELandValue + TBRECust(1).REBldgValue
     LPRINT TAB(8); TBRECust(1).PropDesc3
     LPRINT TAB(8); TBRECust(1).PropDesc4
     LPRINT TAB(8); DistrictName$
     LPRINT ""
     LPRINT ""
     LPRINT TAB(20); LEFT$(TBRECust(1).CustName, 31); TAB(53); "|Applicants to appear before the Board"
     LPRINT TAB(20); LEFT$(TBRECust(1).Address1, 31); TAB(53); "|of Equalization must file from 2/2/98 to 2/27/98."
     LPRINT TAB(20); LEFT$(TBRECust(1).Address2, 31); TAB(53); "|The BOE will complete its work by 4/1/98.  For"
     LPRINT TAB(20); LEFT$(TBRECust(1).Address3, 31); TAB(53); "|further details refer to your local newspaper."
     LPRINT ""
     LPRINT TAB(10); "Hearings will be held at the Lunenburg Human Resources Building, Board Room by appointments only."
     LPRINT TAB(10); "If you wish to make an appointment, please call by December 12, 1997,  (804)696-2546."
     LPRINT TAB(10); "Monday - Friday 9:00-4:00."
     LPRINT ""
     CLOSE
     GOTO 3000


60000  IF taxtype$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBpath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBpath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN

60025 TBBalFile = FREEFILE
      IF taxtype$ = "R" THEN
      OPEN TBpath$ + "TBREBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBpath$ + "TBPPBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF
60030 RETURN

61000 TBTransFile = FREEFILE
      OPEN TBpath$ + "TBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTransFile LEN = LEN(TBTrans(1))
61020 RETURN

62000 COLOR 14: CF = 0: in$ = "": RO = CSRLIN: PO = POS(Xy): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN CF = 1: RETURN
      IF z$ = CHR$(9) THEN CF = 2: RETURN
      IF LEN(z$) < 2 THEN 62065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN CF = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN CF = 12: RETURN

62065 IF ASC(z$) = 0 THEN 62030
62070 IF LEN(z$) > 1 THEN CF = ASC(RIGHT$(z$, 1)) - 57: IF CF > 0 THEN RETURN ELSE CF = CF + 27: RETURN
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
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN

Namesearch:
 LOCATE 8, 1:
 COLOR 11
 PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
 GOSUB 62000
 IF CF = 1 THEN
  ggo = 0
  FOR lll = 8 TO 13: PRINT STRING$(79, 32): NEXT lll
  RETURN
 END IF
 nn$ = in$
 nn = LEN(nn$)
 IF nn > 20 THEN nn = 20
 nn$ = LEFT$(nn$, nn)

 IF taxtype$ = "R" THEN
  TBIDXFile = FREEFILE
  OPEN TBpath$ + "TBRENAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS LastNam$, 4 AS Mrec$
  GOSUB 60000
  FOR LL = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, LL
  IF LEFT$(LastNam$, nn) = nn$ THEN GOTO n1
n3:
  NEXT LL
  FOR lll = 8 TO 13: PRINT STRING$(79, 32): NEXT lll
  ggo = 0
  RETURN

n1:
GET TBFile, CVS(Mrec$), TBRECust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBRECust(1).CustName
COLOR 10
FOR lll = 10 TO 12: LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(Mrec$)
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
  ggo = 1: Record! = CVS(Mrec$)
  FOR lll = 8 TO 13:  LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO n3
GOTO n2
END IF


  TBIDXFile = FREEFILE
  OPEN TBpath$ + "TBPPNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS LastNam$, 4 AS Mrec$
  GOSUB 60000
  FOR LL = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, LL
  IF LEFT$(LastNam$, nn) = nn$ THEN GOTO PPn1
PPn3:
  NEXT LL
  LOCATE 8, 1: PRINT STRING$(79, 32)
  ggo = 0
  RETURN

PPn1:
GET TBFile, CVS(Mrec$), TBPPCust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBPPCust(1).CustName
COLOR 10
FOR lll = 10 TO 12: LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(Mrec$)
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
  ggo = 1: Record! = CVS(Mrec$)
  FOR lll = 8 TO 13: LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO PPn3
GOTO PPn2

GetDistrict:
     DistrictName$ = ""
     District = VAL(TBRECust(1).District)
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
     IF taxtype$ = "R" THEN
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
               fl = 1: GOSUB 62000: IF CF = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
            
     END SELECT

