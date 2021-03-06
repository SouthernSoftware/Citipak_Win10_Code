10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbcustrep
30 REM Customer Report Menu
40 REM Last Revision Date: 8-24-87
50 REM by Nick Picerno
 
   DEFINT A-Z
   DECLARE FUNCTION QPTrim$ (A$)
   DECLARE FUNCTION num2date$ (Num%)
   DECLARE FUNCTION Date2Num% (Dat$)
   DECLARE FUNCTION WaitKey% ()
   DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
   DECLARE SUB PrintTitle (Title$)
   DECLARE SUB PrintLoading ()
   DECLARE SUB ShowCursor ()
   DECLARE SUB HideCursor ()
   
'$INCLUDE: 'lcTAX.BI'
   
60 CLS :
   TBpath$ = "F:\"                    ' change to "F:\" Before compiling
 '  TBpath$ = "c:\qb45\lc\"                    ' change to "F:\" Before compiling
   CLOSE
   p = 0: page = 0
   OPEN "R", 1, TBpath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
   CurrentYear$ = year$
   CLOSE 1
   'ON ERROR GOTO networkerror


     REDIM TransLink!(500)
     REDIM TBRECust(1) AS TBRECust
     REDIM TBRECustBal(1) AS TBREBAL
     REDIM TBPPCust(1) AS TBPPCust
     REDIM TBPPCustBal(1) AS TBPPBal
     REDIM TBTrans(1) AS TBTrans



3000 TaxType$ = "P"
     Record! = 5254
3075 CLOSE : GOSUB 60000: GOSUB 60025
     IF TaxType$ = "R" THEN
       LOCK #TBFile, Record!
       GET TBFile, Record!, TBRECust(1)
       GET TBBalFile, Record!, TBRECustBal(1)
       IF TBRECust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 3000
       ELSE
       
       LOCK #TBFile, Record!
       GET TBFile, Record!, TBPPCust(1)
       GET TBBalFile, Record!, TBPPCustBal(1)
       IF TBPPCust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 3000
     END IF

     IF TaxType$ = "R" THEN
       GOSUB DisplayReal
        ELSE
       GOSUB DisplayPersonal
     END IF
     CLOSE
     GOTO 3000
     
DisplayReal:
RETURN
RealBalance:
RETURN

DisplayPersonal:
FOR ll = 3 TO 23: LOCATE ll, 1: PRINT STRING$(79, 32): NEXT ll
COLOR 11
LOCATE 3, 1: PRINT "Personal Property Customer"
PRINT "Fixing the Account of ..."
COLOR 15
PRINT TBPPCust(1).CustName; : COLOR 11: PRINT TAB(60); "Acct # "; : COLOR 15: PRINT Record!
PRINT TBPPCust(1).Address1
PRINT TBPPCust(1).Address2
PRINT RTRIM$(TBPPCust(1).City); ", "; TBPPCust(1).State; "  "; TBPPCust(1).Zip
FOR Pyr = 2 TO 2
 TBPPCustBal(1).PerPropTaxAmt(Pyr) = 0
  TBPPCustBal(1).FarmEquipAmt(Pyr) = 0
   TBPPCustBal(1).MachToolsAmt(Pyr) = 0
   TBPPCustBal(1).MerchCapAmt(Pyr) = 0
   TBPPCustBal(1).MobileHomeAmt(Pyr) = 0
   TBPPCustBal(1).LateFeeAmt(Pyr) = 0
   TBPPCustBal(1).IntAmt(Pyr) = 0
   TBPPCustBal(1).PenAmt(Pyr) = 0
  NEXT Pyr
  PUT TBBalFile, Record!, TBPPCustBal(1)
  GOTO PersBalance

CLOSE TBBalFile

PRINT
CLOSE
COLOR 11
PRINT "1995 Balance for Property Taxes Has Been Zero'd"
PRINT CHR$(7);
SLEEP 6
END
RETURN
PersBalance:
RETURN

     
RealHistory:
  RETURN

RealHeading:
RETURN

RealTransLink:
 Counter = 0
Realtrlink1:
 IF TR! = 0 THEN RETURN
 Counter = Counter + 1
 TransLink!(Counter) = TR!
 GET TBTransFile, TR!, TBTrans(1)
 TR! = TBTrans(1).TrNextTrans
 GOTO Realtrlink1

PersHistory:
  RETURN

PersHeading:
RETURN

PersTransLink:
 Counter = 0
Perstrlink1:
 IF TR! = 0 THEN RETURN
 Counter = Counter + 1
 TransLink!(Counter) = TR!
 GET TBTransFile, TR!, TBTrans(1)
 TR! = TBTrans(1).TrNextTrans
 GOTO Perstrlink1

60000  IF TaxType$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBpath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBpath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
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
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN

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
 nn$ = in$
 nn = LEN(nn$)
 IF nn > 20 THEN nn = 20
 nn$ = LEFT$(nn$, nn)

 IF TaxType$ = "R" THEN
  TBIDXFile = FREEFILE
  OPEN TBpath$ + "TBRENAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS mrec$
  GOSUB 60000
  FOR ll = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, ll
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO n1
n3:
  NEXT ll
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  ggo = 0
  RETURN

n1:
GET TBFile, CVS(mrec$), TBRECust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBRECust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(mrec$)
n2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO n2
IF in$ = "Q" THEN
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  ggo = 0
  RETURN
END IF
IF in$ = "Y" THEN
  ggo = 1: Record! = CVS(mrec$)
  FOR LLL = 8 TO 13:  LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO n3
GOTO n2
END IF


  TBIDXFile = FREEFILE
  OPEN TBpath$ + "TBPPNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS mrec$
  GOSUB 60000
  FOR ll = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, ll
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO PPn1
PPn3:
  NEXT ll
  LOCATE 8, 1: PRINT STRING$(79, 32)
  ggo = 0
  RETURN

PPn1:
GET TBFile, CVS(mrec$), TBPPCust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBPPCust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(mrec$)
PPn2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO PPn2
IF in$ = "Q" THEN
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  ggo = 0
  RETURN
END IF
IF in$ = "Y" THEN
  ggo = 1: Record! = CVS(mrec$)
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO PPn3
GOTO PPn2

FigureBalance:
     CurBal# = 0: PastDue# = 0
     IF TaxType$ = "R" THEN
     CurBal# = TBRECust(1).CurYrTaxAmount + TBRECust(1).CurYrIntAmount + TBRECust(1).CurYrPenAmount + TBRECust(1).CurYrOthAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR ll = 1 TO 10
     TBRECustBal(1).OthAmt(ll) = 0
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

