'$INCLUDE: 'LCTax.Bi'
  DECLARE FUNCTION Num2Date$ (Num%)
  DECLARE FUNCTION Date2Num% (Dat$)
10 CLS
   CLEAR
   ON ERROR GOTO networkerror
           
  TBPath$ = ""                    ' change to "F:\" Before compiling
'   TBPath$ = "C:\qb45\lc\"

   REDIM TBRECUST(1) AS TBRECUST
   REDIM TBRECustBal(1) AS TBREBAL
   REDIM TBPPCUST(1) AS TBPPCUST
   REDIM TBPPCustBal(1) AS TBPPBal
   REDIM TBTrans(1) AS TBTrans

   CLOSE : OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
    CURRENTYEAR$ = year$: CLOSE
   
     
    taxtype$ = "R"
    GOSUB 60000
    FOR Record! = 1 TO LOF(TBFILE) / LEN(TBRECUST(1))
     GET TBFILE, Record!, TBRECUST(1)
        IF TBRECUST(1).FIRSTTRANS < 0 THEN
         TBRECUST(1).FIRSTTRANS = 0
         TBRECUST(1).LASTTRANS = 0
        END IF
        IF TBRECUST(1).LASTTRANS < 0 THEN
         TBRECUST(1).FIRSTTRANS = 0
         TBRECUST(1).LASTTRANS = 0
        END IF


     PUT TBFILE, Record!, TBRECUST(1)
     NEXT Record!
     STOP
       
60000  IF taxtype$ = "R" THEN
         TBFILE = FREEFILE
         OPEN TBPath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFILE LEN = LEN(TBRECUST(1))
        ELSEIF taxtype$ = "P" THEN
         TBFILE = FREEFILE
         OPEN TBPath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFILE LEN = LEN(TBPPCUST(1))
        ELSEIF taxtype$ = "S" THEN
         TBFILE = FREEFILE
         OPEN TBPath$ + "TBSPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFILE LEN = LEN(TBPPCUST(1))
        END IF
60025 TBBALFILE = FREEFILE
      IF taxtype$ = "R" THEN
      OPEN TBPath$ + "TBREBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBALFILE LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBPath$ + "TBPPBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBALFILE LEN = LEN(TBPPCustBal(1))
      END IF
60030 RETURN
61000 TBTransFile = FREEFILE
      OPEN TBPath$ + "TBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTransFile LEN = LEN(TBTrans(1))
61020 RETURN
        

62000 cf = 0: COLOR 14: in$ = "": RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(z$) < 2 THEN 62065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
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
 nn$ = UCASE$(in$)
 nn = LEN(nn$)
 IF nn > 20 THEN nn = 20
 nn$ = LEFT$(nn$, nn)

 IF taxtype$ = "R" THEN
  TBIDXFile = FREEFILE
  OPEN TBPath$ + "TBRENAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS Mrec$
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
GET TBFILE, CVS(Mrec$), TBRECUST(1)
LOCATE 11, 1: COLOR 11
nme$ = TBRECUST(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(Mrec$)
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
  ggo = 1: Record! = CVS(Mrec$)
  FOR LLL = 8 TO 13:  LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO n3
GOTO n2
END IF

 
  TBIDXFile = FREEFILE
  OPEN TBPath$ + "TBPPNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS Mrec$
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
GET TBFILE, CVS(Mrec$), TBPPCUST(1)
LOCATE 11, 1: COLOR 11
nme$ = TBPPCUST(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(Mrec$)
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
  ggo = 1: Record! = CVS(Mrec$)
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO PPn3
GOTO PPn2



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
     CurBal# = TBRECUST(1).CurYrTaxAmount + TBRECUST(1).CurYrIntAmount + TBRECUST(1).CurYrPenAmount + TBRECUST(1).CurYrOthAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR ll = 1 TO 10
     PastDue# = PastDue# + TBRECustBal(1).TaxAmt(ll) + TBRECustBal(1).IntAmt(ll) + TBRECustBal(1).PenAmt(ll) + TBRECustBal(1).OthAmt(ll) + TBRECustBal(1).LateAmt(ll)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT ll
     ELSE
     CurBal# = TBPPCUST(1).CurYrPersPropTax + TBPPCUST(1).CurYrFarmEquipTax + TBPPCUST(1).CurYrMachToolsTax + TBPPCUST(1).CurYrMerchCapTax + TBPPCUST(1).CurYrMobileHomeTax + TBPPCUST(1).CurYrLateFeeTax + TBPPCUST(1).CurYrIntAmount + TBPPCUST(1). _
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

