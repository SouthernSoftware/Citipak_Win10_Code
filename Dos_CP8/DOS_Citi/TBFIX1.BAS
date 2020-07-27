10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbbilling
30 REM Tax Billing Menu
40 REM Last Revision Date:11-29-88
50 REM by Nick Picerno
60 CLS : CLEAR
   '$INCLUDE: 'lcTAX.BI'
   TBPath$ = "F:\"                    ' change to "F:\" Before compiling

    CLOSE : OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
    CurrentYear$ = year$: CLOSE
    CurrentYear = VAL(CurrentYear$)
    ON ERROR GOTO networkerror

    DIM SHARED yrdesc(11)

    REDIM TBRECust(1) AS TBRECust
    REDIM tbrecustbal(1) AS TBREBAL
    REDIM TBPPCust(1) AS TBPPCust
    REDIM tbppcustbal(1) AS TBPPBal
    REDIM TBTrans(1) AS TBTrans

    WIDTH LPRINT 132


4000 CLS : scren$ = "3.04.00"
     GOSUB 63500: COLOR 11
     LOCATE 1, 30: PRINT "Fix Account Procedure"
     totreal# = 0: totpers# = 0: tottax# = 0
     LOCATE 3, 1: COLOR 11
     PRINT "Start the Procedure (Y/N or <ESC>-End) ..... "; : fl = 1: GOSUB 62210
     IF CF = 1 THEN 10
     IF NOT (in$ = "Y" OR in$ = "N") THEN 4000
     IF in$ = "N" THEN CLS : END
     FOR MainLoop = 1 TO 2
     IF MainLoop = 1 THEN TaxType$ = "R"
     IF MainLoop = 2 THEN TaxType$ = "P"
     CLOSE
     GOSUB 60000
     GOSUB 60025
     IF TaxType$ = "R" THEN
        EndofFile = LOF(TBFile) / LEN(TBRECust(1))
       ELSE
        EndofFile = LOF(TBFile) / LEN(TBPPCust(1))
     END IF

     FOR x = 1 TO EndofFile
      IF TaxType$ = "R" THEN
       GET 1, x, TBRECust(1)
       GET tbbalfile, x, tbrecustbal(1)
      ELSE
       GET 1, x, TBPPCust(1)
       GET tbbalfile, x, tbppcustbal(1)
      END IF
      AcctRecord = x

   LOCATE 15, 1: COLOR 11: PRINT "Process Record # "; x
   IF TaxType$ = "R" THEN
      IF TBRECust(1).Deleted = "Y" THEN 4230
     IF TBRECust(1).LandValue = 0 AND TBRECust(1).BldgValue = 0 AND VAL(TBRECust(1).CurYrBillNumber) > 0 THEN
      GOSUB FigureBalance
      TBalance# = CurBal# + PastDue#
      TBalance# = INT((TBalance# * 100) + .5) / 100
      IF TBalance# > 0 THEN GOSUB PostRealBill
     END IF
      ELSE
       IF TBPPCust(1).Deleted = "Y" THEN 4230
      IF TBPPCust(1).FirstVeh = 0 AND TBPPCust(1).MachTools = 0 AND TBPPCust(1).MobileHome = 0 AND TBPPCust(1).MerchCap = 0 AND TBPPCust(1).FarmEquip = 0 AND VAL(TBPPCust(1).CurYrBillNumber) > 0 THEN
      GOSUB FigureBalance
      TBalance# = CurBal# + PastDue#
      TBalance# = INT((TBalance# * 100) + .5) / 100
      IF TBalance# > 0 THEN GOSUB PostPersBill
       END IF
   END IF

4230 NEXT x
     NEXT MainLoop
     CLOSE
     LOCATE 20, 1
     PRINT "DONE!!": SLEEP 3
     CLS
     END

PostRealBill:


 GET tbbalfile, AcctRecord, tbrecustbal(1)
 tbrecustbal(1).TaxAmt(10) = tbrecustbal(1).TaxAmt(10) + tbrecustbal(1).TaxAmt(9)
 tbrecustbal(1).IntAmt(10) = tbrecustbal(1).IntAmt(10) + tbrecustbal(1).IntAmt(9)
 tbrecustbal(1).PenAmt(10) = tbrecustbal(1).PenAmt(10) + tbrecustbal(1).PenAmt(9)
 tbrecustbal(1).OthAmt(10) = tbrecustbal(1).OthAmt(10) + tbrecustbal(1).OthAmt(9)
 tbrecustbal(1).LateAmt(10) = tbrecustbal(1).LateAmt(10) + tbrecustbal(1).LateAmt(9)
 tbrecustbal(1).TaxNotice(10) = tbrecustbal(1).TaxNotice(9)

 FOR Cnt! = 9 TO 2 STEP -1
 tbrecustbal(1).TaxAmt(Cnt!) = tbrecustbal(1).TaxAmt(Cnt! - 1)
 tbrecustbal(1).IntAmt(Cnt!) = tbrecustbal(1).IntAmt(Cnt! - 1)
 tbrecustbal(1).PenAmt(Cnt!) = tbrecustbal(1).PenAmt(Cnt! - 1)
 tbrecustbal(1).OthAmt(Cnt!) = tbrecustbal(1).OthAmt(Cnt! - 1)
 tbrecustbal(1).LateAmt(Cnt!) = tbrecustbal(1).LateAmt(Cnt! - 1)
 tbrecustbal(1).TaxNotice(Cnt!) = tbrecustbal(1).TaxNotice(Cnt! - 1)
 NEXT Cnt!

 tbrecustbal(1).TaxAmt(1) = TBRECust(1).CurYrTaxAmount
 tbrecustbal(1).IntAmt(1) = TBRECust(1).CurYrIntAmount
 tbrecustbal(1).PenAmt(1) = TBRECust(1).CurYrPenAmount
 tbrecustbal(1).OthAmt(1) = TBRECust(1).CurYrOthAmount
 tbrecustbal(1).LateAmt(1) = 0
 tbrecustbal(1).TaxNotice(1) = TBRECust(1).CurYrBillNumber

 PUT tbbalfile, AcctRecord, tbrecustbal(1)

 'Now to Set Current Stuff
 TBRECust(1).CurYrTaxAmount = 0
 TBRECust(1).CurYrIntAmount = 0
 TBRECust(1).CurYrPenAmount = 0
 TBRECust(1).CurYrOthAmount = 0
 TBRECust(1).CurYrBillNumber = ""



 ' Now Set Billing Information to Null for Next Yr Printing
 TBRECust(1).LBCurYrTaxAmount = 0
 TBRECust(1).LBCurYrBillNumber = ""


 PUT TBFile, AcctRecord, TBRECust(1)
 RETURN


PostPersBill:

 
 GET tbbalfile, AcctRecord, tbppcustbal(1)
 tbppcustbal(1).PerPropTaxAmt(10) = tbppcustbal(1).PerPropTaxAmt(10) + tbppcustbal(1).PerPropTaxAmt(9)
 tbppcustbal(1).FarmEquipAmt(10) = tbppcustbal(1).FarmEquipAmt(10) + tbppcustbal(1).FarmEquipAmt(9)
 tbppcustbal(1).MachToolsAmt(10) = tbppcustbal(1).MachToolsAmt(10) + tbppcustbal(1).MachToolsAmt(9)
 tbppcustbal(1).MerchCapAmt(10) = tbppcustbal(1).MerchCapAmt(10) + tbppcustbal(1).MerchCapAmt(9)
 tbppcustbal(1).MobileHomeAmt(10) = tbppcustbal(1).MobileHomeAmt(10) + tbppcustbal(1).MobileHomeAmt(9)
 tbppcustbal(1).LateFeeAmt(10) = tbppcustbal(1).LateFeeAmt(10) + tbppcustbal(1).LateFeeAmt(9)
 tbppcustbal(1).IntAmt(10) = tbppcustbal(1).IntAmt(10) + tbppcustbal(1).IntAmt(9)
 tbppcustbal(1).PenAmt(10) = tbppcustbal(1).PenAmt(10) + tbppcustbal(1).PenAmt(9)
 tbppcustbal(1).TaxNotice(10) = tbppcustbal(1).TaxNotice(9)


 FOR Cnt! = 9 TO 2 STEP -1
 tbppcustbal(1).PerPropTaxAmt(Cnt!) = tbppcustbal(1).PerPropTaxAmt(Cnt! - 1)
 tbppcustbal(1).FarmEquipAmt(Cnt!) = tbppcustbal(1).FarmEquipAmt(Cnt! - 1)
 tbppcustbal(1).MachToolsAmt(Cnt!) = tbppcustbal(1).MachToolsAmt(Cnt! - 1)
 tbppcustbal(1).MerchCapAmt(Cnt!) = tbppcustbal(1).MerchCapAmt(Cnt! - 1)
 tbppcustbal(1).MobileHomeAmt(Cnt!) = tbppcustbal(1).MobileHomeAmt(Cnt! - 1)
 tbppcustbal(1).LateFeeAmt(Cnt!) = tbppcustbal(1).LateFeeAmt(Cnt! - 1)
 tbppcustbal(1).IntAmt(Cnt!) = tbppcustbal(1).IntAmt(Cnt! - 1)
 tbppcustbal(1).PenAmt(Cnt!) = tbppcustbal(1).PenAmt(Cnt! - 1)
 tbppcustbal(1).TaxNotice(Cnt!) = tbppcustbal(1).TaxNotice(Cnt! - 1)
 NEXT Cnt!

 tbppcustbal(1).PerPropTaxAmt(1) = TBPPCust(1).CurYrPersPropTax
 tbppcustbal(1).FarmEquipAmt(1) = TBPPCust(1).CurYrFarmEquipTax
 tbppcustbal(1).MachToolsAmt(1) = TBPPCust(1).CurYrMachToolsTax
 tbppcustbal(1).MerchCapAmt(1) = TBPPCust(1).CurYrMerchCapTax
 tbppcustbal(1).MobileHomeAmt(1) = TBPPCust(1).CurYrMobileHomeTax
 tbppcustbal(1).LateFeeAmt(1) = TBPPCust(1).CurYrLateFeeTax
 tbppcustbal(1).IntAmt(1) = TBPPCust(1).CurYrIntAmount
 tbppcustbal(1).PenAmt(1) = TBPPCust(1).CurYrPenAmount
 tbppcustbal(1).TaxNotice(1) = TBPPCust(1).CurYrBillNumber

PUT tbbalfile, AcctRecord, tbppcustbal(1)

 'Now to Set Current Stuff

        TBPPCust(1).CurYrBillNumber = ""
        TBPPCust(1).CurYrPersPropTax = 0
        TBPPCust(1).CurYrFarmEquipTax = 0
        TBPPCust(1).CurYrMachToolsTax = 0
        TBPPCust(1).CurYrMerchCapTax = 0
        TBPPCust(1).CurYrMobileHomeTax = 0
        TBPPCust(1).CurYrLateFeeTax = 0
        TBPPCust(1).CurYrIntAmount = 0
        TBPPCust(1).CurYrPenAmount = 0
        

 'Store Information for Posting in Detail
        

 ' Now Set Billing Information to Null for Next Yr Printing
        TBPPCust(1).LBCurYrBillNumber = ""
        TBPPCust(1).LBCurYrPersPropTaxAmount = 0
        TBPPCust(1).LBCurYrFarmEquipTaxAmount = 0
        TBPPCust(1).LBCurYrMachToolsTaxAmount = 0
        TBPPCust(1).LBCurYrMerchCapTaxAmount = 0
        TBPPCust(1).LBCurYrMobileHomeTaxAmount = 0
        TBPPCust(1).LBLateFeeAmount = 0
  PUT TBFile, AcctRecord, TBPPCust(1)
  RETURN



60000  IF TaxType$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TBRTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBPath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN

60025 tbbalfile = FREEFILE
      IF TaxType$ = "R" THEN
      OPEN TBPath$ + "TBRTBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #tbbalfile LEN = LEN(tbrecustbal(1))
      ELSE
      OPEN TBPath$ + "TBPTBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #tbbalfile LEN = LEN(tbppcustbal(1))
      END IF
60030 RETURN

61000 TBTranFile = FREEFILE
      OPEN TBPath$ + "TBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTranFile LEN = LEN(TBTrans(1))
61020 RETURN




62000 COLOR 14: CF = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); scren$: PRINT STRING$(79, "_"): RETURN

Namesearch:
 LOCATE 8, 1:
 COLOR 11
 PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
 GOSUB 62000
 IF CF = 1 THEN
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
  OPEN TBPath$ + "TBRENAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
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
  OPEN TBPath$ + "TBPPNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
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

Deleted:
    LOCATE 23, 1: COLOR 15
    PRINT "This Account Has Been Deleted! Press <ENTER> to Continue"; :
    fl = 1
    GOSUB 62000
    LOCATE 23, 1: PRINT STRING$(79, 32)
    RETURN

FigureBalance:
     CurBal# = 0: PastDue# = 0
     IF TaxType$ = "R" THEN
     CurBal# = TBRECust(1).CurYrTaxAmount + TBRECust(1).CurYrIntAmount + TBRECust(1).CurYrPenAmount + TBRECust(1).CurYrOthAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR ll = 1 TO 10
     PastDue# = PastDue# + tbrecustbal(1).TaxAmt(ll) + tbrecustbal(1).IntAmt(ll) + tbrecustbal(1).PenAmt(ll) + tbrecustbal(1).OthAmt(ll) + tbrecustbal(1).LateAmt(ll)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT ll
     ELSE
     CurBal# = TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax + TBPPCust(1).CurYrIntAmount + TBPPCust(1). _
CurYrPenAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR ll = 1 TO 10
     PastDue# = PastDue# + tbppcustbal(1).PerPropTaxAmt(ll) + tbppcustbal(1).FarmEquipAmt(ll) + tbppcustbal(1).MachToolsAmt(ll) + tbppcustbal(1).MerchCapAmt(ll) + tbppcustbal(1).MobileHomeAmt(ll) + tbppcustbal(1).LateFeeAmt(ll) + tbppcustbal(1). _
IntAmt(ll) + tbppcustbal(1).PenAmt(ll)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT ll


     END IF
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
               fl = 1: GOSUB 62000: IF CF = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
            
     END SELECT
    

