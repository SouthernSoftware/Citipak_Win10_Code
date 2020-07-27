10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbpayment
30 REM Payment Entry Menu
40 REM Last Revision Date: 7-14-87
50 REM by Nick Picerno
  DECLARE FUNCTION num2date$ (Num%)
  DECLARE FUNCTION Date2Num% (Dat$)
60 CLS : CLEAR
   
   '$INCLUDE: 'lcTAX.BI'

     'TBPath$ = "F:\"                    ' change to "F:\" Before compiling
      TBPath$ = "c:\QB45\LC\"                    ' change to "F:\" Before compiling

     CLOSE : OPEN "R", 1, TBPath$ + "tbyear.dat", 4: FIELD 1, 4 AS Year$: GET 1, 1
     CurrentYear$ = Year$: CLOSE
     CurrentYear = VAL(CurrentYear$)
     ON ERROR GOTO networkerror
     DIM SHARED yrdesc(11)
     REDIM TBLog(1) AS TBLog
     REDIM TBRECust(1) AS TBRECust
     REDIM TBRECustBal(1) AS TBREBAL
     REDIM TBPPCust(1) AS TBPPCust
     REDIM TBPPCustBal(1) AS TBPPBal
     REDIM TBTrans(1) AS TBTrans

     PDate$ = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 2)
   
5000 SCREN$ = "4.05.00": CLS : GOSUB 63500
5020 COLOR 11: LOCATE 1, 32: PRINT "Customer Manual Adjustment";
5025 LOCATE 3, 1: COLOR 11
     IF taxtype$ = "R" OR taxtype$ = "P" THEN
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; taxtype$
     ELSE
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN RUN "TBPAYMEN"
     IF NOT (in$ = "R" OR in$ = "P" OR in$ = "A") THEN 5025
     taxtype$ = UCASE$(in$)
     END IF
5030 LOCATE 4, 1: COLOR 11
     PRINT "Enter Customer Account Number (TAB-Last Name : ESC-Menu) .. "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN
      CLOSE : GOTO 10
     END IF
     IF cf = 2 THEN
      GOSUB Namesearch
      IF ggo = 1 THEN 5075
      IF ggo = 0 THEN LOCATE 10, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 62000: GOTO 5000
     END IF
     Record! = VAL(in$)
5075 CLOSE
     GOSUB 60000
     GOSUB 60025
     IF taxtype$ = "R" THEN
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBRECust(1)) THEN 5030
       LOCK #TBFile, Record!
       GET TBFile, Record!, TBRECust(1)
       GET TBBalFile, Record!, TBRECustBal(1)
       IF TBRECust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 5000
       ELSE
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBPPCust(1)) THEN 5030
       LOCK #TBFile, Record!
       GET TBFile, Record!, TBPPCust(1)
       GET TBBalFile, Record!, TBPPCustBal(1)
       IF TBPPCust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 5000
     END IF

     IF taxtype$ = "R" THEN
       GOSUB ProcessRealAdj
        ELSE
       GOSUB ProcessPersonalAdj
     END IF
     CLOSE
     GOTO 10




ProcessRealAdj:

FOR ll = 3 TO 23: LOCATE ll, 1: PRINT STRING$(79, 32): NEXT ll
CLOSE TBFile
CLOSE TBBalFile
GOSUB 60000
GOSUB 60025
GET TBBalFile, Record!, TBRECustBal(1)
LOCATE 3, 1: COLOR 11
PRINT "Account of ... "; TBRECust(1).CustName
PRINT
PRINT " Year"; TAB(15); "Tax Amount"; TAB(30); "Int't Amt"; TAB(44); "Pen'ty Amt"; TAB(63); " Balance"
PRINT STRING$(79, "-")
COLOR 15
PRINT " 1-"; VAL(CurrentYear$); TAB(13); USING "#####,#.##"; TBRECust(1).CurYrTaxAmount;
 PRINT TAB(30); USING "####,#.##"; TBRECust(1).CurYrIntAmount;
  PRINT TAB(44); USING "#####,#.##"; TBRECust(1).CurYrPenAmount;
   PRINT TAB(60); USING "$$#####,#.##"; TBRECust(1).CurYrTaxAmount + TBRECust(1).CurYrIntAmount + TBRECust(1).CurYrPenAmount
FOR Pyr = 1 TO 10
Cnt$ = STR$(Pyr + 1)
Cnt$ = RTRIM$(Cnt$)
IF Pyr >= 9 THEN Cnt$ = LTRIM$(Cnt$)

PRINT Cnt$; "-"; VAL(CurrentYear$) - Pyr;
PRINT TAB(13); USING "#####,#.##"; TBRECustBal(1).TaxAmt(Pyr);
 PRINT TAB(30); USING "####,#.##"; TBRECustBal(1).IntAmt(Pyr);
  PRINT TAB(44); USING "#####,#.##"; TBRECustBal(1).PenAmt(Pyr);
   PRINT TAB(60); USING "$$#####,#.##"; TBRECustBal(1).TaxAmt(Pyr) + TBRECustBal(1).IntAmt(Pyr) + TBRECustBal(1).PenAmt(Pyr)
NEXT Pyr



5100 LOCATE 23, 1: COLOR 11
        PRINT "Enter Selection # to Change (1-11) or (ESC-End) .. "; : fl = -2
        GOSUB 62000
        IF cf = 1 THEN
         GOSUB UpdateRealTransaction
         CLOSE : GOTO 10
        END IF
        ON VAL(in$) GOSUB l1, l2, l3, l4, l5, l6, l7, l8, l9, l10, l11
        PUT TBFile, Record!, TBRECust(1)
        PUT TBBalFile, Record!, TBRECustBal(1)
        GOTO ProcessRealAdj

UpdateRealTransaction:
        ' UPDATE TRANSACTION HISTORY NOW
        GOSUB 60000
        GOSUB 60025

        GET TBFile, Record!, TBRECust(1)
        GET TBBalFile, Record!, TBRECustBal(1)
        GOSUB FigureBalance
        AcctBalance# = CurBal# + PastDue#
        IF AcctBalance# <> 0 THEN
        GOSUB 61000
  TBTrans(1).TransDate = Date2Num%(DATE$)
 TBTrans(1).TransAmt = AcctBalance#
 TBTrans(1).TRTaxYear = CurrentYear$
 TBTrans(1).TRDesc = "Balance Reset"
 TBTrans(1).TRType = 1                   '1=bill 2=payment 3=INT  4= PEN 5=release 6=BEG BAL
 TBTrans(1).TRTax = 0
 TBTrans(1).TrInterest = 0
 TBTrans(1).TrPenalty = 0
 TBTrans(1).TrLate = 0
 TBTrans(1).TrMachTools = 0
 TBTrans(1).TrMobileHome = 0
 TBTrans(1).TrMerchCap = 0
 TBTrans(1).TrFarmEquip = 0
 TBTrans(1).TrCustomerRecord = Record!
 TBTrans(1).TRNextTrans = 0
 TBTrans(1).TrRunningBal = AcctBalance#
 TBTrans(1).TrSource = "R"
 TBTrans(1).TrPad = ""
 PUT TBTranFile, LOF(TBTranFile) / LEN(TBTrans(1)) + 1, TBTrans(1)
 TransRecord! = LOF(TBTranFile) / LEN(TBTrans(1))
 TBRECust(1).FirstTrans = TransRecord!
 TBRECust(1).LastTrans = TransRecord!
 PUT TBFile, Record!, TBRECust(1)
 CLOSE TBTranFile
 CLOSE TBFile
 END IF
 RETURN

l1:
LOCATE 7, 14: fl = -9: GOSUB 62000
Amt# = TBRECust(1).CurYrTaxAmount: CustName$ = TBRECust(1).CustName: Code = 1
TBRECust(1).CurYrTaxAmount = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE 7, 31: fl = -9: GOSUB 62000
Amt# = TBRECust(1).CurYrIntAmount: CustName$ = TBRECust(1).CustName: Code = 2
TBRECust(1).CurYrIntAmount = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE 7, 45: fl = -9: GOSUB 62000
Amt# = TBRECust(1).CurYrPenAmount: CustName$ = TBRECust(1).CustName: Code = 3
TBRECust(1).CurYrPenAmount = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN
l2:
LOCATE 8, 14: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).TaxAmt(1): CustName$ = TBRECust(1).CustName: Code = 1
TBRECustBal(1).TaxAmt(1) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE 8, 31: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).IntAmt(1): CustName$ = TBRECust(1).CustName: Code = 2
TBRECustBal(1).IntAmt(1) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE 8, 45: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).PenAmt(1): CustName$ = TBRECust(1).CustName: Code = 3
TBRECustBal(1).PenAmt(1) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

l3:
Place = 9: Year = 2
LOCATE Place, 14: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).TaxAmt(Year): CustName$ = TBRECust(1).CustName: Code = 1
TBRECustBal(1).TaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 31: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).IntAmt(Year): CustName$ = TBRECust(1).CustName: Code = 2
TBRECustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 45: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).PenAmt(Year): CustName$ = TBRECust(1).CustName: Code = 3
TBRECustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN
l4:
Place = 10: Year = 3
LOCATE Place, 14: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).TaxAmt(Year): CustName$ = TBRECust(1).CustName: Code = 1
TBRECustBal(1).TaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 31: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).IntAmt(Year): CustName$ = TBRECust(1).CustName: Code = 2
TBRECustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 45: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).PenAmt(Year): CustName$ = TBRECust(1).CustName: Code = 3
TBRECustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN
l5:
Place = 11: Year = 4
LOCATE Place, 14: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).TaxAmt(Year): CustName$ = TBRECust(1).CustName: Code = 1
TBRECustBal(1).TaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 31: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).IntAmt(Year): CustName$ = TBRECust(1).CustName: Code = 2
TBRECustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 45: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).PenAmt(Year): CustName$ = TBRECust(1).CustName: Code = 3
TBRECustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

l6:
Place = 12: Year = 5
LOCATE Place, 14: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).TaxAmt(Year): CustName$ = TBRECust(1).CustName: Code = 1
TBRECustBal(1).TaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 31: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).IntAmt(Year): CustName$ = TBRECust(1).CustName: Code = 2
TBRECustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 45: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).PenAmt(Year): CustName$ = TBRECust(1).CustName: Code = 3
TBRECustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

l7:
Place = 13: Year = 6
LOCATE Place, 14: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).TaxAmt(Year): CustName$ = TBRECust(1).CustName: Code = 1
TBRECustBal(1).TaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 31: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).IntAmt(Year): CustName$ = TBRECust(1).CustName: Code = 2
TBRECustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 45: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).PenAmt(Year): CustName$ = TBRECust(1).CustName: Code = 3
TBRECustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

l8:
Place = 14: Year = 7
LOCATE Place, 14: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).TaxAmt(Year): CustName$ = TBRECust(1).CustName: Code = 1
TBRECustBal(1).TaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 31: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).IntAmt(Year): CustName$ = TBRECust(1).CustName: Code = 2
TBRECustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 45: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).PenAmt(Year): CustName$ = TBRECust(1).CustName: Code = 3
TBRECustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

l9:
Place = 15: Year = 8
LOCATE Place, 14: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).TaxAmt(Year): CustName$ = TBRECust(1).CustName: Code = 1
TBRECustBal(1).TaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 31: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).IntAmt(Year): CustName$ = TBRECust(1).CustName: Code = 2
TBRECustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 45: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).PenAmt(Year): CustName$ = TBRECust(1).CustName: Code = 3
TBRECustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

l10:
Place = 16: Year = 9
LOCATE Place, 14: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).TaxAmt(Year): CustName$ = TBRECust(1).CustName: Code = 1
TBRECustBal(1).TaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 31: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).IntAmt(Year): CustName$ = TBRECust(1).CustName: Code = 2
TBRECustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 45: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).PenAmt(Year): CustName$ = TBRECust(1).CustName: Code = 3
TBRECustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

l11:
Place = 17: Year = 10
LOCATE Place, 14: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).TaxAmt(Year): CustName$ = TBRECust(1).CustName: Code = 1
TBRECustBal(1).TaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 31: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).IntAmt(Year): CustName$ = TBRECust(1).CustName: Code = 2
TBRECustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 45: fl = -9: GOSUB 62000
Amt# = TBRECustBal(1).PenAmt(Year): CustName$ = TBRECust(1).CustName: Code = 3
TBRECustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

LogIt:
 CLOSE 14:
 OPEN "R", 14, "TBLog.Dat", LEN(TBLog(1))
 TBLog(1).LogAcct = Record!
 TBLog(1).LogOrigAmt = Amt#
 TBLog(1).LogNewAmt = NewAmt#
 TBLog(1).LogName = CustName$
 TBLog(1).LogCode = Code
 PUT 14, LOF(14) / LEN(TBLog(1)) + 1, TBLog(1)
 CLOSE 14
 RETURN


ProcessPersonalAdj:
 FOR ll = 3 TO 23: LOCATE ll, 1: PRINT STRING$(79, 32): NEXT ll
 CLOSE TBBalFile
 GOSUB 60025
 GET TBBalFile, Record!, TBPPCustBal(1)
 LOCATE 3, 1: COLOR 11
 PRINT "Account of ... "; TBPPCust(1).CustName
 PRINT TAB(13); "Pers"; TAB(22); "Farm"; TAB(31); "Mach"; TAB(40); "Merch"; TAB(49); "Mobile"; TAB(58); "Late"; TAB(67); "Int"; TAB(75); "Pen"
 PRINT " Year"; TAB(13); "Prop"; TAB(22); "Equip"; TAB(31); "Tools"; TAB(40); "Cap."; TAB(49); "Home"; TAB(58); "Fee"; TAB(67); "Amt"; TAB(75); "Amt"
 PRINT STRING$(79, "-")
 COLOR 15
 PRINT " 1-"; VAL(CurrentYear$);
  PRINT TAB(11); USING "#####.##"; TBPPCust(1).CurYrPersPropTax;
  PRINT TAB(20); USING "#####.##"; TBPPCust(1).CurYrFarmEquipTax;
  PRINT TAB(29); USING "#####.##"; TBPPCust(1).CurYrMachToolsTax;
  PRINT TAB(38); USING "#####.##"; TBPPCust(1).CurYrMerchCapTax;
  PRINT TAB(47); USING "#####.##"; TBPPCust(1).CurYrMobileHomeTax;
  PRINT TAB(56); USING "####.##"; TBPPCust(1).CurYrLateFeeTax;
   PRINT TAB(65); USING "####.##"; TBPPCust(1).CurYrIntAmount;
   PRINT TAB(74); USING "####.##"; TBPPCust(1).CurYrPenAmount
   
 FOR Pyr = 1 TO 10
 Cnt$ = STR$(Pyr + 1)
 Cnt$ = RTRIM$(Cnt$)
 IF Pyr >= 9 THEN Cnt$ = LTRIM$(Cnt$)

 PRINT Cnt$; "-"; VAL(CurrentYear$) - Pyr;
 PRINT TAB(11); USING "#####.##"; TBPPCustBal(1).PerPropTaxAmt(Pyr);
  PRINT TAB(20); USING "#####.##"; TBPPCustBal(1).FarmEquipAmt(Pyr);
   PRINT TAB(29); USING "#####.##"; TBPPCustBal(1).MachToolsAmt(Pyr);
    PRINT TAB(38); USING "#####.##"; TBPPCustBal(1).MerchCapAmt(Pyr);
    PRINT TAB(47); USING "#####.##"; TBPPCustBal(1).MobileHomeAmt(Pyr);
    PRINT TAB(56); USING "####.##"; TBPPCustBal(1).LateFeeAmt(Pyr);
    PRINT TAB(65); USING "####.##"; TBPPCustBal(1).IntAmt(Pyr);
    PRINT TAB(74); USING "####.##"; TBPPCustBal(1).PenAmt(Pyr)
 NEXT Pyr



5200 LOCATE 23, 1: COLOR 11
        PRINT "Enter Selection # to Change (1-11) or (ESC-End) .. "; : fl = -2
        GOSUB 62000
        IF cf = 1 THEN
         CLOSE
         GOSUB UpdatePersHistory
         CLOSE : GOTO 10
        END IF
        ON VAL(in$) GOSUB P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11
        PUT TBFile, Record!, TBPPCust(1)
        PUT TBBalFile, Record!, TBPPCustBal(1)
        GOTO ProcessPersonalAdj

P1:
LOCATE 7, 11: fl = -8: GOSUB 62000
Amt# = TBPPCust(1).CurYrPersPropTax: Code = 11
TBPPCust(1).CurYrPersPropTax = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE 7, 20: fl = -8: GOSUB 62000
Amt# = TBPPCust(1).CurYrFarmEquipTax: Code = 12
TBPPCust(1).CurYrFarmEquipTax = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE 7, 29: fl = -8: GOSUB 62000
Amt# = TBPPCust(1).CurYrMachToolsTax: Code = 13
TBPPCust(1).CurYrMachToolsTax = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE 7, 38: fl = -8: GOSUB 62000
Amt# = TBPPCust(1).CurYrMerchCapTax: Code = 14
TBPPCust(1).CurYrMerchCapTax = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE 7, 47: fl = -8: GOSUB 62000
Amt# = TBPPCust(1).CurYrMobileHomeTax: Code = 15
TBPPCust(1).CurYrMobileHomeTax = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE 7, 56: fl = -8: GOSUB 62000
Amt# = TBPPCust(1).CurYrLateFeeTax: Code = 16
TBPPCust(1).CurYrLateFeeTax = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE 7, 65: fl = -7: GOSUB 62000
Amt# = TBPPCust(1).CurYrIntAmount: Code = 17
TBPPCust(1).CurYrIntAmount = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE 7, 74: fl = -7: GOSUB 62000
Amt# = TBPPCust(1).CurYrPenAmount: Code = 18
TBPPCust(1).CurYrPenAmount = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

P2:
Place = 8: Year = 1
LOCATE Place, 11: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).PerPropTaxAmt(Year): Code = 21
TBPPCustBal(1).PerPropTaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 20: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).FarmEquipAmt(Year): Code = 22
TBPPCustBal(1).FarmEquipAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 29: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MachToolsAmt(Year): Code = 23
TBPPCustBal(1).MachToolsAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 38: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MerchCapAmt(Year): Code = 24
TBPPCustBal(1).MerchCapAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 47: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MobileHomeAmt(Year): Code = 25
TBPPCustBal(1).MobileHomeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 56: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).LateFeeAmt(Year): Code = 26
TBPPCustBal(1).LateFeeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 65: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).IntAmt(Year): Code = 27
TBPPCustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 74: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).PenAmt(Year): Code = 18
TBPPCustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

P3:
Place = 9: Year = 2
LOCATE Place, 11: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).PerPropTaxAmt(Year): Code = 31
TBPPCustBal(1).PerPropTaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 20: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).FarmEquipAmt(Year): Code = 32
TBPPCustBal(1).FarmEquipAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 29: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MachToolsAmt(Year): Code = 33
TBPPCustBal(1).MachToolsAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 38: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MerchCapAmt(Year): Code = 34
TBPPCustBal(1).MerchCapAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 47: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MobileHomeAmt(Year): Code = 35
TBPPCustBal(1).MobileHomeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 56: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).LateFeeAmt(Year): Code = 36
TBPPCustBal(1).LateFeeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 65: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).IntAmt(Year): Code = 37
TBPPCustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 74: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).PenAmt(Year): Code = 38
TBPPCustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

P4:
Place = 10: Year = 3
LOCATE Place, 11: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).PerPropTaxAmt(Year): Code = 41
TBPPCustBal(1).PerPropTaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 20: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).FarmEquipAmt(Year): Code = 42
TBPPCustBal(1).FarmEquipAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 29: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MachToolsAmt(Year): Code = 43
TBPPCustBal(1).MachToolsAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 38: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MerchCapAmt(Year): Code = 44
TBPPCustBal(1).MerchCapAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 47: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MobileHomeAmt(Year): Code = 45
TBPPCustBal(1).MobileHomeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 56: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).LateFeeAmt(Year): Code = 46
TBPPCustBal(1).LateFeeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 65: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).IntAmt(Year): Code = 47
TBPPCustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 74: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).PenAmt(Year): Code = 48
TBPPCustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

P5:
Place = 11: Year = 4
LOCATE Place, 11: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).PerPropTaxAmt(Year): Code = 51
TBPPCustBal(1).PerPropTaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 20: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).FarmEquipAmt(Year): Code = 52
TBPPCustBal(1).FarmEquipAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 29: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MachToolsAmt(Year): Code = 53
TBPPCustBal(1).MachToolsAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 38: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MerchCapAmt(Year): Code = 54
TBPPCustBal(1).MerchCapAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 47: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MobileHomeAmt(Year): Code = 55
TBPPCustBal(1).MobileHomeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 56: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).LateFeeAmt(Year): Code = 56
TBPPCustBal(1).LateFeeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 65: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).IntAmt(Year): Code = 57
TBPPCustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 74: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).PenAmt(Year): Code = 58
TBPPCustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

P6:
Place = 12: Year = 5
LOCATE Place, 11: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).PerPropTaxAmt(Year): Code = 61
TBPPCustBal(1).PerPropTaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 20: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).FarmEquipAmt(Year): Code = 62
TBPPCustBal(1).FarmEquipAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 29: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MachToolsAmt(Year): Code = 63
TBPPCustBal(1).MachToolsAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 38: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MerchCapAmt(Year): Code = 64
TBPPCustBal(1).MerchCapAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 47: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MobileHomeAmt(Year): Code = 65
TBPPCustBal(1).MobileHomeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 56: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).LateFeeAmt(Year): Code = 66
TBPPCustBal(1).LateFeeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 65: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).IntAmt(Year): Code = 67
TBPPCustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 74: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).PenAmt(Year): Code = 68
TBPPCustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

P7:
Place = 13: Year = 6

LOCATE Place, 11: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).PerPropTaxAmt(Year): Code = 71
TBPPCustBal(1).PerPropTaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 20: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).FarmEquipAmt(Year): Code = 72
TBPPCustBal(1).FarmEquipAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 29: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MachToolsAmt(Year): Code = 73
TBPPCustBal(1).MachToolsAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 38: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MerchCapAmt(Year): Code = 74
TBPPCustBal(1).MerchCapAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 47: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MobileHomeAmt(Year): Code = 75
TBPPCustBal(1).MobileHomeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 56: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).LateFeeAmt(Year): Code = 76
TBPPCustBal(1).LateFeeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 65: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).IntAmt(Year): Code = 77
TBPPCustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 74: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).PenAmt(Year): Code = 78
TBPPCustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

P8:
Place = 14: Year = 7
LOCATE Place, 11: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).PerPropTaxAmt(Year): Code = 81
TBPPCustBal(1).PerPropTaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 20: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).FarmEquipAmt(Year): Code = 82
TBPPCustBal(1).FarmEquipAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 29: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MachToolsAmt(Year): Code = 83
TBPPCustBal(1).MachToolsAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 38: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MerchCapAmt(Year): Code = 84
TBPPCustBal(1).MerchCapAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 47: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MobileHomeAmt(Year): Code = 85
TBPPCustBal(1).MobileHomeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 56: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).LateFeeAmt(Year): Code = 86
TBPPCustBal(1).LateFeeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 65: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).IntAmt(Year): Code = 87
TBPPCustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 74: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).PenAmt(Year): Code = 88
TBPPCustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

P9:
Place = 15: Year = 8
LOCATE Place, 11: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).PerPropTaxAmt(Year): Code = 91
TBPPCustBal(1).PerPropTaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 20: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).FarmEquipAmt(Year): Code = 92
TBPPCustBal(1).FarmEquipAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 29: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MachToolsAmt(Year): Code = 93
TBPPCustBal(1).MachToolsAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 38: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MerchCapAmt(Year): Code = 94
TBPPCustBal(1).MerchCapAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 47: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MobileHomeAmt(Year): Code = 95
TBPPCustBal(1).MobileHomeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 56: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).LateFeeAmt(Year): Code = 96
TBPPCustBal(1).LateFeeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 65: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).IntAmt(Year): Code = 97
TBPPCustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 74: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).PenAmt(Year): Code = 98
TBPPCustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

P10:
Place = 16: Year = 9
LOCATE Place, 11: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).PerPropTaxAmt(Year): Code = 101
TBPPCustBal(1).PerPropTaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 20: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).FarmEquipAmt(Year): Code = 102
TBPPCustBal(1).FarmEquipAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 29: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MachToolsAmt(Year): Code = 103
TBPPCustBal(1).MachToolsAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 38: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MerchCapAmt(Year): Code = 104
TBPPCustBal(1).MerchCapAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 47: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MobileHomeAmt(Year): Code = 105
TBPPCustBal(1).MobileHomeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 56: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).LateFeeAmt(Year): Code = 106
TBPPCustBal(1).LateFeeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 65: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).IntAmt(Year): Code = 107
TBPPCustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 74: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).PenAmt(Year): Code = 108
TBPPCustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

P11:
Place = 17: Year = 10
LOCATE Place, 11: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).PerPropTaxAmt(Year): Code = 111
TBPPCustBal(1).PerPropTaxAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 20: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).FarmEquipAmt(Year): Code = 112
TBPPCustBal(1).FarmEquipAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 29: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MachToolsAmt(Year): Code = 113
TBPPCustBal(1).MachToolsAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 38: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MerchCapAmt(Year): Code = 114
TBPPCustBal(1).MerchCapAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 47: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).MobileHomeAmt(Year): Code = 115
TBPPCustBal(1).MobileHomeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 56: fl = -8: GOSUB 62000
Amt# = TBPPCustBal(1).LateFeeAmt(Year): Code = 116
TBPPCustBal(1).LateFeeAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 65: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).IntAmt(Year): Code = 117
TBPPCustBal(1).IntAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
LOCATE Place, 74: fl = -7: GOSUB 62000
Amt# = TBPPCustBal(1).PenAmt(Year): Code = 118
TBPPCustBal(1).PenAmt(Year) = VAL(in$)
NewAmt# = VAL(in$)
GOSUB LogIt
RETURN

UpdatePersHistory:
        ' UPDATE TRANSACTION HISTORY NOW
        GOSUB 60000
        GOSUB 60025

        GET TBFile, Record!, TBPPCust(1)
        GET TBBalFile, Record!, TBPPCustBal(1)
        GOSUB FigureBalance
        AcctBalance# = CurBal# + PastDue#
        AcctBalance# = INT((AcctBalance# * 100) + .5) / 100
        IF AcctBalance# <> 0 THEN
        GOSUB 61000
  TBTrans(1).TransDate = Date2Num%(DATE$)
 TBTrans(1).TransAmt = AcctBalance#
 TBTrans(1).TRTaxYear = CurrentYear$
 TBTrans(1).TRDesc = "Balance Reset"
 TBTrans(1).TRType = 1                   '1=bill 2=payment 3=INT  4= PEN 5=release 6=BEG BAL
 TBTrans(1).TRTax = 0
 TBTrans(1).TrInterest = 0
 TBTrans(1).TrPenalty = 0
 TBTrans(1).TrLate = 0
 TBTrans(1).TrMachTools = 0
 TBTrans(1).TrMobileHome = 0
 TBTrans(1).TrMerchCap = 0
 TBTrans(1).TrFarmEquip = 0
 TBTrans(1).TrCustomerRecord = Record!
 TBTrans(1).TRNextTrans = 0
 TBTrans(1).TrRunningBal = AcctBalance#
 TBTrans(1).TrSource = "P"
 TBTrans(1).TrPad = ""
 PUT TBTranFile, LOF(TBTranFile) / LEN(TBTrans(1)) + 1, TBTrans(1)
 TransRecord! = LOF(TBTranFile) / LEN(TBTrans(1))
 TBPPCust(1).FirstTrans = TransRecord!
 TBPPCust(1).LastTrans = TransRecord!
 PUT TBFile, Record!, TBPPCust(1)
 CLOSE TBTranFile
 CLOSE TBFile
 END IF
 RETURN

60000  IF taxtype$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBPath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN
60025 TBBalFile = FREEFILE
      IF taxtype$ = "R" THEN
      OPEN TBPath$ + "TBREBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBPath$ + "TBPPBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF
60030 RETURN

61000 TBTranFile = FREEFILE
      OPEN TBPath$ + "TBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTranFile LEN = LEN(TBTrans(1))
61020 RETURN

62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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
62140 IF z$ = "." AND INSTR(in$, ".") = 0 THEN GOTO 62180
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
 nn$ = in$
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
GET TBFile, CVS(Mrec$), TBRECust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBRECust(1).CustName
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
GET TBFile, CVS(Mrec$), TBPPCust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBPPCust(1).CustName
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

Deleted:
    LOCATE 23, 1: COLOR 15
    PRINT "This Account Has Been Deleted! Press <ENTER> to Continue"; :
    fl = 1
    GOSUB 62000
    LOCATE 23, 1: PRINT STRING$(79, 32)
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

