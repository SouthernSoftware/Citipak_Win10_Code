10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbpayment
30 REM Payment Entry Menu
40 REM Last Revision Date: 7-14-87
50 REM by Nick Picerno
  DECLARE FUNCTION Num2Date$ (Num%)
  DECLARE FUNCTION Date2Num% (Dat$)
60 CLS : CLEAR
   
   '$INCLUDE: 'lcTAX.BI'

    TBPath$ = "F:\"                    ' change to "F:\" Before compiling
    'TBPath$ = "c:\QB45\LC\"                    ' change to "F:\" Before compiling

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
     REDIM TBRealExport(1) AS TBRealExport
     REDIM TBPersExport(1) AS TBPersonalExport

     PDate$ = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 2)

1000 scren$ = "4.01.00": CLS : GOSUB 63500
1020 COLOR 11: LOCATE 1, 32: PRINT "Customer Export Utility";
1025 LOCATE 3, 1: COLOR 11
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN RUN "tbtmenu"
     IF NOT (in$ = "R" OR in$ = "P" OR in$ = "A") THEN 1025
     taxtype$ = UCASE$(in$)
     IF taxtype$ = "P" THEN 2000
     IF taxtype$ = "R" THEN 3000
     GOTO 10
2000 CLS : scren$ = "3.03.00": GOSUB 63500: COLOR 11: LOCATE 1, 25
     PRINT "Extract Personal Property for Export"
     LOCATE 3, 1: COLOR 11
     PRINT "Enter the District Code to Pull (1-10) .. "; : fl = -2
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbCustMa"
     District = VAL(in$): District$ = in$
      IF District < 1 OR District > 10 THEN 3000
      TBFile = FREEFILE
      OPEN "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
      TBEFILE = FREEFILE
      OPEN "TBPERS" + District$ + ".EXP" FOR RANDOM ACCESS READ WRITE SHARED AS #TBEFILE LEN = LEN(TBPersExport(1))
      CLOSE TBEFILE
      FileNam$ = "DEL TBPERS" + District$ + ".EXP"
      SHELL FileNam$
      OPEN "TBPERS" + District$ + ".EXP" FOR RANDOM ACCESS READ WRITE SHARED AS #TBEFILE LEN = LEN(TBPersExport(1))
      
      Cnt = 1
      WHILE Cnt <= LOF(TBFile) / LEN(TBPPCust(1))
        GET TBFile, Cnt, TBPPCust(1)
        LOCATE 8, 1: COLOR 11: PRINT "Processing Record # "; : COLOR 15: PRINT USING "######"; Cnt
        IF VAL(TBPPCust(1).District) = District THEN
        TBPersExport(1).Account = LTRIM$(STR$(Cnt))
        TBPersExport(1).CustName = TBPPCust(1).CustName
        TBPersExport(1).Address1 = TBPPCust(1).Address1
        TBPersExport(1).Address2 = TBPPCust(1).Address2
        TBPersExport(1).City = TBPPCust(1).City
        TBPersExport(1).State = TBPPCust(1).State
        TBPersExport(1).Zip = TBPPCust(1).Zip
        TBPersExport(1).SocSec = TBPPCust(1).SocSec
        TBPersExport(1).SocSec2 = TBPPCust(1).SocSec2
        TBPersExport(1).Short = TBPPCust(1).Short
        TBPersExport(1).PersProp = RTRIM$(STR$(TBPPCust(1).PersProp))
        TBPersExport(1).MachTools = RTRIM$(STR$(TBPPCust(1).MachTools))
        TBPersExport(1).MobileHome = RTRIM$(STR$(TBPPCust(1).MobileHome))
        TBPersExport(1).MerchCap = RTRIM$(STR$(TBPPCust(1).MerchCap))
        TBPersExport(1).FarmEquip = RTRIM$(STR$(TBPPCust(1).FarmEquip))
        TBPersExport(1).Business = TBPPCust(1).Business
        TBPersExport(1).District = TBPPCust(1).District
        TBPersExport(1).LateFile = TBPPCust(1).LateFile
        PUT TBEFILE, (LOF(TBEFILE) / (LEN(TBPersExport(1)))) + 1, TBPersExport(1)
        END IF
        Cnt = Cnt + 1
      WEND
      CLOSE
      GOTO 10


3000 CLS : scren$ = "3.03.00": GOSUB 63500: COLOR 11
     LOCATE 1, 25: PRINT "Extract District Data for Export"
     LOCATE 3, 1: COLOR 11
     PRINT "Enter the District Code to Pull (1-10) .. "; : fl = -2
     GOSUB 62000
     IF cf = 1 THEN 10
     District = VAL(in$): District$ = in$
      IF District < 1 OR District > 10 THEN 3000
      TBFile = FREEFILE
      OPEN TBPath$ + "TBRTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
      TBEFILE = FREEFILE
      OPEN "TBREAL" + District$ + ".EXP" FOR RANDOM ACCESS READ WRITE SHARED AS #TBEFILE LEN = LEN(TBRealExport(1))
      CLOSE TBEFILE
      FileNam$ = "DEL TBREAL" + District$ + ".EXP"
      SHELL FileNam$
      OPEN "TBREAL" + District$ + ".EXP" FOR RANDOM ACCESS READ WRITE SHARED AS #TBEFILE LEN = LEN(TBRealExport(1))
      RecLen = LEN(TBRECust(1))

      Cnt = 1
      WHILE Cnt <= LOF(TBFile) / LEN(TBRECust(1))
        GET TBFile, Cnt, TBRECust(1)
        LOCATE 8, 1: COLOR 11: PRINT "Processing Record # "; : COLOR 15: PRINT USING "######"; Cnt
        IF VAL(TBRECust(1).District) = District THEN
        TBRealExport(1).Account = LTRIM$(STR$(Cnt))
        TBRealExport(1).MapNumber = TBRECust(1).MapNumber
        TBRealExport(1).District = TBRECust(1).District
        TBRealExport(1).CustName = TBRECust(1).CustName
        TBRealExport(1).Address1 = TBRECust(1).Address1
        TBRealExport(1).Address2 = TBRECust(1).Address2
        TBRealExport(1).Address3 = TBRECust(1).Address3
        TBRealExport(1).Zipcode = TBRECust(1).Zipcode
        TBRealExport(1).PropDesc1 = TBRECust(1).PropDesc1
        TBRealExport(1).PropDesc2 = TBRECust(1).PropDesc2
        TBRealExport(1).PropDesc3 = TBRECust(1).PropDesc3
        TBRealExport(1).PropDesc4 = TBRECust(1).PropDesc4
        TBRealExport(1).Zone = TBRECust(1).Zone
        TBRealExport(1).MGCode = TBRECust(1).MGCode
        TBRealExport(1).ClassCode = TBRECust(1).ClassCode
        TBRealExport(1).Tract = TBRECust(1).Tract
        TBRealExport(1).SqFtSize = TBRECust(1).SqFtSize
        TBRealExport(1).Acres = TBRECust(1).Acres
        TBRealExport(1).TransferYr = TBRECust(1).TransferYr
        TBRealExport(1).Comment1 = TBRECust(1).Comment1
        TBRealExport(1).Comment2 = TBRECust(1).Comment2
        TBRealExport(1).Comment3 = TBRECust(1).Comment3
        TBRealExport(1).LandValue = RTRIM$(STR$(TBRECust(1).LandValue))
        TBRealExport(1).BldgValue = RTRIM$(STR$(TBRECust(1).BldgValue))
        TBRealExport(1).LandUseValue = RTRIM$(STR$(TBRECust(1).LandUseValue))
        PUT TBEFILE, (LOF(TBEFILE) / LEN(TBRealExport(1))) + 1, TBRealExport(1)
        END IF
        Cnt = Cnt + 1
      WEND
      CLOSE
      GOTO 10


60000  IF taxtype$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TBRTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBPath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN

60025 TBBalFile = FREEFILE
      IF taxtype$ = "R" THEN
      OPEN TBPath$ + "TBRTBAL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBPath$ + "TBPTBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF
60030 RETURN

61000 TBTranFile = FREEFILE
      OPEN TBPath$ + "TBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTranFile LEN = LEN(TBTrans(1))
61020 RETURN

62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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
63500 COLOR 10: LOCATE 1, 1: PRINT "Sys. Date: "; SYSDATE$; TAB(70); scren$: PRINT STRING$(79, "_"): RETURN
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
  OPEN TBPath$ + "TBRTNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
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
  OPEN TBPath$ + "TBPTNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
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

