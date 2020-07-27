10 DEFINT A-Z
DECLARE FUNCTION Date2Num% (Dat$)
CLEAR , , 4000
'$INCLUDE: 'lcTAX.BI'

        REDIM TBCust(1) AS TBPPCust
        REDIM TBOldCust(1) AS TBPPCustomer

        TBFile = FREEFILE
        OPEN "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBCust(1))

        OldTBFile = FREEFILE
        OPEN "TBPPCUS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #OldTBFile LEN = LEN(TBOldCust(1))

        FOR x = 1 TO LOF(OldTBFile) / LEN(TBOldCust(1))
        GET OldTBFile, x, TBOldCust(1)
        TBCust(1).CustName = TBOldCust(1).CustName
        TBCust(1).Address1 = TBOldCust(1).Address1
        TBCust(1).Address2 = TBOldCust(1).Address2
        TBCust(1).City = TBOldCust(1).City
        TBCust(1).State = TBOldCust(1).State
        TBCust(1).Zip = TBOldCust(1).Zip
        TBCust(1).SocSec = TBOldCust(1).SocSec
        TBCust(1).SocSec2 = TBOldCust(1).SocSec2
        TBCust(1).Short = TBOldCust(1).Short
        TBCust(1).PersProp = VAL(TBOldCust(1).PersProp)
        TBCust(1).MachTools = VAL(TBOldCust(1).MachTools)
        TBCust(1).MobileHome = VAL(TBOldCust(1).MobileHome)
        TBCust(1).MerchCap = VAL(TBOldCust(1).MerchCap)
        TBCust(1).FarmEquip = VAL(TBOldCust(1).FarmEquip)
        TBCust(1).Business = TBOldCust(1).Business
        TBCust(1).LastChangeDate = Date2Num%(MID$(TBOldCust(1).LastChangeDate, 5, 4) + LEFT$(TBOldCust(1).LastChangeDate, 4))
        TBCust(1).District = TBOldCust(1).District
        TBCust(1).LateFile = TBOldCust(1).LateFile      'T if True
        TBCust(1).CurYrBillNumber = TBOldCust(1).BillNumb
        TBCust(1).= VAL(TBOldCust(1).PersPropTax)
        TBCust(1).CurYrFarmEquipTax = VAL(TBOldCust(1).FarmEquipTax)
        TBCust(1).CurYrMachToolsTax = VAL(TBOldCust(1).MachToolsTax)
        TBCust(1).CurYrMerchCapTax = VAL(TBOldCust(1).MerchCapTax)
        TBCust(1).CurYrMobileHomeTax = VAL(TBOldCust(1).MobileHomeTax)
        TBCust(1).CurYrLateFeeTax = VAL(TBOldCust(1).LateFeeTax)
        TBCust(1).SocSec3 = TBOldCust(1).SocSec3
        TBCust(1).FirstTrans = 0
        TBCust(1).LastTrans = 0
        TBCust(1).Pad = ""

        PUT TBFile, VAL(TBOldCust(1).Account), TBCust(1)
        NEXT x
        CLOSE
        STOP

62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 Z$ = INKEY$: IF Z$ = "" THEN 62030
62040 IF Z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF Z$ = CHR$(27) OR Z$ = CHR$(21) THEN cf = 1: RETURN
62065 IF ASC(Z$) = 0 THEN 62030
62070 IF LEN(Z$) > 1 THEN cf = ASC(RIGHT$(Z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
62080 IF Z$ = " " THEN 62130 ELSE IF Z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 62000
62090 IF Z$ <> CHR$(8) THEN 62130 ELSE IF ZL = 0 THEN 62030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 62110 ELSE IF Z$ = "," THEN 62120
62100 IF Z$ = "." THEN ZD = 0: GOTO 62110 ELSE IF Z$ = "+" OR Z$ = "-" THEN ZS = 0
62110 in$ = LEFT$(in$, LEN(in$) - 1)
62120 ZL = ZL - 1: GOTO 62030
62130 IF ABS(fl) = ZL THEN 62030 ELSE IF fl > 0 THEN IF Z$ >= " " AND Z$ <= "z" THEN 62180
62140 IF Z$ = "." AND INSTR(in$, ".") = 0 THEN GOTO 62180
62150 IF Z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (Z$ = "-" OR Z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF Z$ < "0" OR Z$ > "9" THEN 62030
62180 PRINT Z$; : in$ = in$ + Z$: ZL = ZL + 1
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

