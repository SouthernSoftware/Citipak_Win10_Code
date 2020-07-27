10 DEFINT A-Z
   
   DECLARE FUNCTION QPTrim$ (A$)
   DECLARE FUNCTION num2date$ (Num%)
   DECLARE FUNCTION Date2Num% (Dat$)
 
   '$INCLUDE: 'LCTAX.BI'

   'TBpath$ = "F:\"                    ' change to "F:\" Before compiling
   TBpath$ = "c:\qb45\lc\"                    ' change to "F:\" Before compiling
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
     REDIM TBTrans(1) AS TBTrans


     WIDTH LPRINT 132
     ON ERROR GOTO networkerror
     DEFDBL T
     DIM SHARED Fields$(40)
 
110 scren$ = "3.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
130 LOCATE 4, 1: COLOR 11

3000 CLS : scren$ = "3.03.00": GOSUB 63500: COLOR 11: LOCATE 1, 32: PRINT "Print Tax Bills"
     LOCATE 3, 1: COLOR 11
     PRINT "(R)egular Real Estate Bills or (N)ew Bldg Value Bills (R,N or <ESC>-Menu).. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN RUN "tbbillin"
     IF NOT (in$ = "R" OR in$ = "N") THEN 3000
     BillType$ = in$
     LOCATE 4, 1: COLOR 11
     PRINT "Enter the Tax Rate Per $100 (Ex. .50 for 50 cents per $100): "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     taxrate = VAL(in$)
     TaxType$ = "R"
     GOSUB 60000

3001 LOCATE 5, 1: COLOR 11
     PRINT "Print (A)ll or (I)ndividual Bills or (ESC) to End .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN RUN "tbbillin"
     IF NOT (in$ = "A" OR in$ = "I") THEN 3001
     IF in$ = "A" THEN
       GOSUB taxalign
       OPEN TBpath$ + "TBREname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #9 LEN = 24
       FIELD 9, 20 AS lastnam$, 4 AS Mrec$: EndofFile = LOF(9) / 24
       LastRecord = LOF(9) / 24
       OPEN "O", 10, "RTAXBILL.PRN": WIDTH #10, 255
       
   FOR Cnt! = 1 TO 20
    GET 9, Cnt!
    GET TBFile, CVS(Mrec$), TBRECust(1)
    AcctNumber! = CVS(Mrec$)
    IF BillType$ = "R" THEN GOSUB PrintRealBill
    IF BillType$ = "N" THEN GOSUB PrintNBBill
   NEXT Cnt!
   CLOSE
   RUN "tbbillin"
 ELSE
   COLOR 11
   LOCATE 15, 1: PRINT "Individual Printing"
   GOSUB taxalign
mainbody:
   LOCATE 17, 1: PRINT "Enter Customer Acct Number to Reprint (ESC-END) .. "; : fl = -7
   GOSUB 62000
   IF cf = 1 THEN PRINT #10, CHR$(12); : CLOSE : RUN "tbbillin"
     AcctNumber! = VAL(in$)
     
     GET TBFile, AcctNumber!, TBRECust(1)
     IF BillType$ = "R" THEN GOSUB PrintRealBill
     IF BillType$ = "N" THEN GOSUB PrintNBBill
     GOTO mainbody


 END IF

     
taxalign:
     LOCATE 17, 1: COLOR 11
     PRINT "Print Alignment (Y/N) or (ESC-ABORT PROCESS) .. "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO taxalign
     IF in$ = "N" THEN RETURN
     PRINT #10,
     PRINT #10, TAB(18); "XXXXXXXXXXXXXX"; TAB(57); "XX.XX"; TAB(87); "XXXXX"
     PRINT #10,
     PRINT #10,
     PRINT #10,
     PRINT #10, TAB(10); "XXX,XXX"; TAB(19); "XXX,XXX"; TAB(28); "X,XXX,XXX"; TAB(39); "XX,XXX.XX"; TAB(50); "XXX.XX"; TAB(58); "XX,XXX.XX"
     PRINT #10,
     PRINT #10, TAB(10); STRING$(35, "X"); " MAP ########"
     PRINT #10, TAB(10); STRING$(35, "X")
     PRINT #10, TAB(10); STRING$(35, "X")
     PRINT #10,
     PRINT #10,
     PRINT #10,
     PRINT #10, TAB(10); STRING$(35, "X")
     PRINT #10, TAB(10); STRING$(35, "X")
     PRINT #10, TAB(10); STRING$(35, "X")
     PRINT #10, TAB(10); STRING$(35, "X")
     PRINT #10,
     PRINT #10,
     PRINT #10,
     PRINT #10, TAB(55); "Acct # XXXXXXX"
     PRINT #10,
     GOTO taxalign


PrintRealBill:

     IF TBRECust(1).LBCurYrTaxAmount <= 0 THEN RETURN   'No Bill
     IF VAL(LEFT$(TBRECust(1).ClassCode, 1)) = 7 THEN RETURN  'Exempt Bill

     GOSUB GetDistrict

     Penalty! = TBRECust(1).LBCurYrTaxAmount * .1
     Penalty! = INT((Penalty! * 100) + .5) / 100


     TotalTaxAmount@ = TBRECust(1).LBCurYrTaxAmount + Penalty!
     TotalTaxAmount@ = INT((TotalTaxAmount@ * 100) + .5) / 100

     PRINT #10, "COUNTY OF LUNENBURG"
     PRINT #10, "804-696-3354 or 804-696-2516"
     PRINT #10, "BETSY N. LONG, TREASURER OF LUNENBURG COUNTY"
     PRINT #10, "11512 COURTHOUSE RD."
     PRINT #10, "LUNENBURG, VA 23952"
     PRINT #10, "Betsy N. Long"
     PRINT #10, "COUNTY OF LUNENBURG"
     PRINT #10, "11512 Courthouse Rd."
     PRINT #10, "Lunenburg, VA 23952"
     PRINT #10, "Credit cards accepted. Call 804-696-3354 for information."
     PRINT #10,
     PRINT #10, TAB(18); DistrictName$; TAB(57); USING "##.##"; taxrate;
      PRINT #10, TAB(87); TBRECust(1).LBCurYrBillNumber
     PRINT #10,
     PRINT #10,
     PRINT #10,
     PRINT #10, TAB(8); USING "######,#"; TBRECust(1).LandValue;
     PRINT #10, TAB(17); USING "#######,#"; TBRECust(1).BldgValue;
      PRINT #10, TAB(27); USING "#######,#"; TBRECust(1).LandValue + TBRECust(1).BldgValue;
       PRINT #10, TAB(38); USING "######,#.##"; TBRECust(1).LBCurYrTaxAmount;
       PRINT #10, TAB(49); USING "#####.##"; Penalty!;
        PRINT #10, TAB(58); USING "#####,#.##"; TotalTaxAmount@
     
     PRINT #10,
     PRINT #10, TAB(10); TBRECust(1).PropDesc1; " ** "; TBRECust(1).MapNumber
     PRINT #10, TAB(10); TBRECust(1).PropDesc2
     PRINT #10, TAB(10); TBRECust(1).PropDesc3
     PRINT #10,
     PRINT #10,
     PRINT #10,
     PRINT #10, TAB(10); TBRECust(1).CustName
     PRINT #10, TAB(10); TBRECust(1).Address1
     PRINT #10, TAB(10); TBRECust(1).Address2
     PRINT #10, TAB(10); TBRECust(1).Address3
     PRINT #10,
     PRINT #10,
     PRINT #10,
     PRINT #10, TAB(55); "Acct #"; AcctNumber!
     PRINT #10,
     RETURN



PrintNBBill:

     TotalTaxAmount@ = TBRECust(1).BBCurYrTaxAmount
     TotalTaxAmount@ = INT((TotalTaxAmount@ * 100) + .5) / 100


     IF TotalTaxAmount@ <= 0 THEN RETURN   'No Bill
     IF VAL(LEFT$(TBRECust(1).ClassCode, 1)) = 7 THEN RETURN  'Exempt Bill

     GOSUB GetDistrict

     Penalty! = 0





     PRINT #10,
     PRINT #10, TAB(18); DistrictName$; TAB(57); USING "##.##"; taxrate;
      PRINT #10, TAB(87); TBRECust(1).BBCurYrBillNumber
     PRINT #10,
     PRINT #10,
     PRINT #10,
     PRINT #10, TAB(17); USING "#######,#"; TBRECust(1).NewBldgValue;
      PRINT #10, TAB(27); USING "#######,#"; TBRECust(1).NewBldgValue;
       PRINT #10, TAB(38); USING "######,#.##"; TBRECust(1).BBCurYrTaxAmount;
       PRINT #10, TAB(49); USING "#####.##"; Penalty!;
        PRINT #10, TAB(58); USING "#####,#.##"; TotalTaxAmount@

     PRINT #10,
     PRINT #10, TAB(10); TBRECust(1).NewCom1; " ** "; TBRECust(1).MapNumber
     PRINT #10, TAB(10); TBRECust(1).NewCom2
     PRINT #10,
     PRINT #10,
     PRINT #10,
     PRINT #10,
     PRINT #10, TAB(10); TBRECust(1).CustName
     PRINT #10, TAB(10); TBRECust(1).Address1
     PRINT #10, TAB(10); TBRECust(1).Address2
     PRINT #10, TAB(10); TBRECust(1).Address3
     PRINT #10,
     PRINT #10,
     PRINT #10,
     PRINT #10, TAB(55); "Acct #"; AcctNumber!
     PRINT #10,
     RETURN

GetDistrict:
     DistrictName$ = ""
     IF VAL(TBRECust(1).District) = 1 THEN DistrictName$ = "BROWN'S STORE"
     IF VAL(TBRECust(1).District) = 2 THEN DistrictName$ = "COLUMBIAN GROVE"
     IF VAL(TBRECust(1).District) = 3 THEN DistrictName$ = "LEWISTON"
     IF VAL(TBRECust(1).District) = 4 THEN DistrictName$ = "LOCHLEVEN"
     IF VAL(TBRECust(1).District) = 5 THEN DistrictName$ = "PLEASANT GROVE"
     IF VAL(TBRECust(1).District) = 6 THEN DistrictName$ = "PLYMOUTH"
     IF VAL(TBRECust(1).District) = 7 THEN DistrictName$ = "REHOBOTH"
     IF VAL(TBRECust(1).District) = 9 THEN DistrictName$ = "TOWN OF KENBRIDGE"
     IF VAL(TBRECust(1).District) = 10 THEN DistrictName$ = "TOWN OF VICTORIA"
     RETURN

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

