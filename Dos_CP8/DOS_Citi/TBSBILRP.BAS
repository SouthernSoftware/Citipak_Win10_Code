10 DEFINT A-Z
   
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


   TBpath$ = "F:\"                    ' change to "F:\" Before compiling
  
   CLOSE
   OPEN "R", 1, TBpath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
   CurrentYear$ = year$: CLOSE 1
   ON ERROR GOTO networkerror


     REDIM TransLink!(500)
     REDIM LandValue#(99)
     REDIM BldgValue#(99)
     REDIM TaxAmt#(99)
     REDIM Customer!(99)

     DIM SHARED TBRECust(1) AS TBRECust
     REDIM TBPPCust(1) AS TBPPCust
     REDIM TBRECustBal(1) AS TBREBAL
     REDIM TBTrans(1) AS TBTrans


     WIDTH LPRINT 132
   
     DEFDBL T
     DIM SHARED Fields$(40)
 
110 scren$ = "3.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
130 LOCATE 4, 1: COLOR 11

3000 CLS : scren$ = "3.03.00": GOSUB 63500
     COLOR 11: LOCATE 1, 27
     PRINT "Print Supplemental Tax Bill Proof"
     LOCATE 3, 1: COLOR 11
     PRINT "Enter (R)eal or  (P)ersonal (R,P or <ESC>-Cancel) .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN CLOSE : RUN "tbsbilln"
     IF NOT (in$ = "R" OR in$ = "P") THEN 3000
     TaxType$ = in$
     IF TaxType$ = "P" THEN 4000

     COLOR 11
     PRINT "Now Processing the Report "
     GOSUB 60000
     OPEN TBpath$ + "TTSRname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #9 LEN = 24
     FIELD 9, 20 AS lastnam$, 4 AS Mrec$: EndofFile = LOF(9) / 24
     LastRecord = LOF(9) / 24

     OPEN "O", 10, "TBBILL.REP": WIDTH #10, 255
     GOSUB Heading
       
   FOR Cnt! = 1 TO LastRecord
    GET 9, Cnt!
    GET TBFile, CVS(Mrec$), TBRECust(1)
    AcctNumber! = CVS(Mrec$)
    GOSUB PrintLine
   NEXT Cnt!
   PRINT #10, STRING$(79, "-")
   PRINT #10, "Total Amount Billed: "; USING "$$######,#.##"; TotalAmt#
   PRINT #10, "Total Customers    : "; USING "######,#"; Cust!


   CLOSE

   LPTPort% = 1
   EntryPoint% = 1
   ReportFile$ = "TBBILL.REP"
   Header$ = "Billing Proof"
   PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
   KILL ReportFile$

   RUN "tbsbilln"
 

     


PrintLine:

     IF TBRECust(1).LBCurYrTaxAmount <= 0 THEN RETURN   'No Bill
     


     IF lc >= 58 THEN
      PRINT #10, CHR$(12); : GOSUB Heading
     END IF
     PRINT #10, USING "#####"; Cnt!;
      PRINT #10, TAB(7); USING "#####"; AcctNumber!;
       PRINT #10, TAB(15); TBRECust(1).CustName;
       PRINT #10, TAB(50); USING "######,#.##"; TBRECust(1).LBCurYrTaxAmount;
        PRINT #10, TAB(70); TBRECust(1).LBCurYrBillNumber
        lc = lc + 1
        TotalAmt# = TotalAmt# + TBRECust(1).LBCurYrTaxAmount
        Cust! = Cust! + 1
     RETURN

Heading:
Page = Page + 1
PRINT #10, TAB(32); "Tax Bill Listing Proof"
PRINT #10, "Report Date: "; DATE$; TAB(65); "Page #"; Page
PRINT #10, ""
PRINT #10, " TR#"; TAB(8); "Acct#"; TAB(15); "Customer Name"; TAB(53); "Bill Amt"; TAB(70); "Bill #"
PRINT #10, STRING$(79, "=")
lc = 5
RETURN


4000 COLOR 11
     PRINT "Now Processing the Report "
     GOSUB 60000
     OPEN TBpath$ + "TTSPname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #9 LEN = 24
     FIELD 9, 20 AS lastnam$, 4 AS Mrec$: EndofFile = LOF(9) / 24
     LastRecord = LOF(9) / 24

     OPEN "O", 10, "TBPPBILL.REP": WIDTH #10, 255
     GOSUB PersHeading

   FOR Cnt! = 1 TO LastRecord
    GET 9, Cnt!
    GET TBFile, CVS(Mrec$), TBPPCust(1)
    AcctNumber! = CVS(Mrec$)
     IF VAL(TBPPCust(1).LBCurYrBillNumber) > 0 THEN
    
     TaxAmount# = TBPPCust(1).LBCurYrPersPropTaxAmount + TBPPCust(1).LBCurYrFarmEquipTaxAmount + TBPPCust(1).LBCurYrMachToolsTaxAmount + TBPPCust(1).LBCurYrMerchCapTaxAmount + TBPPCust(1).LBCurYrMobileHomeTaxAmount + TBPPCust(1).LBLateFeeAmount


     IF lc >= 58 THEN
      PRINT #10, CHR$(12); : GOSUB PersHeading
     END IF

     PRINT #10, USING "#####"; Cnt!;
      PRINT #10, TAB(7); USING "#####"; AcctNumber!;
      PRINT #10, TAB(15); TBPPCust(1).CustName;
       PRINT #10, TAB(50); USING "######,#.##"; TaxAmount#;
        PRINT #10, TAB(70); TBPPCust(1).LBCurYrBillNumber
        PRINT #10, USING "#####,#.##"; TBPPCust(1).LBCurYrPersPropTaxAmount;
        PRINT #10, TAB(12); USING "#####,#.##"; TBPPCust(1).LBCurYrFarmEquipTaxAmount;
        PRINT #10, TAB(24); USING "#####,#.##"; TBPPCust(1).LBCurYrMachToolsTaxAmount;
        PRINT #10, TAB(36); USING "#####,#.##"; TBPPCust(1).LBCurYrMerchCapTaxAmount;
        PRINT #10, TAB(48); USING "#####,#.##"; TBPPCust(1).LBCurYrMobileHomeTaxAmount;
        PRINT #10, TAB(62); USING "#####,#.##"; TBPPCust(1).LBLateFeeAmount
        IF TBPPCust(1).LBPPTRAVALUE > 0 THEN
         PRINT #10, "PPTRA Value: "; USING "######,#"; TBPPCust(1).LBPPTRAVALUE;
         PRINT #10, "  PPTRA Dis: "; USING "$$#####,#.##"; TBPPCust(1).LBPPTRADiscnt;
          PRINT #10, "    NET BILLED: "; USING "$$######,#.##"; TaxAmount# - TBPPCust(1).LBPPTRADiscnt
         lc = lc + 1
        END IF
        PRINT #10, STRING$(79, "-")
        lc = lc + 3
        TotalAmt# = TotalAmt# + TaxAmount#
        TotalPPTRAValue# = TotalPPTRAValue# + TBPPCust(1).LBPPTRAVALUE
        TotalPPTRADis# = TotalPPTRADis# + TBPPCust(1).LBPPTRADiscnt
        TotalPP# = TotalPP# + TBPPCust(1).PersProp
        Cust! = Cust! + 1
   END IF

   NEXT Cnt!
   PRINT #10, STRING$(79, "-")
   PRINT #10, "Total Amount Billed Before PPTRA: "; USING "$$######,#.##"; TotalAmt#
   PRINT #10, "    Less PPTRA Discounts Allowed: "; USING "$$######,#.##"; TotalPPTRADis#
   PRINT #10, "   Total Amount Billed Customers: "; USING "$$######,#.##"; TotalAmt# - TotalPPTRADis#
   PRINT #10, "          Total Customers Billed: "; USING "######,#"; Cust!
   PRINT #10,
   PRINT #10, "PPTRA Information"
   PRINT #10, "Total PPTRA Value   : "; USING "$$#######,#"; TotalPPTRAValue#
   PRINT #10, "Total PPTRA Discount: "; USING "$$######,#.##"; TotalPPTRADis#



   CLOSE

   LPTPort% = 1
   EntryPoint% = 1
   ReportFile$ = "TBPPBILL.REP"
   Header$ = "Billing Proof"
   PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
   KILL ReportFile$

   RUN "tbsbilln"


PersHeading:
Page = Page + 1
PRINT #10, TAB(27); "Tax Bill Listing Proof (Personal Property)"
PRINT #10, "Report Date: "; DATE$; TAB(65); "Page #"; Page
PRINT #10, ""
PRINT #10, " TR#"; TAB(8); "Acct#"; TAB(15); "Customer Name"; TAB(53); "Bill Amt"; TAB(70); "Bill #"
PRINT #10, "Personal"; TAB(12); "Farm Equip"; TAB(24); "Mach & Tools"; TAB(36); "Merch Cap"; TAB(48); "Mobile Home"; TAB(62); "Late Fee"
PRINT #10, STRING$(79, "=")
lc = 6
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
         OPEN TBpath$ + "TTRSCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBpath$ + "TTPSCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN


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

