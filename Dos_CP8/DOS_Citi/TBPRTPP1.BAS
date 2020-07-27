10 DEFINT A-Z
   

   DECLARE FUNCTION QPTrim$ (a$)
   DECLARE FUNCTION num2date$ (Num%)
   DECLARE FUNCTION Date2Num% (Dat$)
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


   WIDTH LPRINT 136
   ON ERROR GOTO networkerror


   TBpath$ = "F:\"                    ' change to "F:\" Before compiling

  CLOSE
   OPEN "R", 1, TBpath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
   CurrentYear$ = year$: CLOSE 1

    DIM SHARED Fields$(40)

    REDIM TBCust(1)  AS TBPPCust
    REDIM TBVeh(1) AS PPVehType

110 scren$ = "3.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
130 LOCATE 3, 1: COLOR 11
    PRINT "Print Personal Property Bills For"
    PRINT "1- Lunenburg County  2- Kenbridge  3-Victoria or (ESC-Cancel) .. "; : fl = -1
    GOSUB 62000
    IF cf = 1 THEN CLOSE : RUN "tbbillin"
    IF VAL(in$) < 1 OR VAL(in$) > 3 THEN 120
    IF VAL(in$) = 1 THEN 3000
    IF VAL(in$) = 2 THEN 4000
    IF VAL(in$) = 3 THEN 5000


3000 CLS : scren$ = "3.03.00": GOSUB 63500: COLOR 11
     LOCATE 1, 32: PRINT "Print Tax Bills"
     LOCATE 3, 1: COLOR 11
     PRINT "Enter the Pers. Prop. Tax Rate Per $100 (Ex. .50 for 50 cents / $100): "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     PPTaxRate! = VAL(in$)
     LOCATE 4, 1: COLOR 11
     PRINT "Enter the Machinery and Tools Tax Rate Per $100 (ESC-End) : "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     MTtaxrate! = VAL(in$)
     LOCATE 5, 1: COLOR 11
     PRINT "Enter the Merchants Capital Tax Rate Per $100 (ESC-End)   : "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     MCtaxrate! = VAL(in$)
     LOCATE 6, 1: COLOR 11
     PRINT "Enter the Mobile Home Tax Rate Per $100 (ESC-End)         : "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     MHtaxrate! = VAL(in$)
     LOCATE 7, 1: COLOR 11
     PRINT "Enter the Farm Equipment Tax Rate Per $100 (ESC-End)      : "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     FEtaxrate! = VAL(in$)
TaxYearEnter:
     LOCATE 8, 1: COLOR 11
     PRINT "Enter the Tax Year  (ESC-End) (yyyy) : "; : fl = -4
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     IF LEN(in$) <> 4 THEN GOTO TaxYearEnter
     TaxYear$ = in$

    'PPTRA Percent
     IF VAL(TaxYear$) = 1998 THEN Perc! = 12.5
     IF VAL(TaxYear$) = 1999 THEN Perc! = 27.5
     IF VAL(TaxYear$) = 2000 THEN Perc! = 47.5
     IF VAL(TaxYear$) = 2001 THEN Perc! = 70
     IF VAL(TaxYear$) = 2002 THEN Perc! = 100
PenDate:
    LOCATE 9, 1: COLOR 11
    PRINT "Enter the Penalty Date (mm/dd/yyyy) .. "; : fl = 10
    GOSUB 62000
    IF LEN(in$) <> 10 THEN GOTO PenDate
    PenDate$ = in$
IntDate:
    LOCATE 10, 1: COLOR 11
    PRINT "Enter the Date Interest to Begin (mm/dd/yyyy) .. "; : fl = 10
    GOSUB 62000
    IF LEN(in$) <> 10 THEN GOTO IntDate
    IntDate$ = in$


3001 LOCATE 12, 1: COLOR 11
     PRINT "Print (A)ll or (I)ndividual Bills or (ESC) to End .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN RUN "tbbillin"
     IF NOT (in$ = "A" OR in$ = "I") THEN 3001
     IF in$ = "A" THEN
3035  ' GET A SORTED INDEX FIRST
      LOCATE 13, 1
      COLOR 11
      PRINT "Enter the TR# From the Proof Report to Start (0=All:<ESC>-Cancel) .. "; : fl = -5
      GOSUB 62000
      IF cf = 1 THEN CLOSE : GOTO 10
      BegTr# = VAL(in$)



      TBFile = FREEFILE
      OPEN TBpath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBCust(1))
      OPEN TBpath$ + "TBPTname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #9 LEN = 24
      FIELD 9, 20 AS LastNam$, 4 AS Mrec$: EndofFile = LOF(9) / 24
      TBVFile = FREEFILE
      OPEN TBpath$ + "TBTVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))


       ReportFile$ = "TAXPBILL.PRN"
       OPEN "O", 10, "TAXPBILL.PRN"
       WIDTH #10, 255

    LastRecord = LOF(9) / 24
    Count = 1
    WHILE Count <= LastRecord
     GET 9, Count
     IF BegTr# = 0 OR Count >= BegTr# THEN


     LOCATE 15, 1: COLOR 11: PRINT "Processing Acct # "; Count
     GET TBFile, CVS(Mrec$), TBCust(1)

     acctnumber = CVS(Mrec$)
     
     GOSUB PrintBill
     END IF
     Count = Count + 1
     WEND
     CLOSE
'Set Report Defaults
EndAllBills:
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TAXPBILL.PRN"
      Header$ = "Personal Property Tickets"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      CLOSE
      KILL ReportFile$
      RUN "tbbillin"

    ELSE

      COLOR 11
      LOCATE 15, 1: PRINT "Individual Printing"


mainbody:
     'Open Main Tax Files


      TBFile = FREEFILE
      OPEN TBpath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBCust(1))
      TBVFile = FREEFILE
      OPEN TBpath$ + "TBTVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))

     CLS : scren$ = "3.03.00": GOSUB 63500: COLOR 11
     LOCATE 1, 32: PRINT "Print Tax Bills"
     LOCATE 3, 1: COLOR 11
    LOCATE 17, 1: PRINT "Enter Customer Acct Number to Reprint (ESC-END) .. "; : fl = -7
    GOSUB 62000
  IF cf = 1 THEN CLOSE : RUN "tbbillin"
      acctnumber = VAL(in$)
      GET TBFile, acctnumber, TBCust(1)
       ReportFile$ = "TAXPBILL.PRN"
       OPEN "O", 10, "TAXPBILL.PRN"
       WIDTH #10, 255
       GOSUB PrintBill
       CLOSE
       LPTPort% = 1
       EntryPoint% = 1
       ReportFile$ = "TAXPBILL.PRN"
       Header$ = "Personal Property Tickets"
       PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
       CLOSE
       KILL ReportFile$
       GOTO mainbody
  END IF
  


PrintBill:
     IF VAL(TBCust(1).LBCurYrBillNumber) < 1 THEN RETURN     'No Bill
     TotalTaxAmount@ = (TBCust(1).LBCurYrPersPropTaxAmount + TBCust(1).LBCurYrMachToolsTaxAmount + TBCust(1).LBCurYrMerchCapTaxAmount + TBCust(1).LBCurYrMobileHomeTaxAmount + TBCust(1).LBCurYrFarmEquipTaxAmount - TBCust(1).LBPPTRADiscnt)
     TotalTaxAmount@ = INT((TotalTaxAmount@ * 100) + .50001) / 100
     TotalTaxAmount@ = TotalTaxAmount@ + TBCust(1).LBLateFeeAmount
     TotalTaxAmount@ = INT((TotalTaxAmount@ * 100) + .50001) / 100
     
     Penalty@ = TotalTaxAmount@ * .1
     Penalty@ = INT((Penalty@ * 100) + .50001) / 100

     IF TotalTaxAmount@ < .01 AND TBCust(1).LBPPTRADiscnt = 0 THEN RETURN      'No Bill
     

     ' Print the Bill Here
     PRINT #10, "PERSONAL PROPERTY BILL START"
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
     PRINT #10, "Tax Year: "; TaxYear$
     PRINT #10, TAB(85); TBCust(1).LBCurYrBillNumber
     PRINT #10,
     PRINT #10, "Personal Property ";
       PRINT #10, TAB(45); USING "########"; TBCust(1).PersProp;
        PRINT #10, TAB(55); USING "#.##"; PPTaxRate!;
        PRINT #10, TAB(61); USING "######.##"; TBCust(1).LBCurYrPersPropTaxAmount;
        PRINT #10, TAB(71); USING "######.##"; TBCust(1).LBPPTRADiscnt;
        PRINT #10, TAB(81); USING "######.##"; TBCust(1).LBCurYrPersPropTaxAmount - TBCust(1).LBPPTRADiscnt

     PRINT #10, "Machinery & Tools";
      PRINT #10, TAB(45); USING "########"; TBCust(1).MachTools;
       PRINT #10, TAB(55); USING "#.##"; MTtaxrate!;
       PRINT #10, TAB(61); USING "######.##"; TBCust(1).LBCurYrMachToolsTaxAmount;
       PRINT #10, TAB(81); USING "######.##"; TBCust(1).LBCurYrMachToolsTaxAmount

     PRINT #10, "Merchant Capital";
      PRINT #10, TAB(45); USING "########"; TBCust(1).MerchCap;
       PRINT #10, TAB(55); USING "#.##"; MCtaxrate!;
       PRINT #10, TAB(61); USING "######.##"; TBCust(1).LBCurYrMerchCapTaxAmount;
       PRINT #10, TAB(81); USING "######.##"; TBCust(1).LBCurYrMerchCapTaxAmount

    PRINT #10, "Mobile Home ";
      PRINT #10, TAB(45); USING "########"; TBCust(1).MobileHome;
      PRINT #10, TAB(55); USING "#.##"; MHtaxrate!;
       PRINT #10, TAB(61); USING "######.##"; TBCust(1).LBCurYrMobileHomeTaxAmount;
       PRINT #10, TAB(81); USING "######.##"; TBCust(1).LBCurYrMobileHomeTaxAmount

     PRINT #10, "Farm Equipment ";
      PRINT #10, TAB(45); USING "########"; TBCust(1).FarmEquip;
      PRINT #10, TAB(55); USING "#.##"; FEtaxrate!;
       PRINT #10, TAB(61); USING "######.##"; TBCust(1).LBCurYrFarmEquipTaxAmount;
       PRINT #10, TAB(81); USING "######.##"; TBCust(1).LBCurYrFarmEquipTaxAmount
     PRINT #10, "Late Fee "; TAB(61); USING "######.##"; TBCust(1).LBLateFeeAmount;
      PRINT #10, TAB(81); USING "######.##"; TBCust(1).LBLateFeeAmount
     

     PRINT #10, "Vehicle Listing"
      CarCount = 0                      'Set Car Count to Zero
      VehRecord! = TBCust(1).FirstVeh
      WHILE VehRecord! <> 0
        GET TBVFile, VehRecord!, TBVeh(1)
        'Programming Change 10-5-01 to let all vehicles print not just veh qualify
        'IF TBVeh(1).VehTyp <> "D" AND TBVeh(1).VehQ = "Y" AND TBVeh(1).VehTaxYr = VAL(TaxYear$) THEN
         IF TBVeh(1).VehTyp <> "D" AND TBVeh(1).VehTaxYr = VAL(TaxYear$) THEN

         'Figure PPTRADis# Here
          PPTRADis# = 0
        IF TBVeh(1).VehQ = "Y" THEN
          IF TBVeh(1).VehValue > 20000 THEN
           PPTRAVal# = 20000
           ELSE
           PPTRAVal# = TBVeh(1).VehValue
          END IF
          IF PPTRAVal# <= 1000 THEN
           PPTRADis# = ((PPTRAVal# / 100) * PPTaxRate!)
           ELSE
           PPTRADis# = (((PPTRAVal# / 100) * PPTaxRate!) * (Perc! / 100))
          END IF
         END IF
          PPTRADis# = INT((PPTRADis# * 100) + .50001) / 100

         TaxAmt# = (TBVeh(1).VehValue / 100) * PPTaxRate!
         MakeModel$ = RTRIM$(TBVeh(1).VehYear) + " " + RTRIM$(TBVeh(1).VehMake) + " " + RTRIM$(TBVeh(1).VehModel)

         PRINT #10, TAB(1); LEFT$(MakeModel$, 20);
         PRINT #10, TAB(23); LEFT$(TBVeh(1).VehVin, 20);
          PRINT #10, TAB(45); USING "########"; TBVeh(1).VehValue;
          PRINT #10, TAB(55); USING "#.##"; PPTaxRate!;
          PRINT #10, TAB(61); USING "######.##"; TaxAmt#;
          PRINT #10, TAB(71); USING "######.##"; PPTRADis#
          CarCount = CarCount + 1
         IF CarCount >= 4 AND TBVeh(1).VehNext <> 0 THEN
                BlankLines = 16 - CarCount
                FOR LCnt = 1 TO BlankLines: PRINT #10, "": NEXT LCnt
                PRINT #10, TAB(14); "Acct #"; acctnumber
                PRINT #10, TAB(14); TBCust(1).CustName
                PRINT #10, TAB(14); TBCust(1).Address1;
                                PRINT #10, TAB(40); USING "######.##"; TotalTaxAmount@ + TBCust(1).LBPPTRADiscnt;
                                PRINT #10, TAB(50); USING "######.##"; TBCust(1).LBPPTRADiscnt;
                                PRINT #10, TAB(60); USING "######.##"; TotalTaxAmount@
                PRINT #10, TAB(14); TBCust(1).Address2;
                                PRINT #10, TAB(60); USING "######.##"; TotalTaxAmount@
                PRINT #10, TAB(14); RTRIM$(TBCust(1).City); ",  "; TBCust(1).State; "  "; TBCust(1).Zip
                PRINT #10, TAB(60); "Penalty Date: "; PenDate$
                PRINT #10, TAB(60); "Interest Date: "; IntDate$
                PRINT #10, "End of Bill - To Continue Same Account"
                GOSUB AnotherBill
                CarCount = 0
         END IF
        END IF
        VehRecord! = TBVeh(1).VehNext
      WEND
      BlankLines = 16 - CarCount
      FOR LCnt = 1 TO BlankLines: PRINT #10, "": NEXT LCnt
      PRINT #10, TAB(14); "Acct #"; acctnumber
      PRINT #10, TAB(14); TBCust(1).CustName
      PRINT #10, TAB(14); TBCust(1).Address1;
                PRINT #10, TAB(40); USING "######.##"; TotalTaxAmount@ + TBCust(1).LBPPTRADiscnt;
                PRINT #10, TAB(50); USING "######.##"; TBCust(1).LBPPTRADiscnt;
                PRINT #10, TAB(60); USING "######.##"; TotalTaxAmount@
      PRINT #10, TAB(14); TBCust(1).Address2;
                PRINT #10, TAB(60); USING "######.##"; TotalTaxAmount@
      PRINT #10, TAB(14); RTRIM$(TBCust(1).City); ",  "; TBCust(1).State; "  "; TBCust(1).Zip
      PRINT #10, TAB(60); "Penalty Date: "; PenDate$
      PRINT #10, TAB(60); "Interest Date: "; IntDate$
      PRINT #10, "End of Bill"
      RETURN


AnotherBill:
     ' Print the Bill Here
     PRINT #10, "PERSONAL PROPERTY BILL START"
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
     PRINT #10, TAB(85); TBCust(1).LBCurYrBillNumber
     PRINT #10, "Tax Year: "; TaxYear$
     PRINT #10, "Personal Property "
     PRINT #10, "Machinery & Tools "
     PRINT #10, "Merchant Capital"
     PRINT #10, "Mobile Home"
     PRINT #10, "Farm Equipment "
     PRINT #10, "Late Fee  "
     PRINT #10, "Vehicle Listing Continued"
     RETURN




4000 CLS : scren$ = "3.03.00": GOSUB 63500: COLOR 11
     LOCATE 1, 27: PRINT "Print Tax Bills - Kenbridge"
     LOCATE 3, 1: COLOR 11
     PRINT "Enter the Pers. Prop. Tax Rate Per $100 (Ex. .50 for 50 cents / $100): "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     PPTaxRate! = VAL(in$)
     LOCATE 4, 1: COLOR 11
     PRINT "Enter the Machinery and Tools Tax Rate Per $100 (ESC-End) : "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     MTtaxrate! = VAL(in$)
     LOCATE 5, 1: COLOR 11
     PRINT "Enter the Merchants Capital Tax Rate Per $100 (ESC-End)   : "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     MCtaxrate! = VAL(in$)
     LOCATE 6, 1: COLOR 11
     PRINT "Enter the Mobile Home Tax Rate Per $100 (ESC-End)         : "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     MHtaxrate! = VAL(in$)
     LOCATE 7, 1: COLOR 11
     PRINT "Enter the Farm Equipment Tax Rate Per $100 (ESC-End)      : "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     FEtaxrate! = VAL(in$)


TaxYearEnter4:

     LOCATE 8, 1: COLOR 11
     PRINT "Enter the Tax Year  (ESC-End) (yyyy) : "; : fl = -4
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     IF LEN(in$) <> 4 THEN GOTO TaxYearEnter4
     TaxYear$ = in$
     
   'PPTRA Percent
     IF VAL(TaxYear$) = 1998 THEN Perc! = 12.5
     IF VAL(TaxYear$) = 1999 THEN Perc! = 27.5
     IF VAL(TaxYear$) = 2000 THEN Perc! = 47.5
     IF VAL(TaxYear$) = 2001 THEN Perc! = 70
     IF VAL(TaxYear$) = 2002 THEN Perc! = 100

PenDate4:
    LOCATE 9, 1: COLOR 11
    PRINT "Enter the Penalty Date (mm/dd/yyyy) .. "; : fl = 10
    GOSUB 62000
    IF LEN(in$) <> 10 THEN GOTO PenDate4
    PenDate$ = in$
IntDate4:
    LOCATE 10, 1: COLOR 11
    PRINT "Enter the Date Interest to Begin (mm/dd/yyyy) .. "; : fl = 10
    GOSUB 62000
    IF LEN(in$) <> 10 THEN GOTO IntDate4
    IntDate$ = in$

4001 LOCATE 12, 1: COLOR 11
     PRINT "Print (A)ll or (I)ndividual Bills or (ESC) to End .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN RUN "tbbillin"
     IF NOT (in$ = "A" OR in$ = "I") THEN 4001
     IF in$ = "A" THEN
       
4035  ' GET A SORTED INDEX FIRST
      TBFile = FREEFILE
      OPEN TBpath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBCust(1))
      OPEN TBpath$ + "TBPTname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #9 LEN = 24
      FIELD 9, 20 AS LastNam$, 4 AS Mrec$: EndofFile = LOF(9) / 24
      TBVFile = FREEFILE
      OPEN TBpath$ + "TBTVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))


       ReportFile$ = "TAXPBILL.PRN"
       OPEN "O", 10, "TAXPBILL.PRN"
       WIDTH #10, 255


      LastRecord = LOF(9) / 24
       Count = 1
    WHILE Count <= LastRecord
     GET 9, Count

     GET TBFile, CVS(Mrec$), TBCust(1)

     acctnumber = CVS(Mrec$)

     GOSUB PrintBillKenbridge
     Count = Count + 1
     WEND
     CLOSE

'Set Report Defaults
KBEndAllBills:
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TAXPBILL.PRN"
      Header$ = "Personal Property Tickets"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      CLOSE
      KILL ReportFile$
      RUN "tbbillin"

    ELSE

      COLOR 11
      LOCATE 15, 1: PRINT "Individual Printing"
      

KBmainbody:
     'Open Main Tax Files
      TBFile = FREEFILE
      OPEN TBpath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBCust(1))
      TBVFile = FREEFILE
      OPEN TBpath$ + "TBTVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))

     CLS : scren$ = "3.03.00": GOSUB 63500: COLOR 11
     LOCATE 1, 27: PRINT "Print Tax Bills-Kenbridge"
     LOCATE 3, 1: COLOR 11
    LOCATE 17, 1: PRINT "Enter Customer Acct Number to Reprint (ESC-END) .. "; : fl = -7
    GOSUB 62000
  IF cf = 1 THEN CLOSE : RUN "tbbillin"
      acctnumber = VAL(in$)
      GET TBFile, acctnumber, TBCust(1)
       ReportFile$ = "TAXPBILL.PRN"
       OPEN "O", 10, "TAXPBILL.PRN"
       WIDTH #10, 255
       GOSUB PrintBillKenbridge
       CLOSE
       LPTPort% = 1
       EntryPoint% = 1
       ReportFile$ = "TAXPBILL.PRN"
       Header$ = "Personal Property Tickets"
       PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
       CLOSE
       KILL ReportFile$
       GOTO KBmainbody
  END IF



PrintBillKenbridge:
     IF VAL(TBCust(1).TBCurYrBillNumber) < 1 THEN RETURN     'No Bill
     IF VAL(TBCust(1).District) <> 9 THEN RETURN'Must be 9 for Kenbridge
     TotalTaxAmount@ = (TBCust(1).TBCurYrPersPropTaxAmount + TBCust(1).TBCurYrMachToolsTaxAmount + TBCust(1).TBCurYrMerchCapTaxAmount + TBCust(1).TBCurYrMobileHomeTaxAmount + TBCust(1).TBCurYrFarmEquipTaxAmount - TBCust(1).TBPPTRADiscnt)
     TotalTaxAmount@ = INT((TotalTaxAmount@ * 100) + .50001) / 100
     TotalTaxAmount@ = TotalTaxAmount@ + TBCust(1).TBLateFeeAmount
     TotalTaxAmount@ = INT((TotalTaxAmount@ * 100) + .50001) / 100

     Penalty@ = TotalTaxAmount@ * .1
     Penalty@ = INT((Penalty@ * 100) + .50001) / 100

     IF TotalTaxAmount@ < .01 THEN RETURN       'No Bill



     PRINT #10, "PERSONAL PROPERTY BILL START"
     PRINT #10, "TOWN OF KENBRIDGE"
     PRINT #10, "804-676-2452"
     PRINT #10, "TREASURER, TOWN OF KENBRIDGE"
     PRINT #10, "100 HIGH STREET   PO BOX 478"
     PRINT #10, "KENBRIDGE, VA 23944"
     PRINT #10, ""
     PRINT #10, "TOWN OF KENBRIDGE"
     PRINT #10, "100 High Street   PO Box 478"
     PRINT #10, "Kenbridge, VA 23944"
     PRINT #10, "Call 804-676-2452 for information."
     PRINT #10, "Tax Year: "; TaxYear$
     
     PRINT #10, TAB(85); TBCust(1).TBCurYrBillNumber
     PRINT #10,
     PRINT #10, "Personal Property ";
       PRINT #10, TAB(45); USING "########"; TBCust(1).PersProp;
       PRINT #10, TAB(55); USING "#.##"; PPTaxRate!;
       PRINT #10, TAB(61); USING "######.##"; TBCust(1).TBCurYrPersPropTaxAmount;
       PRINT #10, TAB(71); USING "######.##"; TBCust(1).TBPPTRADiscnt;
       PRINT #10, TAB(81); USING "######.##"; TBCust(1).TBCurYrPersPropTaxAmount - TBCust(1).TBPPTRADiscnt

    PRINT #10, "Machinery & Tools";
      PRINT #10, TAB(45); USING "########"; 0;
      PRINT #10, TAB(55); USING "#.##"; MTtaxrate!;
       PRINT #10, TAB(61); USING "######.##"; TBCust(1).TBCurYrMachToolsTaxAmount;
       PRINT #10, TAB(81); USING "######.##"; TBCust(1).TBCurYrMachToolsTaxAmount

     PRINT #10, "Merchant Capital"; TAB(45); USING "########"; 0;
      PRINT #10, TAB(55); USING "#.##"; MCtaxrate!;
       PRINT #10, TAB(61); USING "######.##"; TBCust(1).TBCurYrMerchCapTaxAmount;
       PRINT #10, TAB(81); USING "######.##"; TBCust(1).TBCurYrMerchCapTaxAmount

    PRINT #10, "Mobile Home "; TAB(45); USING "########"; TBCust(1).MobileHome;
      PRINT #10, TAB(55); USING "#.##"; MHtaxrate!;
       PRINT #10, TAB(61); USING "######.##"; TBCust(1).TBCurYrMobileHomeTaxAmount;
       PRINT #10, TAB(81); USING "######.##"; TBCust(1).TBCurYrMobileHomeTaxAmount

     PRINT #10, "Farm Equipment "; TAB(45); USING "########"; 0;
      PRINT #10, TAB(55); USING "#.##"; FEtaxrate!;
       PRINT #10, TAB(61); USING "######.##"; TBCust(1).TBCurYrFarmEquipTaxAmount;
       PRINT #10, TAB(81); USING "######.##"; TBCust(1).TBCurYrFarmEquipTaxAmount

     PRINT #10, "Late Fee "; TAB(61); USING "######.##"; 0;
      PRINT #10, TAB(81); USING "######.##"; 0


     PRINT #10, "Vehicle Listing"
      CarCount = 0                      'Set Car Count to Zero
      VehRecord! = TBCust(1).FirstVeh
      WHILE VehRecord! <> 0
        GET TBVFile, VehRecord!, TBVeh(1)
        IF TBVeh(1).VehTyp <> "D" AND TBVeh(1).VehQ = "Y" AND TBVeh(1).VehTaxYr = VAL(TaxYear$) THEN

         'Figure PPTRADis# Here
         IF TBVeh(1).VehQ = "Y" THEN
          IF TBVeh(1).VehValue > 20000 THEN
           PPTRAVal# = 20000
           ELSE
           PPTRAVal# = TBVeh(1).VehValue
          END IF
          IF PPTRAVal# <= 1000 THEN
           PPTRADis# = ((PPTRAVal# / 100) * PPTaxRate!)
           ELSE
           PPTRADis# = (((PPTRAVal# / 100) * PPTaxRate!) * (Perc! / 100))
          END IF
         END IF
          PPTRADis# = INT((PPTRADis# * 100) + .50001) / 100

         TaxAmt# = (TBVeh(1).VehValue / 100) * PPTaxRate!

         MakeModel$ = RTRIM$(TBVeh(1).VehYear) + " " + RTRIM$(TBVeh(1).VehMake) + " " + RTRIM$(TBVeh(1).VehModel)

         PRINT #10, TAB(1); LEFT$(MakeModel$, 20);
          PRINT #10, TAB(23); LEFT$(TBVeh(1).VehVin, 20);
          PRINT #10, TAB(45); USING "########"; TBVeh(1).VehValue;
          PRINT #10, TAB(55); USING "#.##"; PPTaxRate!;
          PRINT #10, TAB(61); USING "######.##"; TaxAmt#;
          PRINT #10, TAB(71); USING "######.##"; PPTRADis#
         CarCount = CarCount + 1
         IF CarCount >= 4 AND TBVeh(1).VehNext <> 0 THEN
                BlankLines = 16 - CarCount
                FOR LCnt = 1 TO BlankLines: PRINT #10, "": NEXT LCnt
                PRINT #10, TAB(14); "Acct #"; acctnumber
                PRINT #10, TAB(14); TBCust(1).CustName
                PRINT #10, TAB(14); TBCust(1).Address1;
                 PRINT #10, TAB(40); USING "######.##"; TotalTaxAmount@ + TBCust(1).TBPPTRADiscnt;
                 PRINT #10, TAB(50); USING "######.##"; TBCust(1).TBPPTRADiscnt;
                 PRINT #10, TAB(60); USING "######.##"; TotalTaxAmount@
                PRINT #10, TAB(14); TBCust(1).Address2;
                 PRINT #10, TAB(60); USING "######.##"; TotalTaxAmount@
                PRINT #10, TAB(14); RTRIM$(TBCust(1).City); ",  "; TBCust(1).State; "  "; TBCust(1).Zip
                PRINT #10, TAB(60); "Penalty Date: "; PenDate$
                PRINT #10, TAB(60); "Interest Date: "; IntDate$
                PRINT #10, "End of Bill - To Continue Same Account"
                GOSUB KBAnotherBill
                CarCount = 0
         END IF
        END IF
        VehRecord! = TBVeh(1).VehNext
      WEND
      BlankLines = 16 - CarCount
      FOR LCnt = 1 TO BlankLines: PRINT #10, "": NEXT LCnt
      PRINT #10, TAB(14); "Acct #"; acctnumber
      PRINT #10, TAB(14); TBCust(1).CustName
      PRINT #10, TAB(14); TBCust(1).Address1;
                 PRINT #10, TAB(40); USING "######.##"; TotalTaxAmount@ + TBCust(1).TBPPTRADiscnt;
                 PRINT #10, TAB(50); USING "######.##"; TBCust(1).TBPPTRADiscnt;
                 PRINT #10, TAB(60); USING "######.##"; TotalTaxAmount@
      PRINT #10, TAB(14); TBCust(1).Address2;
                 PRINT #10, TAB(60); USING "######.##"; TotalTaxAmount@
      PRINT #10, TAB(14); RTRIM$(TBCust(1).City); ",  "; TBCust(1).State; "  "; TBCust(1).Zip
      PRINT #10, TAB(60); "Penalty Date: "; PenDate$
      PRINT #10, TAB(60); "Interest Date: "; IntDate$
      PRINT #10, "End of Bill"
     RETURN


KBAnotherBill:
   ' Print the Bill Here
     PRINT #10, "PERSONAL PROPERTY BILL START"
     PRINT #10, "TOWN OF KENBRIDGE"
     PRINT #10, "804-676-2452"
     PRINT #10, "TREASURER, TOWN OF KENBRIDGE"
     PRINT #10, "100 HIGH STREET   PO BOX 478"
     PRINT #10, "KENBRIDGE, VA 23944"
     PRINT #10, ""
     PRINT #10, "TOWN OF KENBRIDGE"
     PRINT #10, "100 High Street   PO Box 478"
     PRINT #10, "Kenbridge, VA 23944"
     PRINT #10, "Call 804-676-2452 for information."
     PRINT #10, "Tax Year: "; TaxYear$
     PRINT #10, TAB(85); TBCust(1).TBCurYrBillNumber
     PRINT #10,
     PRINT #10, "Personal Property "
     PRINT #10, "Machinery & Tools "
     PRINT #10, "Merchant Capital"
     PRINT #10, "Mobile Home"
     PRINT #10, "Farm Equipment "
     PRINT #10, "Late Fee  "
     PRINT #10, "Vehicle Listing Continued"
     RETURN

5000 CLS : scren$ = "3.03.00": GOSUB 63500: COLOR 11
     LOCATE 1, 27: PRINT "Print Tax Bills - Victoria"
     LOCATE 3, 1: COLOR 11
     PRINT "Enter the Pers. Prop. Tax Rate Per $100 (Ex. .50 for 50 cents / $100): "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     PPTaxRate! = VAL(in$)
     LOCATE 4, 1: COLOR 11
     PRINT "Enter the Machinery and Tools Tax Rate Per $100 (ESC-End) : "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     MTtaxrate! = VAL(in$)
     LOCATE 5, 1: COLOR 11
     PRINT "Enter the Merchants Capital Tax Rate Per $100 (ESC-End)   : "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     MCtaxrate! = VAL(in$)
     LOCATE 6, 1: COLOR 11
     PRINT "Enter the Mobile Home Tax Rate Per $100 (ESC-End)         : "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     MHtaxrate! = VAL(in$)
     LOCATE 7, 1: COLOR 11
     PRINT "Enter the Farm Equipment Tax Rate Per $100 (ESC-End)      : "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     FEtaxrate! = VAL(in$)

TaxYearEnter5:
     LOCATE 8, 1: COLOR 11
     PRINT "Enter the Tax Year  (ESC-End) (yyyy) : "; : fl = -4
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     IF LEN(in$) <> 4 THEN GOTO TaxYearEnter5
     TaxYear$ = in$
    'PPTRA Percent
     IF VAL(TaxYear$) = 1998 THEN Perc! = 12.5
     IF VAL(TaxYear$) = 1999 THEN Perc! = 27.5
     IF VAL(TaxYear$) = 2000 THEN Perc! = 47.5
     IF VAL(TaxYear$) = 2001 THEN Perc! = 70
     IF VAL(TaxYear$) = 2002 THEN Perc! = 100


PenDate5:
     LOCATE 9, 1: COLOR 11
     PRINT "Enter the Penalty Date (mm/dd/yyyy) .. "; : fl = 10
     GOSUB 62000
     IF LEN(in$) <> 10 THEN GOTO PenDate5
     PenDate$ = in$
IntDate5:
     LOCATE 10, 1: COLOR 11
     PRINT "Enter the Date Interest to Begin (mm/dd/yyyy) .. "; : fl = 10
     GOSUB 62000
     IF LEN(in$) <> 10 THEN GOTO IntDate5
     IntDate$ = in$
5001 LOCATE 12, 1: COLOR 11
     PRINT "Print (A)ll or (I)ndividual Bills or (ESC) to End .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN RUN "tbbillin"
     IF NOT (in$ = "A" OR in$ = "I") THEN 5001
     IF in$ = "A" THEN
       
5035  ' GET A SORTED INDEX FIRST
      TBFile = FREEFILE
      OPEN TBpath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBCust(1))
      OPEN TBpath$ + "TBPTname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #9 LEN = 24
      FIELD 9, 20 AS LastNam$, 4 AS Mrec$: EndofFile = LOF(9) / 24
      TBVFile = FREEFILE
      OPEN TBpath$ + "TBTVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))


       ReportFile$ = "TAXPBILL.PRN"
       OPEN "O", 10, "TAXPBILL.PRN"
       WIDTH #10, 255


      LastRecord = LOF(9) / 24
       Count = 1
    WHILE Count <= LastRecord
     GET 9, Count

     GET TBFile, CVS(Mrec$), TBCust(1)

     acctnumber = CVS(Mrec$)

     GOSUB PrintBillVictoria
     Count = Count + 1
     WEND
     CLOSE

'Set Report Defaults
VCEndAllBills:
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TAXPBILL.PRN"
      Header$ = "Personal Property Tickets"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      CLOSE
      KILL ReportFile$
      RUN "tbbillin"

    ELSE

      COLOR 11
      LOCATE 15, 1: PRINT "Individual Printing"
      

VCmainbody:
     'Open Main Tax Files
      TBFile = FREEFILE
      OPEN TBpath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBCust(1))
      TBVFile = FREEFILE
      OPEN TBpath$ + "TBTVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))

     CLS : scren$ = "3.03.00": GOSUB 63500: COLOR 11
     LOCATE 1, 27: PRINT "Print Tax Bills-Victoria"
     LOCATE 3, 1: COLOR 11
     LOCATE 17, 1: PRINT "Enter Customer Acct Number to Reprint (ESC-END) .. "; : fl = -7
    GOSUB 62000
  IF cf = 1 THEN CLOSE : RUN "tbbillin"
      acctnumber = VAL(in$)
      GET TBFile, acctnumber, TBCust(1)
       ReportFile$ = "TAXPBILL.PRN"
       OPEN "O", 10, "TAXPBILL.PRN"
       WIDTH #10, 255
       GOSUB PrintBillVictoria
       CLOSE
       LPTPort% = 1
       EntryPoint% = 1
       ReportFile$ = "TAXPBILL.PRN"
       Header$ = "Personal Property Tickets"
       PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
       CLOSE
       KILL ReportFile$
       GOTO VCmainbody
  END IF



PrintBillVictoria:
     IF VAL(TBCust(1).TBCurYrBillNumber) < 1 THEN RETURN     'No Bill
     IF VAL(TBCust(1).District) <> 10 THEN RETURN'Must be 10 for Victoria
     TotalTaxAmount@ = (TBCust(1).TBCurYrPersPropTaxAmount + TBCust(1).TBCurYrMachToolsTaxAmount + TBCust(1).TBCurYrMerchCapTaxAmount + TBCust(1).TBCurYrMobileHomeTaxAmount + TBCust(1).TBCurYrFarmEquipTaxAmount - TBCust(1).TBPPTRADiscnt)
     TotalTaxAmount@ = INT((TotalTaxAmount@ * 100) + .50001) / 100
     TotalTaxAmount@ = TotalTaxAmount@ + TBCust(1).TBLateFeeAmount
     TotalTaxAmount@ = INT((TotalTaxAmount@ * 100) + .50001) / 100

     Penalty@ = TotalTaxAmount@ * .1
     Penalty@ = INT((Penalty@ * 100) + .50001) / 100

     IF TotalTaxAmount@ < .01 THEN RETURN       'No Bill



   ' Print the Bill Here
     PRINT #10, "PERSONAL PROPERTY BILL START"
     PRINT #10, "TOWN OF VICTORIA"
     PRINT #10, "804-696-2343"
     PRINT #10, "TREASURER, TOWN OF VICTORIA"
     PRINT #10, "PO BOX V"
     PRINT #10, "VICTORIA, VA 23974"
     PRINT #10, ""
     PRINT #10, "TOWN OF VICTORIA"
     PRINT #10, "PO Box V"
     PRINT #10, "Victoria, VA 23974"
     PRINT #10, "Call 804-696-2343 for information."
     PRINT #10, "Tax Year: "; TaxYear$
     PRINT #10, TAB(85); TBCust(1).TBCurYrBillNumber
     PRINT #10,
     PRINT #10, "Personal Property ";
       PRINT #10, TAB(45); USING "########"; TBCust(1).PersProp;
       PRINT #10, TAB(55); USING "#.##"; PPTaxRate!;
       PRINT #10, TAB(61); USING "######.##"; TBCust(1).TBCurYrPersPropTaxAmount;
       PRINT #10, TAB(71); USING "######.##"; TBCust(1).TBPPTRADiscnt;
       PRINT #10, TAB(81); USING "######.##"; TBCust(1).TBCurYrPersPropTaxAmount - TBCust(1).TBPPTRADiscnt

    PRINT #10, "Machinery & Tools";
      PRINT #10, TAB(45); USING "########"; 0;
      PRINT #10, TAB(55); USING "#.##"; MTtaxrate!;
       PRINT #10, TAB(61); USING "######.##"; TBCust(1).TBCurYrMachToolsTaxAmount;
       PRINT #10, TAB(81); USING "######.##"; TBCust(1).TBCurYrMachToolsTaxAmount

     PRINT #10, "Merchant Capital"; TAB(45); USING "########"; 0;
      PRINT #10, TAB(55); USING "#.##"; MCtaxrate!;
       PRINT #10, TAB(61); USING "######.##"; TBCust(1).TBCurYrMerchCapTaxAmount;
       PRINT #10, TAB(81); USING "######.##"; TBCust(1).TBCurYrMerchCapTaxAmount

    PRINT #10, "Mobile Home "; TAB(45); USING "########"; TBCust(1).MobileHome;
      PRINT #10, TAB(55); USING "#.##"; MHtaxrate!;
       PRINT #10, TAB(61); USING "######.##"; TBCust(1).TBCurYrMobileHomeTaxAmount;
       PRINT #10, TAB(81); USING "######.##"; TBCust(1).TBCurYrMobileHomeTaxAmount

     PRINT #10, "Farm Equipment "; TAB(45); USING "########"; 0;
      PRINT #10, TAB(55); USING "#.##"; FEtaxrate!;
       PRINT #10, TAB(61); USING "######.##"; TBCust(1).TBCurYrFarmEquipTaxAmount;
       PRINT #10, TAB(81); USING "######.##"; TBCust(1).TBCurYrFarmEquipTaxAmount

     PRINT #10, "Late Fee "; TAB(61); USING "######.##"; 0;
      PRINT #10, TAB(81); USING "######.##"; 0


     PRINT #10, "Vehicle Listing"
      CarCount = 0                      'Set Car Count to Zero
      VehRecord! = TBCust(1).FirstVeh
      WHILE VehRecord! <> 0
        GET TBVFile, VehRecord!, TBVeh(1)
        IF TBVeh(1).VehTyp <> "D" AND TBVeh(1).VehQ = "Y" AND TBVeh(1).VehTaxYr = VAL(TaxYear$) THEN

         'Figure PPTRADis# Here
         IF TBVeh(1).VehQ = "Y" THEN
          IF TBVeh(1).VehValue > 20000 THEN
           PPTRAVal# = 20000
           ELSE
           PPTRAVal# = TBVeh(1).VehValue
          END IF
          IF PPTRAVal# <= 1000 THEN
           PPTRADis# = ((PPTRAVal# / 100) * PPTaxRate!)
           ELSE
           PPTRADis# = (((PPTRAVal# / 100) * PPTaxRate!) * (Perc! / 100))
          END IF
         END IF
          PPTRADis# = INT((PPTRADis# * 100) + .50001) / 100

         TaxAmt# = (TBVeh(1).VehValue / 100) * PPTaxRate!

         MakeModel$ = RTRIM$(TBVeh(1).VehYear) + " " + RTRIM$(TBVeh(1).VehMake) + " " + RTRIM$(TBVeh(1).VehModel)

         PRINT #10, TAB(1); LEFT$(MakeModel$, 20);
          PRINT #10, TAB(23); LEFT$(TBVeh(1).VehVin, 20);
          PRINT #10, TAB(45); USING "########"; TBVeh(1).VehValue;
          PRINT #10, TAB(55); USING "#.##"; PPTaxRate!;
          PRINT #10, TAB(61); USING "######.##"; TaxAmt#;
          PRINT #10, TAB(71); USING "######.##"; PPTRADis#
         CarCount = CarCount + 1
         IF CarCount >= 4 AND TBVeh(1).VehNext <> 0 THEN
                BlankLines = 16 - CarCount
                FOR LCnt = 1 TO BlankLines: PRINT #10, "": NEXT LCnt
                PRINT #10, TAB(14); "Acct #"; acctnumber
                PRINT #10, TAB(14); TBCust(1).CustName
                PRINT #10, TAB(14); TBCust(1).Address1;
                 PRINT #10, TAB(40); USING "######.##"; TotalTaxAmount@ + TBCust(1).TBPPTRADiscnt;
                 PRINT #10, TAB(50); USING "######.##"; TBCust(1).TBPPTRADiscnt;
                 PRINT #10, TAB(60); USING "######.##"; TotalTaxAmount@
                PRINT #10, TAB(14); TBCust(1).Address2;
                 PRINT #10, TAB(60); USING "######.##"; TotalTaxAmount@
                PRINT #10, TAB(14); RTRIM$(TBCust(1).City); ",  "; TBCust(1).State; "  "; TBCust(1).Zip
                PRINT #10, TAB(60); "Penalty Date: "; PenDate$
                PRINT #10, TAB(60); "Interest Date: "; IntDate$
                PRINT #10, "End of Bill - To Continue Same Account"
                GOSUB KBAnotherBill
                CarCount = 0
         END IF
        END IF
        VehRecord! = TBVeh(1).VehNext
      WEND
      BlankLines = 16 - CarCount
      FOR LCnt = 1 TO BlankLines: PRINT #10, "": NEXT LCnt
      PRINT #10, TAB(14); "Acct #"; acctnumber
      PRINT #10, TAB(14); TBCust(1).CustName
      PRINT #10, TAB(14); TBCust(1).Address1;
                 PRINT #10, TAB(40); USING "######.##"; TotalTaxAmount@ + TBCust(1).TBPPTRADiscnt;
                 PRINT #10, TAB(50); USING "######.##"; TBCust(1).TBPPTRADiscnt;
                 PRINT #10, TAB(60); USING "######.##"; TotalTaxAmount@
      PRINT #10, TAB(14); TBCust(1).Address2;
                 PRINT #10, TAB(60); USING "######.##"; TotalTaxAmount@
      PRINT #10, TAB(14); RTRIM$(TBCust(1).City); ",  "; TBCust(1).State; "  "; TBCust(1).Zip
      PRINT #10, TAB(60); "Penalty Date: "; PenDate$
      PRINT #10, TAB(60); "Interest Date: "; IntDate$
      PRINT #10, "End of Bill"
     RETURN

VCAnotherBill:
     ' Print the Bill Here
     PRINT #10, "PERSONAL PROPERTY BILL START"
     PRINT #10, "TOWN OF VICTORIA"
     PRINT #10, "804-696-2343"
     PRINT #10, "TREASURER, TOWN OF VICTORIA"
     PRINT #10, "PO BOX V"
     PRINT #10, "VICTORIA, VA 23974"
     PRINT #10, ""
     PRINT #10, "TOWN OF VICTORIA"
     PRINT #10, "PO Box V"
     PRINT #10, "Victoria, VA 23974"
     PRINT #10, "Call 804-696-2343 for information."
     PRINT #10, "Tax Year: "; TaxYear$
     PRINT #10, TAB(85); TBCust(1).TBCurYrBillNumber
     PRINT #10,
     PRINT #10, "Personal Property "
     PRINT #10, "Machinery & Tools "
     PRINT #10, "Merchant Capital"
     PRINT #10, "Mobile Home"
     PRINT #10, "Farm Equipment "
     PRINT #10, "Late Fee  "
     PRINT #10, "Vehicle Listing Continued"
     RETURN




GetDistrict:
     ' Only Towns of Kenbridge and Victoria Need to Be Defined for Personal Property
     DistrictName$ = ""
     IF VAL(TBCust(1).District) = 9 THEN DistrictName$ = "KENBRIDGE"
     IF VAL(TBCust(1).District) = 10 THEN DistrictName$ = "VICTORIA"
     RETURN

60000  IF taxtype$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBpath$ + "TBRTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBpath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         CLOSE TBVFile
         TBVFile = FREEFILE
         OPEN TBpath$ + "TBTVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))
        END IF
        RETURN

60025 TBBalFile = FREEFILE
      IF taxtype$ = "R" THEN
      OPEN TBpath$ + "TBRTBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBpath$ + "TBPTBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
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

