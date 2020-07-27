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
    REDIM tbcust(1)  AS TBPPCust
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
     PPtaxrate! = VAL(in$)
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
     IF VAL(TaxYear$) >= 2002 THEN Perc! = 70
PenDate:
    LOCATE 9, 1: COLOR 11
    PRINT "Enter the Due Date (mm/dd/yyyy) .. "; : fl = 10
    GOSUB 62000
    IF LEN(in$) <> 10 THEN GOTO PenDate
    PenDate$ = in$


3001 LOCATE 12, 1: COLOR 11
     PRINT "Press <Enter> to Create File for Total Billing or (ESC) to End .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN RUN "tbbillin"
3035 ' GET A SORTED INDEX FIRST
      LOCATE 13, 1
      COLOR 11
      PRINT "Enter the TR# From the Proof Report to Start (0=All:<ESC>-Cancel) .. "; : fl = -5
      GOSUB 62000
      IF cf = 1 THEN CLOSE : GOTO 10
      BegTr# = VAL(in$)


      TBFile = FREEFILE
      OPEN TBpath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(tbcust(1))
      OPEN TBpath$ + "TBPTname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #9 LEN = 24
      FIELD 9, 20 AS LastNam$, 4 AS Mrec$: EndofFile = LOF(9) / 24
      TBVFile = FREEFILE
      OPEN TBpath$ + "TBTVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))


      ReportFile$ = "LCPP" + TaxYear$ + ".TXT"
      OPEN "O", 10, ReportFile$
      WIDTH #10, 255
      PRINT "Now Creating File for Total Billing Called: "; ReportFile$

    LastRecord = LOF(9) / 24
    Count = 1
    WHILE Count <= LastRecord
     GET 9, Count
     IF BegTr# = 0 OR Count >= BegTr# THEN

     LOCATE 15, 1: COLOR 11: PRINT "Processing Acct # "; Count
     GET TBFile, CVS(Mrec$), tbcust(1)

     acctnumber = CVS(Mrec$)
     
     GOSUB PrintBill
     END IF
     Count = Count + 1
     WEND
     CLOSE

EndAllBills:
      CLOSE
      LOCATE 19, 1
      COLOR 11
      PRINT "Please note file name for total billing, press enter to continue "; : fl = 1
      GOSUB 62000
      RUN "tbbillin"

''''STARTS HERE BOB
PrintBill:
     IF VAL(tbcust(1).LBCurYrBillNumber) < 1 THEN RETURN     'No Bill
     TotalTaxAmount@ = (tbcust(1).LBCurYrPersPropTaxAmount + tbcust(1).LBCurYrMachToolsTaxAmount + tbcust(1).LBCurYrMerchCapTaxAmount + tbcust(1).LBCurYrMobileHomeTaxAmount + tbcust(1).LBCurYrFarmEquipTaxAmount - tbcust(1).LBPPTRADiscnt)
     TotalTaxAmount@ = INT((TotalTaxAmount@ * 100) + .50001) / 100
     TotalTaxAmount@ = TotalTaxAmount@ + tbcust(1).LBLateFeeAmount
     TotalTaxAmount@ = INT((TotalTaxAmount@ * 100) + .50001) / 100

     IF TotalTaxAmount@ < .01 AND tbcust(1).LBPPTRADiscnt = 0 THEN RETURN      'No Bill

     
     'First Determine Pers Prop of Cars
      VehRecord! = tbcust(1).FirstVeh
      VehTaxValue# = 0
      WHILE VehRecord! <> 0
        GET TBVFile, VehRecord!, TBVeh(1)
         IF TBVeh(1).VehTyp <> "D" AND TBVeh(1).VehTaxYr = VAL(TaxYear$) THEN
           VehTaxValue# = VehTaxValue# + TBVeh(1).VehValue
          END IF
          VehRecord! = TBVeh(1).VehNext
       WEND

      'IF INSTR(tbcust(1).CustName, "ADAMS JAMES") > 0 THEN STOP

     '***** TOTAL BILLING OUTPUT FILE HERE *****
      PRINT #10, tbcust(1).CustName;
      PRINT #10, tbcust(1).Address1;
      PRINT #10, tbcust(1).Address2;
      PRINT #10, tbcust(1).City;
      PRINT #10, tbcust(1).State;
      PRINT #10, tbcust(1).Zip;
      PRINT #10, USING "######"; acctnumber;
      PRINT #10, tbcust(1).SocSec;
      PRINT #10, tbcust(1).SocSec2;
      PRINT #10, PenDate$;
      PRINT #10, USING "########.##"; TotalTaxAmount@ + tbcust(1).LBPPTRADiscnt;
      PRINT #10, USING "########.##"; tbcust(1).LBPPTRADiscnt;
      PRINT #10, USING "########.##"; TotalTaxAmount@;
      
'Detail Lines Start Here

     PRINT #10, "MACHINERY & TOOLS   ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; tbcust(1).MachTools;
     PRINT #10, USING "#.##"; MTtaxrate!;
     PRINT #10, USING "######.##"; tbcust(1).LBCurYrMachToolsTaxAmount;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).LBCurYrMachToolsTaxAmount;

     PRINT #10, "MERCHANT CAPITAL    ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; tbcust(1).MerchCap;
     PRINT #10, USING "#.##"; MCtaxrate!;
     PRINT #10, USING "######.##"; tbcust(1).LBCurYrMerchCapTaxAmount;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).LBCurYrMerchCapTaxAmount;
 
     PRINT #10, "MOBILE HOME         ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; tbcust(1).MobileHome;
     PRINT #10, USING "#.##"; MHtaxrate!;
     PRINT #10, USING "######.##"; tbcust(1).LBCurYrMobileHomeTaxAmount;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).LBCurYrMobileHomeTaxAmount;

     PRINT #10, "FARM EQUIPMENT      ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; tbcust(1).FarmEquip;
     PRINT #10, USING "#.##"; FEtaxrate!;
     PRINT #10, USING "######.##"; tbcust(1).LBCurYrFarmEquipTaxAmount;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).LBCurYrFarmEquipTaxAmount;

     PRINT #10, "OTHER PERS PROPERTY ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; (tbcust(1).PersProp - VehTaxValue#);
     PRINT #10, USING "#.##"; PPtaxrate!;
     PRINT #10, USING "######.##"; ((tbcust(1).PersProp - VehTaxValue#) / 100) * PPtaxrate!;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; ((tbcust(1).PersProp - VehTaxValue#) / 100) * PPtaxrate!;
     
     PRINT #10, "LATE FEE            ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; 0;
     PRINT #10, USING "#.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).LBLateFeeAmount;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).LBLateFeeAmount;

'''''CAR PROCESSING STARTS HERE
carprocess:
      CarCount = 0
      VehRecord! = tbcust(1).FirstVeh
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
           PPTRADis# = ((PPTRAVal# / 100) * PPtaxrate!)
           ELSE
           PPTRADis# = (((PPTRAVal# / 100) * PPtaxrate!) * (Perc! / 100))
          END IF
        END IF
          PPTRADis# = INT((PPTRADis# * 100) + .50001) / 100

         TaxAmt# = (TBVeh(1).VehValue / 100) * PPtaxrate!
         TaxAmt# = INT((TaxAmt# * 100) + .50001) / 100
         NetAmt# = TaxAmt# - PPTRADis#


         MakeModel$ = RTRIM$(TBVeh(1).VehYear) + " " + RTRIM$(TBVeh(1).VehMake) + " " + RTRIM$(TBVeh(1).VehModel)
         IF LEN(MakeModel$) < 20 THEN MakeModel$ = MakeModel$ + STRING$(20 - LEN(MakeModel$), " ")
        
         PRINT #10, LEFT$(MakeModel$, 20);
         PRINT #10, LEFT$(TBVeh(1).VehVin, 17);
         PRINT #10, USING "########"; TBVeh(1).VehValue;
         PRINT #10, USING "#.##"; PPtaxrate!;
         PRINT #10, USING "######.##"; TaxAmt#;
         PRINT #10, USING "######.##"; PPTRADis#;
         PRINT #10, USING "######.##"; NetAmt#;
         CarCount = CarCount + 1
        END IF
        VehRecord! = TBVeh(1).VehNext
      WEND
'''''CAR PROCESSING END

'''''BLANK RECORD FILLING????
      FOR Blk = CarCount + 1 TO 68
         PRINT #10, STRING$(76, " ");
      NEXT Blk
      PRINT #10, STRING$(76, " ")
      RETURN
'''''BLANK FILLING END
''''''ENDS HERE


4000 CLS : scren$ = "3.03.00": GOSUB 63500: COLOR 11
     LOCATE 1, 27: PRINT "Print Tax Bills - Kenbridge"
     LOCATE 3, 1: COLOR 11
     PRINT "Enter the Pers. Prop. Tax Rate Per $100 (Ex. .50 for 50 cents / $100): "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     PPtaxrate! = VAL(in$)
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
KBTaxYearEnter:
     LOCATE 8, 1: COLOR 11
     PRINT "Enter the Tax Year  (ESC-End) (yyyy) : "; : fl = -4
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     IF LEN(in$) <> 4 THEN GOTO KBTaxYearEnter
     TaxYear$ = in$

    'PPTRA Percent
     IF VAL(TaxYear$) = 1998 THEN Perc! = 12.5
     IF VAL(TaxYear$) = 1999 THEN Perc! = 27.5
     IF VAL(TaxYear$) = 2000 THEN Perc! = 47.5
     IF VAL(TaxYear$) = 2001 THEN Perc! = 70
     IF VAL(TaxYear$) >= 2002 THEN Perc! = 70
KBPenDate:
    LOCATE 9, 1: COLOR 11
    PRINT "Enter the Due Date (mm/dd/yyyy) .. "; : fl = 10
    GOSUB 62000
    IF LEN(in$) <> 10 THEN GOTO KBPenDate
    PenDate$ = in$


4001 LOCATE 12, 1: COLOR 11
     PRINT "Press <Enter> to Create File for Total Billing or (ESC) to End .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN RUN "tbbillin"

4035 ' GET A SORTED INDEX FIRST
      LOCATE 13, 1
      COLOR 11
      BegTr# = 1



      TBFile = FREEFILE
      OPEN TBpath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(tbcust(1))
      OPEN TBpath$ + "TBPTname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #9 LEN = 24
      FIELD 9, 20 AS LastNam$, 4 AS Mrec$: EndofFile = LOF(9) / 24
      TBVFile = FREEFILE
      OPEN TBpath$ + "TBTVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))


       ReportFile$ = "KBPP" + TaxYear$ + ".TXT"
       OPEN "O", 10, ReportFile$
       WIDTH #10, 255
       PRINT "Now Creating File (Kenbridge) for Total Billing Called: "; ReportFile$

    LastRecord = LOF(9) / 24
    Count = 1
    WHILE Count <= LastRecord
     GET 9, Count
     IF BegTr# = 0 OR Count >= BegTr# THEN


     LOCATE 15, 1: COLOR 11: PRINT "Processing Acct # "; Count
     GET TBFile, CVS(Mrec$), tbcust(1)

     acctnumber = CVS(Mrec$)

     GOSUB KBPrintBill
     END IF
     Count = Count + 1
     WEND
     CLOSE

KBEndAllBills:
      CLOSE
      LOCATE 19, 1
      COLOR 11
      PRINT "Please note file name for total billing, press enter to continue "; : fl = 1
      GOSUB 62000
      RUN "tbbillin"


KBPrintBill:
     IF VAL(tbcust(1).TBCurYrBillNumber) < 1 THEN RETURN     'No Bill
     IF VAL(tbcust(1).District) <> 9 THEN RETURN'Must be 9 for Kenbridge
     

     TotalTaxAmount@ = (tbcust(1).TBCurYrPersPropTaxAmount + tbcust(1).TBCurYrMachToolsTaxAmount + tbcust(1).TBCurYrMerchCapTaxAmount + tbcust(1).TBCurYrMobileHomeTaxAmount + tbcust(1).TBCurYrFarmEquipTaxAmount - tbcust(1).TBPPTRADiscnt)
     TotalTaxAmount@ = INT((TotalTaxAmount@ * 100) + .50001) / 100

     TotalTaxAmount@ = TotalTaxAmount@ + tbcust(1).TBLateFeeAmount
     TotalTaxAmount@ = INT((TotalTaxAmount@ * 100) + .50001) / 100

     IF TotalTaxAmount@ < .01 AND tbcust(1).TBPPTRADiscnt = 0 THEN RETURN      'No Bill
     
     'First Determine Pers Prop of Cars
      VehRecord! = tbcust(1).FirstVeh
      VehTaxValue# = 0
      WHILE VehRecord! <> 0
        GET TBVFile, VehRecord!, TBVeh(1)
         IF TBVeh(1).VehTyp <> "D" AND TBVeh(1).VehTaxYr = VAL(TaxYear$) THEN
           VehTaxValue# = VehTaxValue# + TBVeh(1).VehValue
          END IF

           VehRecord! = TBVeh(1).VehNext
       WEND

     '***** TOTAL BILLING OUTPUT FILE HERE *****
      PRINT #10, tbcust(1).CustName;
      PRINT #10, tbcust(1).Address1;
      PRINT #10, tbcust(1).Address2;
      PRINT #10, tbcust(1).City;
      PRINT #10, tbcust(1).State;
      PRINT #10, tbcust(1).Zip;
      PRINT #10, USING "######"; acctnumber;
      PRINT #10, tbcust(1).SocSec;
      PRINT #10, tbcust(1).SocSec2;
      PRINT #10, PenDate$;
      PRINT #10, USING "########.##"; TotalTaxAmount@ + tbcust(1).TBPPTRADiscnt;
      PRINT #10, USING "########.##"; tbcust(1).TBPPTRADiscnt;
      PRINT #10, USING "########.##"; TotalTaxAmount@;


      'Detail Lines Start Here

     PRINT #10, "MACHINERY & TOOLS   ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; tbcust(1).MachTools;
     PRINT #10, USING "#.##"; MTtaxrate!;
     PRINT #10, USING "######.##"; tbcust(1).TBCurYrMachToolsTaxAmount;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).TBCurYrMachToolsTaxAmount;

     PRINT #10, "MERCHANT CAPITAL    ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; tbcust(1).MerchCap;
     PRINT #10, USING "#.##"; MCtaxrate!;
     PRINT #10, USING "######.##"; tbcust(1).TBCurYrMerchCapTaxAmount;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).TBCurYrMerchCapTaxAmount;

     PRINT #10, "MOBILE HOME         ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; tbcust(1).MobileHome;
     PRINT #10, USING "#.##"; MHtaxrate!;
     PRINT #10, USING "######.##"; tbcust(1).TBCurYrMobileHomeTaxAmount;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).TBCurYrMobileHomeTaxAmount;

     PRINT #10, "FARM EQUIPMENT      ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; tbcust(1).FarmEquip;
     PRINT #10, USING "#.##"; FEtaxrate!;
     PRINT #10, USING "######.##"; tbcust(1).TBCurYrFarmEquipTaxAmount;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).TBCurYrFarmEquipTaxAmount;

     PRINT #10, "OTHER PERS PROPERTY ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; (tbcust(1).PersProp - VehTaxValue#);
     PRINT #10, USING "#.##"; PPtaxrate!;
     PRINT #10, USING "######.##"; ((tbcust(1).PersProp - VehTaxValue#) / 100) * PPtaxrate!;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; ((tbcust(1).PersProp - VehTaxValue#) / 100) * PPtaxrate!;

     PRINT #10, "LATE FEE            ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; 0;
     PRINT #10, USING "#.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).TBLateFeeAmount;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).TBLateFeeAmount;

KBcarprocess:
      CarCount = 0
      VehRecord! = tbcust(1).FirstVeh
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
           PPTRADis# = ((PPTRAVal# / 100) * PPtaxrate!)
           ELSE
           PPTRADis# = (((PPTRAVal# / 100) * PPtaxrate!) * (Perc! / 100))
          END IF
        END IF
          PPTRADis# = INT((PPTRADis# * 100) + .50001) / 100

         TaxAmt# = (TBVeh(1).VehValue / 100) * PPtaxrate!
         TaxAmt# = INT((TaxAmt# * 100) + .50001) / 100
         NetAmt# = TaxAmt# - PPTRADis#


         MakeModel$ = RTRIM$(TBVeh(1).VehYear) + " " + RTRIM$(TBVeh(1).VehMake) + " " + RTRIM$(TBVeh(1).VehModel)
         IF LEN(MakeModel$) < 20 THEN MakeModel$ = MakeModel$ + STRING$(20 - LEN(MakeModel$), " ")

         PRINT #10, LEFT$(MakeModel$, 20);
         PRINT #10, LEFT$(TBVeh(1).VehVin, 17);
         PRINT #10, USING "########"; TBVeh(1).VehValue;
         PRINT #10, USING "#.##"; PPtaxrate!;
         PRINT #10, USING "######.##"; TaxAmt#;
         PRINT #10, USING "######.##"; PPTRADis#;
         PRINT #10, USING "######.##"; NetAmt#;
         CarCount = CarCount + 1
        END IF
        VehRecord! = TBVeh(1).VehNext
      WEND
        FOR Blk = CarCount + 1 TO 68
         PRINT #10, STRING$(76, " ");
        NEXT Blk
         PRINT #10, STRING$(76, " ")
       RETURN
     RETURN



5000 CLS : scren$ = "3.03.00": GOSUB 63500: COLOR 11
     LOCATE 1, 27: PRINT "Print Tax Bills -Victoria"
     LOCATE 3, 1: COLOR 11
     PRINT "Enter the Pers. Prop. Tax Rate Per $100 (Ex. .50 for 50 cents / $100): "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     PPtaxrate! = VAL(in$)
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
VCTaxYearEnter:
     LOCATE 8, 1: COLOR 11
     PRINT "Enter the Tax Year  (ESC-End) (yyyy) : "; : fl = -4
     GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "tbbillin"
     IF LEN(in$) <> 4 THEN GOTO VCTaxYearEnter
     TaxYear$ = in$

    'PPTRA Percent
     IF VAL(TaxYear$) = 1998 THEN Perc! = 12.5
     IF VAL(TaxYear$) = 1999 THEN Perc! = 27.5
     IF VAL(TaxYear$) = 2000 THEN Perc! = 47.5
     IF VAL(TaxYear$) = 2001 THEN Perc! = 70
     IF VAL(TaxYear$) >= 2002 THEN Perc! = 70
VCPenDate:
    LOCATE 9, 1: COLOR 11
    PRINT "Enter the Due Date (mm/dd/yyyy) .. "; : fl = 10
    GOSUB 62000
    IF LEN(in$) <> 10 THEN GOTO VCPenDate
    PenDate$ = in$


5001 LOCATE 12, 1: COLOR 11
     PRINT "Press <Enter> to Create File for Total Billing or (ESC) to End .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN RUN "tbbillin"
5035 ' GET A SORTED INDEX FIRST
      LOCATE 13, 1
      COLOR 11
      BegTr# = 1



      TBFile = FREEFILE
      OPEN TBpath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(tbcust(1))
      OPEN TBpath$ + "TBPTname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #9 LEN = 24
      FIELD 9, 20 AS LastNam$, 4 AS Mrec$: EndofFile = LOF(9) / 24
      TBVFile = FREEFILE
      OPEN TBpath$ + "TBTVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))


       ReportFile$ = "VCPP" + TaxYear$ + ".TXT"
       OPEN "O", 10, ReportFile$
       WIDTH #10, 255
       PRINT "Now Creating File (Victoria) for Total Billing Called: "; :
       COLOR 15: PRINT ReportFile$
       COLOR 11

    LastRecord = LOF(9) / 24
    Count = 1
    WHILE Count <= LastRecord
     GET 9, Count
     IF BegTr# = 0 OR Count >= BegTr# THEN


     LOCATE 15, 1: COLOR 11: PRINT "Processing Acct # "; Count
     GET TBFile, CVS(Mrec$), tbcust(1)

     acctnumber = CVS(Mrec$)

     GOSUB VCPrintBill
     END IF
     Count = Count + 1
     WEND
     CLOSE

VCEndAllBills:
      CLOSE
      LOCATE 19, 1
      COLOR 11
      PRINT "Please note file name for total billing, press enter to continue "; : fl = 1
      GOSUB 62000
      RUN "tbbillin"




VCPrintBill:
     IF VAL(tbcust(1).TBCurYrBillNumber) < 1 THEN RETURN     'No Bill
     IF VAL(tbcust(1).District) <> 10 THEN RETURN'Must be 10 for Victoria

     TotalTaxAmount@ = (tbcust(1).TBCurYrPersPropTaxAmount + tbcust(1).TBCurYrMachToolsTaxAmount + tbcust(1).TBCurYrMerchCapTaxAmount + tbcust(1).TBCurYrMobileHomeTaxAmount + tbcust(1).TBCurYrFarmEquipTaxAmount - tbcust(1).TBPPTRADiscnt)
     TotalTaxAmount@ = INT((TotalTaxAmount@ * 100) + .50001) / 100

     TotalTaxAmount@ = TotalTaxAmount@ + tbcust(1).TBLateFeeAmount
     TotalTaxAmount@ = INT((TotalTaxAmount@ * 100) + .50001) / 100

     IF TotalTaxAmount@ < .01 AND tbcust(1).TBPPTRADiscnt = 0 THEN RETURN      'No Bill

     'First Determine Pers Prop of Cars
      VehRecord! = tbcust(1).FirstVeh
      VehTaxValue# = 0
      WHILE VehRecord! <> 0
        GET TBVFile, VehRecord!, TBVeh(1)
         IF TBVeh(1).VehTyp <> "D" AND TBVeh(1).VehTaxYr = VAL(TaxYear$) THEN
           VehTaxValue# = VehTaxValue# + TBVeh(1).VehValue
          END IF

           VehRecord! = TBVeh(1).VehNext
       WEND

     '***** TOTAL BILLING OUTPUT FILE HERE *****
      PRINT #10, tbcust(1).CustName;
      PRINT #10, tbcust(1).Address1;
      PRINT #10, tbcust(1).Address2;
      PRINT #10, tbcust(1).City;
      PRINT #10, tbcust(1).State;
      PRINT #10, tbcust(1).Zip;
      PRINT #10, USING "######"; acctnumber;
      PRINT #10, tbcust(1).SocSec;
      PRINT #10, tbcust(1).SocSec2;
      PRINT #10, PenDate$;
      PRINT #10, USING "########.##"; TotalTaxAmount@ + tbcust(1).TBPPTRADiscnt;
      PRINT #10, USING "########.##"; tbcust(1).TBPPTRADiscnt;
      PRINT #10, USING "########.##"; TotalTaxAmount@;


      'Detail Lines Start Here

     PRINT #10, "MACHINERY & TOOLS   ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; tbcust(1).MachTools;
     PRINT #10, USING "#.##"; MTtaxrate!;
     PRINT #10, USING "######.##"; tbcust(1).TBCurYrMachToolsTaxAmount;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).TBCurYrMachToolsTaxAmount;

     PRINT #10, "MERCHANT CAPITAL    ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; tbcust(1).MerchCap;
     PRINT #10, USING "#.##"; MCtaxrate!;
     PRINT #10, USING "######.##"; tbcust(1).TBCurYrMerchCapTaxAmount;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).TBCurYrMerchCapTaxAmount;

     PRINT #10, "MOBILE HOME         ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; tbcust(1).MobileHome;
     PRINT #10, USING "#.##"; MHtaxrate!;
     PRINT #10, USING "######.##"; tbcust(1).TBCurYrMobileHomeTaxAmount;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).TBCurYrMobileHomeTaxAmount;

     PRINT #10, "FARM EQUIPMENT      ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; tbcust(1).FarmEquip;
     PRINT #10, USING "#.##"; FEtaxrate!;
     PRINT #10, USING "######.##"; tbcust(1).TBCurYrFarmEquipTaxAmount;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).TBCurYrFarmEquipTaxAmount;

     PRINT #10, "OTHER PERS PROPERTY ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; (tbcust(1).PersProp - VehTaxValue#);
     PRINT #10, USING "#.##"; PPtaxrate!;
     PRINT #10, USING "######.##"; ((tbcust(1).PersProp - VehTaxValue#) / 100) * PPtaxrate!;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; ((tbcust(1).PersProp - VehTaxValue#) / 100) * PPtaxrate!;

     PRINT #10, "LATE FEE            ";
     PRINT #10, STRING$(17, " ");
     PRINT #10, USING "########"; 0;
     PRINT #10, USING "#.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).TBLateFeeAmount;
     PRINT #10, USING "######.##"; 0;
     PRINT #10, USING "######.##"; tbcust(1).TBLateFeeAmount;

VCcarprocess:
      CarCount = 0
      VehRecord! = tbcust(1).FirstVeh
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
           PPTRADis# = ((PPTRAVal# / 100) * PPtaxrate!)
           ELSE
           PPTRADis# = (((PPTRAVal# / 100) * PPtaxrate!) * (Perc! / 100))
          END IF
        END IF
          PPTRADis# = INT((PPTRADis# * 100) + .50001) / 100

         TaxAmt# = (TBVeh(1).VehValue / 100) * PPtaxrate!
         TaxAmt# = INT((TaxAmt# * 100) + .50001) / 100
         NetAmt# = TaxAmt# - PPTRADis#


         MakeModel$ = RTRIM$(TBVeh(1).VehYear) + " " + RTRIM$(TBVeh(1).VehMake) + " " + RTRIM$(TBVeh(1).VehModel)
         IF LEN(MakeModel$) < 20 THEN MakeModel$ = MakeModel$ + STRING$(20 - LEN(MakeModel$), " ")

         PRINT #10, LEFT$(MakeModel$, 20);
         PRINT #10, LEFT$(TBVeh(1).VehVin, 17);
         PRINT #10, USING "########"; TBVeh(1).VehValue;
         PRINT #10, USING "#.##"; PPtaxrate!;
         PRINT #10, USING "######.##"; TaxAmt#;
         PRINT #10, USING "######.##"; PPTRADis#;
         PRINT #10, USING "######.##"; NetAmt#;
         CarCount = CarCount + 1
        END IF
        VehRecord! = TBVeh(1).VehNext
      WEND
        FOR Blk = CarCount + 1 TO 68
         PRINT #10, STRING$(76, " ");
        NEXT Blk
         PRINT #10, STRING$(76, " ")
       RETURN
     RETURN


GetDistrict:
     ' Only Towns of Kenbridge and Victoria Need to Be Defined for Personal Property
     DistrictName$ = ""
     IF VAL(tbcust(1).District) = 9 THEN DistrictName$ = "KENBRIDGE"
     IF VAL(tbcust(1).District) = 10 THEN DistrictName$ = "VICTORIA"
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

