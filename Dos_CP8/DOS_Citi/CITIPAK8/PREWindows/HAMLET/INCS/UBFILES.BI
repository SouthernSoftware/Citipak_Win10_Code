REM New Utility Billing Record Layout and Definitions
REM Updated 3-27-96



TYPE ubsetuprectype
   TownName AS STRING * 35
   CityDef AS STRING * 18
   StDef AS STRING * 2
   ZipDef AS STRING * 10
   RevName(1 TO 10) AS STRING * 25
   RateCode(1 TO 10) AS STRING * 1
   TaxRate(1 TO 10) AS SINGLE
   Metered(1 TO 10) AS STRING * 1
   Mult(1 TO 10) AS SINGLE
   Unit(1 TO 10) AS STRING * 1
   DepositRevenue(1 TO 3) AS INTEGER
   DepositRefund(1 TO 3) AS STRING * 1
   HandheldType AS STRING * 1
    ' 0 = none
    ' 1 = PC/3000
    ' 2 = Sensus
   UseEpp AS STRING * 1
   UseWorkord AS STRING * 1
   UseEstread AS STRING * 1
   UseBankDraft AS STRING * 1
   UseSeqNumb AS STRING * 1
   UseBillCycle AS STRING * 1
   UseBarCode AS STRING * 1
   LowMtrVar AS SINGLE
   HighMtrVar AS SINGLE
   Unused AS STRING * 62
END TYPE


TYPE UBCustIndexRecType
  RecNum AS INTEGER
END TYPE

TYPE UBCustRecType
    BOOK     AS STRING * 2
    SEQNUMB  AS STRING * 6
    ASTATUS  AS STRING * 10
    SEARCHN  AS STRING * 10
    FULLNAME AS STRING * 35
    CADDR1   AS STRING * 35
    CADDR2   AS STRING * 35
    CITY     AS STRING * 18
    STATE    AS STRING * 2
    ZIPCODE  AS STRING * 10
    SOSEC    AS STRING * 11
    DRVLIC   AS STRING * 16
    OPENDATE AS INTEGER
    BILLCOM  AS STRING * 30
    PAYCOM   AS STRING * 30
    CASHONLY AS STRING * 1
    HPHONE   AS STRING * 14
    WPHONE   AS STRING * 14
    LATEFEE  AS STRING * 1
    CUTOFFYN AS STRING * 1
'Page 2
    BILLCYCL AS INTEGER
    CUSTTYPE AS STRING * 3
    ZONE     AS STRING * 3
    POSTRTE  AS STRING * 3
    METERSEQ AS LONG
    BILLCOPY AS INTEGER
    PRORATE  AS INTEGER
    TAXEXPT  AS STRING * 1
    SRCIT    AS STRING * 1
    BARCODE  AS STRING * 1
    EPP      AS STRING * 1
    EPPAMT   AS SINGLE
    BILLWHO  AS STRING * 1
    OWNER    AS STRING * 23
    OWNADD1  AS STRING * 23
    OWNADD2  AS STRING * 23
    OWNADD3  AS STRING * 23
    OWNSOSEC AS STRING * 11
    DRAFTYPE AS STRING * 1
    BankName AS STRING * 15
    BANKLOC  AS STRING * 15
    TRANSIT  AS STRING * 9
    BANKACCT AS STRING * 15
'Page 3
    FRDESC1  AS STRING * 20
    FRAMT1   AS DOUBLE
    FRFREQ1  AS STRING * 1
    REVSRC1  AS INTEGER
    FRDESC2  AS STRING * 20
    FRAMT2   AS DOUBLE
    FRFREQ2  AS STRING * 1
    REVSRC2  AS INTEGER
    FRDESC3  AS STRING * 20
    FRAMT3   AS DOUBLE
    FRFREQ3  AS STRING * 1
    REVSRC3  AS INTEGER
    FRDESC4  AS STRING * 20
    FRAMT4   AS DOUBLE
    FRFREQ4  AS STRING * 1
    REVSRC4  AS INTEGER

    REVNAM1  AS STRING * 20
    RC1      AS STRING * 4
    RMTRTYP1 AS STRING * 1
    REVNAM2  AS STRING * 20
    RC2      AS STRING * 4
    RMTRTYP2 AS STRING * 1
    REVNAM3  AS STRING * 20
    RC3      AS STRING * 4
    RMTRTYP3 AS STRING * 1
    REVNAM4  AS STRING * 20
    RC4      AS STRING * 4
    RMTRTYP4 AS STRING * 1
    REVNAM5  AS STRING * 20
    RC5      AS STRING * 4
    RMTRTYP5 AS STRING * 1
    REVNAM6  AS STRING * 20
    RC6      AS STRING * 4
    RMTRTYP6 AS STRING * 1
    REVNAM7  AS STRING * 20
    RC7      AS STRING * 4
    RMTRTYP7 AS STRING * 1
    REVNAM8  AS STRING * 20
    RC8      AS STRING * 4
    RMTRTYP8 AS STRING * 1
    REVNAM9  AS STRING * 20
    RC9      AS STRING * 4
    RMTRTYP9 AS STRING * 1
    REVNAM10 AS STRING * 20
    RC10     AS STRING * 4
    RMTRTY10 AS STRING * 1

'Page 4
    SERADDR  AS STRING * 35
    MTRNUM1  AS STRING * 12
    CURRD1   AS DOUBLE
    CRDATE1  AS INTEGER
'Past read date
'Past read amt
'est flag   y/n
    MULT1    AS INTEGER
    MTRTYP1  AS STRING * 1
    MTRUNT1  AS STRING * 1
    USER1    AS INTEGER
    INDATE1  AS INTEGER
    MTRNUM2  AS STRING * 12
    CURRD2   AS DOUBLE
    CRDATE2  AS INTEGER
    MULT2    AS INTEGER
    MTRTYP2  AS STRING * 1
    MTRUNT2  AS STRING * 1
    USER2    AS INTEGER
    INDATE2  AS INTEGER
    MTRNUM3  AS STRING * 12
    CURRD3   AS DOUBLE
    CRDATE3  AS INTEGER
    MULT3    AS INTEGER
    MTRTYP3  AS STRING * 1
    MTRUNT3  AS STRING * 1
    USER3    AS INTEGER
    INDATE3  AS INTEGER
    MTRNUM4  AS STRING * 12
    CURRD4   AS DOUBLE
    CRDATE4  AS INTEGER
    MULT4    AS INTEGER
    MTRTYP4  AS STRING * 1
    MTRUNT4  AS STRING * 1
    USER4    AS INTEGER
    INDATE4  AS INTEGER
    MTRNUM5  AS STRING * 12
    CURRD5   AS DOUBLE
    CRDATE5  AS INTEGER
    MULT5    AS INTEGER
    MTRTYP5  AS STRING * 1
    MTRUNT5  AS STRING * 1
    USER5    AS INTEGER
    INDATE5  AS INTEGER
    MTRNUM6  AS STRING * 12
    CURRD6   AS DOUBLE
    CRDATE6  AS INTEGER
    MULT6    AS INTEGER
    MTRTYP6  AS STRING * 1
    MTRUNT6  AS STRING * 1
    USER6    AS INTEGER
    INDATE6  AS INTEGER
    MTRNUM7  AS STRING * 12
    CURRD7   AS DOUBLE
    CRDATE7  AS INTEGER
    MULT7    AS INTEGER
    MTRTYP7  AS STRING * 1
    MTRUNT7  AS STRING * 1
    USER7    AS INTEGER
    INDATE7  AS INTEGER
    AMTOWED  AS SINGLE
    AMTPD    AS SINGLE
    PAYMENT  AS SINGLE
    REVSRCE  AS INTEGER
    REFMFEE  AS DOUBLE
    MFEENORF AS DOUBLE
'end of actual quick screen form
    LockedFlag AS INTEGER
    PendingRec AS INTEGER
END TYPE



TYPE ubmeterrectype
   Location AS STRING * 9
   SerialNumb AS STRING * 15
   Status AS STRING * 1         'A-Active : V-Vacant : L-Locked : P-Pending : R-Removed
   CurCustAccount AS SINGLE
   PendCustAccount AS SINGLE
   ServiceAddress AS STRING * 35
   RevenueSource AS INTEGER
   HaveIBeenRead AS STRING * 1  ' Y/N
   CurReading AS DOUBLE
   CurReadDate AS STRING * 6
   CurReadEst AS STRING * 1     ' Y/N
   PastReading AS DOUBLE
   PastReadDate AS STRING * 6
   PastReadEst AS STRING * 1    ' Y/N
   SequenceNumb AS DOUBLE
   Multiplier AS SINGLE
   Users AS INTEGER             ' Number of Minimums
   SourcePumpNumber AS INTEGER
   UnitofMeasure AS STRING * 1
   SensusType AS STRING * 1     ' B=Touch Read  M=Manual   if applicable
   HHNote1 AS STRING * 20
   HHNote2 AS STRING * 20
   HHNote3 AS STRING * 20
   NextMeter AS SINGLE
   Unused AS STRING * 73
END TYPE

TYPE ubworkorderrectype
   CustomerAcct AS SINGLE
   WorkOrderDate AS STRING * 6
   WorkCodeType AS STRING * 2
   WorkOrderOperator AS INTEGER
   WorkOrderLocation AS STRING * 9  ' Should Be Blank or Valid Meter Location
   WorkOrderToDoLine(1 TO 12) AS STRING * 60
   WorkOrderDoneLine(1 TO 12) AS STRING * 60
   WorkOrderCompleteDate AS STRING * 6
   WorkOrderCompletedBy AS STRING * 30
   WorkOrderComplete AS STRING * 1  'Y/N
   NextWorkOrderRecord AS SINGLE
   Unused AS STRING * 544
END TYPE



TYPE ubmessagerectype
  Messline(1 TO 23) AS STRING * 70
  Messdate(1 TO 23) AS STRING * 6
  Unused AS STRING * 300
END TYPE

TYPE ubdraftrectype
    BankName AS STRING * 30
    BankLocation AS STRING * 30
    BankTransit AS STRING * 9       'Must be 9 digits or record is invalid
    BankAcctNumber AS STRING * 17
    CustomerAcct AS SINGLE
    PreNote AS STRING * 1       'Rem Y/N to indicate prenote sent
    Unused AS STRING * 165
END TYPE

TYPE ubraterectype
    RateCode AS STRING * 2
    RateDesc AS STRING * 30
    MinAmount AS SINGLE
    MinUnit AS DOUBLE
    StepBeg(1 TO 9) AS DOUBLE
    StepEnd(1 TO 9) AS DOUBLE
    StepAmt(1 TO 9) AS SINGLE
    UnitofMeasure AS STRING * 1
    OverUnit AS DOUBLE
    OverAmount AS SINGLE
    RevenueCode AS INTEGER
    Unused AS STRING * 17
END TYPE

