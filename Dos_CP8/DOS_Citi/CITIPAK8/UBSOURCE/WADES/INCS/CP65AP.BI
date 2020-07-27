'AP.BI  Contains current Citipak AP*.DAT file structures.

'apcontrl.dat
'chkjour.dat
'apckproc.dat
'APDIST.DAT
'APIF.dat
'apinvdis.dat
'apledger.dat
'apledgr1.dat
'appo.dat
'apvenddx.dat
'apvendor.dat
'apvendor1.dat
'crchk.dat
'tempinv.dat

'OPEN "apvenddx.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 14
TYPE Vendor7IdxRecType
  VendorCode AS STRING * 10           '10 AS act$
  RecNum AS SINGLE                    '4 AS REC$
END TYPE

'OPEN "apvendor.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 204
TYPE Vendor6RecType
 VenName AS STRING * 30  '30 AS venname$,
 Address AS STRING * 30  '30 AS address$,
 City AS STRING * 22     '22 AS CitY$,
 St AS STRING * 2        '2 AS St$,
 Zip AS STRING * 10      '10 AS ZiP$,
 Payto AS STRING * 30    '30 AS payto$,
 payadd AS STRING * 30   '30 AS payadd$,
 paycity AS STRING * 22  '22 AS paycity$,
 payst  AS STRING * 2    '2 AS payst$,
 payzip AS STRING * 10   '10 AS payzip$,
 curbal AS DOUBLE        '8 AS CURBAL$,
 first  AS SINGLE        '4 AS first$,
 last   AS SINGLE        '4 AS last$
END TYPE


TYPE Vendor7RecType
 VenName AS STRING * 30  '30 AS venname$,
 Address AS STRING * 30  '30 AS address$,
 City AS STRING * 22     '22 AS CitY$,
 St AS STRING * 2        '2 AS St$,
 Zip AS STRING * 10      '10 AS ZiP$,
 Payto AS STRING * 30    '30 AS payto$,
 payadd AS STRING * 30   '30 AS payadd$,
 paycity AS STRING * 22  '22 AS paycity$,
 payst  AS STRING * 2    '2 AS payst$,
 payzip AS STRING * 10   '10 AS payzip$,
 curbal AS DOUBLE        '8 AS CURBAL$,
 first  AS SINGLE        '4 AS first$,
 last   AS SINGLE        '4 AS last$
END TYPE



TYPE VendorLVRecType
 VenName AS STRING * 40  '30 AS venname$,
 Address AS STRING * 40  '30 AS address$,
 City AS STRING * 22     '22 AS CitY$,
 St AS STRING * 2        '2 AS St$,
 Zip AS STRING * 10      '10 AS ZiP$,
 Payto AS STRING * 40    '30 AS payto$,
 payadd AS STRING * 40   '30 AS payadd$,
 paycity AS STRING * 22  '22 AS paycity$,
 payst  AS STRING * 2    '2 AS payst$,
 payzip AS STRING * 10   '10 AS payzip$,
 curbal AS DOUBLE        '8 AS CURBAL$,
 first  AS SINGLE        '4 AS first$,
 last   AS SINGLE        '4 AS last$
 Addr1  AS STRING * 40
 pad AS STRING * 100
END TYPE

'OPEN "apvendor.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 384
TYPE VendorSetRecType
  VNum AS STRING * 10                 '10 AS VENNUMBER$
  VName AS STRING * 30                '30 AS venname$
  Addr1 AS STRING * 30                '30 AS addRESS$
  City AS STRING * 22                 '22 AS CitY$
  State AS STRING * 2                 '2 AS St$
  Zip AS STRING * 10                  '10 AS ZiP$
  PayToName AS STRING * 30            '30 AS payto$
  PayToAddr AS STRING * 30            '30 AS payadd$
  PayToCity AS STRING * 22            '22 AS paycity$
  PaytoState AS STRING * 2            '2 AS payst$
  PaytoZip AS STRING * 10             '10 AS payzip$
  CurrBal AS DOUBLE                   '8 AS curbal$
  FrstTran AS SINGLE                  '4 AS first$
  LastTran AS SINGLE                  '4 AS last$
  Addr2 AS STRING * 30                '30 AS address2$
  PaytoAddr2 AS STRING * 30           '30 AS paddress2$
  pad AS STRING * 110
END TYPE

TYPE APVendorGQRecType
  VenNum AS INTEGER        '  2 AS VENNUMBER$,
  NotUsed AS STRING * 8    '  8 AS notused$,
  VenName AS STRING * 30   ' 30 AS venname$,
  Address AS STRING * 30   ' 30 AS addRESS$,
  City AS STRING * 22      ' 22 AS CitY$,
  St AS STRING * 2         '  2 AS St$,
  Zip AS STRING * 10       ' 10 AS ZiP$,
  Payto AS STRING * 30     ' 30 AS payto$,
  payadd AS STRING * 30    ' 30 AS payadd$,
  paycity AS STRING * 22   ' 22 AS paycity$,
  payst AS STRING * 2      '  2 AS payst$,
  payzip AS STRING * 10    ' 10 AS payzip$,
  curbal AS STRING * 8     '  8 AS curbal$,
  first AS SINGLE          '  4 AS first$,
  last AS SINGLE           '  4 AS last$,
  address2 AS STRING * 30  ' 30 AS address2$,
  Paddress2 AS STRING * 30 ' 30 AS paddress2$
  pad AS STRING * 110
END TYPE

'OPEN "apvendr1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #7 LEN = 256
TYPE Vendor1RecType
  Phone AS STRING * 12                '12 AS TEL$
  YTDPurch AS SINGLE                  '4 AS YTDPUR$
  YTDPay AS SINGLE                    '4 AS ytdpay$
  FedID AS STRING * 12                '12 AS FEDID$
  CoCode AS STRING * 3                '3 AS CTYCODE$
  StCode AS STRING * 2                '2 AS STCODE$
  FrstPO AS SINGLE                    '4 AS POFIRST$
  LastPO AS SINGLE                    '4 AS POLAST$
  FrstAPDist AS SINGLE                '4 AS firstAPdist$
  Get1099 AS STRING * 1               'some vers don't have 1099 field
  filler AS STRING * 206              '207 AS bk$
END TYPE
'--see ap.bi for ledger rec types
'OPEN "R", 2, "apledger.dat", 96
'TYPE APLedger7RecType
'  RecNo AS INTEGER                 '2 AS vennum$
'  TrDate AS STRING * 8             '8 AS INVDATE$
'  InvNum AS STRING * 25           '25 AS invnum$
'  PONum AS STRING * 10             '10 AS ponum$
'  Amt AS DOUBLE                    '8 AS amount$
'  PayCode AS INTEGER               '2 AS PAYCODE$
'  FirstDist AS SINGLE              '4 AS fdist$
'  LastDist AS SINGLE               '4 AS ldist$
'  Ndata AS SINGLE                  '4 AS ndata$
'  CoTaxCode AS STRING * 3          '3 AS tctycode$
'  StTaxCode AS STRING * 3          '3 AS tstcode$
'  CoTaxAmt AS SINGLE               '4 AS ctaxamt$
'  StTaxAmt AS SINGLE               '4 AS staxamt$
'END TYPE

'OPEN "R", 3, "appo.dat", 64
TYPE PORecType
  PODate AS STRING * 8             '8 AS podate$
  POVendorNumb AS STRING * 10      '10 AS povend$
  PONumb AS STRING * 10            '10 AS ponumB$
  POCost AS STRING * 10            '8 AS pocost$
  POFrstDist AS SINGLE             '4 AS podfirst$
  POLastDist AS SINGLE             '4 AS podLAST$
  PONext AS SINGLE                 '4 AS PONEXT$
  Empty AS STRING * 16             '16 AS empty$
END TYPE


'OPEN "R", 5, "APIF.dat", 85
TYPE APIFRecType
  GLAcct AS STRING * 9       '9 AS tranacct$
  TranDate AS STRING * 8     '8 AS trandate$
  Desc AS STRING * 20        '20 AS trandesc$
  CrAmt AS DOUBLE            '8 AS cramt$
  DrAmt AS DOUBLE            '8 AS dramt$
  Ref AS STRING * 8          '8 AS detail$
  Src AS STRING * 8          '8 AS source$
  filler AS STRING * 16      '4 AS nexttr$
END TYPE

'OPEN "R", 6, "tempinv.dat", 128
TYPE TempInvRecType
  TmpInvDate AS STRING * 8      '8 AS tidate$
  TmpInvVendNum AS STRING * 10  '10 AS TIVEND$
  TmpInvDet AS STRING * 25      '25 AS tidet$
  TmpInvCost AS DOUBLE          '8 AS ticost$
  TmpPayCode AS INTEGER         '2 AS TIPAY$
  TmpInvNote AS STRING * 10     '10 AS tinote$
  TmpFrstDist AS SINGLE         '4 AS TFDIST$
  TmpLastDist AS SINGLE         '4 AS tldist$
  StTax AS SINGLE               '4 AS STAX$
  CtyTax AS SINGLE              '4 AS CTAX$
  DistDate AS STRING * 8        '8 AS DISTdate$
  DueDate AS STRING * 8         '8 AS duedate$:
END TYPE

'OPEN "chkjour.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #8 LEN = 64
TYPE CheckJrnlRecType
  ChkNumb AS SINGLE          '4 AS checknumber$
  ChkAmt AS DOUBLE           '8 AS checkamount$
  ChkName AS STRING * 30     '30 AS checkname$
  VendorNum AS STRING * 10   '10 AS vnumber$
END TYPE

'OPEN "apckproc.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #9 LEN = 16
TYPE APCkProcRecType
  VRec AS SINGLE             '4 AS vrec$
  JRec AS INTEGER            '2 AS jrec$
  Void AS STRING * 1         '1 AS void$
  ChkNum AS SINGLE           '4 AS apcknum$
END TYPE

'OPEN "R", 11, "apinvdis.dat", 13
TYPE APVendorDefDistRecType
  APAcct AS STRING * 9          '9 AS apacct$
  APDefPct AS SINGLE            '4 AS apperc$
END TYPE

'OPEN "crchk.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #12 LEN = 64
TYPE OutCheckRecType
   CheckNum AS SINGLE         '4 AS chknum$
   CheckDate AS STRING * 8    '8 AS chkdate$
   CheckDesc AS STRING * 30   '30 AS chkdesc$
   CheckAmt AS SINGLE         '4 AS chkamt$
   CheckSrc AS STRING * 2      '2 AS chksource$
   Blank AS STRING * 16       '16 AS nul$
END TYPE

'OPEN "apledgr1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #12 LEN = 32
TYPE APLedger1RecType
   APDistDate AS STRING * 8         '8 AS apdistdate$
   APDueDate AS STRING * 8          '8 AS apduedate$
END TYPE

'OPEN "apcontrl.dat" FOR RANDOM ACCESS READ WRITE AS #13 LEN = 64
TYPE APControlRecType
   APAcctNumb AS STRING * 7               '7 AS apnumb$
   CashAcctNumb AS STRING * 7             '7 AS cashnumb$
   StaTaxAcctNum AS STRING * 9            '9 AS stsalestax$
   CtyTaxAcctNum AS STRING * 9            '9 AS CTSALESTAX$
   NotUsed AS STRING * 32                 '32 AS notused$
END TYPE

'OPEN "r", 15, "APDIST.DAT", 33
TYPE nAPDistRecType
   TranRecNum  AS SINGLE         '4 AS trrec$
   DistAcctNum AS STRING * 9     '9 AS daacct$
   DistCrAmt   AS DOUBLE         '8 AS DCRAMT$
   DistDrAmt    AS DOUBLE        '8 AS DDRAMT$
   NextDist    AS SINGLE         '4 AS DNEX$
END TYPE


