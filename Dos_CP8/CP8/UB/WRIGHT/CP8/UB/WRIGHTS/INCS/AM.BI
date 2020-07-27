'OPEN "apvendor.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 204
TYPE VendorAMRecType
  VenName AS STRING * 30      '30 AS venname$
  Address AS STRING * 30      '30 AS address$
  City AS STRING * 30         '22 AS CitY$
  State AS STRING * 2         '2 AS St$
  Zip AS STRING * 10          '10 AS ZiP$
  PayTo AS STRING * 30        '30 AS payto$
  PayAdd AS STRING * 30       '30 AS payadd$
  PayCity AS STRING * 22      '22 AS paycity$
  PaySt AS STRING * 2         '2 AS payst$
  PayZip AS STRING * 10       '10 AS payzip$
  CurBal AS DOUBLE            '8 AS curbal$
  First AS SINGLE             '4 AS first$
  Last AS SINGLE              '4 AS last$
END TYPE

