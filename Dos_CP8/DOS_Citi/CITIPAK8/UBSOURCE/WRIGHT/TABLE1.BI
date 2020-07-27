TYPE BladenRecType
  acctnum         AS STRING * 6
  status          AS STRING * 1

  Cycle           AS STRING * 2
  Fill1           AS STRING * 1

  Book            AS STRING * 2
  Location        AS STRING * 4

  pad1            AS STRING * 120
  closedate       AS STRING * 10
  depbal          AS STRING * 10
  catcode1        AS STRING * 10
  addrkey         AS STRING * 15
  addrdesc        AS STRING * 33
  numunits        AS STRING * 2
  meterloc        AS STRING * 20
  meterunit       AS STRING * 2
  meterid         AS STRING * 10
  prevread        AS STRING * 10
  prevdate        AS STRING * 10
  curread         AS STRING * 10
  curdate         AS STRING * 10
  readmsg         AS STRING * 30
  estbill         AS STRING * 1
  serialno        AS STRING * 16
  dials           AS STRING * 1
  multi           AS STRING * 6
  avguse          AS STRING * 10
  usage           AS STRING * 10
  readstatus      AS STRING * 1
  special         AS STRING * 1
  cFirstName      AS STRING * 10
  CMidI           AS STRING * 1
  cLastName       AS STRING * 15
  BusiName        AS STRING * 33
  LandLord        AS STRING * 33
  BillAddr1       AS STRING * 33
  BillAddr2       AS STRING * 33
  BillCity        AS STRING * 20
  BillST          AS STRING * 2
  BillZip         AS STRING * 10
  LandAddr1       AS STRING * 33
  LandAddr2       AS STRING * 33
  Contact         AS STRING * 33
  refno           AS STRING * 10
  AcctStartDate   AS STRING * 10
  AcctStopDate    AS STRING * 10
  LateCharge      AS STRING * 1
  comment         AS STRING * 60
  lastPayDate     AS STRING * 10
  lastPayAmount   AS STRING * 10
  CrLf            AS STRING * 2
END TYPE

