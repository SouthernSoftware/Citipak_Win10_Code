' CLOSE 1: OPEN "R", 1, "wbcust.dat", 512

TYPE PineBluffCustType
  LName   AS STRING * 25
  FName   AS STRING * 25
  Addr1   AS STRING * 30
  SAddr   AS STRING * 30
  City    AS STRING * 20
  State   AS STRING * 2
  Zip     AS STRING * 9
  Phone   AS STRING * 12   '???
  Pad1    AS STRING * 173
  Addr2   AS STRING * 30
  Pad2    AS STRING * 156
END TYPE

