DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)

  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubowner.BI'

  DIM Cust  AS NewUBCustRecType

  DIM Owner AS UBOwnerRecType
  OwnerLen = LEN(Owner)
  
  NumOfRecs = FileSize("UBCUST.DAT") \ LEN(Cust)

  'KILL "UBOWNER.DAT"

  OPEN "UBOWNER.DAT" FOR RANDOM AS #1 LEN = OwnerLen
  FOR Cnt = 1 TO NumOfRecs
    PUT #1, Cnt, Owner
  NEXT
  CLOSE

