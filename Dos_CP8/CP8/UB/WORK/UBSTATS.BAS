DEFINT A-Z
DECLARE SUB KILLFILE (filename$)
CONST true = -1, False = NOT true

'$INCLUDE: 'ub2wind5.bi'
'$INCLUDE: 'wadecust.bi'

  DIM NUBCust AS NewUBCustRecType
  DIM DelCnt&, ACnt%, ICnt%, OCnt%, BCnt%
  NRecLen = LEN(NUBCust)

  CLS

  nubCustFile = FREEFILE
  OPEN "ubcust.dat" FOR RANDOM SHARED AS nubCustFile LEN = NRecLen
  NumORecs& = LOF(nubCustFile) / NRecLen
  FOR cnt& = 1 TO NumORecs&
    LOCATE 5, 1
    PRINT "Processing:"; cnt&; "Of"; NumORecs&;
    GET #nubCustFile, cnt&, NUBCust
    IF NUBCust.delflag <> 0 THEN
      DelCnt& = DelCnt& + 1
    ELSE
      SELECT CASE NUBCust.Status
      CASE "I"
        ICnt = ICnt + 1
      CASE "A"
        ACnt = ACnt + 1
      CASE "B"
        BCnt = BCnt + 1
      CASE ELSE
        OCnt = OCnt + 1
      END SELECT
    END IF

    'PUT #nUBCustFile, cnt&, NUBCust
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT " Deleted:"; DelCnt&
  PRINT
  PRINT "  Active:"; ACnt
  PRINT "InActive:"; ICnt
  PRINT " Balance:"; BCnt
  PRINT "   Other:"; OCnt






END

