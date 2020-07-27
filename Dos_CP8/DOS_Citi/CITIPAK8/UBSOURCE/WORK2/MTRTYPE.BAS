DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
  
  '$INCLUDE: 'newcust.BI'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))


  CLS
'  PRINT
'  PRINT
'  PRINT "Processing Reading Flags. . ."
'  PRINT
'  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    'LOCATE , 1
    'PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    IF UBCustRec(1).Serv(2).RMTRTYPE = "S" THEN
      IF LEN(QPTrim$(UBCustRec(1).Serv(1).RMTRTYPE)) = 0 THEN
         'UBCustRec(1).LocMeters(2).MTRNUM = UBCustRec(1).LocMeters(1).MTRNUM
         'UBCustRec(1).LocMeters(2).MTRMulti = UBCustRec(1).LocMeters(1).MTRMulti
         'UBCustRec(1).LocMeters(2).MTRType = UBCustRec(1).LocMeters(1).MTRType
         'UBCustRec(1).LocMeters(2).MTRUnit = UBCustRec(1).LocMeters(1).MTRUnit
         'UBCustRec(1).LocMeters(2).NumUser = UBCustRec(1).LocMeters(1).NumUser
         'UBCustRec(1).LocMeters(2).InsDate = UBCustRec(1).LocMeters(1).InsDate
         'UBCustRec(1).LocMeters(2).CurRead = UBCustRec(1).LocMeters(1).CurRead
         'UBCustRec(1).LocMeters(2).PrevRead = UBCustRec(1).LocMeters(1).PrevRead
         'UBCustRec(1).LocMeters(2).CurDate = UBCustRec(1).LocMeters(1).CurDate
         'UBCustRec(1).LocMeters(2).PastDate = UBCustRec(1).LocMeters(1).PastDate
         'UBCustRec(1).LocMeters(2).ReadFlag = UBCustRec(1).LocMeters(1).ReadFlag
         'UBCustRec(1).LocMeters(2).AvgUse = UBCustRec(1).LocMeters(1).AvgUse
         'UBCustRec(1).LocMeters(2).UseCnt = UBCustRec(1).LocMeters(1).UseCnt
         '
         'UBCustRec(1).LocMeters(1).MTRNUM = UBCustRec(1).LocMeters(3).MTRNUM
         'UBCustRec(1).LocMeters(1).MTRMulti = UBCustRec(1).LocMeters(3).MTRMulti
         'UBCustRec(1).LocMeters(1).MTRType = UBCustRec(1).LocMeters(3).MTRType
         'UBCustRec(1).LocMeters(1).MTRUnit = UBCustRec(1).LocMeters(3).MTRUnit
         'UBCustRec(1).LocMeters(1).NumUser = UBCustRec(1).LocMeters(3).NumUser
         'UBCustRec(1).LocMeters(1).InsDate = UBCustRec(1).LocMeters(3).InsDate
         'UBCustRec(1).LocMeters(1).CurRead = UBCustRec(1).LocMeters(3).CurRead
         'UBCustRec(1).LocMeters(1).PrevRead = UBCustRec(1).LocMeters(3).PrevRead
         'UBCustRec(1).LocMeters(1).CurDate = UBCustRec(1).LocMeters(3).CurDate
         'UBCustRec(1).LocMeters(1).PastDate = UBCustRec(1).LocMeters(3).PastDate
         'UBCustRec(1).LocMeters(1).ReadFlag = UBCustRec(1).LocMeters(3).ReadFlag
         'UBCustRec(1).LocMeters(1).AvgUse = UBCustRec(1).LocMeters(3).AvgUse
         'UBCustRec(1).LocMeters(1).UseCnt = UBCustRec(1).LocMeters(3).UseCnt
         SWAP UBCustRec(1).LocMeters(2), UBCustRec(1).LocMeters(1)
         PRINT UBCustRec(1).CustName
         PUT UBFile, Cnt&, UBCustRec(1)
      END IF
    END IF
    'FOR MCnt = 1 TO 7
      'IF UBCustRec(1).LocMeters(2).MTRType = "S" THEN
      '  STOP
      'END IF
    'NEXT
    'PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE UBFile

'PRINT
'PRINT
'PRINT
'PRINT "Done."

