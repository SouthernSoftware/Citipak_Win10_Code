DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)

  '$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.

  '$INCLUDE: 'PRFiles.bi'
  '$INCLUDE: 'PREmpRec.bi'



  CONST False = 0, True = NOT False

  DIM SHARED Emp1Rec   AS EmpData1Type
  DIM SHARED Emp2Rec   AS EmpData2Type
  DIM SHARED Emp3Rec   AS EmpData3Type
  DIM SHARED Emp3RecB  AS EmpData3Type

  Emp1RecLen = LEN(Emp1Rec)
  Emp2RecLen = LEN(Emp2Rec)
  Emp3RecLen = LEN(Emp3Rec)


  E1NumOfRec = FileSize(EmpData1Name) \ Emp1RecLen
  E2NumOfRec = FileSize(EmpData2Name) \ Emp2RecLen
  E3NumOfRec = FileSize(EmpData3Name) \ Emp3RecLen


PRINT E1NumOfRec
PRINT E2NumOfRec
PRINT E3NumOfRec

