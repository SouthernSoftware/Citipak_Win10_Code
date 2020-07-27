DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

'$INCLUDE: 'PREmpRec.bi'
'$INCLUDE: 'PRFiles.bi'

  REDIM TEmp3Rec(1) AS EmpData3Type
  Emp3RecLen = LEN(TEmp3Rec(1))

  REDIM TEmp1Rec(1) AS EmpData1Type
  Emp1RecLen = LEN(TEmp1Rec(1))

  OPEN EmpData1Name FOR RANDOM AS #2 LEN = Emp1RecLen
  OPEN EmpData3Name FOR RANDOM AS #1 LEN = Emp3RecLen

  NumOfRec = LOF(2) \ Emp1RecLen

  FOR Cnt = 1 TO NumOfRec

    GET #2, Cnt, TEmp1Rec(1)
    EmpNumb = VAL(QPTrim$(TEmp1Rec(1).EMPNO))
    SELECT CASE EmpNumb
      CASE 50, 125, 135, 140
        GET #1, Cnt, TEmp3Rec(1)
        SELECT CASE EmpNumb
          CASE 50
            TEmp3Rec(1).YTDSocGrossPay = 13105.97
            TEmp3Rec(1).YTDMedGrossPay = 13105.97
          CASE 125
            TEmp3Rec(1).YTDSocGrossPay = 30973.89
            TEmp3Rec(1).YTDMedGrossPay = 30973.89
          CASE 135
            TEmp3Rec(1).YTDSocGrossPay = 14615.9
            TEmp3Rec(1).YTDMedGrossPay = 14615.9
          CASE 140
            TEmp3Rec(1).YTDSocGrossPay = 10612.26
            TEmp3Rec(1).YTDMedGrossPay = 10612.26
          END SELECT
        PUT #1, Cnt, TEmp3Rec(1)
        'GET #1, Cnt, TEmp3Rec(1)
    END SELECT
  NEXT
  CLOSE

