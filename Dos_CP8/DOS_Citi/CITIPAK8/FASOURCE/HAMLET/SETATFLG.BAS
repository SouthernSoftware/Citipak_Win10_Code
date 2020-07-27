DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newfa.BI'

  CLS
  REDIM FAItemRec(1) AS FAItemRecType
  FAItemRecLen = LEN(FAItemRec(1))

  FAFile = FREEFILE
  OPEN FAItemFile FOR RANDOM AS FAFile LEN = FAItemRecLen
  NumOfRec = LOF(FAFile) / FAItemRecLen
  PRINT "Setting GL-Accounts & Depreciation Flag. . ."
  PRINT
  PRINT

  FOR zz = 1 TO NumOfRec
    LOCATE , 1: PRINT "Processing:"; zz; " of"; NumOfRec;
    GET FAFile, zz, FAItemRec(1)
    ItemOrg# = FAItemRec(1).OrgCost
    Dept$ = QPTrim$(FAItemRec(1).IDEPT)
    IF ItemOrg# < 2500 THEN
      FAItemRec(1).DEPYN = "N"
    END IF
    IF LEN(Dept$) > 0 THEN
      SELECT CASE Dept$
      CASE "0410", "0420", "0440", "0500", "0510", "0530", "0560", "0580", "0611", "0620", "0630", "0660"
        FAItemRec(1).GLACCT = "10-228-0000"
      CASE "0570"
        FAItemRec(1).GLACCT = "11-228-0000"
      CASE "0720", "0810", "0820", "0830", "0840"
        FAItemRec(1).GLACCT = "30-228-0000"
      END SELECT
    END IF

    PUT FAFile, zz, FAItemRec(1)
  NEXT

  CLOSE
  PRINT
  PRINT
  PRINT "Done."
  CALL KillFile("setATflg.exe")

