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
      CASE "410", "420", "440", "500", "510", "530", "560"
        FAItemRec(1).GLACCT = ""
      CASE "570"
      CASE "580", "611", "620", "630", "660"
      CASE "720", "810", "820", "830", "840"
      CASE ELSE



    PUT FAFile, zz, FAItemRec(1)
  NEXT

  CLOSE
  PRINT
  PRINT
  PRINT "Done."
  CALL KillFile("setdflag.exe")

