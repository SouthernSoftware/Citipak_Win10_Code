DEFINT A-Z



  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))

'*********Setup the owner control arrays
  OwnNumFlds = LibNumberOfFields("UB", "UBOWNER")
  REDIM OwnFrm(1) AS FormInfo
  REDIM OwnForm$(OwnNumFlds, 2)
  REDIM OwnFld(OwnNumFlds) AS FieldInfo

  OwnStartEL = 0
  LibGetFldDef "UB", "UBOWNER", OwnStartEL, OwnFld(), OwnForm$(), ErrCode
  OwnForm$(0, 0) = SPACE$(UBOwnerRecLen)
'***************



LoadOwnerInfo:
  UBFile = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBFile LEN = UBOwnerRecLen
  GET UBFile, RecNo&, UBOwnerRec(1)
  CLOSE
RETURN

CONST BlankLng& = -2147483647                   '       "        long int.
