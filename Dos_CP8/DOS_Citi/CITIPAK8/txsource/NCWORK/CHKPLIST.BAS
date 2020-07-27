DEFINT A-Z

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxCONST.BI'
  '$INCLUDE: 'txpaymnt.BI'
  ''$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'cmfiles.BI'

  REDIM PayList(1 TO 1)      AS PayListType
  PayListLen = LEN(PayList(1))

  PayListFile = FREEFILE
  OPEN "TAXLOP98.DAT" FOR RANDOM AS PayListFile LEN = PayListLen
  FOR Cnt = 1 TO LOF(PayListFile) \ PayListLen
    GET #PayListFile, Cnt, PayList(1)
    PRINT PayList(1).Principle1;
    PRINT PayList(1).Interest1
  NEXT

