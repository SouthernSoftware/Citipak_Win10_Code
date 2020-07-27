DECLARE FUNCTION NRound# (N#)
DECLARE SUB ReIndexSystem ()
DEFINT A-Z

DECLARE FUNCTION NRound# (DblNum#)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'PageInfo.BI'                  'Form Page info
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  ''$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  ''$INCLUDE: 'ubcust.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubLoc.bi'
  '$INCLUDE: 'ubTrans.BI'
  '$INCLUDE: 'oubTypes.BI'

  CONST False = 0, True = NOT False
  
  REDIM UBCustRec(1) AS UBCustRecTypeV1
  UBCustRecLen = LEN(UBCustRec(1))
  REDIM UBLocaRec(1 TO 2) AS UBLocationRecTypeV1
  UBLocaRecLen = LEN(UBLocaRec(1))
  REDIM NUBCustRec(1 TO 2) AS NewUBCustRecType
  NUBCustRecLen = LEN(NUBCustRec(1))
  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))

  CLS
'  PRINT "New:"; NUBCustRecLen
'  PRINT "Old:"; UBCustRecLen + UBLocaRecLen
'  END

  UBLocaFile = FREEFILE
  OPEN "UBLOCA.DAT" FOR RANDOM SHARED AS UBLocaFile LEN = UBLocaRecLen
  UBCustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
  NUBCustFile = FREEFILE
  OPEN "NUBCUST.DAT" FOR RANDOM SHARED AS NUBCustFile LEN = NUBCustRecLen
  UBTransFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTransFile LEN = UBTransRecLen

'  NUBTransFile = FREEFILE
'  OPEN "NUBTRANS.DAT" FOR RANDOM SHARED AS NUBTransFile LEN = UBTransRecLen

  NumOfLocaRecs = LOF(UBLocaFile) / UBLocaRecLen
  NumOfCustRecs = LOF(UBCustFile) / UBCustRecLen

  FOR RecCnt = 1 TO NumOfLocaRecs
    NewRecCnt& = NewRecCnt& + 1
    GET UBLocaFile, RecCnt, UBLocaRec(1)
    IF UBLocaRec(1).CustRec > 0 THEN
      GotCustFlag = True
      GET UBCustFile, UBLocaRec(1).CustRec, UBCustRec(1)
      GOSUB DoSetVariables
      IF UBCustRec(1).LastTrans > 0 THEN
        GOSUB DoTransFix
      END IF
    ELSE
      GotCustFlag = False
      GOSUB DoSetVariables
    END IF
    PUT NUBCustFile, NewRecCnt&, NUBCustRec(1)
    LOCATE 12, 1: PRINT RecCnt;
  NEXT
  LSET UBLocaRec(1) = UBLocaRec(2)
  FOR RecCnt = 1 TO NumOfCustRecs
    GET UBCustFile, RecCnt, UBCustRec(1)
    IF UBCustRec(1).CustFrstLocRec = 0 AND UBCustRec(1).LastTrans > 0 THEN
      GotCustFlag = True
      NewRecCnt& = NewRecCnt& + 1
      WithoutLoca = WithoutLoca + 1
      GOSUB DoSetVariables
      GOSUB DoTransFix
      PUT NUBCustFile, NewRecCnt&, NUBCustRec(1)
      LOCATE 16, 1: PRINT RecCnt;
    END IF
  NEXT
  LOCATE 21, 1: PRINT "Without Location Converted:"; WithoutLoca
CLOSE
  KILL "UBLOCA.DAT"
  KILL "UBCUST.DAT"
  NAME "NUBCUST.DAT" AS "UBCUST.DAT"
  LOCATE 23, 1: PRINT "Press any key to continue. . .";
  a$ = INPUT$(1)
  'CALL WaitForAction
  ERASE UBCustRec, UBLocaRec, NUBCustRec, UBTransRec
  'ReIndexSystem

END 'Program

DoTransFix:
   TCnt = 0
   ThisTrans& = UBCustRec(1).LastTrans
   DO
     TCnt = TCnt + 1
     GET UBTransFile, ThisTrans&, UBTransRec(1)
     UBTransRec(1).CustAcctNo = NewRecCnt&
     PUT UBTransFile, ThisTrans&, UBTransRec(1)
     ThisTrans& = UBTransRec(1).PrevTrans
     LOCATE 15, 1: PRINT TCnt;
   LOOP WHILE ThisTrans& > 0
RETURN


DoSetVariables:
  LSET NUBCustRec(1) = NUBCustRec(2)
  NUBCustRec(1).BOOK = UBLocaRec(1).BOOK
  NUBCustRec(1).SEQNUMB = UBLocaRec(1).SEQNUMB
  NUBCustRec(1).SERVADDR = UBLocaRec(1).SERVADDR
  NUBCustRec(1).BILLCYCL = UBLocaRec(1).BILLCYCL
  NUBCustRec(1).ZONE = UBLocaRec(1).ZONE
  NUBCustRec(1).SEQ = UBLocaRec(1).SEQ
  NUBCustRec(1).PUMPCODE = UBLocaRec(1).PUMPCODE
  NUBCustRec(1).USERCODE1 = UBLocaRec(1).USERCODE1
  NUBCustRec(1).USERCODE2 = UBLocaRec(1).USERCODE2
  NUBCustRec(1).HHMSG1 = ""
  NUBCustRec(1).HHMSG2 = ""
  NUBCustRec(1).HHMSG3 = ""
  FOR SCnt = 1 TO 15
    LSET NUBCustRec(1).Serv(SCnt) = UBLocaRec(1).Serv(SCnt)
  NEXT
  FOR SCnt = 1 TO 4
    LSET NUBCustRec(1).FlatRates(SCnt) = UBLocaRec(1).FlatRates(SCnt)
  NEXT
  NUBCustRec(1).MFEE1 = UBLocaRec(1).MFEE1
  NUBCustRec(1).MFEE2 = UBLocaRec(1).MFEE2
  FOR SCnt = 1 TO 7
    LSET NUBCustRec(1).LocMeters(SCnt) = UBLocaRec(1).LocMeters(SCnt)
  NEXT
  'FOR SCnt = 1 TO 15
  '  NUBCustRec(1).PrevRevAmts(SCnt) = 0
  'NEXT

  IF GotCustFlag THEN
    NUBCustRec(1).STATUS = UBCustRec(1).STATUS
    NUBCustRec(1).OPENDATE = UBCustRec(1).OPENDATE
    NUBCustRec(1).SEARCH = UBCustRec(1).SEARCH
    NUBCustRec(1).CUSTNAME = UBCustRec(1).CUSTNAME
    NUBCustRec(1).ADDR1 = UBCustRec(1).ADDR1
    NUBCustRec(1).ADDR2 = UBCustRec(1).ADDR2
    NUBCustRec(1).CITY = UBCustRec(1).CITY
    NUBCustRec(1).STATE = UBCustRec(1).STATE
    NUBCustRec(1).ZIPCODE = UBCustRec(1).ZIPCODE
    NUBCustRec(1).HPHONE = UBCustRec(1).HPHONE
    NUBCustRec(1).WPHONE = UBCustRec(1).WPHONE
    NUBCustRec(1).SOSEC = UBCustRec(1).SOSEC
    NUBCustRec(1).DRVLIC = UBCustRec(1).DRVLIC
    NUBCustRec(1).CUSTTYPE = UBCustRec(1).CUSTTYPE
    NUBCustRec(1).BILLCOPY = UBCustRec(1).BILLCOPY
    NUBCustRec(1).PostRte = UBCustRec(1).PostRte
    NUBCustRec(1).CASHONLY = UBCustRec(1).CASHONLY
    NUBCustRec(1).LATEFEE = UBCustRec(1).LATEFEE
    NUBCustRec(1).CUTOFFYN = UBCustRec(1).CUTOFFYN
    NUBCustRec(1).TAXEXPT = UBCustRec(1).TAXEXPT
    NUBCustRec(1).SRCIT = UBCustRec(1).SRCIT
    NUBCustRec(1).EPP = UBCustRec(1).EPP
    NUBCustRec(1).EPPAMT = UBCustRec(1).EPPAMT
    NUBCustRec(1).USEDRAFT = UBCustRec(1).DRAFTYPE
    NUBCustRec(1).BANKNAME = UBCustRec(1).BANKNAME
    NUBCustRec(1).BANKLOC = UBCustRec(1).BANKLOC
    NUBCustRec(1).TRANSIT = UBCustRec(1).TRANSIT
    NUBCustRec(1).BANKACCT = UBCustRec(1).BANKACCT
    NUBCustRec(1).BILLCMNT = UBCustRec(1).BILLCMNT
    NUBCustRec(1).PAYCMNT = UBCustRec(1).PAYCMNT
    FOR SCnt = 1 TO 2
      LSET NUBCustRec(1).Monthly(SCnt) = UBCustRec(1).Monthly(SCnt)
    NEXT
    NUBCustRec(1).CustPIN = UBCustRec(1).CustPIN
    NUBCustRec(1).LastTrans = UBCustRec(1).LastTrans
    NUBCustRec(1).CurrBalance = UBCustRec(1).CurrBalance
    NUBCustRec(1).PrevBalance = UBCustRec(1).PrevBalance
    FOR SCnt = 1 TO 15
      NUBCustRec(1).CurrRevAmts(SCnt) = NRound#(UBCustRec(1).CurrRevAmts(SCnt))
    NEXT
    NUBCustRec(1).DepositAmt = UBCustRec(1).DepositAmt
    NUBCustRec(1).DelFlag = UBCustRec(1).DelFlag
    NUBCustRec(1).PreNoteFlag = UBCustRec(1).PreNoteFlag
    NUBCustRec(1).WOLastTrans = UBCustRec(1).WOLastTrans
  ELSE
    NUBCustRec(1).STATUS = "I"
    NUBCustRec(1).OPENDATE = -32767
    NUBCustRec(1).SEARCH = "VACANT"
    NUBCustRec(1).CUSTNAME = "VACANT"
    NUBCustRec(1).ADDR1 = ""
    NUBCustRec(1).ADDR2 = ""
    NUBCustRec(1).CITY = ""
    NUBCustRec(1).STATE = ""
    NUBCustRec(1).ZIPCODE = ""
    NUBCustRec(1).HPHONE = ""
    NUBCustRec(1).WPHONE = ""
    NUBCustRec(1).SOSEC = ""
    NUBCustRec(1).DRVLIC = ""
    NUBCustRec(1).CUSTTYPE = ""
    NUBCustRec(1).BILLCOPY = 1
    NUBCustRec(1).PostRte = ""
    NUBCustRec(1).CASHONLY = "N"
    NUBCustRec(1).LATEFEE = "Y"
    NUBCustRec(1).CUTOFFYN = "Y"
    NUBCustRec(1).TAXEXPT = "N"
    NUBCustRec(1).SRCIT = "N"
    NUBCustRec(1).EPP = "N"
    NUBCustRec(1).EPPAMT = 0
    NUBCustRec(1).USEDRAFT = "N"
    NUBCustRec(1).BANKNAME = ""
    NUBCustRec(1).BANKLOC = ""
    NUBCustRec(1).TRANSIT = ""
    NUBCustRec(1).BANKACCT = ""
    NUBCustRec(1).BILLCMNT = ""
    NUBCustRec(1).PAYCMNT = ""
    FOR SCnt = 1 TO 2
      NUBCustRec(1).Monthly(SCnt).AMTOWED = 0
      NUBCustRec(1).Monthly(SCnt).TotAmtPD = 0
      NUBCustRec(1).Monthly(SCnt).PayAmt = 0
      NUBCustRec(1).Monthly(SCnt).RevSource = 0
    NEXT
    NUBCustRec(1).CustPIN = 0
    NUBCustRec(1).LastTrans = 0
    NUBCustRec(1).CurrBalance = 0
    NUBCustRec(1).PrevBalance = 0
    FOR SCnt = 1 TO 15
      NUBCustRec(1).CurrRevAmts(SCnt) = 0
    NEXT
    NUBCustRec(1).DepositAmt = 0
    NUBCustRec(1).DelFlag = 0
    NUBCustRec(1).PreNoteFlag = 0
    NUBCustRec(1).WOLastTrans = 0

  END IF
RETURN

  '****************************************************************************
  'Rounds a double precision value to nearest hundreth
  '****************************************************************************
FUNCTION NRound# (N#)
  NRound# = INT(N# * 100 + .5) / 100
END FUNCTION

