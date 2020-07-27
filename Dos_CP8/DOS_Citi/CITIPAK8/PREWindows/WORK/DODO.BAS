DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DEFINT A-Z

DEFINT A-Z
SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, ActiveOnly%)
  
  ShowProcessingScrn "Searching Customers Info."

  STATIC Choice, LastSEARCH$
  
  IF LastSEARCH$ <> SEARCH$ THEN
    LastSEARCH$ = SEARCH$
    Choice = 1
  END IF
  
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  'DisplayUBScrn "SHOWSCRH"
  
  REDIM MChoice(1 TO 1) AS FLen2
  REDIM UBCustRec(1) AS NewUBCustRecType
  
  MaxBlockCnt = 1024            'Max Buff size: 16384   (1024*16)
  REDIM UBCustSN(1 TO MaxBlockCnt) AS nUBCustReIndexRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBCustSNLen = LEN(UBCustSN(1))
  
  SearchLen = LEN(SEARCH$)
  Match = False
  
'This search reads 1024 recs at a pop into a search buffer
'Basics file i/o CAN NOT read in this way.
'For a Cust file of "5000" recs
'Basic will do "5000" disk accesses.
'Our search ONLY NEEDS 5!
'5000 Recs, BlockSize 1024.
' 4reads * 1024 = 4096 Recs
'   5000 - 4096 = 904 Odd Recs
' 1read  * 904 Recs
' 4 + 1 = 5 Reads!
'This is many many times faster than Basic can do.
  
  FOpenS "UBCUSTSN.DAT", C1Handle               'open data file
  FOpenS "UBCUST.DAT", R1Handle 'open data file
  
  NumOfCust& = FLof&(C1Handle) / UBCustSNLen
  
'************************************
  NumChunks& = NumOfCust& \ MaxBlockCnt
'****DO NOT CHANGE THE DIVISION HERE!
  OddRecs = NumOfCust& MOD MaxBlockCnt
  
  IF NumChunks& = 0 THEN        'if the actual cust count is less than
    MaxBlockCnt = OddRecs       'the work buffer
    NumChunks& = 1
    OddRecs = 0
  END IF
  
  BlockSize = UBCustSNLen * MaxBlockCnt
  
  'Find matching record
  FOR CCnt& = 1 TO NumChunks&
    FGetRTA C1Handle, UBCustSN(1), CCnt&, BlockSize
    FOR RecCnt = 1 TO MaxBlockCnt
      UBSearchN$ = LEFT$(UBCustSN(RecCnt).SearchName, SearchLen)
      IF (SEARCH$ = UBSearchN$) THEN
        IF LEN(QPTrim$(UBCustSN(RecCnt).DelFlag)) GOTO DelSkip2
        IF (ActiveOnly = 0) OR ((ActiveOnly = True) AND ((UBCustSN(RecCnt).Status = "A"))) THEN
          WhatRec& = UBCustSN(RecCnt).RecNum
          'IF WhatRec& = 3129 THEN STOP
          GOSUB CustLoadEM2
        ELSEIF (ActiveOnly = 1) AND (UBCustSN(RecCnt).Status = "I") THEN
          WhatRec& = UBCustSN(RecCnt).RecNum
          GOSUB CustLoadEM2
        END IF
      END IF
DelSkip2:
    NEXT
    ShowPctCompL CCnt&, NumChunks&
    'ShowSearchWheel 12, 44
  NEXT
  
  IF OddRecs > 0 THEN                  'this is always less than max (1024)
    BlockSize = UBCustSNLen * OddRecs  'Adj Block size to get last chunk
    FGetA C1Handle, UBCustSN(1), BlockSize
    FOR RecCnt = 1 TO OddRecs   'search'em
      UBSearchN$ = LEFT$(UBCustSN(RecCnt).SearchName, SearchLen)
      IF (SEARCH$ = UBSearchN$) THEN
        IF LEN(QPTrim$(UBCustSN(RecCnt).DelFlag)) GOTO DelSkip3
        IF (ActiveOnly = 0) OR ((ActiveOnly = True) AND ((UBCustSN(RecCnt).Status = "A"))) THEN
          WhatRec& = UBCustSN(RecCnt).RecNum
          GOSUB CustLoadEM2
        ELSEIF (ActiveOnly = 1) AND (UBCustSN(RecCnt).Status = "I") THEN
          WhatRec& = UBCustSN(RecCnt).RecNum
          GOSUB CustLoadEM2
        END IF
      END IF
DelSkip3:
    NEXT
  END IF
  
  FClose C1Handle               'close files
  FClose R1Handle
  
  IF DCnt = 0 THEN
    RecNo& = 0
    GOTO ExitSearch2
  ELSE

    SortT MChoice(1), DCnt, 0, 64, 0, 64
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Customer/Owner       Service Address        Location No.  S"
    '--Find max menu width
    
    '--Center Menu within Screen
    
    Row = 4
    Col = ((80 - 60) \ 2) - 1
    
    IF CLSFlag THEN
      Row = 4
      BoxBot = 17               'limit the box length
      BlockClear
    ELSE
      Row = 6
      BoxBot = 14               'limit the box length to go no lower than line 20
      RestScrn ScrnArray()
    END IF
    
    LOCATE Row, Col, 0
    
    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(DCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
    LOOP UNTIL RecNo& > 0
  END IF
  
ExitSearch2:
  RestScrn ScrnArray()
  
  ERASE ScrnArray, UBCustRec, MChoice, UBCustSN
  
  EXIT SUB
  
CustLoadEM2:
  
  FGetRTA R1Handle, UBCustRec(1), WhatRec&, UBCustRecLen
  
  DCnt = DCnt + 1
  'IF FRE(-1) < 5000 THEN STOP
  REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
  LSET MChoice(DCnt).V = LEFT$(QPTrim$(UBCustRec(1).CustName), 20)
  MID$(MChoice(DCnt).V, 22, 22) = LEFT$(QPTrim$(UBCustRec(1).SERVADDR), 25)
  MID$(MChoice(DCnt).V, 48, 9) = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb
  MID$(MChoice(DCnt).V, 59, 1) = UBCustRec(1).Status
  MID$(MChoice(DCnt).V, 61) = MKL$(WhatRec&)
  
RETURN
  
END SUB

