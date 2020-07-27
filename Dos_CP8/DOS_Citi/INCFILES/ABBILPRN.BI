'011303

'  TName$ = QPTrim$(UBCustRec(1).CustName)
'  CNLen = LEN(TName$)

  TName$ = QPTrim$(UBCustRec(1).CustName)
  CNLen = LEN(TName$)
  SPos = INSTR(TName$, ",")
 
  IF SPos > 0 THEN 'AND SPos < CNLen THEN
    MID$(TName$, SPos, 1) = " "
  END IF
  CName$ = TName$
  
'  SPos = INSTR(TName$, " ")
'  IF SPos > 0 AND SPos < CNLen THEN
'    CName$ = LEFT$(TName$, SPos - 1) + "," + MID$(TName$, SPos)
'  ELSE
'    CName$ = TName$
'  END IF

  Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

  IF UBBillRec(1).TransAmt MOD 10 <> 0 THEN
    DidFifTeen = True
  ELSE
    DidFifTeen = False
  END IF

  TBCnt = VAL(UBCustRec(1).UserCode1)

  AlarmCnt = VAL(UBCustRec(1).UserCode1)

  IF DidFifTeen THEN
    AlarmCnt = AlarmCnt - 1
  END IF

  'TACnt = AlarmCnt '+ 3

  'IF AlarmCnt > 4 THEN
  '  AlarmCnt = AlarmCnt - 4
  'ELSE
  '  AlarmCnt = 0
  'END IF

  CityStaZip$ = QPTrim$(UBCustRec(1).City) + " " + QPTrim$(UBCustRec(1).State) + " " + QPTrim$(UBCustRec(1).ZipCode)
  PRINT #UBRpt, "~"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); CName$; TAB(52); "***** ALARM INVOICE *****"
  PRINT #UBRpt, TAB(8); QPTrim$(UBCustRec(1).Addr1); TAB(52); "REFERENCE# "; QPTrim$(UBCustRec(1).HPhone)
  PRINT #UBRpt, TAB(8); CityStaZip$; TAB(52); BillDate$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); "IN "; QPTrim$(Message$); " THE DARLINGTON COMMUNICATIONS CENTER RECEIVED EMERGENCY"
  PRINT #UBRpt, TAB(8); "ALARMS FROM THE LOCATION BELOW. FURTHER INVESTIGATION BY"
  PRINT #UBRpt, TAB(8); "RESPONDING PERSONNEL INDICATED THAT THIS WAS A FALSE ALARM(S),"
  PRINT #UBRpt, TAB(8); "ACCORDING TO OUR RECORDS, YOU OWN AND OPERATE THIS ALARM."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(12); "LOCATION: "; QPTrim$(UBCustRec(1).SERVADDR); TAB(60); "ACCT: "; CustAcctNo&
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(11); "YTD FALSE: "; QPTrim$(UBCustRec(1).PUMPCODE)
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(12); "DATE/TIME                 DATE/TIME               DATE/TIME"
  PRINT #UBRpt, TAB(8); STRING$(72, "-")
  PRINT #UBRpt, TAB(8); QPTrim$(UBCustRec(1).HHMsg1); TAB(35); QPTrim$(UBCustRec(1).HHMsg2); TAB(59); QPTrim$(UBCustRec(1).HHMsg3)
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  IF DidFifTeen THEN
    PRINT #UBRpt, TAB(12); "    3       FIRST THREE FALSE ALARMS         0.00"
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(12); "    1       FOURTH FALSE ALARM @15.00       15.00"
  ELSE
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
  END IF
  PRINT #UBRpt,
  IF AlarmCnt > 0 THEN
    BAmt# = 20# * AlarmCnt
    PRINT #UBRpt, TAB(16); QPTrim$(STR$(AlarmCnt)); TAB(24); USING "FIVE AND ABOVE @$20.00 EACH ######.##"; BAmt#
  ELSE
    PRINT #UBRpt, TAB(12); "            FIVE AND ABOVE @$20.00 EACH      0.00"
  END IF
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(12); USING "            PREVIOUS BALANCE DUE         #####.##"; Previous#
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(12); USING "               TOTAL BALANCE DUE         #####.##"; TotalAmt#
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); "IF YOU HAVE ANY QUESTIONS CONCERNING THIS STATEMENT PLEASE"
  PRINT #UBRpt, TAB(8); "CONTACT LIBBY WEST AT (843) 398-4920."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); "NO. ALARMS BILLED FOR ="; QPTrim$(Message$); "  "; STR$(TBCnt)
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(25); "***** PAYMENT DUE UPON RECEIPT *****"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(20); "PLEASE REMIT TO:  DARLINGTON COUNTY TREASURER"
  PRINT #UBRpt, TAB(20); "                  1 PUBLIC SQUARE, ROOM 203"
  PRINT #UBRpt, TAB(20); "                  DARLINGTON, S.C. 29532"
  PRINT #UBRpt, CHR$(12)



