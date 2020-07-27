DEFINT A-Z
DECLARE SUB CheckCustFile ()
DECLARE SUB CheckSetupFile ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Num$)

CONST False = 0, True = NOT False

  STACK 5000

  '$INCLUDE: 'newcust.bi'
  ''$INCLUDE: 'ubowner.BI'
  ''$INCLUDE: 'ubsetup.BI'

  REDIM UBCust(1)   AS NewUBCustRecType
  UBCustLen = LEN(UBCust(1))


  'DIM SHARED OSetup  AS oUBSetupRecType
  'DIM SHARED NSetup  AS UBSetupRecType
  'DIM SHARED CnvrtCust AS INTEGER

  'OSetLen = LEN(OSetup)
  'NSetLen = LEN(NSetup)

  CLS

  LOCATE , 4
  PRINT "UBUTIL v1.01  1994-98 Southern Software."
  PRINT

  Cmd$ = COMMAND$

  IF INSTR(Cmd$, "HELP") OR LEN(Cmd$) = 0 THEN
    PRINT
    PRINT TAB(1); "  Syntax:  UBUTIL1 COMMAND <arg1> <arg2...>"
    PRINT
    PRINT TAB(1); "COMMANDS:"
    PRINT
    PRINT TAB(10); "   BStatus  <AcctNO>"
    PRINT
    PRINT TAB(10); "     TDate  <FromDate> <ToDate>"
    PRINT
    PRINT TAB(10); "Resequence"
    PRINT
    PRINT
    PRINT
    PRINT TAB(5); "Please contact Southern Software for Assistance."
    GOTO AbortExit
  END IF

  IF INSTR(Cmd$, "BSTATUS") THEN
    ParmPos = INSTR(Cmd$, "BSTATUS")
    Acct$ = QPTrim$(MID$(Cmd$, ParmPos + 7))
    IF LEN(Acct$) > 0 THEN
      ACCTNO = QPValI(Acct$)
      IF ACCTNO > 0 THEN
        CustFile = FREEFILE
        OPEN "UBCUST.DAT" FOR RANDOM AS #CustFile LEN = UBCustLen
        GET #CustFile, ACCTNO, UBCust(1)
        UBCust(1).Status = "B"
        PUT #CustFile, ACCTNO, UBCust(1)
        CLOSE CustFile
        PRINT "  Changed Account:"; ACCTNO; "to B Status"
      END IF
    ELSE
      PRINT "  No Account Number for B Status"
    END IF
  END IF


AbortExit:
  CLOSE
  END

