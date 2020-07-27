DEFINT A-Z

DECLARE FUNCTION MakeBucks$ (Bucks$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION FUsing$ (Bucks$, NoFormat$)

  '$INCLUDE: 'HAMCNVRT.bi'                    'A/R FILE LAYOUTS
  '$INCLUDE: 'ARV85.BI'

  REDIM ARCustRec(1) AS ARCustRecType
  REDIM OldHam(1) AS OHamRecType
  CustLen = LEN(ARCustRec(1))
  OHamLen = LEN(OldHam(1))

  OPEN "privseq" FOR RANDOM AS #1 LEN = OHamLen
  OPEN "ARCUST.DAT" FOR RANDOM AS #2 LEN = CustLen
'  OPEN "CUST.LST" FOR OUTPUT AS #3
  OPEN "SCRN:" FOR OUTPUT AS #3

  NumRec = LOF(1) \ OHamLen
  CLS
  FOR Cnt = 1 TO NumRec
    GET #1, Cnt, OldHam(1)
    IF OldHam(1).PVStatus = "A" THEN
      REDIM ARCustRec(1) AS ARCustRecType
      ARCustRec(1).CUSTNUMB = OldHam(1).AcctNo
      TName$ = QPTrim$(OldHam(1).CustName)
      ARCustRec(1).BILLNAME = TName$
      ARCustRec(1).CustName = TName$
      CharPos = INSTR(TName$, " ")
      DO WHILE CharPos > 0
        TName$ = LEFT$(TName$, CharPos - 1) + MID$(TName$, CharPos + 1)
        CharPos = INSTR(TName$, " ")
      LOOP

      ARCustRec(1).SORTNAME = TName$
'     ARCustRec(1).ADDRESS1 = QPTrim$(OldHam(1).CustName)
'      ARCustRec(1).ADDRESS2 = Form$(5, 0)
      A1Len = LEN(QPTrim$(OldHam(1).MailAddr1))
      A2Len = LEN(QPTrim$(OldHam(1).MailAddr2))
      A3Len = LEN(QPTrim$(OldHam(1).MailAddr3))
      'IF A2Len > 0 AND A3Len > 0 THEN
        CLS
        ZipLine = 0
        REDIM Addr$(1 TO 3)
        Addr$(1) = QPTrim$(OldHam(1).MailAddr1)
        Addr$(2) = QPTrim$(OldHam(1).MailAddr2)
        Addr$(3) = QPTrim$(OldHam(1).MailAddr3)
        FOR ZCnt = 1 TO 3
          IF LEN(Addr$(ZCnt)) > 0 THEN
            Zip$ = QPTrim$(RIGHT$(Addr$(ZCnt), 5))
            IF LEN(Zip$) = 5 THEN
              IF LEFT$(Zip$, 1) = "-" THEN
                Zip$ = MID$(Addr$(ZCnt), (LEN(Addr$(ZCnt)) - 9), 5)
              END IF
              IF VAL(Zip$) > 0 THEN
                ZipLine = ZCnt
                ARCustRec(1).ZipCode = Zip$
                CharPos = INSTR(Addr$(ZipLine), ".")
                DO WHILE CharPos > 0
                  Addr$(ZipLine) = LEFT$(Addr$(ZipLine), CharPos - 1) + MID$(Addr$(ZipLine), CharPos + 1)
                  CharPos = INSTR(Addr$(ZipLine), ".")
                LOOP
                CharPos = INSTR(Addr$(ZipLine), ",")
                DO WHILE CharPos > 0
                  Addr$(ZipLine) = LEFT$(Addr$(ZipLine), CharPos - 1) + MID$(Addr$(ZipLine), CharPos + 1)
                  CharPos = INSTR(Addr$(ZipLine), ",")
                LOOP

                CityST$ = QPTrim$(LEFT$(Addr$(ZCnt), LEN(Addr$(ZCnt)) - LEN(Zip$)))
                ARCustRec(1).State = RIGHT$(CityST$, 2)
                ARCustRec(1).City = LEFT$(CityST$, LEN(CityST$) - 2)
                EXIT FOR
              END IF
            END IF
          END IF
        NEXT

        SELECT CASE ZipLine
        CASE 1

        CASE 2
           ARCustRec(1).ADDRESS1 = Addr$(1)
        CASE 3
           ARCustRec(1).ADDRESS1 = Addr$(1)
           ARCustRec(1).ADDRESS2 = Addr$(2)
        'Addr$(3) = QPTrim$(OldHam(1).MailAddr3)

        CASE ELSE
        END SELECT

        'ARCustRec (1)
        'if instr(OldHam(1).MailAddr1
        PRINT #3, "------"
        PRINT #3, OldHam(1).CustName
        PRINT #3, OldHam(1).MailAddr1
        PRINT #3, OldHam(1).MailAddr2
        PRINT #3, OldHam(1).MailAddr3
        PRINT #3, OldHam(1).BusiAddr1
        PRINT #3, "ZipLine:"; ZipLine

        'IF ZipLine = 0 THEN
        '  DO
        '    E$ = INKEY$
        '  LOOP UNTIL LEN(E$)
        'END IF

        IF E$ = "x" THEN EXIT FOR
       ' DCnt = DCnt + 1
       ' DO: LOOP UNTIL LEN(INKEY$) > 0
      'END IF
      'AcctNo, OldHam(1).LicTypeDesc, OldHam(1).PVStatus, ;
      'PRINT MakeBucks$(OldHam(1).AnnCharge)
      PUT #2, , ARCustRec(1)
    ELSE
    END IF
  NEXT
  CLOSE
END


'      DCnt = DCnt + 1
'      IF DCnt > 22 THEN
'        DO
'          a$ = INKEY$
'        LOOP UNTIL LEN(a$) > 0
'        IF a$ = CHR$(27) THEN
'          EXIT FOR
'        END IF
'        CLS
'        DCnt = 0
'      END IF

FUNCTION MakeBucks$ (Bucks$)

  TempBuck$ = "0.00"
  Bucks$ = QPTrim$(Bucks$)
  BuckLen = LEN(Bucks$)
  IF BuckLen > 2 THEN
    Cash# = VAL(LEFT$(Bucks$, BuckLen - 2))
    Cent# = VAL("." + RIGHT$(Bucks$, 2))
    TempBuck$ = FUsing$(STR$(Cash# + Cent#), "#######.##")
  END IF
  MakeBucks$ = TempBuck$

END FUNCTION

