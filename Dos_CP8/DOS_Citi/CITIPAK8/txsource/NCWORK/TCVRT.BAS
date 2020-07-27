DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

'$INCLUDE: 'TaxCust.BI'

  ON ERROR GOTO ErrorTrap

  REDIM Old(1)   AS TBCustType
  REDIM New(1)   AS TaxCustType
  OldLen = LEN(Old(1))
  NewLen = LEN(New(1))

  OPEN "TBCUST.DAT" FOR RANDOM AS #1 LEN = OldLen
  OPEN "TAXCUST.DAT" FOR RANDOM AS #2 LEN = NewLen
  Recs& = LOF(1) \ OldLen

  FOR Cnt& = 1 TO Recs&
    LOCATE 5, 2
    PRINT Cnt&; " of"; Recs&
    GET #1, Cnt&, Old(1)
    REDIM New(1)   AS TaxCustType
    New(1).ACCT = VAL(Old(1).ACCT)
    New(1).OPENDATE = 6553

    First$ = Old(1).FirstName
    Last$ = Old(1).LName

    Comma = INSTR(First$, ",")
    IF Comma > 0 THEN
      MID$(First$, Comma) = " "
    END IF

    Comma = INSTR(Last$, ",")
    IF Comma > 0 THEN
      MID$(Last$, Comma) = " "
    END IF

    New(1).FNAME = QPTrim$(First$)
    New(1).LName = QPTrim$(Last$)
    New(1).SNAME = QPTrim$(Last$)

    New(1).HPHONE = Old(1).Phone
    New(1).Addr1 = Old(1).Addr1
    New(1).Addr2 = Old(1).Addr2
    New(1).City = Old(1).City
    New(1).State = Old(1).State

    New(1).Active = "Y"
    New(1).TaxExempt = "N"
    New(1).Interest = "Y"
    New(1).Penalty = "Y"
    
    DashPos = INSTR(Old(1).Zip, "-")
    IF DashPos > 0 THEN
      Zip$ = LEFT$(Old(1).Zip, DashPos - 1) + MID$(Old(1).Zip, DashPos + 1)
    ELSE
      Zip$ = Old(1).Zip
    END IF

    New(1).Zip = Zip$
    New(1).FileVer = CustRecVer
    New(1).PIN = Cnt&
    PUT #2, Cnt&, New(1)
here:
  NEXT

  CLOSE

  OPEN "TAXCPIN.DAT" FOR RANDOM AS #1 LEN = 4
  PUT #1, 1, Cnt&
  CLOSE

PRINT
PRINT
PRINT ErrorCnt

END

ErrorTrap:
ErrorCnt = ErrorCnt + 1
RESUME here

