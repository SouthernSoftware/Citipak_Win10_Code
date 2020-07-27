DEFINT A-Z
DECLARE FUNCTION GetCustCnt& ()
DECLARE FUNCTION QPTrim$ (Text$)

'$INCLUDE: 'TaxCust.BI'
'$INCLUDE: 'TAXCONST.BI'

  CONST False = 0, True = NOT False
  Cmd$ = QPTrim$(COMMAND$)
  PRINT "NC Tax Customer Undelete v1.0  Southern Software"
  PRINT
  IF LEN(Cmd$) = 0 THEN
    GOSUB ShowSyntax
    GOTO DoneINHere
  ELSE
    AcctNum& = VAL(Cmd$)
  END IF

  CustCnt& = GetCustCnt&

  IF AcctNum& < 1 OR AcctNum& > CustCnt& THEN

    PRINT "  ERROR: Invalid account number:"; AcctNum&
    PRINT
    GOSUB ShowSyntax
    GOTO DoneINHere
  END IF

  REDIM TaxCust(1)  AS TaxCustType
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure
  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  GET TaxFile, AcctNum&, TaxCust(1)
  TaxCust(1).Deleted = 0
  PUT TaxFile, AcctNum&, TaxCust(1)
  CLOSE TaxFile
  PRINT "TAX ACCOUNT:"; AcctNum&; "UNDELETED."
DoneINHere:
END

ShowSyntax:
PRINT " Syntax:  TAXUNDEL <CustNumber>"
PRINT "Example:  TAXUNDEL 300"
PRINT "          UNDELETES ACCOUNT 300"
RETURN

FUNCTION GetCustCnt&

  REDIM TaxCust(1)  AS TaxCustType
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure

  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  NumOfRecs& = LOF(TaxFile) \ TaxCustLen
  CLOSE TaxFile

  ERASE TaxCust

  GetCustCnt& = NumOfRecs&

END FUNCTION

