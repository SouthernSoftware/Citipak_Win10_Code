DECLARE SUB CustHistoryRpt ()
DEFINT A-Z
DECLARE SUB ClearBack ()
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Round# (DblNumber#)
DECLARE FUNCTION ParseBillNum$ (Text$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB ShowCursor ()
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION Monitor% ()

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  CONST False = 0, True = NOT False
  
  CrLf$ = CHR$(13) + CHR$(10)
  
  TYPE HistRecInfoType
    TranRec    AS LONG
    TranType   AS INTEGER
    TranDate   AS INTEGER
    BelongTo   AS LONG
    Printed    AS INTEGER
  END TYPE

CustHistoryRpt

SUB CustHistoryRpt
  
  SHARED Choice$()
  
  REDIM TaxTran(1 TO 2) AS TaxTransactionType
  REDIM TaxCustRec(1) AS TaxCustType
  REDIM HistRecs(1 TO 1) AS HistRecInfoType
  
  HistRecLen = LEN(HistRecs(1))
  
  TaxCustRecLen = LEN(TaxCustRec(1))
  TaxTranRecLen = LEN(TaxTran(1))
  
  RptFile$ = "TaxCHIST.RPT"     'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)
  
TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
   
   GET TaxTran, 2634, TaxTran(1)
   TaxTran(1).Amount = 0
   TaxTran(1).Revenue.Principle1Pd = 0
   TaxTran(1).Description = "VOIDED PAYMENT"
   TaxTran(1).Posted2GL = "Y"
   TaxTran(1).Altered = 1
   PUT TaxTran, 2634, TaxTran(1)

   GET TaxTran, 1910, TaxTran(1)
   TaxTran(1).Revenue.Principle1Pd = TaxTran(1).Revenue.Principle1
   PUT TaxTran, 1910, TaxTran(1)


    
      CLOSE
      END
END SUB

