DEFINT A-Z

TYPE APLedger81RecType             'version for Troy's paid supply list
   VIN          AS INTEGER
   VendorCode   AS STRING * 10
   VRecNum      AS INTEGER
   TrDate       AS INTEGER
   GLDistDate   AS INTEGER
   DUEDATE      AS INTEGER
   TrCode       AS INTEGER      '1=Invoice, 4=PO, 3=Check, -3=Void Check, -4=Cleared PO
   DOCNum       AS STRING * 25
   PONUM        AS STRING * 20
   PAYCODE      AS INTEGER
   PrintCode    AS INTEGER
   PDCheckNum   AS LONG
   PDCheckDate  AS INTEGER
   Comment      AS STRING * 31          'Reduced to 31 bytes to allow for dept to be stored (wrightsville beach)
   DeptNumb     AS LONG
   PSLFlag      AS STRING * 1
   Get1099      AS STRING * 1
   Amt          AS DOUBLE
   FrstDist     AS LONG
   LastDist     AS LONG
   NextTrans    AS LONG
   TaxAmt       AS DOUBLE
   Pad           AS STRING * 2
END TYPE

'--A/P Ledger
TYPE APLedgerRecType  'version 8.0
   VIN        AS INTEGER
   VendorCode AS STRING * 10
   VRecNum    AS INTEGER
   TrDate     AS INTEGER
   GLDistDate AS INTEGER
   DUEDATE    AS INTEGER
   TrCode     AS INTEGER
   DOCNum     AS STRING * 25
   PONUM      AS STRING * 20
   PAYCODE    AS INTEGER
   PrintCode  AS INTEGER
   PDCheckNum   AS LONG
   PDCheckDate  AS INTEGER
   MiscCode     AS STRING * 23
   Amt          AS DOUBLE
   FrstDist     AS LONG
   LastDist   AS LONG
   NextTrans  AS LONG
END TYPE

DIM ORec AS APLedgerRecType
DIM NRec AS APLedger81RecType

OLen = LEN(ORec)
NLen = LEN(NRec)

OPEN "APLEDGER.dat" FOR RANDOM AS #1 LEN = OLen
OPEN "NLEDGER.dat" FOR RANDOM AS #2 LEN = NLen

   FOR Cnt = 1 TO LOF(1) \ OLen
   GET #1, Cnt, ORec
   NRec.VIN = ORec.VIN
   NRec.VendorCode = ORec.VendorCode
   NRec.VRecNum = ORec.VRecNum
   NRec.TrDate = ORec.TrDate
   NRec.GLDistDate = ORec.GLDistDate
   NRec.DUEDATE = ORec.DUEDATE
   NRec.TrCode = ORec.TrCode
   NRec.DOCNum = ORec.DOCNum
   NRec.PONUM = ORec.PONUM
   NRec.PAYCODE = ORec.PAYCODE
   NRec.PrintCode = ORec.PrintCode
   NRec.PDCheckNum = ORec.PDCheckNum
   NRec.PDCheckDate = ORec.PDCheckDate
   NRec.Comment = ORec.MiscCode
'   NRec.DeptNumb
'   NRec.PSLFlag
'   NRec.Get1099
   NRec.Amt = ORec.Amt
   NRec.FrstDist = ORec.FrstDist
   NRec.LastDist = ORec.LastDist
   NRec.NextTrans = ORec.NextTrans
'   NRec.TaxAmt
'   NRec.Pad
   PUT #2, Cnt, NRec
 NEXT
 CLOSE

