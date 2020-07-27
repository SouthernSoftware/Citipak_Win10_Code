DEFINT A-Z
DECLARE SUB ShowNoCodes ()
DECLARE SUB OpenARCustIdxFile (NumOfARIdxRecs%, ARIdxFile%)
DECLARE SUB OpenARCust (NumOfArRecs%, ARFile%)
DECLARE SUB SortARNameIndex ()
DECLARE SUB AddCustomer ()
DECLARE SUB EditCustomer ()
DECLARE SUB PrintCustomer ()
DECLARE SUB printhelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB HideCursor ()
DECLARE SUB QPrint (x$, Colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()

'$INCLUDE: 'DefCnf.BI'
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)

TYPE Struct
 who AS STRING * 14
 RecNum AS INTEGER
END TYPE



  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ARV85.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
   REDIM SHARED ARCust(1) AS ARCustRecType
   REDIM SHARED ARCustRec(1) AS ARCustRecType
   REDIM ARCatCodeRec(1) AS ARCatCodeRecType



   STACK 8000
   CONST False = 0, True = NOT False

 ShowCursor
CLS : PRINT "Now Updating Your Accounts : Please Wait!"
   CatCodeRecLen = LEN(ARCatCodeRec(1))
   CatFile = FREEFILE
   OPEN "ARCODE.DAT" FOR RANDOM AS CatFile LEN = CatCodeRecLen
   NumOfCatRecs = LOF(CatFile) \ CatCodeRecLen

 OpenARCust NumOfArRecs, ARFile
 FOR Cnt! = 1 TO NumOfArRecs
 GET ARFile, Cnt!, ARCustRec(1)

  IF VAL(ARCustRec(1).BillCat1) > 0 THEN
   BillCat$ = ARCustRec(1).BillCat1
   GOSUB GetDesc
   IF Desc$ = "NO" THEN
   ELSE
   ARCustRec(1).Desc1 = Desc$
   ARCustRec(1).Fee1 = ARCatCodeRec(1).Fee
   END IF
   ELSE
   ARCustRec(1).BillCat1 = ""
  END IF
  IF VAL(ARCustRec(1).BillCat2) > 0 THEN
   BillCat$ = ARCustRec(1).BillCat2
   GOSUB GetDesc
   IF Desc$ = "NO" THEN
   ELSE
   ARCustRec(1).Desc2 = Desc$
   ARCustRec(1).Fee2 = ARCatCodeRec(1).Fee
   END IF
   ELSE
   ARCustRec(1).BillCat2 = ""

  END IF
  IF VAL(ARCustRec(1).BillCat3) > 0 THEN
   BillCat$ = ARCustRec(1).BillCat3
   GOSUB GetDesc
   IF Desc$ = "NO" THEN
   ELSE
   ARCustRec(1).Desc3 = Desc$
   ARCustRec(1).Fee3 = ARCatCodeRec(1).Fee
   END IF
   ELSE
   ARCustRec(1).BillCat3 = ""
  END IF
  IF VAL(ARCustRec(1).BillCat4) > 0 THEN
   BillCat$ = ARCustRec(1).BillCat4
   GOSUB GetDesc
   IF Desc$ = "NO" THEN
   ELSE
   ARCustRec(1).Desc4 = Desc$
   ARCustRec(1).Fee4 = ARCatCodeRec(1).Fee
   END IF
   ELSE
   ARCustRec(1).BillCat4 = ""
  END IF
  IF VAL(ARCustRec(1).BillCat5) > 0 THEN
   BillCat$ = ARCustRec(1).BillCat5
   GOSUB GetDesc
   IF Desc$ = "NO" THEN
   ELSE
   ARCustRec(1).Desc5 = Desc$
   ARCustRec(1).Fee5 = ARCatCodeRec(1).Fee
   END IF
   ELSE
   ARCustRec(1).BillCat5 = ""
  END IF

   FeeAmt# = ARCustRec(1).Fee1 + ARCustRec(1).Fee2 + ARCustRec(1).Fee3 + ARCustRec(1).Fee4 + ARCustRec(1).Fee5
   ARCustRec(1).FeeAmt = FeeAmt#
   PUT ARFile, Cnt!, ARCustRec(1)
  NEXT Cnt!
  CLOSE
  END

GetDesc:
  FOR Cnt1 = 1 TO NumOfCatRecs
  GET CatFile, Cnt1, ARCatCodeRec(1)
  IF ARCatCodeRec(1).CatCode = BillCat$ THEN Desc$ = ARCatCodeRec(1).CodeDesc: RETURN
  NEXT Cnt1
  Desc$ = "NO"
  RETURN

SUB OpenARCust (NumOfArRecs%, ARFile%)
  CLOSE ARFile
  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen

END SUB

