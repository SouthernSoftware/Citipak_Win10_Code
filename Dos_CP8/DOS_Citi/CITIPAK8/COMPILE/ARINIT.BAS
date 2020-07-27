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
  '$INCLUDE: 'AR.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
   DIM SHARED ARCust(1) AS ARCustRecType
   DIM SHARED ARCustRec(1) AS ARCustRecType
   DIM SHARED ARCustIdxRec(1) AS ARCustIDXRecType
   DIM SHARED ARTransRec(1) AS ARTransRecType

   STACK 8000
   CONST False = 0, True = NOT False

   '--Dim the choice array to the number of menu items
   REDIM Mchoice$(1 TO 4)

   Mchoice$(1) = "Add New Customer"
   Mchoice$(2) = "Edit Existing Customer"
   Mchoice$(3) = "Print Customer Listing"
   Mchoice$(4) = "Exit to OS"

   MaxLen = 0     'Set menu width to zero
   BoxBot = 17    'limit the box length to go no lower than line 20
   Action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

   '--Find max menu width
   FOR Cnt = 1 TO UBOUND(Mchoice$)
     TLen = LEN(Mchoice$(Cnt))
     IF TLen > MaxLen THEN
       MaxLen = TLen
     END IF
   NEXT

   '--Center Menu within Screen
   Row = ((25 - (UBOUND(Mchoice$))) \ 2)
   Col = ((80 - MaxLen) \ 2) - 2
   help$ = "Add/Edit/Print Customers"
   
   

      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      LibFile2Scrn "AR.QSL", "MENUBAK", MonoCode, -1, ErrorCode

      TitleBox 3, Col, MaxLen + 3, "Customer Maintenance ", Cnf
      TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

      ShowCursor
   LOCATE 15, 20: PRINT "Re-Linking Data : Please Wait!!!"
 OpenARCust NumOfArRecs, ARFile
 ARTransRecLen = LEN(ARTransRec(1))
 ARTransFile = FREEFILE
 OPEN "ARTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARTransFile LEN = ARTransRecLen
 NumOfTransRecs = LOF(ARTransFile) \ ARTransRecLen
 NextTransRec = NumOfTransRecs + 1
  FOR TCnt! = 1 TO NumOfTransRecs
   GET #ARTransFile, TCnt!, ARTransRec(1)

   IF Num2Date$(ARTransRec(1).TransDate) = "05-05-1998" OR Num2Date$(ARTransRec(1).TransDate) = "05-05-1999" THEN
   GOSUB CheckRecord
   IF Yes$ = "Y" THEN
    ARTransRec(1).CustomerNumber = "0"
    PUT #ARTransFile, TCnt!, ARTransRec(1)
   END IF
   END IF
 NEXT TCnt!
 FOR TCnt! = 1 TO NumOfTransRecs
   GET #ARTransFile, TCnt!, ARTransRec(1)
   GOSUB CheckRecord
   IF Yes$ = "Y" THEN
   ARTransRec(1).NextTrans = 0
   PUT #ARTransFile, TCnt!, ARTransRec(1)
   END IF
 NEXT TCnt!

 
  FOR Cnt! = 1 TO NumOfArRecs
  GET ARFile, Cnt!, ARCustRec(1)
  GOSUB Checkrecord1
  IF Yes$ = "Y" THEN
   ARCustRec(1).FirstTrans = 0
   ARCustRec(1).LastTrans = 0
   ARCustRec(1).AcctBal = 0
   PUT ARFile, Cnt!, ARCustRec(1)
  END IF
  NEXT Cnt!

  IF NumOfTransRecs <= 0 THEN CLOSE : RUN "ARMENU"
  FOR TCnt! = 1 TO NumOfTransRecs
   GET #ARTransFile, TCnt!, ARTransRec(1)
   CustRecord = VAL(ARTransRec(1).CustomerNumber)
   GOSUB CheckRecord
   IF Yes$ = "Y" THEN
   IF CustRecord > 0 THEN
   GET ARFile, CustRecord, ARCustRec(1)
   RunBalance# = ARCustRec(1).AcctBal
   IF ARTransRec(1).TransType = 1 OR ARTransRec(1).TransType = 9 THEN
   RunBalance# = RunBalance# + ARTransRec(1).TransAmount
   END IF
   IF ARTransRec(1).TransType = 2 THEN
   RunBalance# = RunBalance# - ARTransRec(1).TransAmount
   END IF


   IF ARCustRec(1).FirstTrans = 0 THEN
    ARCustRec(1).FirstTrans = TCnt!
    ARCustRec(1).LastTrans = TCnt!
    ARCustRec(1).AcctBal = RunBalance#
    PUT ARFile, CustRecord, ARCustRec(1)
    ELSE
    PrevRecord = ARCustRec(1).LastTrans
    ARCustRec(1).LastTrans = TCnt!
    ARCustRec(1).AcctBal = RunBalance#
    PUT ARFile, CustRecord, ARCustRec(1)
    GET ARTransFile, PrevRecord, ARTransRec(1)
    ARTransRec(1).NextTrans = TCnt!
    PUT ARTransFile, PrevRecord, ARTransRec(1)
  END IF
  END IF
  END IF
  NEXT TCnt!
  CLOSE : RUN "ARMENU"

   RUN "armenu"
CheckRecord:
Yes$ = "N"
Tr! = VAL(ARTransRec(1).CustomerNumber)
IF Tr! = 335 OR Tr! = 342 OR Tr! = 344 THEN Yes$ = "Y": RETURN
IF Tr! = 346 OR Tr! = 347 OR Tr! = 350 THEN Yes$ = "Y": RETURN
IF Tr! = 353 OR Tr! = 360 OR Tr! = 364 THEN Yes$ = "Y": RETURN
IF Tr! = 336 OR Tr! = 341 OR Tr! = 423 THEN Yes$ = "Y": RETURN
IF Tr! = 345 OR Tr! = 348 OR Tr! = 349 THEN Yes$ = "Y": RETURN
IF Tr! = 352 OR Tr! = 478 OR Tr! = 392 THEN Yes$ = "Y": RETURN
IF Tr! = 357 OR Tr! = 359 OR Tr! = 363 THEN Yes$ = "Y": RETURN
IF Tr! = 365 OR Tr! = 367 THEN Yes$ = "Y": RETURN
IF Tr! = 332 OR Tr! = 338 OR Tr! = 333 THEN Yes$ = "Y": RETURN
IF Tr! = 340 OR Tr! = 475 OR Tr! = 362 THEN Yes$ = "Y": RETURN
IF Tr! = 331 OR Tr! = 337 OR Tr! = 334 THEN Yes$ = "Y": RETURN
IF Tr! = 339 OR Tr! = 343 OR Tr! = 351 THEN Yes$ = "Y": RETURN
IF Tr! = 474 OR Tr! = 354 OR Tr! = 356 THEN Yes$ = "Y": RETURN
IF Tr! = 358 OR Tr! = 361 THEN Yes$ = "Y": RETURN
Yes$ = "N"
RETURN

Checkrecord1:
Yes$ = "N"
Tr! = Cnt!
IF Tr! = 335 OR Tr! = 342 OR Tr! = 344 THEN Yes$ = "Y": RETURN
IF Tr! = 346 OR Tr! = 347 OR Tr! = 350 THEN Yes$ = "Y": RETURN
IF Tr! = 353 OR Tr! = 360 OR Tr! = 364 THEN Yes$ = "Y": RETURN
IF Tr! = 336 OR Tr! = 341 OR Tr! = 423 THEN Yes$ = "Y": RETURN
IF Tr! = 345 OR Tr! = 348 OR Tr! = 349 THEN Yes$ = "Y": RETURN
IF Tr! = 352 OR Tr! = 478 OR Tr! = 392 THEN Yes$ = "Y": RETURN
IF Tr! = 357 OR Tr! = 359 OR Tr! = 363 THEN Yes$ = "Y": RETURN
IF Tr! = 365 OR Tr! = 367 THEN Yes$ = "Y": RETURN
IF Tr! = 332 OR Tr! = 338 OR Tr! = 333 THEN Yes$ = "Y": RETURN
IF Tr! = 340 OR Tr! = 475 OR Tr! = 362 THEN Yes$ = "Y": RETURN
IF Tr! = 331 OR Tr! = 337 OR Tr! = 334 THEN Yes$ = "Y": RETURN
IF Tr! = 339 OR Tr! = 343 OR Tr! = 351 THEN Yes$ = "Y": RETURN
IF Tr! = 474 OR Tr! = 354 OR Tr! = 356 THEN Yes$ = "Y": RETURN
IF Tr! = 358 OR Tr! = 361 THEN Yes$ = "Y": RETURN
Yes$ = "N"
RETURN

SUB OpenARCust (NumOfArRecs%, ARFile%)
  CLOSE ARFile

  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen
  'FOR x = 1 TO NumOfArRecs
  'GET ARFile, x, ARCust(1)
  'PRINT ARCust(1).Custnumb; TAB(15); ARCust(1).FirstTrans
  'SLEEP 1
  'NEXT x
  'STOP

END SUB

SUB OpenARCustIdxFile (NumOfARIdxRecs, ARIdxFile)
  CLOSE ARIdxFile
  ARCustIdxRecLen = LEN(ARCustIdxRec(1))
  ARIdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS ARIdxFile LEN = ARCustIdxRecLen
  NumOfARIdxRecs = LOF(ARIdxFile) \ ARCustIdxRecLen
END SUB

