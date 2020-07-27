DEFINT A-Z
DECLARE SUB ShowNoCodes ()
DECLARE SUB OpenDCCustIdxFile (NumOfARIdxRecs%, ARIdxFile%)
DECLARE SUB OpenDCCust (NumOfDCRecs%, DCFile%)
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
  '$INCLUDE: 'DC.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
   DIM SHARED DCCust(1) AS DCCustRecType
   DIM SHARED DCCustRec(1) AS DCCustRecType
   DIM SHARED DCCustIdxRec(1) AS DCCustIDXRecType
   DIM SHARED DCTransRec(1) AS DCTransRecType

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
 OpenDCCust NumOfDCRecs, DCFile
 DCTransRecLen = LEN(DCTransRec(1))
 DCTransFile = FREEFILE
 OPEN "DCTrans.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCTransFile LEN = DCTransRecLen
 NumOfTransRecs = LOF(DCTransFile) \ DCTransRecLen


  

  FOR Cnt! = 1 TO NumOfDCRecs
  GET DCFile, Cnt!, DCCustRec(1)
   DCCustRec(1).FirstTrans = 0
   DCCustRec(1).LastTrans = 0
  PUT DCFile, Cnt!, DCCustRec(1)
  NEXT Cnt!
  IF NumOfTransRecs <= 0 THEN CLOSE : RUN "DCMENU"
  FOR TCnt! = 1 TO NumOfTransRecs
   GET #DCTransFile, TCnt!, DCTransRec(1)
   CustRecord = VAL(DCTransRec(1).CustomerNumber)
   GET DCFile, CustRecord, DCCustRec(1)
   IF DCCustRec(1).FirstTrans = 0 THEN
    DCCustRec(1).FirstTrans = TCnt!
    DCCustRec(1).LastTrans = TCnt!
    PUT DCFile, CustRecord, DCCustRec(1)
    ELSE
    PrevRecord = DCCustRec(1).LastTrans
    DCCustRec(1).LastTrans = TCnt!
    PUT DCFile, CustRecord, DCCustRec(1)
    GET DCTransFile, PrevRecord, DCTransRec(1)
    DCTransRec(1).NextTrans = TCnt!
    PUT DCTransFile, PrevRecord, DCTransRec(1)
  END IF
  NEXT TCnt!
  CLOSE : RUN "dcMENU"

   RUN "armenu"

SUB OpenDCCust (NumOfDCRecs%, DCFile%)
  CLOSE DCFile

  DCCustRecLen = LEN(DCCustRec(1))
  DCFile = FREEFILE
  OPEN "DCCust.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile LEN = DCCustRecLen
  NumOfDCRecs = LOF(DCFile) \ DCCustRecLen
  'FOR x = 1 TO NumOfDCRecs
  'GET DCFile, x, DCCust(1)
  'PRINT DCCust(1).Custnumb; TAB(15); DCCust(1).FirstTrans
  'SLEEP 1
  'NEXT x
  'STOP

END SUB

SUB OpenDCCustIdxFile (NumOfARIdxRecs, ARIdxFile)
  CLOSE ARIdxFile
  DCCustIdxRecLen = LEN(DCCustIdxRec(1))
  ARIdxFile = FREEFILE
  OPEN "DCCust.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS ARIdxFile LEN = DCCustIdxRecLen
  NumOfARIdxRecs = LOF(ARIdxFile) \ DCCustIdxRecLen
END SUB

