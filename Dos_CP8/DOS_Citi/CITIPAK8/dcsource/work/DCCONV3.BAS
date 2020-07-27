DEFINT A-Z
DECLARE SUB ShowNoCodes ()
DECLARE SUB OpenDCCustIdxFile (NumOfARIdxRecs%, ARIdxFile%)
DECLARE SUB OpenDCCust (NumOfDCRecs%, DCFile%, DCCustRecLen%)
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

   
   DIM SHARED DCCustRec(1) AS DCCustRecType

   DIM SHARED DCOldCustRec(1) AS DCOldCustRecType

   DCCustRecLen1 = LEN(DCCustRec(1))

   

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

 OpenDCCust NumOfDCRecs, DCFile, DCCustRecLen

  
  DCFile1 = FREEFILE
  OPEN "DCCust.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile1 LEN = DCCustRecLen1
  
  FOR Cnt! = 1 TO NumOfDCRecs
  GET DCFile, Cnt!, DCOldCustRec(1)
    DCCustRec(1).CUSTNUMB = DCOldCustRec(1).CUSTNUMB
    DCCustRec(1).SORTNAME = DCOldCustRec(1).SORTNAME
    DCCustRec(1).BILLNAME = DCOldCustRec(1).BILLNAME
    DCCustRec(1).ADDRESS1 = DCOldCustRec(1).ADDRESS1
    DCCustRec(1).ADDRESS2 = DCOldCustRec(1).ADDRESS2
    DCCustRec(1).city = DCOldCustRec(1).city
    DCCustRec(1).STATE = DCOldCustRec(1).STATE
    DCCustRec(1).ZIPCODE = DCOldCustRec(1).ZIPCODE
    DCCustRec(1).SOSEC = ""
    DCCustRec(1).DRVLIC = ""
    DCCustRec(1).DATEOPED = -32767
    DCCustRec(1).CASHONLY = "N"
    DCCustRec(1).resident = "Y"
    DCCustRec(1).Owner = "Y"
    DCCustRec(1).HPHONE = ""
    DCCustRec(1).WPHONE = ""
    DCCustRec(1).LICENSE = ""
    DCCustRec(1).Valid = 1
    DCCustRec(1).AcctBal = 0
    DCCustRec(1).Deleted = "N"
    DCCustRec(1).FirstTrans = 0
    DCCustRec(1).LastTrans = 0
    DCCustRec(1).FirstCar = 0
    DCCustRec(1).LastCar = 0
    DCCustRec(1).SocSec1 = ""
    DCCustRec(1).OtherName = ""
    DCCustRec(1).RoomtoGrow = ""
    DCCustRec(1).FirstTrans = 0
    DCCustRec(1).LastTrans = 0
  PUT DCFile1, Cnt!, DCCustRec(1)
  NEXT Cnt!
  CLOSE : RUN "DCMENU"
  CLOSE : RUN "dcMENU"

   RUN "armenu"

SUB OpenDCCust (NumOfDCRecs%, DCFile%, DCCustRecLen%)
  CLOSE DCFile

  DCCustRecLen = LEN(DCOldCustRec(1))
  DCFile = FREEFILE
  OPEN "DCCust.OLD" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile LEN = DCCustRecLen
  NumOfDCRecs = LOF(DCFile) \ DCCustRecLen

END SUB

