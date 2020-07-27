DEFINT A-Z
DECLARE SUB SetLicense ()
DECLARE SUB ShowNoCodes ()
DECLARE SUB OpenARCustIdxFile (NumOfARIdxRecs%, ARIdxFile%)
DECLARE SUB OpenARCustFile (NumOfArRecs%, ARFile%)
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
  ''$INCLUDE: 'AR.bi'
  '$INCLUDE: 'ARV85.bi'                        'A/R FILE LAYOUTS
  '$INCLUDE: 'GL.bi'
   DIM SHARED ARCust(1) AS ARCustRecType
   DIM SHARED ARCustRec(1) AS ARCustRecType
   DIM SHARED ARCustRec1(1) AS AROLDCustRecType
   DIM SHARED ARCustIdxRec(1) AS ARCustIDXRecType

   STACK 8000
   CONST False = 0, True = NOT False

   'Convert Writ Beech Bus Lic Data

   AddCustomer

   RUN "armenu"

SUB AddCustomer

mainbody:

  OpenARCustFile NumOfArRecs, ARFile

  ARCustRecLen1 = LEN(ARCustRec1(1))
  ARFile1 = FREEFILE
  OPEN "AROCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile1 LEN = ARCustRecLen1


    FOR Cnt! = 1 TO LOF(ARFile1) / LEN(ARCustRec1(1))
    GET ARFile1, Cnt!, ARCustRec1(1)
     ARCustRec(1).SORTNAME = ARCustRec1(1).SORTNAME
     ARCustRec(1).BILLNAME = ARCustRec1(1).BILLNAME
     ARCustRec(1).ADDRESS1 = ARCustRec1(1).ADDRESS1
     ARCustRec(1).ADDRESS2 = ARCustRec1(1).ADDRESS2
     ARCustRec(1).CITY = ARCustRec1(1).CITY
     ARCustRec(1).STATE = ARCustRec1(1).STATE
     ARCustRec(1).ZIPCODE = ARCustRec1(1).ZIPCODE
     ARCustRec(1).CustName = ARCustRec1(1).CustName
     ARCustRec(1).Contact = ARCustRec1(1).Contact
     ARCustRec(1).BILLCAT1 = ARCustRec1(1).BILLCAT
     ARCustRec(1).Desc1 = ""
     ARCustRec(1).REV1 = 0
     ARCustRec(1).Fee1 = 0
     ARCustRec(1).BILLCAT2 = ""
     ARCustRec(1).Desc2 = ""
     ARCustRec(1).REV2 = 0
     ARCustRec(1).Fee2 = 0
     ARCustRec(1).BILLCAT3 = ""
     ARCustRec(1).Desc3 = ""
     ARCustRec(1).REV3 = 0
     ARCustRec(1).Fee3 = 0
     ARCustRec(1).BILLCAT4 = ""
     ARCustRec(1).Desc4 = ""
     ARCustRec(1).REV4 = 0
     ARCustRec(1).Fee4 = 0
     ARCustRec(1).BILLCAT5 = ""
     ARCustRec(1).Desc5 = ""
     ARCustRec(1).REV5 = 0
     ARCustRec(1).Fee5 = 0
     ARCustRec(1).WPHONE = ARCustRec1(1).WPHONE
     ARCustRec(1).FeeAmt = 0
     ARCustRec(1).LICENSE = ARCustRec1(1).LICENSE
     ARCustRec(1).VALID = ARCustRec1(1).VALID
     ARCustRec(1).AcctBal = 0
     ARCustRec(1).IssuanceFee = ARCustRec1(1).IssuanceFee
     ARCustRec(1).CustLocation = ARCustRec1(1).CustLocation
     ARCustRec(1).FirstTrans = 0
     ARCustRec(1).LastTrans = 0
     ARCustRec(1).Deleted = ARCustRec1(1).Deleted
     ARCustRec(1).IssueLicense = "Y"
     ARCustRec(1).RoomtoGrow = ""
     ARCustRec(1).CUSTNUMB = ARCustRec1(1).CUSTNUMB
     PUT ARFile, Cnt!, ARCustRec(1)
    NEXT Cnt!
    CLOSE
    END
     
END SUB

SUB OpenARCustFile (NumOfArRecs, ARFile)
  CLOSE ARFile
  
  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  END SUB

SUB OpenARCustIdxFile (NumOfARIdxRecs, ARIdxFile)
  CLOSE ARIdxFile
  ARCustIdxRecLen = LEN(ARCustIdxRec(1))
  ARIdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS ARIdxFile LEN = ARCustIdxRecLen
  NumOfARIdxRecs = LOF(ARIdxFile) \ ARCustIdxRecLen
END SUB

SUB ShowNoCodes
  LibName$ = "AR"
  ScrnName$ = "ARNOCODE"
  help$ = "NEW A/R Customer Entry"
  LOCATE 1, 1, 0


  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F


  PRINT CHR$(7);
  ShowCursor
  LibFile2Scrn "AR.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  Action = 1
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE EscKey
     EXIT SUB
    END SELECT

  LOOP




END SUB

SUB SortARNameIndex
  SHARED Mchoice$


  size = 2500
  Start = 1               'start at array element 1
  Dir = 0                 'sort direction - use anything else for descending
  SSize = 16               'total size of each TYPE element
  MOff = 0                'offset into the TYPE for the key element
  MSize = 7              'size of the key element - coded as follows:
                        '   -1 = integer
                        '   -2 = long integer
                        '   -3 = single precision
                        '   -4 = double precision
                        '   +N = TYPE array/fixed-length string of length N

  DIM array(1 TO size)  AS Struct
  help$ = "Sorting Customer Index"
  printhelp help$

  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen

  ARCustIdxRecLen = LEN(ARCustIdxRec(1))
  ARIdxFile = FREEFILE
  OPEN "ARCUST.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS ARIdxFile LEN = ARCustIdxRecLen
  
 FOR Cnt = 1 TO NumOfArRecs
    GET ARFile, Cnt, ARCustRec(1)
     array(Cnt).who = ARCustRec(1).SORTNAME + "    "
     array(Cnt).RecNum = Cnt
 NEXT Cnt

 SortT array(Start), NumOfArRecs, Dir, SSize, MOff, MSize

 FOR Cnt = 1 TO NumOfArRecs
   ARCustIdxRec(1).IDXNAME = array(Cnt).who
   ARCustIdxRec(1).IDXRECORD = array(Cnt).RecNum
   PUT ARIdxFile, Cnt, ARCustIdxRec(1)
 NEXT Cnt
 CLOSE ARFile
 CLOSE ARIdxFile
END SUB

