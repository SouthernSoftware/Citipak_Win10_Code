DEFINT A-Z
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB RelinkAbstracts ()
DECLARE SUB MortCode ()
DECLARE SUB TAXSetup ()
DECLARE SUB TAXCustomerMenu ()
DECLARE SUB ClearBack ()
DECLARE SUB SendDist2GL ()
DECLARE SUB UBMiscMenu ()
DECLARE SUB UBBillMenu ()
DECLARE SUB UBCustomerMenu ()
DECLARE SUB ClearScrn ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB QPrintRC (Text$, Row, Col, WhatColor)

'$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)

  CONST False = 0, True = NOT False

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'

  DIM SHARED TaxSetUpRec AS TaxMasterType

  STACK 5000

TYPE Struct
 who AS STRING * 14
 RecNum AS INTEGER
END TYPE


  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 4)

  MChoice$(1) = " Tax Default Setup"
  MChoice$(2) = " Mortgage Code Maintenance"
  MChoice$(3) = " Relink Abstracts"
  MChoice$(4) = " Exit to OS"

  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight

  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT

  '--Center Menu within Screen
  Row = ((23 - (UBOUND(MChoice$))) \ 2) + 1
  Col = ((80 - MaxLen) \ 2) - 1

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0

    ClearBack

    TitleBox 3, Col, MaxLen + 3, "Tax Billing Setup Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select ", Cnf

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1
       TAXSetup
    CASE 2
       MortCode
    CASE 3
       RelinkAbstracts

    CASE 4
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP

  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF
  IF Exist("Taxmenu" + Ext$) THEN
    RUN "TaxMenu"
  ELSE
    HideCursor
    ClearScrn
  END IF

END

  SUB MortCode
  SHARED MChoice$
  ClearBack

MainBody:
  size = 500
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

  REDIM array(1 TO size)  AS Struct

  LibName$ = "TAX"
  ScrnName$ = "MORTCODE"

  LOCATE 1, 1, 0
  help$ = "Add Mortgage Codes V8.0"
  ShowCursor

  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp help$

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

  Action = 1
  NoProcess = False

  REDIM MortCodeRec(1) AS MortCodeRecType
  MortCodeRecLen = LEN(MortCodeRec(1))
  MortFile = FREEFILE
  OPEN "TAXMORT.DAT" FOR RANDOM AS MortFile LEN = MortCodeRecLen
  NumOfCatRecs = LOF(MortFile) \ MortCodeRecLen

  IF NumOfCatRecs > 0 THEN

  GOSUB SortMortCodes

  REDIM MChoice$(1 TO NumOfCatRecs)
  FOR Cnt = 1 TO NumOfCatRecs
    GET MortFile, array(Cnt).RecNum, MortCodeRec(1)
    MChoice$(Cnt) = SPACE$(50)
    LSET MChoice$(Cnt) = MortCodeRec(1).MortCode
    MID$(MChoice$(Cnt), 5) = MortCodeRec(1).BName
  NEXT Cnt

   MaxLen = 50     'Set menu width to zero
   BoxBot = 17    'limit the box length to go no lower than line 20
   Action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

   TText$ = SPACE$(MaxLen + 4)
   LSET TText$ = " Code   Description      ** Press <ESC> to Add Code"

   '--Center Menu within Screen
   Row = 8
   Col = 15
   
   DO

      '--Set upper left corner of menu, turn off the cursor
      help$ = "Add or Edit Catagories"

      LOCATE Row, Col, 0
      LibFile2Scrn "DC", "MENUBAK", MonoCode, -1, ErrorCode
            'printhelp help$
      ShowCursor
      QPrintRC TText$, Row - 1, Col, 112
      VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo = 0
        ExitFlag = True
      ELSE
        RecNo = array(Choice).RecNum
        ExitFlag = True
      END IF
   LOOP UNTIL ExitFlag
  END IF


  IF RecNo > 0 THEN
    GET MortFile, RecNo, MortCodeRec(1)
    Form$(1, 0) = MortCodeRec(1).MortCode
    Form$(2, 0) = MortCodeRec(1).BName
    Form$(3, 0) = MortCodeRec(1).Add1
    Form$(4, 0) = MortCodeRec(1).Add2
    Form$(5, 0) = MortCodeRec(1).Add3
    Form$(6, 0) = MortCodeRec(1).Contact
    Form$(7, 0) = MortCodeRec(1).Phone
  END IF

  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  'printhelp help$

 Action = 1

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode

    CASE F7KEY
     CLOSE MortFile
     EXIT SUB

    CASE F10Key

      IF LEFT$(Form$(1, 0), 2) = "  " THEN EXIT SUB
      ' check for duplicate account number
      IF RecNo = 0 THEN
       GOSUB CheckCodes
       ELSE
       ValidCode = True
       NoProcess = False
       AcctRecNo = True
      END IF

      IF ValidCode = False THEN
       LOCATE 14, 18: PRINT "DUPLICATE CODE ENTERED : PLEASE ENTER NEW CODE";
       PRINT CHR$(7);
       SLEEP 1
       Frm(1).FldNo = 1
       Action = 1
       LOCATE 14, 18: PRINT "                                               "
       Done = False
       NoProcess = True
      END IF
    

     IF NoProcess = False THEN
      MortCodeRec(1).MortCode = Form$(1, 0)
      MortCodeRec(1).BName = Form$(2, 0)
      MortCodeRec(1).Add1 = Form$(3, 0)
      MortCodeRec(1).Add2 = Form$(4, 0)
      MortCodeRec(1).Add3 = Form$(5, 0)
      MortCodeRec(1).Contact = Form$(6, 0)
      MortCodeRec(1).Phone = Form$(7, 0)

        IF RecNo > 0 THEN
           PUT MortFile, RecNo, MortCodeRec(1)
         ELSE
           PUT MortFile, LOF(MortFile) \ MortCodeRecLen + 1, MortCodeRec(1)
         END IF
         CLOSE MortFile
         GOTO MainBody
       END IF
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done

  EXIT SUB


CheckCodes:
 FOR Cnt = 1 TO NumOfCatRecs
 GET MortFile, Cnt, MortCodeRec(1)
 IF Form$(1, 0) = MortCodeRec(1).MortCode THEN ValidCode = False: RETURN
 NEXT Cnt
 ValidCode = True
 RETURN

SortMortCodes:
 FOR Cnt = 1 TO NumOfCatRecs
    GET MortFile, Cnt, MortCodeRec(1)
    array(Cnt).who = MortCodeRec(1).MortCode + STRING$(11, " ")
    array(Cnt).RecNum = Cnt
 NEXT Cnt
 SortT array(Start), NumOfCatRecs, Dir, SSize, MOff, MSize
RETURN

END SUB

SUB RelinkAbstracts

  ClearBack
  ShowProcessingScrn "Relinking Customer Information."

  REDIM TaxCust(1) AS TaxCustType
  TaxCustLen = LEN(TaxCust(1))          'Length of Cust Record Structure

  REDIM PersRec(1 TO 2) AS PersonalRecType
  PersRecLen = LEN(PersRec(1))

  REDIM RealRec(1 TO 2) AS PropertyRecType
  RealRecLen = LEN(RealRec(1))

  RealFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS RealFile LEN = RealRecLen
  NumRealRecs& = LOF(RealFile) / RealRecLen

  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  NumPersRecs& = LOF(PersFile) / PersRecLen
  
  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  NumOfCRecs& = LOF(TaxFile) \ TaxCustLen

  QPrintRC "Clearing Personal Abstracts Pointers", 11, 23, -1

  FOR Cnt& = 1 TO NumPersRecs&           'clear pers pointers
    GET #PersFile, Cnt&, PersRec(1)
    PersRec(1).NextRec = 0
    PUT #PersFile, Cnt&, PersRec(1)
    ShowPctComp Cnt&, NumPersRecs&
  NEXT

  QPrintRC "  Clearing Real Abstracts Pointers  ", 11, 23, -1
  FOR Cnt& = 1 TO NumRealRecs&           'clear real pointers
    GET #RealFile, Cnt&, RealRec(1)
    RealRec(1).NextRec = 0
    PUT #RealFile, Cnt&, RealRec(1)
    ShowPctComp Cnt&, NumRealRecs&
  NEXT

  QPrintRC "    Reading Customer Information    ", 11, 23, -1
  REDIM CPins(1 TO NumOfCRecs&) AS LONG
  FOR Cnt& = 1 TO NumOfCRecs&
    GET #TaxFile, Cnt&, TaxCust(1)     'read the customers PINS
    CPins(Cnt&) = TaxCust(1).PIN
    TaxCust(1).FirstPropRec = 0        'zero old pointers
    TaxCust(1).FirstPersRec = 0        'zero old pointers
    PUT #TaxFile, Cnt&, TaxCust(1)     'update cust file
    ShowPctComp Cnt&, NumOfCRecs&
  NEXT
  
'Process the Personal abstracts
  QPrintRC "    Relinking Personal Abstracts    ", 11, 23, -1
  FOR PersCnt& = 1 TO NumPersRecs&
    GET #PersFile, PersCnt&, PersRec(1)
    ThisPin& = PersRec(1).CustPin&
    FOR CCnt& = 1 TO NumOfCRecs&             'Look for pin in customer
      IF ThisPin& = CPins(CCnt&) THEN        'found matching pin
        GET #TaxFile, CCnt&, TaxCust(1)      'get the cust rec
        IF TaxCust(1).FirstPersRec > 0 THEN  'if there are others
          ThisPersRec& = TaxCust(1).FirstPersRec
          DO
            GET #PersFile, ThisPersRec&, PersRec(2)
            IF PersRec(2).NextRec = 0 THEN
              PersRec(2).NextRec = PersCnt&  'point to this pers rec
              PUT #PersFile, ThisPersRec&, PersRec(2)
              EXIT DO
            END IF
            ThisPersRec& = PersRec(2).NextRec
          LOOP
        ELSE                                   'no first personal rec
          TaxCust(1).FirstPersRec = PersCnt&   'point cust to this pers rec
          PUT #TaxFile, CCnt&, TaxCust(1)      'update cust file
          PersRec(1).NextRec = 0               'set pers next pointer to 0
          PUT #PersFile, PersCnt&, PersRec(1)  'update pers file
        END IF
        EXIT FOR   'done with this pers rec
      END IF
    NEXT
    ShowPctComp PersCnt&, NumPersRecs&
  NEXT

'Process the Real abstracts
  QPrintRC "      Relinking Real Abstracts      ", 11, 23, -1
  FOR RealCnt& = 1 TO NumRealRecs&
    GET #RealFile, RealCnt&, RealRec(1)
    ThisPin& = RealRec(1).CustPin&
    FOR CCnt& = 1 TO NumOfCRecs&             'Look for pin in customer
      IF ThisPin& = CPins(CCnt&) THEN        'found matching pin
        GET #TaxFile, CCnt&, TaxCust(1)      'get the cust rec
        IF TaxCust(1).FirstPropRec > 0 THEN  'if there are others
          ThisRealRec& = TaxCust(1).FirstPropRec
          DO
            GET #RealFile, ThisRealRec&, RealRec(2)
            IF RealRec(2).NextRec = 0 THEN
              RealRec(2).NextRec = RealCnt&  'point to this real rec
              PUT #RealFile, ThisRealRec&, RealRec(2)
              EXIT DO
            END IF
            ThisRealRec& = RealRec(2).NextRec
          LOOP
        ELSE                                   'no first personal rec
          TaxCust(1).FirstPropRec = RealCnt&   'point cust to this pers rec
          PUT #TaxFile, CCnt&, TaxCust(1)      'update cust file
          PersRec(1).NextRec = 0               'set pers next pointer to 0
          PUT #RealFile, RealCnt&, RealRec(1)  'update pers file
        END IF
        EXIT FOR   'done with this pers rec
      END IF
    NEXT
    ShowPctComp RealCnt&, NumRealRecs&
  NEXT

  CLOSE

END SUB

SUB TAXSetup

  ClearBack
  SHARED Choice$()

  

  LibName$ = "TAX"
  ScrnName$ = "TAXSETUP"
  help$ = "Cash Management Operator Entry"
  LOCATE 1, 1, 0

  help$ = "Add Misc Codes V8.0"
  ShowCursor
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  

  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 11)
  REDIM Fld(NumFlds) AS FieldInfo

  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F

  REDIM Choice$(0 TO 3, 0 TO 1)
  Choice$(0, 0) = "5"
  Choice$(1, 0) = "NC"
  Choice$(2, 0) = "VA"
  Choice$(3, 0) = "GA"

  Choice$(0, 1) = "6"
  Choice$(1, 1) = "POSTCARD"
  Choice$(2, 1) = "LASER"
  Choice$(3, 1) = "MULTI-PART"
  


  Action = 1
  NoProcess = False


    TxFile = FREEFILE
    OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetUpRec)
    IF LOF(TxFile) > 0 THEN
     GET TxFile, 1, TaxSetUpRec
    Form$(1, 0) = TaxSetUpRec.Name
    Form$(2, 0) = TaxSetUpRec.Add1
    Form$(3, 0) = TaxSetUpRec.Add2
     Form$(4, 0) = TaxSetUpRec.Add3
     Form$(5, 0) = TaxSetUpRec.TaxSt
     Form$(6, 0) = TaxSetUpRec.TaxForm

    Form$(7, 0) = STR$(TaxSetUpRec.CurRate)
    Form$(8, 0) = STR$(TaxSetUpRec.PastRate)
    Form$(9, 0) = STR$(TaxSetUpRec.PenRate)
    CLOSE TxFile
  END IF



   DO
     EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE F7KEY
    CASE F10Key
      SaveFlag = 1
      Done = True
      CASE ESC
      Done = True
    END SELECT


   LOOP UNTIL Done
   IF SaveFlag THEN
    TxFile = FREEFILE
    OPEN "TAXSETUP.DAT" FOR RANDOM AS #TxFile LEN = LEN(TaxSetUpRec)

    TaxSetUpRec.Name = Form$(1, 0)
    TaxSetUpRec.Add1 = Form$(2, 0)
    TaxSetUpRec.Add2 = Form$(3, 0)
    TaxSetUpRec.Add3 = Form$(4, 0)
    TaxSetUpRec.TaxSt = Form$(5, 0)
    TaxSetUpRec.TaxForm = Form$(6, 0)
                        
    TaxSetUpRec.CurRate = Value(Form$(7, 0), ecode)
    TaxSetUpRec.PastRate = Value(Form$(8, 0), ecode)
    TaxSetUpRec.PenRate = Value(Form$(9, 0), ecode)

    TaxSetUpRec.Padding = ""
    PUT TxFile, 1, TaxSetUpRec
    CLOSE TxFile

   END IF
   
END SUB

