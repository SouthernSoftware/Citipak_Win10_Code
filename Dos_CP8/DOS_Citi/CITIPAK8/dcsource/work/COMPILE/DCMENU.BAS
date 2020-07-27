DEFINT A-Z
DECLARE SUB WaitForAction ()
DECLARE SUB DisplayDCScrn (ScrnName$)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ClearBack ()
DECLARE SUB RelinkDCStuff ()
DECLARE SUB SortDCNameIndex ()
DECLARE SUB DecalUtil ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB HideCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf AS ANY)
DECLARE FUNCTION Exist% (FileName$)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'DC.BI'

  STACK 8000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 6)
  
  MChoice$(1) = "Decal Catagory Maintenance"
  MChoice$(2) = "Customer Maintenance"
  MChoice$(3) = "Customer Reports"
  MChoice$(4) = "Purchase Decal"
  'MChoice$(5) = "Decal Utilities"
  MChoice$(5) = "Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight

'  GOTO Here:

  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 1

  'RelinkDCStuff
  Help$ = "Vehicle Decals"
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    LibFile2Scrn "DC.QSL", "MENUBAK", MonoCode, -1, ErrorCode
    
    TitleBox 3, Col, MaxLen + 3, "Vehicle Decals v3.0", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf
    
    IF ky$ = CHR$(27) THEN EXIT DO              'choice = 0
'
    SELECT CASE Choice
    CASE 1
      RUN "dccode"
    CASE 2
      RUN "dccust"
    CASE 3
      RUN "dcreport"
    CASE 4
      RUN "dcpay"
'    CASE 5
'      DecalUtil
    CASE 5
      HideCursor
      CLS
      END
    END SELECT
  LOOP

  IF Exist("CPMENU.EXE") THEN
    RUN "cpmenu"
  END IF
  HideCursor
  CLS
END

Here:

'  Null$ = CHR$(0)
'  REDIM TmpCustRec(1 TO 2) AS DCCustRecType
'  DCCustRecLen = LEN(TmpCustRec(1))
'
'  DCFile = FREEFILE
'  OPEN "DCCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCFile LEN = DCCustRecLen
'  NumOfDCRecs = LOF(DCFile) \ DCCustRecLen
'  FOR Cnt = 1343 TO 2057
'    'GET DCFile, Cnt, TmpCustRec(1)
'    TmpCustRec(1).Deleted = "Y"
'    PUT DCFile, Cnt, TmpCustRec(1)
'  NEXT
'  CLOSE

SUB DecalUtil

  REDIM MChoice$(1 TO 3)

  MChoice$(1) = "Relink Decal Information"
  MChoice$(2) = "Reindex Customer File"
  MChoice$(3) = "Exit to OS"

  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 1

  Help$ = "Vehicle Decals"

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    LibFile2Scrn "DC.QSL", "MENUBAK", MonoCode, -1, ErrorCode

    TitleBox 3, Col, MaxLen + 3, "Decal Utilities Menu", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf

    IF ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1
      RelinkDCStuff
    CASE 2
      SortDCNameIndex
    CASE 3
      HideCursor
      CLS
      END

    END SELECT
  LOOP

'  IF Exist("CPMENU.EXE") THEN
'    RUN "cpmenu"
'  END IF
'  HideCursor
'  CLS
'  END

END SUB

