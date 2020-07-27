DEFINT A-Z
DECLARE SUB SmallPause ()
DECLARE FUNCTION ICRound# (N#)
DECLARE SUB ICShowProcessingScrn (RptTitle$)
DECLARE SUB ICShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB PostICTransactions ()
DECLARE SUB DisplayICScrn (ScrnName$)
DECLARE FUNCTION SelectTRRec% ()
DECLARE SUB TransEntry (TRRec%)
DECLARE SUB ClearBack ()
DECLARE SUB ILookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)
DECLARE SUB CustomerListing ()
DECLARE SUB TransactionJournal ()
DECLARE SUB OpenARCustIdxFile (NumOfARIdxRecs%, ARIdxFile%)
DECLARE SUB ARFixMess ()
DECLARE SUB ShowNoCode ()
DECLARE SUB OpenARCustFile (NumOfArRecs, ARFile)
DECLARE SUB printhelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB HideCursor ()
DECLARE SUB QPrint (x$, Colr%, Page%)
DECLARE SUB QPrintRC (t$, r%, c%, clr%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE FUNCTION GetICItemCnt& ()
DECLARE FUNCTION GetTRRecCnt% ()
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB ICVertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'IC.bi'
  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE

  TYPE FLen2
    v AS STRING * 64
  END TYPE
  
  CONST False = 0, True = NOT False
  
  STACK 8000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 5)
  
  MChoice$(1) = "Add NEW Transaction"
  MChoice$(2) = "Edit EXISTING Transaction"
  MChoice$(3) = "Print Transaction Edit List"
  MChoice$(4) = "Post Transactions"
  MChoice$(5) = "Exit to OS"
  
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
  Col = ((80 - MaxLen) \ 2)
  help$ = "Main Reports Menu"
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    ClearBack
    
    TitleBox 3, Col, MaxLen + 3, "Transaction Menu", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO                    'choice = 0
    
    SELECT CASE Choice
    CASE 1
      TransEntry 0
    CASE 2
      IF GetICItemCnt& > 0 AND GetTRRecCnt > 0 THEN
        TRRec = SelectTRRec
        IF TRRec > 0 THEN
          TransEntry TRRec
        END IF
      END IF
    CASE 3
      TransactionJournal
    CASE 4
      PostICTransactions
    CASE 5
      HideCursor
      CLS
      END
    END SELECT
  LOOP
  RUN "icmenu"

FUNCTION GetTRRecCnt
  REDIM ICTrans(1) AS ICTransRecType
  ICTRRecLen = LEN(ICTrans(1))
  ICTRFile = FREEFILE
  OPEN "ICTRANST.DAT" FOR RANDOM SHARED AS ICTRFile LEN = ICTRRecLen
  NumTRRecs = LOF(ICTRFile) \ ICTRRecLen
  CLOSE ICTRFile
  GetTRRecCnt = NumTRRecs
END FUNCTION

SUB PostICTransactions

  LibFile$ = "IC"

  REDIM ItemRec(1) AS ItemRecType
  ItemRecLen = LEN(ItemRec(1))

  REDIM ICTrans(1) AS ICTransRecType
  ICTRRecLen = LEN(ICTrans(1))

  IF GetTRRecCnt < 1 THEN
    ClearBack
    Ok = MsgBox(LibFile$, "NON2POST")
    GOTO PostExit
  END IF
  ClearBack
  Ok = MsgBox(LibFile$, "POSTTRAN")
  SELECT CASE Ok
  CASE 1   'F10 Yep, ready post'em
    OK2Post = True
  CASE 2   'Esc nope, don't post
    OK2Post = False
    Ok = MsgBox(LibFile$, "POSTCAN")
  END SELECT

  IF NOT OK2Post THEN
    GOTO PostExit
  END IF
  
  ClearBack
  PostDate = Date2Num(DATE$)

  ICTRData = FREEFILE
  OPEN "ICTRANST.DAT" FOR RANDOM SHARED AS ICTRData LEN = ICTRRecLen
  Num2Post = LOF(ICTRData) \ ICTRRecLen

  ICTRFile = FREEFILE
  OPEN "ICTRANS.DAT" FOR RANDOM SHARED AS ICTRFile LEN = ICTRRecLen
  
  ICItem = FREEFILE
  OPEN "ICITEM.DAT" FOR RANDOM SHARED AS ICItem LEN = ItemRecLen
  NumItemRec = LOF(ICItem) \ ItemRecLen

  NextRec& = (LOF(ICTRFile) \ ICTRRecLen)

  ICShowProcessingScrn "Posting Transactions"

  FOR Cnt = 1 TO Num2Post
    GET #ICTRData, Cnt, ICTrans(1)
    IF ICTrans(1).DelFlag = 0 THEN
      WhatItem = ICTrans(1).ItemRec
      IF WhatItem <= 0 OR WhatItem > NumItemRec THEN
        'this should never happen, but. . .
        BadTR = True
        'Show an error screen here.
        GOTO PostExit
      ELSE
        GET #ICItem, WhatItem, ItemRec(1)
      END IF

      'we have item & transrec by here
      SELECT CASE ICTrans(1).TransType   'what kind of trans is it
        CASE 1   'up adj
          TotQtyOnHnd# = ICRound#(ItemRec(1).QtyOnHand + ICTrans(1).TransQuantity)
          ItemRec(1).QtyOnHand = TotQtyOnHnd#   'adj qty on hand
          TotUnitYTD# = ICRound#(ItemRec(1).QtyOnHand + ItemRec(1).UsedYTD)
          TCost# = ICRound#(ICTrans(1).TransCost + ItemRec(1).COST)
          ItemRec(1).COST = ICRound#(TCost# / TotUnitYTD#)
          
        CASE 2   'dn adj
          TotQtyOnHnd# = ICRound#(ItemRec(1).QtyOnHand - ICTrans(1).TransQuantity)
          ItemRec(1).QtyOnHand = TotQtyOnHnd#   'adj qty on hand
        CASE 3   'reorder
          TotQtyOnOrd# = ICRound#(ItemRec(1).QtyOnOrd + ICTrans(1).TransQuantity)
          ItemRec(1).QtyOnOrd = TotQtyOnOrd#     'adj qty on order
        CASE 4   'used
          TotQtyOnHnd# = ICRound#(ItemRec(1).QtyOnHand - ICTrans(1).TransQuantity)
          ItemRec(1).QtyOnHand = TotQtyOnHnd#    'adj qty on hand
          TotUsed# = ICRound#(ItemRec(1).UsedYTD + ICTrans(1).TransQuantity)
          ItemRec(1).UsedYTD = TotUsed#     'adj used YTD total
        CASE 5   'purchased
          TotQtyOnHnd# = ICRound#(ItemRec(1).QtyOnHand + ICTrans(1).TransQuantity)
          ItemRec(1).QtyOnHand = TotQtyOnHnd#   'Adj qty on hand
          TotQtyOnOrd# = ICRound#(ItemRec(1).QtyOnOrd - ICTrans(1).TransQuantity)
          ItemRec(1).QtyOnOrd = TotQtyOnOrd#    'Adj qty on order

        CASE ELSE

      END SELECT
      'now put the updated Item rec back...
      ICTrans(1).PostedDate = PostDate
      NextRec& = NextRec& + 1
      ICTrans(1).PrevTran = ItemRec(1).LastTran  'point this trans to previous
      ItemRec(1).LastTran = NextRec&             'set item last trans to this
      PUT #ICItem, WhatItem, ItemRec(1)          'write item back to file
      PUT #ICTRFile, NextRec&, ICTrans(1)        'write the new transaction
    END IF
    ICShowPctComp Cnt, Num2Post
  NEXT

  CLOSE

  ClearBack
  DisplayICScrn "UPDATEOK"
  WaitForAction

PostExit:

END SUB

FUNCTION SelectTRRec

  REDIM ICTrans(1) AS ICTransRecType
  ICTRRecLen = LEN(ICTrans(1))

  REDIM TItemRec(1) AS ItemRecType
  ItemRecLen = LEN(TItemRec(1))

  REDIM TransList(1 TO 1) AS FLen2
  
  ItemFile = FREEFILE
  OPEN "ICITEM.DAT" FOR RANDOM SHARED AS ItemFile LEN = ItemRecLen

  ICTRFile = FREEFILE
  OPEN "ICTRANST.DAT" FOR RANDOM SHARED AS ICTRFile LEN = ICTRRecLen
  NumTRRecs = LOF(ICTRFile) \ ICTRRecLen
  FOR Cnt = 1 TO NumTRRecs
    GET ICTRFile, Cnt, ICTrans(1)

'    ICTrans(1).DelFlag = False
'    PUT ICTRFile, Cnt, ICTrans(1)

    IF ICTrans(1).DelFlag = False THEN
      TRCnt = TRCnt + 1
      REDIM PRESERVE TransList(1 TO TRCnt) AS FLen2
      GET ItemFile, ICTrans(1).ItemRec, TItemRec(1)
      SELECT CASE ICTrans(1).TransType
      CASE 1
        TRType$ = " Up-Adj "
      CASE 2
        TRType$ = " Dn-Adj "
      CASE 3
        TRType$ = " Reorder"
      CASE 4
        TRType$ = "  Used  "
      CASE 5
        TRType$ = "Purchase"
      END SELECT
      LSET TransList(TRCnt).v = FUsing$(STR$(Cnt), "####")
      MID$(TransList(TRCnt).v, 6) = FUsing$(STR$(TItemRec(1).ITEMNUMBER), "######")
      MID$(TransList(TRCnt).v, 16, 15) = ICTrans(1).TransNote
      MID$(TransList(TRCnt).v, 35) = Num2Date(ICTrans(1).TransDate)
      MID$(TransList(TRCnt).v, 48) = TRType$
      MID$(TransList(TRCnt).v, 61) = MKI$(Cnt)
    END IF
  NEXT
  CLOSE ICTRFile, ItemFile

  IF TRCnt = 0 THEN
    SelectTRRec = 0
    EXIT FUNCTION
  END IF

  '----- Give menu of items and let User make choices
  Action = 1                                  'Set "Initialize Menu" Flag
                                              '  for multi-tasking mode
  BoxBot = 16                                  'No limit on bottom line
  Choice = 1
  ClearBack

  MaxLen = 58
  ColPos = ((80 - MaxLen) / 2) - 1

  TitleBox 20, ColPos, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
  Text$ = SPACE$(MaxLen + 4)
  LSET Text$ = " TRNo.   ItemNo.      TRNote          TRDate      TRType"
  QPrintRC Text$, 4, ColPos, 112
  LOCATE 5, ColPos, 0
  DO
    ICVertMenuT TransList(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    SELECT CASE Ky$
    CASE CHR$(27)
      SelectTRRec = 0
      ExitFlag = True
    CASE CHR$(13)
      SelectTRRec = CVI(MID$(TransList(Choice).v, 61))
      ExitFlag = True
    END SELECT

  LOOP UNTIL ExitFlag

END FUNCTION

SUB TransactionJournal

  REDIM TRType$(1 TO 5)
  TRType$(1) = " UP-ADJ."
  TRType$(2) = " DN-ADJ."
  TRType$(3) = " Reorder"
  TRType$(4) = "  Used  "
  TRType$(5) = "Purchase"
  
  MaxLines = 55

  RPTFileName$ = "ICTRJOUR.RPT"

  REDIM ItemRec(1) AS ItemRecType
  ItemRecLen = LEN(ItemRec(1))

  REDIM ICTrans(1) AS ICTransRecType
  ICTRRecLen = LEN(ICTrans(1))
  
  ICTRFile = FREEFILE
  OPEN "ICTRANST.DAT" FOR RANDOM SHARED AS ICTRFile LEN = ICTRRecLen
  NumTR = LOF(ICTRFile) \ ICTRRecLen

  RPTFile = FREEFILE
  OPEN RPTFileName$ FOR OUTPUT AS RPTFile
  
  ClearBack
  ICShowProcessingScrn "Transaction Edit List."
  GOSUB JournalHeader
  FOR Cnt = 1 TO NumTR
    GET #ICTRFile, Cnt, ICTrans(1)
    IF ICTrans(1).DelFlag = False THEN
      ACTRCnt = ACTRCnt + 1
      PRINT #RPTFile, Num2Date$(ICTrans(1).TransDate); TAB(13); TRType$(ICTrans(1).TransType); USING "########.###"; TAB(22); ICTrans(1).TransQuantity;
      PRINT #RPTFile, USING "########.##"; TAB(35); ICTrans(1).TransCost;
      PRINT #RPTFile, TAB(50); LEFT$(ICTrans(1).TransNote, 30)
      LineCnt = LineCnt + 1
      IF LineCnt >= MaxLines THEN
        PRINT #RPTFile, CHR$(12)
        GOSUB JournalHeader
      END IF
    '  SmallPause
      ShowPctComp Cnt, NumTR
    END IF
  NEXT
  PRINT #RPTFile, STRING$(79, "-")
  PRINT #RPTFile, "  Transactions:"; ACTRCnt
  PRINT #RPTFile, CHR$(12)

  CLOSE

  PrintRptFile "Transaction Edit List.", RPTFileName$, LPTPort%, RetCode%, 0


EXIT SUB

JournalHeader:
  Page = Page + 1
  PRINT #RPTFile, TAB(25); "Inventory Transaction Edit List"
  PRINT #RPTFile, "Report Date: "; DATE$; TAB(69); "Page:"; Page
  PRINT #RPTFile, " TR Date     TR Type       TR QTY     TR COST        Transaction Note"
  PRINT #RPTFile, STRING$(79, "-")
  LineCnt = 4
RETURN

END SUB

SUB TransEntry (TRRec)
  REDIM ItemRec(1) AS ItemRecType
  ItemRecLen = LEN(ItemRec(1))

  SHARED Choice$()

mainbody:
  REDIM TempScrn(0)

  LibName$ = "IC"
  ScrnName$ = "TRANS"

  'IF TRRec > 0 THEN
  '  help$ = "Editing Transation Entry"
  'ELSE
  '  help$ = "NEW Transation Entry"
  'END IF
  CursorOff
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  'Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  
  Action = 1
  FirstTime = True

  REM check for code file


  REDIM Choice$(0 TO 5, 0 TO 1)

  Choice$(0, 0) = "3"
  Choice$(1, 0) = "1) Upward Adjustment"
  Choice$(2, 0) = "2) Downward Adjustment"
  Choice$(3, 0) = "3) Item Reorder"
  Choice$(4, 0) = "4) Item Used"
  Choice$(5, 0) = "5) Item Purchased"

  '1=Up Adj 2=Dwn Adj 3=Redorder 4=Used 5=Purchased

  Frm(1).FldNo = 1

  ClearBack
  DisplayICScrn ScrnName$
  ShowCursor
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      IF TRRec > 0 THEN
        GOSUB GetTRRec
        Fld(1).Protected = True
        ModeMsg$ = "EDIT MODE"
        LibFile2Scrn "IC", "DELBTN", MonoCode%, -2, ErrorCode%
      ELSE
        LSET Form$(2, 0) = DATE$
        ModeMsg$ = "ADD MODE"
      END IF
      Action = 1
      QPrintRC ModeMsg$, 22, 6, 11
    END IF

    IF Frm(1).PrevFld = 1 AND Frm(1).FldNo <> 1 THEN
      TempItem& = VAL(Form$(1, 0))
      IF TempItem& > 0 THEN
        GOSUB GetItem
        IF ItemRecNo& = 0 THEN
          Frm(1).FldNo = 1
        ELSE
          Fld(1).Protected = True
        END IF
      ELSE
        SaveScrn TempScrn()
        ClearBack
        ILookUp TempItem&, "Item", False, True, False
        IF TempItem& = 0 THEN
          Frm(1).FldNo = 1
        ELSE
          GOSUB GetItem
          IF ItemRecNo& = 0 THEN
            Frm(1).FldNo = 1
          ELSE
            Fld(1).Protected = True
          END IF
        END IF
        Action = 1
        RestScrn TempScrn()
      END IF
    END IF

    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 30 TO 43            '--F2 HH Code list
          PressButton F3Key, 22, 30, 43
        END SELECT
      END SELECT
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF Frm(1).FldNo > 1 THEN
        GOSUB SaveRecord
        IF NOT BadField THEN
          ExitFlag = True
        END IF
      END IF
    CASE F3Key
      IF TRRec > 0 THEN
        GOSUB DelTrans
      END IF
    CASE EscKey
      NeedtoSort = True         ' set to true for testing
      ExitFlag = True
    END SELECT
  LOOP UNTIL ExitFlag

EXIT SUB

SaveRecord:
'  put validation and parsing to actual trans record here.

  BadField = False
  REDIM ICTrans(1) AS ICTransRecType
  ICTRRecLen = LEN(ICTrans(1))
  ICTrans(1).ItemRec = ItemRecNo&

  TRDate = Date2Num(Form$(2, 0))
  IF TRDate < 0 THEN
    BadField = True
    Frm(1).FldNo = 2
    GOTO SaveTRReturn
  END IF
  ICTrans(1).TransDate = Date2Num(Form$(2, 0))

  TR$ = QPTrim$(Form$(3, 0))
  IF LEN(TR$) = 0 THEN
    BadField = True
    Frm(1).FldNo = 3
    GOTO SaveTRReturn
  END IF
  TRType = VAL(LEFT$(TR$, 1))
  ICTrans(1).TransType = TRType

  TR$ = QPTrim$(Form$(4, 0))
  IF LEN(TR$) = 0 THEN
    BadField = True
    Frm(1).FldNo = 4
    GOTO SaveTRReturn
  END IF
  ICTrans(1).TransQuantity = Value#(TR$, ECode)

  TR$ = QPTrim$(Form$(5, 0))
  IF LEN(TR$) = 0 THEN
    BadField = True
    Frm(1).FldNo = 5
    GOTO SaveTRReturn
  END IF
  ICTrans(1).TransCost = Value#(TR$, ECode)
  
  ICTrans(1).TransNote = QPTrim$(Form$(6, 0))
  
  ICTRFile = FREEFILE
  OPEN "ICTRANST.DAT" FOR RANDOM SHARED AS ICTRFile LEN = ICTRRecLen
  NumTRRecs = LOF(ICTRFile) \ ICTRRecLen
  NextRecord = NumTRRecs + 1
  IF TRRec > 0 THEN                      'updating an existing transaction
    PUT ICTRFile, TRRec, ICTrans(1)
  ELSE                                   'creating a new transaction
    EntryDate = Date2Num(DATE$)
    ICTrans(1).EntryDate = EntryDate     'store TR creation date. will be
    PUT ICTRFile, NextRecord, ICTrans(1) 'useful in future!
  END IF
  CLOSE ICTRFile

SaveTRReturn:
  IF BadField THEN
    Ok = MsgBox("IC", "BADDATA")
    Action = 1
  ELSE
    DisplayICScrn "UPDATEOK"
    WaitForAction
  END IF
RETURN

GetItem:
  BadItem = False
  ItemFile = FREEFILE
  OPEN "ICITEM.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ItemFile LEN = ItemRecLen
  NumOfItemRecs = LOF(ItemFile) \ ItemRecLen
  IF TempItem& > NumOfItemRecs THEN
    ItemRecNo& = 0
    BadItem = True
    GOTO GetItemRet
  ELSE
    GET #ItemFile, TempItem&, ItemRec(1)
    LSET Form$(1, 0) = QPTrim$(STR$(ItemRec(1).ITEMNUMBER))
    LSET Form$(7, 0) = QPTrim$(STR$(ItemRec(1).COST))
    LSET Form$(8, 0) = ItemRec(1).Unit
    LSET Form$(9, 0) = QPTrim$(STR$(ItemRec(1).QtyOnHand))
    LSET Form$(10, 0) = QPTrim$(STR$(ItemRec(1).QtyOnOrd))
    LSET Form$(11, 0) = QPTrim$(STR$(ItemRec(1).ReOrder))
    LSET Form$(12, 0) = QPTrim$(STR$(ItemRec(1).UsedYTD))
    LSET Form$(13, 0) = ItemRec(1).DESC1
    LSET Form$(14, 0) = ItemRec(1).DESC2
  END IF
  Action = 1
  ItemRecNo& = TempItem&

GetItemRet:
  CLOSE
RETURN

GetTRRec:
  REDIM ICTrans(1) AS ICTransRecType
  ICTRRecLen = LEN(ICTrans(1))

  ICTRFile = FREEFILE
  OPEN "ICTRANST.DAT" FOR RANDOM SHARED AS ICTRFile LEN = ICTRRecLen
  GET ICTRFile, TRRec, ICTrans(1)
  CLOSE ICTRFile

  'ICTrans(1).ItemRec = ItemRecNo&
  LSET Form$(2, 0) = Num2Date(ICTrans(1).TransDate)
  LSET Form$(3, 0) = Choice$(ICTrans(1).TransType, 0)
  LSET Form$(4, 0) = STR$(ICTrans(1).TransQuantity)
  LSET Form$(5, 0) = STR$(ICTrans(1).TransCost)
  LSET Form$(6, 0) = ICTrans(1).TransNote
  TempItem& = ICTrans(1).ItemRec
  GOSUB GetItem

RETURN

DelTrans:
  Ok = MsgBox("IC", "DELTRANS")
  SELECT CASE Ok
  CASE 1       'they are sure? ok kill it.
    ICTRFile = FREEFILE
    OPEN "ICTRANST.DAT" FOR RANDOM SHARED AS ICTRFile LEN = ICTRRecLen
    GET ICTRFile, TRRec, ICTrans(1)
    ICTrans(1).DelFlag = True
    PUT ICTRFile, TRRec, ICTrans(1)
    CLOSE
    ExitFlag = True
  CASE ELSE    'nope don't
    Action = 1
  END SELECT
RETURN

END SUB

