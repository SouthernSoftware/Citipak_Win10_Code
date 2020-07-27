DECLARE SUB GetFYDates (FY1BegDate%, FY1EndDate%, FY2BegDate%, FY2EndDate%)
DECLARE SUB fixunformatted ()
DEFINT A-Z
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB FileView (FileName$, Ky, Action, FVI AS ANY, SEG Array)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, colr%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE SUB QPrint (X$, colr%, page%)
DECLARE SUB QPrintRC (T$, r%, c%, Clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
DECLARE FUNCTION Value# (E$, ErrCode%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
'DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ClrBG ()
DECLARE SUB MnuTitle (Title$, Row, Col, MaxLen, colr)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)

  TYPE FLen
    V AS STRING * 42
  END TYPE

'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'GL.BI'                        'GL File Types
'$INCLUDE: 'SCRAUX.BI'

DECLARE FUNCTION FldNum% (FldName$, Fld() AS FieldInfo)
DECLARE SUB EdBgtEntry ()
DECLARE SUB OpenBgtEditFile (BgtEditFileNum%, NumEdTrans%)
DECLARE SUB NewBgtEntry ()
DECLARE FUNCTION GetNextRec% (FileNum%, NumRecs%, CurrRec%, NextRec%)
DECLARE SUB PrnEditList ()
DECLARE SUB PrnBgtTran ()
DECLARE SUB DisplayTotals ()
DECLARE SUB PostTrans ()
DECLARE FUNCTION EDCheck4Acct% (Form$(), AcctFld%)

DIM SHARED TotLen, FundLen, AcctLen, DetLen
DIM SHARED AcctIdx AS GLAcctIndexType
DIM SHARED Acct AS GLAcctRecType
DIM SHARED BgtEdit AS TrEditRecType
DIM SHARED Trans AS GLTransRecType

CONST False = 0, True = NOT False

'REDIM SHARED Choice$(0, 0)                'Used to hold multiple choices

'$INCLUDE: 'SetCnf.bi'                    'Set monitor and color information
'$INCLUDE: 'init.bi'                      'Get GL system info

'TextCursor -2, -2                         'Define the Mouse cursor
'ShowCursor                                'Show it

'fixunformattedbgt
'END

'***Master Menu

   DIM mChoice$(1 TO 6)

   mChoice$(1) = "Enter/Edit Budget Transactions"
   mChoice$(2) = "Print Transaction Register"
   mChoice$(3) = "Post Entries to Budget"
   mChoice$(4) = STRING$(30, CHR$(196))
   mChoice$(5) = "Budget Preparation"
   mChoice$(6) = "Quit"

   MaxLen = 0                           'Set menu width to zero
   BoxBot = 20                            'limit the box to line 20
   Action = 0                             '0 means stay in the menu
   Choice = 1                             'Pre-load choice to highlight

   FOR Cnt = 1 TO UBOUND(mChoice$)        'Find max menu width
     TLen = LEN(mChoice$(Cnt))
       IF TLen > MaxLen THEN
         MaxLen = TLen
       END IF
   NEXT

   '--Center Menu within Screen
   Row = ((25 - (UBOUND(mChoice$))) \ 2) - 1
   Col = ((80 - MaxLen) \ 2) - 2

   ShowCursor

   DO
      LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, Attribute, ErrorCode
      'TitleBox 3, Col, MaxLen + 3, "Budget Processing", Cnf
      'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

      Title$ = "Budget Processing "
      MnuTitle Title$, Row, Col, MaxLen, 15
      
      PrintTitle User$
      PrintHelp "Budget Processing"
      
      LOCATE Row, Col, 0
      VertMenu mChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

      IF Ky$ = CHR$(27) THEN EXIT DO

      SELECT CASE Choice
         CASE 1
            ClrBG
            EdBgtEntry
         CASE 2
            ClrBG
            PrnEditList
         CASE 3
            ClrBG
            PostTrans
         'CASE 4
         '   CLS : END
         CASE 5
            ClrBG
            RUN "BGTPREP"
         CASE 6
            CLS : END
         CASE ELSE
      END SELECT

   LOOP

ClrBG
HideCursor
DispLoad
RUN "glmenu"

SUB DisplayTotals

   OpenBgtEditFile BgtEditFileNum, NumEdTrans

   TotDr# = 0: TotCr# = 0                    'reset totals to zero

   FOR Cnt = 1 TO NumEdTrans                 're-summarize totals
      GET BgtEditFileNum, Cnt, BgtEdit
      IF NOT BgtEdit.Deleted THEN
         TotDr# = TotDr# + BgtEdit.DrAmt
         TotCr# = TotCr# + BgtEdit.CrAmt
      END IF
   NEXT
   'STOP
   CLOSE BgtEditFileNum

   TotDr$ = FUsing$(STR$(TotDr#), "##########,.##")
   TotCr$ = FUsing$(STR$(TotCr#), "##########,.##")

   QPrintRC TotDr$, 18, 23, 15               'print the file totals
   QPrintRC TotCr$, 18, 53, 15               ' on the screen
   
END SUB

SUB EdBgtEntry

   GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate

   SHARED Choice$()                          'Multiple choice array
   'SHARED Cnf AS Config                      'System config & colors

   REDIM Form$(0, 0)                         'Holds data from form
   REDIM Fld(1) AS FieldInfo                 'Field editing array

   REDIM Choice$(0 TO 2, 1)                  'Choices for entry type field
   Choice$(0, 0) = "5"                       'Field Number
   Choice$(1, 0) = "Increase "
   Choice$(2, 0) = "Decrease"

   EdStatus$ = "Editing Budget Items"
   NewStatus$ = "New Budget Items"

   OpenBgtEditFile BgtEditFileNum, NumEdTrans  'Opens the Bgt Edit File
   
                                             'determine edit mode on entry
   IF NumEdTrans > 0 THEN                    'If there are any records
      
                                             'Call GetNextRec to get the
                                             'first record not deleted
      CurrRec = 0: NextRec = 1
      RecNum = GetNextRec(BgtEditFileNum, NumEdTrans, CurrRec, NextRec)
   
      IF RecNum = 0 THEN
         CLOSE                               'If All records are deleted..
         KILL "BgtED.DAT"                     '..init file to zero
         EMode = 1                           '..Set mode to New
      ELSE                                   'Otherwise...
         EMode = 0                           '..set Mode to Edit
      END IF
      
   ELSE                                      'No trans.. Set Add Mode
      EMode = 1
   
   END IF
   
   'call sub in GLAUX module to load screen & form
   LibLoadDisplayForm "GL.QSL", "EDBGT", Form$(), Fld()
   
   TrDate = FldNum("TRDATE", Fld())          'Assign field nums to variables
   AcctNum = FldNum("ACCTNUM", Fld())
   AcctName = FldNum("ACCTNAME", Fld())
   Desc = FldNum("DESC", Fld())
   Amt = FldNum("AMT", Fld())
   EType = FldNum("ETYPE", Fld())
   Ref = FldNum("REF", Fld())

   Form$(TrDate, 0) = DATE$
   Action = 1

   DIM Frm AS FormInfo                       'Dim the form info array

   CLOSE                                     'Close file while doing input

   IF RecNum > 0 THEN                        'if we have pending transactions
      GOSUB Rec2Form                         'parse it to the form
      DisplayTotals
      Action = 1                             'Init form edit
   END IF
   
   GOSUB ShowMode

   GOSUB ShowRecNum

   DO                                        'Edit the form

      EditForm Form$(), Fld(), Frm, Cnf, Action
      
      IF Frm.PrevFld <> Frm.FldNo THEN       'user is moving around in form

         IF Frm.FldNo = AcctNum THEN         'if we're back on acct number
                                             ' trim out dashes B4 editing
            Form$(AcctNum, 0) = StripAcct$(Form$(AcctNum, 0))
            Frm.FldNo = AcctNum              'Put user back on acct field
            SaveField AcctNum, Form$(), Fld(), 0
            Action = 1

         END IF

         SELECT CASE Frm.PrevFld             'user is moving around in form
            CASE AcctNum                     'user just left acct field
                                             'look up the account number
               LSET Form$(AcctNum, 0) = FmtAcct$(RTRIM$(Form$(AcctNum, 0)), FundLen%, AcctLen%, DetLen%)
               GOSUB CheckForValidAcctNum
               IF AcctOK THEN
                 RecordNum = FindAcct(Form$(AcctNum, 0))
                 AcctType$ = GetAcctType(RecordNum)
               END IF

            CASE ELSE
               'Other field validation stuff goes here

         END SELECT

      END IF
      
      SELECT CASE Frm.KeyCode                'Respond to the key presses
         CASE F10Key  'Save
            Frm.FldNo = 3
            Action = 1
            BadFld = 0
            GOSUB CheckFlds
            IF BadFld = 0 THEN
               GOSUB SaveChanges
               IF EMode = 1 THEN GOSUB ClearForm
               DisplayTotals
               Frm.FldNo = 1
               Action = 1
            END IF

            GOSUB ShowMode

         CASE F2Key  'switch Edit Mode

           IF EMode = 1 THEN
               IF RecNum > 0 THEN
                  EMode = 0    'In Add mode, switch to edit mode
               END IF

            ELSE
               EMode = 1  'In Edit mode, switch to add
               GOSUB ResetForm
               Record$ = "New Record"
               QPrintRC Record$, 25, 60, -1
            END IF

            GOSUB ShowMode

            IF EMode = 0 AND RecNum > 0 THEN  'were in add mode, show user list recs to edit
               GOSUB Find
               GOSUB Rec2Form
               GOSUB ShowRecNum
            END IF
            
         CASE F3Key  'Delete
            IF EMode = 0 THEN
              Ok = MsgBox%("GL.QSL", "OK2DELTR")
              IF Ok = 2 THEN                   'Button #2 = Yes
                 OpenBgtEditFile BgtEditFileNum, NumEdTrans
                 BgtEdit.Deleted = -1
                 PUT BgtEditFileNum, RecNum, BgtEdit
                 BgtEdit.Deleted = 0
                 CLOSE
                 GOSUB ClearForm
                 DisplayTotals
                 GOSUB Find
                 GOSUB Rec2Form
              END IF

              Frm.FldNo = TrDate 'here
              GOSUB ShowRecNum

           END IF

         CASE CtrlL  'Ctrl+Left=Prev Rec
            OpenBgtEditFile BgtEditFileNum, NumEdTrans
            CurrRec = RecNum: NextRec = -1
            RecNum = GetNextRec(BgtEditFileNum, NumEdTrans, CurrRec, NextRec)
            CLOSE
            GOSUB ShowRecNum
            IF RecNum > 0 THEN
              GOSUB Rec2Form
            END IF

         CASE CtrlR  'Ctrl+Right=Next Rec
            OpenBgtEditFile BgtEditFileNum, NumEdTrans
            CurrRec = RecNum: NextRec = 1
            RecNum = GetNextRec(BgtEditFileNum, NumEdTrans, CurrRec, NextRec)
            CLOSE
            GOSUB ShowRecNum
            IF RecNum > 0 THEN
              GOSUB Rec2Form
            END IF

         CASE F5Key  'F5=Find
            IF EMode = 0 THEN
               OpenBgtEditFile BgtEditFileNum, NumEdTrans
               IF NumEdTrans > 0 THEN
                  GOSUB Find                       'Show user list of transactions
                  GOSUB Rec2Form                   'parse it to the form
                  ELSE
                  BEEP
               END IF
               CLOSE                            'Close file while getting input
            END IF
            GOSUB ShowRecNum

         CASE ELSE

      END SELECT

   LOOP UNTIL Frm.KeyCode = ESC               'Loop until user presses Esc

EXIT SUB


ShowMode:
   Status$ = SPACE$(20)

   SELECT CASE EMode                         'assign status strings
      CASE 0                                 'based on editing mode
         F2BtnCap$ = "  F2=New  "
         F2BtnColor = 112
         F3BtnColor = 116
         F5BtnColor = 112
         LSET Status$ = EdStatus$
      CASE 1
         F2BtnCap$ = "  F2=Edit "
         IF RecNum > 0 THEN
            F2BtnColor = 112
         ELSE
            F2BtnColor = 115
         END IF
         F3BtnColor = 115
         F5BtnColor = 115
         LSET Status$ = NewStatus$
   END SELECT

   MPaintBox 25, 28, 25, 38, F2BtnColor       'Button disable/enable
   MPaintBox 25, 41, 25, 51, F5BtnColor
   MPaintBox 14, 56, 14, 68, F3BtnColor

   QPrintRC Status$, 25, 2, -1
   QPrintRC F2BtnCap$, 25, 28, -1
   
RETURN


CheckFlds:
   FOR Cnt = 1 TO 6                          'Check for blank fields
      IF LEN(LTRIM$(RTRIM$(Form$(Cnt, 0)))) = 0 THEN
         PLAY "L32O3EC"
         Ok = MsgBox("GL.QSL", "NOBLANK")
         BadFld = Cnt
         Frm.FldNo = BadFld
         EXIT FOR
      END IF
   NEXT                           ' FY1BegDate, FY1EndDate
   BTrDate = Date2Num(Form$(TrDate, 0))
   IF BTrDate < FY1BegDate OR BTrDate > FY1EndDate THEN
      Ok = MsgBox("GL", "BADDATE")
      BadFld = TrDate
      Frm.FldNo = BadFld
   END IF
   Action = 1

RETURN


ClearForm:
   FOR FieldNum = 1 TO 7
      Form$(FieldNum, 0) = ""
   NEXT
   Form$(TrDate, 0) = Num2Date(BgtEdit.TrDate)
   Form$(Desc, 0) = BgtEdit.Desc
   Form$(Ref, 0) = BgtEdit.Ref
   Frm.FldNo = 1
   Action = 1

RETURN


ResetForm:
   FOR FieldNum = 1 TO 7
      Form$(FieldNum, 0) = ""
   NEXT
   Form$(TrDate, 0) = DATE$
   Frm.FldNo = 1
   Action = 1

RETURN


SaveChanges:                                 'Save changes to the record
   OpenBgtEditFile BgtEditFileNum, NumEdTrans

   IF EMode = 1 THEN
      RecNum = NumEdTrans + 1     'adding new or editing existing
      'GOSUB ClearForm
   END IF

   BgtEdit.Deleted = 0
   'IF MID$(BgtEdit.AcctNum, FundLen + 1, 1) <> "-" THEN
   IF MID$(Form$(AcctNum, 0), FundLen + 1, 1) <> "-" THEN
     BgtEdit.AcctNum = FmtAcct$(Form$(AcctNum, 0), FundLen, AcctLen, DetLen)
   ELSE
     BgtEdit.AcctNum = Form$(AcctNum, 0)
   END IF
   BgtEdit.AcctName = Form$(AcctName, 0)
   BgtEdit.TrDate = Date2Num(Form$(TrDate, 0))'Store date as integer
   BgtEdit.Desc = Form$(Desc, 0)
   BgtEdit.EType = Form$(EType, 0)
   BgtEdit.Ref = Form$(Ref, 0)
   BgtEdit.Src = "BG" + ConvDateStr$(DATE$)

   Amt# = Value(Form$(Amt, 0), ErrCode)      'Convert Amt$ to Double
   IF ErrCode THEN
      CLS
      CLOSE
      PRINT "An error occured in the value function!"
      EXIT SUB
   END IF
                                             'Only one amount field in form!
   SELECT CASE BgtEdit.EType                  'Is amount decrease or increase
      CASE "D"
         SELECT CASE AcctType$
            CASE "R"
               BgtEdit.DrAmt = Amt#
               BgtEdit.CrAmt = 0
            CASE "E"
               BgtEdit.CrAmt = Amt#
               BgtEdit.DrAmt = 0
         END SELECT
      CASE "I"
         SELECT CASE AcctType$
            CASE "R"
               BgtEdit.DrAmt = 0
               BgtEdit.CrAmt = Amt#
            CASE "E"
               BgtEdit.CrAmt = 0
               BgtEdit.DrAmt = Amt#
         END SELECT
   
   END SELECT
   
   PUT BgtEditFileNum, RecNum, BgtEdit         'write it
   CLOSE BgtEditFileNum                       'close the file

RETURN


Rec2Form:                                    'Set form fields to rec values
   OpenBgtEditFile BgtEditFileNum, NumEdTrans
   GET BgtEditFileNum, RecNum, BgtEdit

   ARecNum = FindAcct(BgtEdit.AcctNum)

   AcctType$ = GetAcctType(ARecNum)
   
   Form$(TrDate, 0) = Num2Date(BgtEdit.TrDate)'Date stored as integer
   Form$(AcctNum, 0) = BgtEdit.AcctNum
   Form$(AcctName, 0) = BgtEdit.AcctName
   Form$(Desc, 0) = BgtEdit.Desc
   Form$(Ref, 0) = BgtEdit.Ref

   SELECT CASE BgtEdit.EType                  'set entry type fld to Dr or Cr
      CASE "D"
         Form$(EType, 0) = "Decrease"
         SELECT CASE AcctType$
            CASE "R"
               Form$(Amt, 0) = STR$(BgtEdit.DrAmt)
            CASE "E"
               Form$(Amt, 0) = STR$(BgtEdit.CrAmt)
         END SELECT
      CASE "I"
         Form$(EType, 0) = "Increase"
         SELECT CASE AcctType$
            CASE "R"
               Form$(Amt, 0) = STR$(BgtEdit.CrAmt)
            CASE "E"
               Form$(Amt, 0) = STR$(BgtEdit.DrAmt)
         END SELECT
      CASE ELSE
   END SELECT
   CLOSE
   Action = 1                                'force form edit to re-init

RETURN


ShowRecNum:
   IF RecNum > 0 THEN
      Record$ = "Record:     "
      MID$(Record$, 8) = STR$(RecNum)
      QPrintRC Record$, 25, 60, -1
   ELSE
      Record$ = "New Record  "
      QPrintRC Record$, 25, 60, -1
   END IF

RETURN


Find:                                        'Find Transaction
   OpenBgtEditFile BgtEditFileNum, NumEdTrans
   
   REDIM TransList$(1 TO NumEdTrans)         'Dim array to hold trans list
   Active = 0                                'Counter for Active Transactions

   FOR Cnt = 1 TO NumEdTrans                 'Read the file

      GET BgtEditFileNum, Cnt, BgtEdit
      IF NOT BgtEdit.Deleted THEN             'if active add to list
         Active = Active + 1

         ARecNum = FindAcct(BgtEdit.AcctNum)
         AcctType$ = GetAcctType(ARecNum)

         List$ = SPACE$(75)
         LSET List$ = ConvDateStr1(Num2Date(BgtEdit.TrDate))
         MID$(List$, 10) = BgtEdit.AcctNum
         MID$(List$, 28) = BgtEdit.Desc
         SELECT CASE BgtEdit.EType
            CASE "D"
               SELECT CASE AcctType$
                  CASE "R"
                     Amt# = BgtEdit.DrAmt
                     EType$ = "Dr"
                  CASE "E"
                     Amt# = BgtEdit.CrAmt
                     EType$ = "Cr"
               END SELECT

            CASE "I"
               SELECT CASE AcctType$
                  CASE "R"
                     Amt# = BgtEdit.CrAmt
                     EType$ = "Cr"
                  CASE "E"
                     Amt# = BgtEdit.DrAmt
                     EType$ = "Dr"
               END SELECT

         END SELECT

         MID$(List$, 48) = FUsing$(STR$(Amt#), "########,.##")
         MID$(List$, 63) = EType$
         MID$(List$, 70) = STR$(Cnt)         'Transaction record number

         TransList$(Active) = List$
         
      END IF
      
   NEXT
   CLOSE
   IF Active = 0 THEN
      KILL "BGTED.DAT"                     '..init file to zero
      EXIT SUB
   END IF

   REDIM PRESERVE TransList$(1 TO Active)    'Redim list to only active trans

   LOCATE 6, 6                               'set up vertmenu
   Action = 0
   MaxLen = 65
   BoxBot = 20

   ListDesc$ = "   Date     Account          Description                 Amt  Dr/Cr  "
   QPrintRC ListDesc$, 5, 6, Cnf.ActivCh
   MPaintBox 6, 75, 7, 76, 8              'to fix shadow

   DO                                        'call vertmenu
      VertMenu TransList$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
         QPrintRC STRING$(80, CHR$(177)), 5, 1, 3
         CLOSE
         EXIT SUB
         'RETURN
      END IF
      IF Choice > 0 THEN EXIT DO
   LOOP
                                             
   QPrintRC STRING$(80, CHR$(177)), 5, 1, 3
   RecNum = VAL(MID$(TransList$(Choice), 70, 5)) 'Get the record number
   MPaintBox 6, 75, 7, 76, 3              'to fix shadow
   ERASE TransList$                          'Clean up after ourselves
   
RETURN

CheckForValidAcctNum:
  AcctFld = Frm.PrevFld
  SELECT CASE EDCheck4Acct%(Form$(), AcctFld%)
  CASE True
    AcctOK = True
    SaveField AcctFld, Form$(), Fld(), BadFld   'Write the flag to the form$
    SaveField AcctFld + 1, Form$(), Fld(), BadFld               'Write the flag to the form$
    Action = 2
  CASE ELSE
    AcctOK = False
    Frm.FldNo = AcctFld      'Put user back on acct field
    LSET Form$(Frm.FldNo, 0) = ""
    LSET Form$(Frm.FldNo + 1, 0) = ""
    Action = 2
  END SELECT

  RETURN


 
END SUB

FUNCTION EDCheck4Acct% (Form$(), AcctFld%)

  SHARED TotLen

  List$ = SPACE$(65)

  Lookup$ = QPTrim$(UCASE$(Form$(AcctFld, 0)))
  Lookuplen = LEN(Lookup$)

  REDIM ScrBuf(2000)
  MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen

  OpenAcctIdx AcctIdxFileNum, NumIdxRecs
  CLOSE AcctIdxFileNum
  REDIM TempAcctIdx(1 TO NumIdxRecs) AS GLAcctIndexType

  ElSize = LEN(TempAcctIdx(1))
  FGetAH "GLACCT.IDX", TempAcctIdx(1), ElSize, NumIdxRecs

  OpenAcctFile AcctFileNum, NumAccts

  REDIM AcctList(1 TO NumIdxRecs) AS FLen       'Dim array to hold trans list
  FOR Cnt = 1 TO NumIdxRecs     'Read the file
    IF LEFT$(TempAcctIdx(Cnt).AcctNum, Lookuplen) = Lookup$ THEN
      GET AcctFileNum, TempAcctIdx(Cnt).RecNum, Acct
      IF Acct.Typ = "E" OR Acct.Typ = "R" THEN
        MatchAccts = MatchAccts + 1
        LSET AcctList(MatchAccts).V = List$
        AcctList(MatchAccts).V = TempAcctIdx(Cnt).AcctNum
        MID$(AcctList(MatchAccts).V, TotLen + 4) = Acct.Title
        MID$(AcctList(MatchAccts).V, 41) = MKI$(TempAcctIdx(Cnt).RecNum)
      END IF
    END IF
  NEXT

  CLOSE AcctFileNum

  SELECT CASE MatchAccts
  CASE 0
    PLAY "L32O3EC"              'Acct NOT found
    Ok = MsgBox("GL.QSL", "NOACCT")
    EDCheck4Acct = False
  CASE 1
    AcctRecNum = CVI(RIGHT$(AcctList(1).V, 2))
    GOSUB GetAcctInfo
    EDCheck4Acct = True
  CASE ELSE
    Row = 5
    Col = 15
    LOCATE Row, Col
    Action = 0
    MaxLen = 39
    BoxBot = 18
    ListDescTxt$ = "Account No.   Description"
    ListDesc$ = SPACE$(MaxLen + 4)
    MID$(ListDesc$, 3) = ListDescTxt$
    QPrintRC ListDesc$, Row - 1, Col, 112       'Cnf.ActivCh
    DO          'call vertmenu to display matching recs to choose from
      VertMenuT AcctList(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        EDCheck4Acct = False
        EXIT DO
      END IF
      IF Choice > 0 THEN
        AcctRecNum = CVI(RIGHT$(AcctList(Choice).V, 2))
        GOSUB GetAcctInfo
        EDCheck4Acct = True
        EXIT DO
      END IF
    LOOP
  END SELECT

  MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen

  ERASE AcctList, TempAcctIdx, ScrBuf


  EXIT FUNCTION

GetAcctInfo:
  OpenAcctFile AcctFileNum, NumAccts
  GET AcctFileNum, AcctRecNum, Acct
  CLOSE AcctFileNum
  LSET Form$(AcctFld, 0) = Acct.Num
  LSET Form$(AcctFld + 1, 0) = Acct.Title

'  LSET Form$(AcctFld + 2, 0) = MKI$(AcctRecNum)
RETURN



END FUNCTION

SUB fixunformattedbgt

   CLS
   LOCATE 1, 1, 0
   QPrintRC "Fix Unformatted Account", 1, 1, -1
   QPrintRC "Esc=Cancel, any other key to continue.", 2, 1, -1
   k$ = INPUT$(1)
   IF k$ = CHR$(27) THEN
      EXIT SUB
   END IF

   'CLS

   OpenBgtEditFile BgtEditFileNum, NumEdTrans  'Opens the Bgt Edit File
   
   FOR Cnt = 1 TO NumEdTrans
      Done! = (Cnt / NumEdTrans) * 100

      LOCATE 4, 1, 0
      PRINT USING "Processing File. ###% Complete."; Done!

      GET BgtEditFileNum, Cnt, BgtEdit
      AcctNum$ = StripAcct$(BgtEdit.AcctNum)
      BgtEdit.AcctNum = FmtAcct$(AcctNum$, FundLen%, AcctLen%, DetLen%)
      PUT BgtEditFileNum, Cnt, BgtEdit

   NEXT

   CLOSE

END SUB

'
'***************************************************************************
'Finds the next undeleted record.
'Call with NextRec value of -1 for previous record, +1 for the next record.
'If a record is not found, the function returns the value of CurrRec.
'***************************************************************************
'
FUNCTION GetNextRec (FileNum, NumRecs, CurrRec, NextRec)

   Found = 0
   Rec = CurrRec

   DO

      Rec = Rec + NextRec                'Set file pointer to next record

      IF Rec > NumRecs OR Rec <= 0 THEN  'test for beg or end of file
         Found = 0                       'if no more records then get out
         EXIT DO
      END IF

      GET FileNum, Rec, BgtEdit           'Get the record

      IF BgtEdit.Deleted = 0 THEN         'Ok if not deleted
         Found = 1
         EXIT DO                         'Get out of loop when we find one
      END IF

   LOOP

   IF Found = 0 THEN
      GetNextRec = CurrRec
   ELSE
      GetNextRec = Rec
   END IF

END FUNCTION

SUB NewBgtEntry

   SHARED TotLen                             'Max account number length
   SHARED Choice$()                          'Multiple choice array
   SHARED Cnf AS Config
   REDIM Form$(0, 0)
   REDIM Fld(0) AS FieldInfo

   REDIM Choice$(0 TO 2, 1)                  'Choices for entry type field
   Choice$(0, 0) = "5"                       'field #5
   Choice$(1, 0) = "Debit "
   Choice$(2, 0) = "Credit"
   
   LibLoadDisplayForm "GL.QSL", "NEWJE", Form$(), Fld()

   TrDate = FldNum("TRDATE", Fld())          'Assign field number to vars
   AcctNum = FldNum("ACCTNUM", Fld())
   AcctName = FldNum("ACCTNAME", Fld())
   Desc = FldNum("DESC", Fld())
   Amt = FldNum("AMT", Fld())
   EType = FldNum("ETYPE", Fld())
   Ref = FldNum("REF", Fld())
   
   DIM Frm AS FormInfo
  
   OpenBgtEditFile BgtEditFileNum, NumEdTrans
   
   CLOSE BgtEditFileNum                       'Close file while user is editing

   IF NumEdTrans > 0 THEN                    'If we have any records
      DisplayTotals                          'Display File Totals
   END IF
   
   Form$(TrDate, 0) = DATE$                  'set date field to sys date
   Action = 1                                'init editform

   DO                                        'Start form editing loop

      EditForm Form$(), Fld(), Frm, Cnf, Action

      IF BadFld > 0 THEN                     'if user attempted to save
         Frm.FldNo = BadFld                  'rec with blank flds
         BadFld = 0                          'move them to blank field
      END IF

      IF Frm.PrevFld <> Frm.FldNo THEN       'user is moving around in form
      
         IF Frm.FldNo = AcctNum THEN         'if we're back on acct number
                                             ' trim out dashes B4 editing
            Form$(AcctNum, 0) = StripAcct$(Form$(AcctNum, 0))
            Frm.FldNo = AcctNum              'Put user back on acct field
            SaveField AcctNum, Form$(), Fld(), 0
            Action = 1
         
         END IF

         SELECT CASE Frm.PrevFld
            CASE AcctNum                     'user just left acct field

               'look up the account number
               Form$(AcctNum, 0) = FmtAcct$(RTRIM$(Form$(AcctNum, 0)), FundLen%, AcctLen%, DetLen%)
               RecordNum = FindAcct(Form$(AcctNum, 0))
               IF RecordNum > 0 THEN         'Eureka!
                  Form$(AcctName, 0) = GetAcctTitle$(RecordNum)
                  PrintArray 0, 0, Form$(), Fld()
                  
               ELSE
                  PLAY "L32O3EC"             'Acct NOT found
                  Ok = MsgBox("GL.QSL", "NOACCT")
                  Form$(AcctNum, 0) = StripAcct$(Form$(AcctNum, 0))
                  LSET Form$(AcctName, 0) = ""
                  Frm.FldNo = AcctNum        'Put user back on acct field
                  SaveField AcctNum, Form$(), Fld(), 0
                  Action = 1

               END IF

            CASE ELSE
               'Other field validation stuff
      
         END SELECT

      END IF
      
      SELECT CASE Frm.KeyCode
         CASE F10Key
            BadFld = 0
            GOSUB CheckSave
            IF BadFld = 0 THEN
               GOSUB SaveTran
            END IF
            DisplayTotals

         CASE F3Key
            Form$(TrDate, 0) = ""
            Form$(AcctNum, 0) = ""
            Form$(AcctName, 0) = ""
            Form$(Amt, 0) = ""
            Form$(EType, 0) = ""
            Form$(Desc, 0) = ""
            Form$(Ref, 0) = ""

            Frm.FldNo = TrDate
            Action = 1

         CASE ELSE

      END SELECT
   
   LOOP UNTIL Frm.KeyCode = ESC

EXIT SUB


SaveTran:
   OpenBgtEditFile BgtEditFileNum, NumEdTrans
   NumEdTrans = NumEdTrans + 1

   BgtEdit.Deleted = 0
   BgtEdit.TrDate = Date2Num(Form$(TrDate, 0))
   BgtEdit.AcctRec = RecordNum
   BgtEdit.AcctNum = Form$(AcctNum, 0)
   BgtEdit.AcctName = Form$(AcctName, 0)
   BgtEdit.EType = Form$(EType, 0)
   BgtEdit.Desc = Form$(Desc, 0)
   BgtEdit.Ref = Form$(Ref, 0)
   BgtEdit.Src = "Bgt" + ConvDateStr$(DATE$)
   
   Amt# = Value(Form$(Amt, 0), ErrCode)
   
   SELECT CASE BgtEdit.EType                  'Debit or Credit
      CASE "D", ""
         BgtEdit.DrAmt = Amt#
         BgtEdit.CrAmt = 0
      CASE "C"
         BgtEdit.CrAmt = Amt#
         BgtEdit.DrAmt = 0
   END SELECT
   
   PUT BgtEditFileNum, NumEdTrans, BgtEdit

   Form$(TrDate, 0) = Num2Date(BgtEdit.TrDate)
   Form$(AcctNum, 0) = ""
   Form$(AcctName, 0) = ""
   Form$(Amt, 0) = ""
   Form$(EType, 0) = ""
   Form$(Desc, 0) = BgtEdit.Desc
   Form$(Ref, 0) = BgtEdit.Ref

   Frm.FldNo = TrDate
   
   CLOSE BgtEditFileNum

RETURN


CheckSave:
   FOR Cnt = 1 TO 6                          'Check for blank fields
      IF LEN(LTRIM$(RTRIM$(Form$(Cnt, 0)))) = 0 THEN
         PLAY "L32O3EC"
         Ok = MsgBox("GL.QSL", "NOBLANK")
         BadFld = Cnt
         EXIT FOR
      END IF
   NEXT
   Action = 1
RETURN


END SUB

SUB OpenBgtEditFile (BgtEditFileNum, NumEdTrans)

   BgtEdLen = LEN(BgtEdit)
   BgtEditFileNum = FREEFILE
   OPEN "BGTED.dat" FOR RANDOM ACCESS READ WRITE SHARED AS BgtEditFileNum LEN = BgtEdLen
   NumEdTrans = LOF(BgtEditFileNum) \ BgtEdLen

END SUB

SUB PostTrans
   
   SHARED Acct AS GLAcctRecType

   '--Paint a background
   FOR Cnt = 2 TO 24
      QPrintRC STRING$(80, CHR$(177)), Cnt, 1, 3
   NEXT
   
   '--verify that there are transactions and they are in balance.
   OpenBgtEditFile BgtEditFileNum, NumEdTrans
   
   '--summarize the file totals
   FOR Cnt = 1 TO NumEdTrans
      GET BgtEditFileNum, Cnt, BgtEdit
      IF NOT BgtEdit.Deleted THEN
         Active = Active + 1
         TotDr# = TotDr# + Round#(BgtEdit.DrAmt)
         TotCr# = TotCr# + Round#(BgtEdit.CrAmt)
      END IF
   NEXT

   CLOSE

   '--if no active transactions tell user and get out
   IF Active = 0 THEN
      Ok = MsgBox%("GL.QSL", "NOTRANS")
      EXIT SUB
   END IF

   Ok = MsgBox%("GL.QSL", "POSTGJ")       'Ask user if sure ready to post
   IF Ok = 1 THEN EXIT SUB   '1=No
   
   IF Round#(TotDr#) <> Round#(TotCr#) THEN               'Transactions out of balance
      Ok = MsgBox%("GL.QSL", "OUTOFBAL")  'ask user if ok to post
      IF Ok = 1 THEN EXIT SUB             'No = button 1
   END IF
  
   TotDr# = 0                             'init totals to zero
   TotCr# = 0
   Active = 0                             'Counter for Active Transactions
   
   OpenBgtEditFile BgtEditFileNum, NumEdTrans
   OpenAcctFile AcctFile, NumAccts

   DIM BgtTrans AS GLTransRecType
   BgtTransFile = FREEFILE
   BgtTransRecLen = LEN(BgtTrans)
   OPEN "BgtTrans.DAT" FOR RANDOM AS BgtTransFile LEN = BgtTransRecLen

   NumBgtTrans = LOF(BgtTransFile) \ BgtTransRecLen
   
   FOR Cnt = 1 TO NumEdTrans              'Assign edit file to trans format
      GET BgtEditFileNum, Cnt, BgtEdit

      IF NOT BgtEdit.Deleted THEN
         Active = Active + 1
         AcctRec = FindAcct(BgtEdit.AcctNum)
         IF AcctRec > 0 THEN
            GET AcctFile, AcctRec, Acct

            SELECT CASE Acct.Typ
               CASE "E"
                  Acct.Bgt = Round#(Acct.Bgt + BgtEdit.DrAmt - BgtEdit.CrAmt)
               CASE "R"
                  Acct.Bgt = Round#(Acct.Bgt + BgtEdit.CrAmt - BgtEdit.DrAmt)

            END SELECT

            PUT AcctFile, AcctRec, Acct

            BgtTrans.AcctRec = BgtEdit.AcctRec
            BgtTrans.AcctNum = BgtEdit.AcctNum
            BgtTrans.TrDate = BgtEdit.TrDate
            BgtTrans.Desc = BgtEdit.Desc
            BgtTrans.Ref = BgtEdit.Ref
            BgtTrans.DrAmt = BgtEdit.DrAmt
            BgtTrans.CrAmt = BgtEdit.CrAmt
            BgtTrans.Src = "BG" + ConvDateStr$(DATE$)
            BgtTrans.NextTran = 0

            NumBgtTrans = NumBgtTrans + 1

            PUT BgtTransFile, NumBgtTrans, BgtTrans
            
            '--------------------------------Start linking here
            IF Acct.FrstBTran = 0 THEN       'if first trans for this acct,
               Acct.FrstBTran = NumBgtTrans  'assign first & last pointers to
               Acct.LastBTran = NumBgtTrans  'this transaction
               PUT AcctFile, AcctRec, Acct

            ELSE                             'otherwise
                                             'in the account file..
               Prev = Acct.LastBTran         'remember the prev trans pointer,
               Acct.LastBTran = NumBgtTrans  'reset last trans to this trans
               PUT AcctFile, AcctRec, Acct

                                             'In the trans file...
               GET BgtTransFile, Prev, BgtTrans  'Get the last transaction
               BgtTrans.NextTran = NumBgtTrans     'reset pointer to this trans
               PUT BgtTransFile, Prev, BgtTrans

           END IF

         ELSE
            BadTrans = BadTrans + 1
        
         END IF
      END IF
   NEXT
   
   CLOSE
   
   IF BadTrans > 0 THEN BEEP
   KILL "BGTED.dat"

END SUB

SUB PrnBgtTran
   
   LOCATE 1, 1, 0
   
   QPrintRC "Budget Transaction List", 12, 1, 15
   QPrintRC "Esc=Cancel, any other key to continue.", 13, 1, 15

   k$ = INPUT$(1)
   IF k$ = CHR$(27) THEN
      EXIT SUB
   END IF

   CLS

   DIM BgtTrans AS GLTransRecType
   
   TransRecLen = LEN(Trans)
   BgtTransFile = FREEFILE
   OPEN "BGTTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS BgtTransFile LEN = TransRecLen
   NumTrans& = LOF(BgtTransFile) \ TransRecLen
   
   PrnFile = FREEFILE
   OPEN "Trans.prn" FOR OUTPUT AS #PrnFile

   'print field names
   ToPrint$ = SPACE$(100)
   LSET ToPrint$ = "Acct Number"
   MID$(ToPrint$, 18) = "Date"
   MID$(ToPrint$, 29) = "Description"
   MID$(ToPrint$, 46) = "Reference"
   MID$(ToPrint$, 56) = "Debit"
   MID$(ToPrint$, 69) = "Credit"
   PRINT #PrnFile, ToPrint$

   FOR Cnt& = 1 TO NumTrans&
      Done! = (Cnt& / NumTrans&) * 100

      LOCATE 12, 1, 0
      PRINT USING "Processing File. ###% Complete."; Done!

      GET BgtTransFile, Cnt&, Trans

         ToPrint$ = SPACE$(100)

         LSET ToPrint$ = Trans.AcctNum
         MID$(ToPrint$, 18) = Num2Date$(Trans.TrDate)
         MID$(ToPrint$, 29) = LEFT$(Trans.Desc, 15)
         MID$(ToPrint$, 46) = Trans.Ref
         MID$(ToPrint$, 56) = FUsing(STR$(Trans.DrAmt), "#######,.##")
         MID$(ToPrint$, 69) = FUsing(STR$(Trans.CrAmt), "#######,.##")
         MID$(ToPrint$, 80) = STR$(Cnt&)
         MID$(ToPrint$, 90) = STR$(Trans.NextTran)

         PRINT #PrnFile, ToPrint$

    NEXT

   CLOSE

   CLS
   QPrintRC "The file TRANS.PRN has been created in the Citipak Directory.", 12, 1, 15
   QPrintRC "Press any key to exit.", 13, 1, 15
   k$ = INPUT$(1)

   'SHELL "list Trans.prn"

END SUB

SUB PrnEditList

   SHARED Choice$()                          'Multiple choice array
   SHARED Cnf AS Config                      'System config & colors
   REDIM Form$(0, 0)                         'Holds data from form
   REDIM Fld(0) AS FieldInfo                 'Field editing array

   REDIM Choice$(0 TO 4, 1)                  'Choices for entry type field
   Choice$(0, 0) = "2"                       'Field Number
   Choice$(1, 0) = "Screen"
   Choice$(2, 0) = "Prn #1"
   Choice$(3, 0) = "Prn #2"

   CommaFmt$ = "########,.##"
   
   '--call sub in GLAUX module to load screen & form
   LibLoadDisplayForm "GL.QSL", "PRNBGT", Form$(), Fld()
   PrintHelp "Print General Journal Register"

   DIM Frm AS FormInfo                       'Dim the form info array
   
   Action = 1

   DO                                        'Get the inputs

      EditForm Form$(), Fld(), Frm, Cnf, Action

         SELECT CASE Frm.KeyCode             'Check the user's key presses

         CASE F10Key                         'SaveButton
            'Lookfor$ = QPTrim$(Form$(1, 0))  'Fld 1 is protected
            IF LEN(LTRIM$(RTRIM$(Form$(2, 0)))) = 0 THEN 'Fld 2 = OutputDevice
               Dev$ = "S"
               LPTNo = 1
            ELSE
               Dev$ = LEFT$(Form$(2, 0), 1)
               LPTNo = VAL(RIGHT$(RTRIM$(Form$(2, 0)), 1))
         END IF

         CASE ESC
            EXIT SUB

         END SELECT

   LOOP UNTIL Frm.KeyCode = F10Key

   LOCATE 1, 1, 0

   OpenBgtEditFile BgtEditFileNum, NumEdTrans

   PrnFileNum = FREEFILE
   OPEN "BGTREG.PRN" FOR OUTPUT AS #PrnFileNum
   GOSUB PrintBgtHeader

   '--Report Vars
   MaxLines = 55
   Title$ = " Budget Transaction Register"
   CrLF$ = CHR$(13) + CHR$(10)
   FF$ = CHR$(12)

   FOR I = 1 TO NumEdTrans
      GET BgtEditFileNum, I, BgtEdit
      'IF I = 33 THEN STOP
      IF NOT BgtEdit.Deleted THEN
         'Add test for fund code here

         '--First Line
         ToPrint$ = SPACE$(80)
         LSET ToPrint$ = Num2Date(BgtEdit.TrDate)
         MID$(ToPrint$, 13) = BgtEdit.AcctNum
         MID$(ToPrint$, 27) = BgtEdit.AcctName
         MID$(ToPrint$, 53) = FUsing$(STR$(BgtEdit.DrAmt), CommaFmt$)
         MID$(ToPrint$, 68) = FUsing$(STR$(BgtEdit.CrAmt), CommaFmt$)
         PRINT #PrnFileNum, ToPrint$
         LineCnt = LineCnt + 1
         IF LineCnt > MaxLines THEN
            PRINT #PrnFileNum, FF$
            GOSUB PrintBgtHeader
         END IF

         '--2nd Line
         ToPrint$ = SPACE$(80)
         MID$(ToPrint$, 13) = BgtEdit.Desc
         MID$(ToPrint$, 35) = BgtEdit.Ref
         PRINT #PrnFileNum, ToPrint$
         LineCnt = LineCnt + 1
         IF LineCnt > MaxLines THEN
            PRINT #PrnFileNum, FF$
            GOSUB PrintBgtHeader
         END IF

         '--3rd line is blank
         PRINT #PrnFileNum,
         LineCnt = LineCnt + 1
         IF LineCnt > MaxLines THEN
            PRINT #PrnFileNum, FF$
            GOSUB PrintBgtHeader
         END IF
         
         TotDr# = TotDr# + BgtEdit.DrAmt
         TotCr# = TotCr# + BgtEdit.CrAmt

      END IF

   NEXT

   'LineCnt = LineCnt + 1
   'IF LineCnt > 45 THEN
   '   PRINT #PrnFileNum, FF$
   '   GOSUB PrintBgtHeader
   'END IF
   
   PRINT #PrnFileNum, STRING$(80, "-")

   ToPrint$ = SPACE$(80)
   LSET ToPrint$ = "File Totals"
   MID$(ToPrint$, 53) = FUsing$(STR$(TotDr#), CommaFmt$)
   MID$(ToPrint$, 68) = FUsing$(STR$(TotCr#), CommaFmt$)
   PRINT #PrnFileNum, ToPrint$
   
   PRINT #PrnFileNum, FF$
   CLOSE
   
   SELECT CASE Dev$
      CASE "S"
         EntryPoint = 2
      CASE "P"
         EntryPoint = 5
   END SELECT

   PrintRptFile Title1$, "BGTREG.PRN", LPTNo, RetCode%, EntryPoint
   KILL "BgtREG.PRN"

EXIT SUB


PrintBgtHeader:
  PRINT #PrnFileNum, "Budget Register"
  PRINT #PrnFileNum,
  PRINT #PrnFileNum, "Date        Description           Reference"
  PRINT #PrnFileNum, "            G/L Account                                     Debit         Credit"
  PRINT #PrnFileNum, "--------------------------------------------------------------------------------"
  LineCnt = 5
RETURN

END SUB

SUB unPost

  DIM BgtTrans AS GLTransRecType
  BgtTransFile = FREEFILE
  BgtTransRecLen = LEN(BgtTrans)
  OPEN "BgtTrans.DAT" FOR RANDOM AS BgtTransFile LEN = BgtTransRecLen
  NumBgtTrans = LOF(BgtTransFile) \ BgtTransRecLen
  
  NewRec = 0
  TDate = Date2Num("07/01/1996")

  FOR Cnt = 1 TO NumBgtTrans
    GET BgtTransFile, Cnt, BgtTrans
    IF BgtTrans.TrDate >= TDate THEN
      GOSUB Post2New
    END IF
  NEXT

  CLOSE

EXIT SUB

Post2New:
  NewRec = NewRec + 1

RETURN

END SUB

