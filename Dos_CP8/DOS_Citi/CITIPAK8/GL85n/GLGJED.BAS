DEFINT A-Z

'DECLARE SUB CalcFields (StartOfForm%, FldNo%, Form$(), Fld() AS ANY)
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
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (t$, R%, c%, Clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortStr2 (BYVAL Address%, Size%, Direction%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
DECLARE FUNCTION Value# (E$, ErrCode%)
DECLARE FUNCTION FldNum% (FldName$, Fld() AS ANY)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
'DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
'DECLARE SUB Prn2Scrn2 (TitleText$(), FileName$, TopRow%, BotRow%, Clr%, PrnFlag%)

DECLARE SUB MnuTitle (Title$, Row, Col, Wdth, TxtColor)
DECLARE SUB ClrBG ()

DECLARE SUB Post2GL (FileName$, BadTrans%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)

'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'GL.BI'                        'GL File Types
'$INCLUDE: 'SCRAUX.BI'

CONST False = 0, True = NOT False

DECLARE SUB EdGJEntry ()
DECLARE SUB OpenGJEditFile (GJEditFileNum%, NumEdTrans%)
DECLARE SUB NewGJEntry ()
DECLARE FUNCTION GetNextRec% (FileNum%, NumRecs%, CurrRec%, NextRec%)
DECLARE SUB PrnEditList ()
DECLARE SUB DisplayTotals ()
DECLARE SUB PostTrans ()
DECLARE SUB GetFundList (FundList$(), NumFunds%)
DECLARE FUNCTION EDCheck4Acct% (Form$(), AcctFld%)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB GetPostDates (LPDate, HPEndDate)

DIM SHARED TotLen, FundLen, AcctLen, DetLen
DIM SHARED AcctIdx AS GLAcctIndexType
DIM SHARED Acct AS GLAcctRecType
DIM SHARED GJEdit AS TrEditRecType
DIM SHARED Trans AS GLTransRecType

TYPE FLen
  v AS STRING * 42
END TYPE

'REDIM SHARED Choice$(0, 0)                'Used to hold multiple choices

'$INCLUDE: 'SetCnf.bi'                    'Set monitor and color information
'$INCLUDE: 'init.bi'                      'Get GL system info
''$INCLUDE: 'keycode.bi'                   'Define keycodes

TextCursor -2, -2                         'Define the Mouse cursor
ShowCursor                                'Show it

'***Master Menu

   DIM mChoice$(1 TO 4)

   mChoice$(1) = "Enter/Edit General Journal Items"
   mChoice$(2) = "Print Journal Register"
   mChoice$(3) = "Post Journal Entries"
   mChoice$(4) = "Quit"

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

   'Center Menu within Screen
   Row = ((25 - (UBOUND(mChoice$))) \ 2) - 1
   Col = ((80 - MaxLen) \ 2) - 2

   ShowCursor

   DO
      LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, Attribute, ErrorCode
      'TitleBox 3, Col, MaxLen + 3, "General Journal Menu ", Cnf
      'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

      PrintTitle User$
      Title$ = "General Journal Menu  "
      MnuTitle Title$, Row, Col, MaxLen, 15
      PrintHelp "General Journal Processing"

      LOCATE Row, Col, 0
      VertMenu mChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

      IF Ky$ = CHR$(27) THEN EXIT DO

      SELECT CASE Choice
         CASE 1
            ClrBG
            EdGJEntry
         CASE 2
            ClrBG
            PrnEditList
         CASE 3
            ClrBG
            PostTrans
         CASE 4
            CLS : END
         CASE ELSE
      END SELECT

   LOOP

ClrBG
HideCursor
DispLoad
RUN "glmenu"

SUB DisplayTotals

   OpenGJEditFile GJEditFileNum, NumEdTrans

   TotDr# = 0: TotCr# = 0                    'reset totals to zero

   FOR Cnt = 1 TO NumEdTrans                 're-summarize totals
      GET GJEditFileNum, Cnt, GJEdit
      IF NOT GJEdit.Deleted THEN
         TotDr# = TotDr# + GJEdit.DrAmt
         TotCr# = TotCr# + GJEdit.CrAmt
      END IF
   NEXT

   CLOSE GJEditFileNum

   TotDr$ = FUsing$(STR$(TotDr#), "##########,.##")
   TotCr$ = FUsing$(STR$(TotCr#), "##########,.##")

   QPrintRC TotDr$, 18, 23, 15               'print the file totals
   QPrintRC TotCr$, 18, 53, 15               ' on the screen


END SUB

FUNCTION EDCheck4Acct% (Form$(), AcctFld%)

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
      MatchAccts = MatchAccts + 1
      LSET AcctList(MatchAccts).v = List$
      AcctList(MatchAccts).v = TempAcctIdx(Cnt).AcctNum
      MID$(AcctList(MatchAccts).v, TotLen + 4) = Acct.Title
      MID$(AcctList(MatchAccts).v, 41) = MKI$(TempAcctIdx(Cnt).RecNum)
    END IF
  NEXT

  CLOSE AcctFileNum

  SELECT CASE MatchAccts
  CASE 0
    PLAY "L32O3EC"              'Acct NOT found
    Ok = MsgBox("GL.QSL", "NOACCT")
    EDCheck4Acct = False
  CASE 1
    AcctRecNum = CVI(RIGHT$(AcctList(1).v, 2))
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
        AcctRecNum = CVI(RIGHT$(AcctList(Choice).v, 2))
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

SUB EdGJEntry

   GetPostDates LPDate, HPDate

   SHARED Choice$()                          'Multiple choice array
   REDIM Form$(0, 0)                         'Holds data from form
   REDIM Fld(1) AS FieldInfo                 'Field editing array

   REDIM Choice$(0 TO 2, 1)                  'Choices for entry type field
   Choice$(0, 0) = "5"                       'Field Number
   Choice$(1, 0) = "Debit "
   Choice$(2, 0) = "Credit"

   EdStatus$ = "Editing Journal Items"
   NewStatus$ = "New Journal Items"

   OpenGJEditFile GJEditFileNum, NumEdTrans  'Opens the GJ Edit File
   
   '--determine edit mode on entry
   IF NumEdTrans > 0 THEN
      
      '--Call GetNextRec to get the first record not deleted
      CurrRec = 0: NextRec = 1
      RecNum = GetNextRec(GJEditFileNum, NumEdTrans, CurrRec, NextRec)
   
      IF RecNum = 0 THEN
         CLOSE                               'If All records are deleted..
         KILL "GLGJED.DAT"                     '..init file to zero
         EMode = 1                           '..Set mode to New
      ELSE                                   'Otherwise...
         EMode = 0                           '..set Mode to Edit
      END IF
      
   ELSE                                      'No trans.. Set Add Mode
      EMode = 1
   
   END IF
   
   'call sub in GLAUX module to load screen & form
   LibLoadDisplayForm "GL.QSL", "EDJE", Form$(), Fld()
   
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
      FirstTime = -1
      DisplayTotals
      Action = 1                             'Init form edit
   END IF
   
   GOSUB ShowMode
   GOSUB ShowRecNum

   DO                                        'Edit the form

      EditForm Form$(), Fld(), Frm, Cnf, Action
      IF FirstTime THEN
        FirstTime = 0
        SaveField AcctNum, Form$(), Fld(), 0
      END IF
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
              GOSUB CheckForValidAcctNum:

            CASE ELSE
               'Other field validation stuff goes here

         END SELECT

      END IF
      
      SELECT CASE Frm.KeyCode                'Respond to the key presses
         CASE F10Key  'Save
            BadFld = 0
            GOSUB CheckFlds
            'GOSUB CheckForValidAcctNum
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
                 OpenGJEditFile GJEditFileNum, NumEdTrans
                 GJEdit.Deleted = -1
                 PUT GJEditFileNum, RecNum, GJEdit
                 GJEdit.Deleted = 0
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
            OpenGJEditFile GJEditFileNum, NumEdTrans
            CurrRec = RecNum: NextRec = -1
            RecNum = GetNextRec(GJEditFileNum, NumEdTrans, CurrRec, NextRec)
            CLOSE
            GOSUB ShowRecNum
            IF RecNum > 0 THEN
              GOSUB Rec2Form
            END IF

         CASE CtrlR  'Ctrl+Right=Next Rec
            OpenGJEditFile GJEditFileNum, NumEdTrans
            CurrRec = RecNum: NextRec = 1
            RecNum = GetNextRec(GJEditFileNum, NumEdTrans, CurrRec, NextRec)
            CLOSE
            GOSUB ShowRecNum
            IF RecNum > 0 THEN
              GOSUB Rec2Form
            END IF

         CASE F5Key  'F5=Find
            IF EMode = 0 THEN
               OpenGJEditFile GJEditFileNum, NumEdTrans
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

   LOOP UNTIL Frm.KeyCode = Esc               'Loop until user presses Esc

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

   NEXT

   TDate = Date2Num(Form$(TrDate, 0))
   IF TDate < LPDate OR TDate > HPDate THEN
      Ok = MsgBox("GL.QSL", "BADDATE")
      BadFld = TrDate
      Frm.FldNo = BadFld
      'EXIT FOR
   END IF
   Action = 1

RETURN


ClearForm:
   FOR FieldNum = 1 TO 7
      Form$(FieldNum, 0) = ""
   NEXT
   Form$(TrDate, 0) = Num2Date(GJEdit.TrDate)
   Form$(Desc, 0) = GJEdit.Desc
   Form$(Ref, 0) = GJEdit.Ref
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
   OpenGJEditFile GJEditFileNum, NumEdTrans

   IF EMode = 1 THEN
      RecNum = NumEdTrans + 1     'adding new or editing existing
      'GOSUB ClearForm
   END IF

   GJEdit.Deleted = 0
   GJEdit.AcctNum = Form$(AcctNum, 0)
   IF MID$(GJEdit.AcctNum, FundLen + 1, 1) <> "-" THEN
     GJEdit.AcctNum = FmtAcct$(GJEdit.AcctNum, FundLen, AcctLen, DetLen)
   END IF
   
   GJEdit.AcctName = Form$(AcctName, 0)

   GJEdit.TrDate = Date2Num(Form$(TrDate, 0))'Store date as integer
   GJEdit.Desc = Form$(Desc, 0)
   GJEdit.EType = Form$(EType, 0)
   GJEdit.Ref = Form$(Ref, 0)
   GJEdit.Src = "GJ" + ConvDateStr$(DATE$)

   Amt# = Value(Form$(Amt, 0), ErrCode)      'Convert Amt$ to Double
   IF ErrCode THEN
      CLS
      CLOSE
      PRINT "An error occured in the value function!"
      Oh.Shit! = True
      EXIT SUB
   END IF
                                             'Only one amount field in form!
   SELECT CASE GJEdit.EType                  'Is amount debit or credit
      CASE "D"                               'Debit
         GJEdit.DrAmt = Amt#                 'assign amt to debit amt
         GJEdit.CrAmt = 0
      CASE "C"                               'Credit
         GJEdit.CrAmt = Amt#                 'assign amt to credit amt
         GJEdit.DrAmt = 0
   END SELECT
   
   PUT GJEditFileNum, RecNum, GJEdit         'write it
   CLOSE GJEditFileNum                       'close the file

RETURN


Rec2Form:                                    'Set form fields to rec values
   OpenGJEditFile GJEditFileNum, NumEdTrans
   GET GJEditFileNum, RecNum, GJEdit

   Form$(TrDate, 0) = Num2Date(GJEdit.TrDate)'Date stored as integer
   IF MID$(GJEdit.AcctNum, FundLen + 1, 1) = "-" THEN
     Form$(AcctNum, 0) = GJEdit.AcctNum
   ELSE
     Form$(AcctNum, 0) = FmtAcct$(GJEdit.AcctNum, FundLen, AcctLen, DetLen)
   END IF
   Form$(AcctName, 0) = GJEdit.AcctName
   Form$(Desc, 0) = GJEdit.Desc
   Form$(Ref, 0) = GJEdit.Ref

   

   SELECT CASE GJEdit.EType                  'set entry type fld to Dr or Cr
      CASE "D"
         Form$(EType, 0) = "Debit "
         Form$(Amt, 0) = STR$(GJEdit.DrAmt)
      CASE "C"
         Form$(EType, 0) = "Credit"
         Form$(Amt, 0) = STR$(GJEdit.CrAmt)
      CASE ELSE
   END SELECT
   CLOSE
   Action = 1                                'force form edit to re-init
   Frm.FldNo = 1
   FirstTime = -1
RETURN


ShowRecNum:
   Record$ = "Record:     "
   MID$(Record$, 8) = STR$(RecNum)
   QPrintRC Record$, 25, 60, -1

RETURN


Find:                                        'Find Transaction
'add input box to get sort yes/no here

   SortEm = MsgBox("GL", "SORTEM")

   OpenGJEditFile GJEditFileNum, NumEdTrans
   
   REDIM TransList$(1 TO NumEdTrans)         'Dim array to hold trans list
   Active = 0                                'Counter for Active Transactions

   FOR Cnt = 1 TO NumEdTrans                 'Read the file
      GET GJEditFileNum, Cnt, GJEdit
      IF NOT GJEdit.Deleted THEN             'if active add to list
         Active = Active + 1

         List$ = SPACE$(75)

         LSET List$ = ConvDateStr1(Num2Date(GJEdit.TrDate))
         MID$(List$, 10) = GJEdit.AcctNum
         MID$(List$, 28) = GJEdit.Desc
         SELECT CASE GJEdit.EType
            CASE "D"
               Amt# = GJEdit.DrAmt
               EType$ = "Dr"
            CASE "C"
               Amt# = GJEdit.CrAmt
               EType$ = "Cr"
         END SELECT
         MID$(List$, 48) = FUsing$(STR$(Amt#), "########,.##")
         MID$(List$, 63) = EType$
         MID$(List$, 70) = STR$(Cnt)         'Transaction record number

         TransList$(Active) = List$
         
      END IF
      
   NEXT
   CLOSE
   IF Active = 0 THEN
      KILL "GLGJED.DAT"                     '..init file to zero
      EXIT SUB
   END IF

   REDIM PRESERVE TransList$(1 TO Active)    'Redim list to only active trans

   IF Active > 1 AND SortEm = 1 THEN
     SortStr2 VARPTR(TransList$(1)), Active, 0
   END IF

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
  SELECT CASE EDCheck4Acct%(Form$(), AcctFld)
  CASE True
    SaveField AcctFld, Form$(), Fld(), BadFld   'Write the flag to the form$
    SaveField AcctFld + 1, Form$(), Fld(), BadFld               'Write the flag to the form$
    Action = 2
  CASE ELSE
    Frm.FldNo = AcctFld      'Put user back on acct field
    LSET Form$(Frm.FldNo, 0) = ""
    LSET Form$(Frm.FldNo + 1, 0) = ""
    Action = 2
  END SELECT

RETURN


 
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

      GET FileNum, Rec, GJEdit           'Get the record

      IF GJEdit.Deleted = 0 THEN         'Ok if not deleted
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

SUB NewGJEntry

   SHARED TotLen                             'Max account number length
   SHARED Choice$()                          'Multiple choice array
   'SHARED Cnf AS Config
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
  
   OpenGJEditFile GJEditFileNum, NumEdTrans
   
   CLOSE GJEditFileNum                       'Close file while user is editing

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
   
   LOOP UNTIL Frm.KeyCode = Esc

EXIT SUB


SaveTran:
   OpenGJEditFile GJEditFileNum, NumEdTrans
   NumEdTrans = NumEdTrans + 1

   GJEdit.Deleted = 0
   GJEdit.TrDate = Date2Num(Form$(TrDate, 0))
   GJEdit.AcctRec = RecordNum
   GJEdit.AcctNum = Form$(AcctNum, 0)
   GJEdit.AcctName = Form$(AcctName, 0)
   GJEdit.EType = Form$(EType, 0)
   GJEdit.Desc = Form$(Desc, 0)
   GJEdit.Ref = Form$(Ref, 0)
   GJEdit.Src = "GJ" + ConvDateStr$(DATE$)
   
   Amt# = Value(Form$(Amt, 0), ErrCode)
   
   SELECT CASE GJEdit.EType                  'Debit or Credit
      CASE "D", ""
         GJEdit.DrAmt = Amt#
         GJEdit.CrAmt = 0
      CASE "C"
         GJEdit.CrAmt = Amt#
         GJEdit.DrAmt = 0
   END SELECT
   
   PUT GJEditFileNum, NumEdTrans, GJEdit

   Form$(TrDate, 0) = Num2Date(GJEdit.TrDate)
   Form$(AcctNum, 0) = ""
   Form$(AcctName, 0) = ""
   Form$(Amt, 0) = ""
   Form$(EType, 0) = ""
   Form$(Desc, 0) = GJEdit.Desc
   Form$(Ref, 0) = GJEdit.Ref

   Frm.FldNo = TrDate
   
   CLOSE GJEditFileNum

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

SUB OpenGJEditFile (GJEditFileNum, NumEdTrans)

   GJEdLen = LEN(GJEdit)
   GJEditFileNum = FREEFILE
   OPEN "GLGJED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJEditFileNum LEN = GJEdLen
   NumEdTrans = LOF(GJEditFileNum) \ GJEdLen

END SUB

SUB PostTrans
   
   '--Paint a background
   FOR Cnt = 2 TO 24
      QPrintRC STRING$(80, CHR$(177)), Cnt, 1, 3
   NEXT
   
   '--verify that there are transactions and they are in balance.
   OpenGJEditFile GJEditFileNum, NumEdTrans
   
   '--summarize the file totals
   FOR Cnt = 1 TO NumEdTrans
      GET GJEditFileNum, Cnt, GJEdit
      IF NOT GJEdit.Deleted THEN
         Active = Active + 1
         TotDr# = TotDr# + Round#(GJEdit.DrAmt)
         TotCr# = TotCr# + Round#(GJEdit.CrAmt)
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
   
   TotDr# = 0                             'init totals to zero
   TotCr# = 0
   Active = 0                             'Counter for Active Transactions

   IF TotDr# <> TotCr# THEN               'Transactions out of balance
      Ok = MsgBox%("GL.QSL", "OUTOFBAL")  'ask user if ok to post
      IF Ok = 1 THEN EXIT SUB             'No = button 1
   END IF

   
   Active = 0                             'Reset Active counter for posting
   OpenGJEditFile GJEditFileNum, NumEdTrans

   DIM Tr2Post AS GLTransRecType
   OPEN "GJ2POST.DAT" FOR RANDOM AS #2 LEN = LEN(Tr2Post)

   FOR Cnt = 1 TO NumEdTrans              'Assign edit file to trans format
      GET GJEditFileNum, Cnt, GJEdit
      IF NOT GJEdit.Deleted THEN
         Active = Active + 1
         Tr2Post.AcctRec = GJEdit.AcctRec
         Tr2Post.AcctNum = GJEdit.AcctNum
         Tr2Post.TrDate = GJEdit.TrDate
         Tr2Post.Desc = GJEdit.Desc
         Tr2Post.Ref = GJEdit.Ref
         Tr2Post.DrAmt = GJEdit.DrAmt
         Tr2Post.CrAmt = GJEdit.CrAmt
         Tr2Post.Src = "GJ" + ConvDateStr$(DATE$)
         PUT #2, Active, Tr2Post

      END IF
   NEXT
   
   CLOSE
   
   Post2GL "GJ2POST.DAT", BadTrans%       'common post & link sub in GLAUX.BAS
   
   KILL "GLGJED.DAT"                    'kill the temp files
   KILL "GJ2POST.DAT"

   IF BadTrans <> 0 THEN                  'posting problem
   CLS
      PRINT "Error: One or more transactions were not posted."
      PRINT "Press any key to view posting log"
      k$ = INPUT$(1)
      PrintRptFile "Posting Log", "GLUTIL.LOG", LPTNo, RetCode%, 2
   END IF

END SUB

SUB PrnEditList

   SHARED User$
   SHARED Choice$()                          'Multiple choice array
   REDIM Form$(0, 0)                         'Holds data from form
   REDIM Fld(0) AS FieldInfo                 'Field editing array
   DIM Frm AS FormInfo                       'Dim the form info array

   REDIM Choice$(0 TO 3, 1)                  'Choices for entry type field
   Choice$(0, 0) = "2"                       'Form Field Number for picklist
   Choice$(1, 0) = "Screen"
   Choice$(2, 0) = "Prn #1"
   Choice$(3, 0) = "Prn #2"
   
   LibLoadDisplayForm "GL.QSL", "PRNGJ", Form$(), Fld()
   PrintHelp "Print General Journal Register"

   Action = 1

   DO

      EditForm Form$(), Fld(), Frm, Cnf, Action

         SELECT CASE Frm.KeyCode

         CASE F10Key  '--SaveButton
            '--Fld 2 = OutputDevice
            IF LEN(LTRIM$(RTRIM$(Form$(2, 0)))) = 0 THEN
               Dev$ = "S"
               LPTNo = 1
            ELSE
               Dev$ = LEFT$(Form$(2, 0), 1)
               LPTNo = VAL(RIGHT$(RTRIM$(Form$(2, 0)), 1))
         END IF

         CASE Esc
            EXIT SUB

         END SELECT

   LOOP UNTIL Frm.KeyCode = F10Key

   LOCATE 1, 1, 0
  
   '--Get a list of active funds
   REDIM FundList$(1)
   GetFundList FundList$(), NumFunds
   REDIM FundDr#(1 TO NumFunds)
   REDIM FundCr#(1 TO NumFunds)
   
   OpenGJEditFile GJEditFileNum, NumEdTrans

   PrnFileNum = FREEFILE
   PrnFileName$ = "GJREG.PRN"
   OPEN PrnFileName$ FOR OUTPUT AS #PrnFileNum

   '--Report Variables
   MaxLines = 55
   RptTitle$ = " General Journal Register"
   CRLF$ = CHR$(13) + CHR$(10)
   FF$ = CHR$(12)
   GOSUB PrintGJHeader
   CommaFmt$ = "#########,.##"
   
   '--Start of printing loop
   FOR I = 1 TO NumEdTrans
      GET GJEditFileNum, I, GJEdit

      IF NOT GJEdit.Deleted THEN
         IF UCASE$(LEFT$(GJEdit.AcctName, 9)) = "UNDEFINED" THEN
          GJEdit.AcctNum = ""
          PUT GJEditFileNum, I, GJEdit
         END IF
         '--First Line
         ToPrint$ = SPACE$(80)
         LSET ToPrint$ = Num2Date(GJEdit.TrDate)
         MID$(ToPrint$, 13) = GJEdit.Desc
         MID$(ToPrint$, 35) = GJEdit.Ref
         PRINT #PrnFileNum, ToPrint$
         LineCnt = LineCnt + 1
         IF LineCnt > MaxLines THEN
           PRINT #PrnFileNum, FF$
           GOSUB PrintGJHeader
         END IF

         '--2nd Line
         ToPrint$ = SPACE$(80)
         IF UCASE$(LEFT$(GJEdit.AcctName, 9)) = "UNDEFINED" THEN
         MID$(ToPrint$, 4) = "???"
         END IF
         MID$(ToPrint$, 13) = GJEdit.AcctNum
         MID$(ToPrint$, 27) = GJEdit.AcctName
         MID$(ToPrint$, 53) = FUsing$(STR$(GJEdit.DrAmt), CommaFmt$)
         MID$(ToPrint$, 68) = FUsing$(STR$(GJEdit.CrAmt), CommaFmt$)
         PRINT #PrnFileNum, ToPrint$
         LineCnt = LineCnt + 1
         IF LineCnt > MaxLines THEN
           PRINT #PrnFileNum, FF$
           GOSUB PrintGJHeader
         END IF

         '--3rd line (Blank)
         PRINT #PrnFileNum,
         LineCnt = LineCnt + 1
         IF LineCnt > MaxLines THEN
           PRINT #PrnFileNum, FF$
           GOSUB PrintGJHeader
         END IF

         '--Sum total debits and credits
         TotDr# = TotDr# + GJEdit.DrAmt
         TotCr# = TotCr# + GJEdit.CrAmt

         '--Sum into proper fund
         Found = False
         FOR Fund = 1 TO NumFunds
           FundNum$ = LEFT$(GJEdit.AcctNum, FundLen)
           IF FundNum$ = FundList$(Fund) THEN
             Found = True
             FundDr#(Fund) = FundDr#(Fund) + Round#(GJEdit.DrAmt)
             FundCr#(Fund) = FundCr#(Fund) + Round#(GJEdit.CrAmt)
             EXIT FOR
           END IF
         NEXT


      END IF
   NEXT

   PRINT #PrnFileNum, STRING$(80, "-")
   LineCnt = LineCnt + 1
   IF LineCnt > MaxLines THEN
     PRINT #PrnFileNum, FF$
     GOSUB PrintGJHeader
   END IF

   ToPrint$ = SPACE$(80)
   LSET ToPrint$ = "File Totals"
   MID$(ToPrint$, 53) = FUsing$(STR$(TotDr#), CommaFmt$)
   MID$(ToPrint$, 68) = FUsing$(STR$(TotCr#), CommaFmt$)
   PRINT #PrnFileNum, ToPrint$
   LineCnt = LineCnt + 1
   IF LineCnt > MaxLines THEN
     PRINT #PrnFileNum, FF$
     GOSUB PrintGJHeader
   END IF

   '--Print Summary by fund
   TranCashTot# = 0
   FundOutofBal = False
   FOR Fund = 1 TO NumFunds
      IF FundDr#(Fund) <> 0 OR FundCr#(Fund) <> 0 THEN
         IF Round(FundDr#(Fund)) <> Round(FundCr#(Fund)) THEN FundOutofBal = True
         ToPrint$ = SPACE$(80)
         MID$(ToPrint$, 4) = "Fund# " + FundList$(Fund)
         MID$(ToPrint$, 53) = FUsing$(STR$(FundDr#(Fund)), CommaFmt$)
         MID$(ToPrint$, 68) = FUsing$(STR$(FundCr#(Fund)), CommaFmt$)
         PRINT #PrnFileNum, ToPrint$
         LineCnt = LineCnt + 1
         IF LineCnt > MaxLines THEN
           PRINT #PrnFileNum, FF$
           GOSUB PrintGJHeader
         END IF
      END IF

   NEXT

   IF FundOutofBal THEN
      '--skip a line
      ToPrint$ = SPACE$(80)
      PRINT #PrnFileNum, ToPrint$
      LineCnt = LineCnt + 1
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintGJHeader
      END IF

      '--Tell user they're screwing up
      ToPrint$ = SPACE$(80)
      MID$(ToPrint$, 4) = "Entries are not in balance!"
      PRINT #PrnFileNum, ToPrint$
      LineCnt = LineCnt + 1
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintGJHeader
      END IF

   END IF

   PRINT #PrnFileNum, FF$

   CLOSE

   SELECT CASE Dev$
      CASE "S"
          EntryPoint = 2
      CASE "P"
          EntryPoint = 5
   END SELECT

  PrintRptFile RptTitle$, PrnFileName$, LPTNo, RetCode%, EntryPoint
  KILL PrnFileName$

EXIT SUB


PrintGJHeader:
  PRINT #PrnFileNum, User$
  PRINT #PrnFileNum, "General Journal Register"
  PRINT #PrnFileNum,
  PRINT #PrnFileNum, "Date        Description           Reference"
  PRINT #PrnFileNum, "            G/L Account                                     Debit         Credit"
  PRINT #PrnFileNum, "--------------------------------------------------------------------------------"
  LineCnt = 5
RETURN


END SUB

