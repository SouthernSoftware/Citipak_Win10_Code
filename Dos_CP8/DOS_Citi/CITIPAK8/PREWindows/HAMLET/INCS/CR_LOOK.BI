DEFINT A-Z

'$INCLUDE: 'KeyCode.BI'                  'Form Page info
  
'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'pageinfo.BI'                  'Form Page info
'$INCLUDE: 'PRO7.BI'                      'QuickPak Pro Routines
'$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'RPTPRN.BI'                    'Report printer routines
'$INCLUDE: 'GL.BI'                        'GL File Types
'$INCLUDE: 'CASHJRNL.BI'
  
DECLARE SUB EdEntry (JrnlCode%, AcctFlds())
DECLARE SUB OpenCJEditFile (CJEditFile%, NumEdTrans%)
DECLARE SUB NewCJEntry ()
DECLARE FUNCTION GetNextRec% (FileNum%, NumRecs%, CurrRec%, NextRec%)
DECLARE SUB PrnEditList ()
DECLARE SUB DisplayTotals ()
DECLARE SUB PostTrans (CJType)
DECLARE SUB Check4BlankFlds (Form$(), BadFld%)
DECLARE FUNCTION ListRecs% ()
DECLARE SUB ResetForm (Form$(), NumFlds%)
DECLARE SUB ShowEdMode (EdMode%)
DECLARE SUB Rec2Form (Form$(), RecordNumber%, Fld() AS ANY)
DECLARE SUB SaveTrans (Form$(), RecordNumber%, EdMode%)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB GetFundList (FundList$(), NumFunds%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)

CONST False = 0, True = NOT False
  
DIM SHARED TotLen, FundLen, AcctLen, DetLen, CashAcct$
DIM SHARED AcctIdx AS GLAcctIndexType
DIM SHARED acct AS GLAcctRecType
DIM SHARED CJEdit AS CJEditRecType
'DIM SHARED Trans AS GLTransRecType

REDIM SHARED Choice$(0, 0)    'Used to hold multiple choices
  
'$INCLUDE: 'SetCnf.bi'                    'Set monitor and color information
'$INCLUDE: 'init.bi'                      'Get GL system info


'Get acct field numbers

REDIM AcctFlds(1 TO 36) AS INTEGER
FGetAH "ACCFLD.DAT" AcctFlds(1), 2, 36

'FOR Cnt = 1 TO 36
'  READ X
'  AcctFlds(Cnt) = X
'NEXT

'Account Field Numbers
'DATA 8,12,16,20,24,28,32,36,40,44,48,52
'DATA 57,61,65,69,73,77,81,85,89,93,97,101
'DATA 106,110,114,118,122,126,130,134,138,142,146,150




TextCursor -2, -2             'Define the Mouse cursor
ShowCursor    'Show it
  
'***Master Menu
  
DIM MChoice$(1 TO 3)
  
  'mChoice$(1) = "Enter New Journal Entries"
  MChoice$(1) = "Cash Reciepts Journal"
  MChoice$(2) = "Print Cash Reciepts Journal"
  MChoice$(3) = "Post Entries"
'  MChoice$(4) = "Exit"
  
  MaxWidth = 0  'Set menu width to zero
  BoxBot = 18   'limit the box to line 20
  Action = 0    '0 means stay in the menu
  Choice = 1    'Pre-load choice to highlight
  
  FOR Cnt = 1 TO UBOUND(MChoice$)               'Find max menu width
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxWidth THEN
      MaxWidth = TLen
    END IF
  NEXT
  
  'Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxWidth) \ 2) - 2
  
  ShowCursor

  DO
    LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, -1, ErrorCode
    
    PrintTitle User$
    PrintHelp "Cash Reciepts Journal"
    
    LOCATE Row, Col, 0
    VertMenu MChoice$(), Choice, MaxWidth, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO
    
    SELECT CASE Choice
      CASE 1
        EdEntry 1, AcctFlds()
      CASE 2
        PrnEditList
      CASE 3
        PostTrans 1
 '     CASE 4
 '       CLS : END
 '     CASE ELSE
    END SELECT
    
  LOOP
  
  HideCursor
  DispLoad
 RUN "glmenu"

END

SUB Check4BlankFlds (Form$(), BadFld)

IF Date2Num(Form$(2, 0)) < -32000 THEN
  BadFld = 2
  EXIT SUB
END IF

FOR FieldNum = 3 TO 5

  IF LEN(QPTrim$(Form$(FieldNum, 0))) = 0 THEN
    BadFld = FieldNum
    EXIT FOR
  END IF

NEXT

END SUB

SUB EdEntry (CJType, AcctFlds())
  
  SHARED Cnf AS Config
  REDIM Frm(1)    AS FormInfo

  CommaFormat$ = "#########,.##"
  NumOfScrns = 3
  
  REDIM FormName$(1 TO NumOfScrns)
  
  SELECT CASE CJType
  CASE 1
    FOR Scr = 1 TO NumOfScrns
      FormName$(Scr) = "CR" + LTRIM$(STR$(Scr))
    NEXT
    Header = 6
    TotalAmtFld = 5
  CASE 2
    FOR Scr = 1 TO NumOfScrns
      FormName$(Scr) = "CD" + LTRIM$(STR$(Scr))
    NEXT
    Header = 6
    TotalAmtFld = 5
  END SELECT

  '-- Get total number of fields
  NumFlds = -1
  FOR Scr = 1 TO NumOfScrns
    NumFlds = NumFlds + LibNumberOfFields("GL", FormName$(Scr)) + 1
  NEXT
  
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  REDIM PageInfo(1 TO NumOfScrns) AS PageInfoType
  
  'for each page, get first and last fields
  StartEl = 0   'Load first form at array start
  FOR Scr = 1 TO NumOfScrns     'Get field defs
    LibGetFldDef "GL", FormName$(Scr), StartEl, Fld(), Form$(), ErrCode
    PageInfo(Scr).PageNo = Scr
    PageInfo(Scr).FirstFld = StartEl + 1
    StartEl = StartEl + Fld(StartEl).Fields + 1 'Inc StartEl to next
    PageInfo(Scr).LastFld = StartEl - 1
  NEXT

  FOR F = 1 TO NumFlds          'For all fields on all pages
    LSET Form$(F, 0) = ""       'Clear each field
  NEXT

'  FOR F = 1 TO NumFlds
'      IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN
'         Cnt = Cnt + 1
'         AcctFlds(Cnt) = F
'      END IF
'  NEXT
'
'  FOR Cnt = 1 TO 36
'    PRINT AcctFlds(Cnt);
'  NEXT

  Frm(1).PageDisable = True
  
  '--get first active record number
  OpenCJEditFile CJEditFile, NumEdTrans
  CurrRec = 0: NextRec = 1
  RecNum = GetNextRec(CJEditFile, NumEdTrans, CurrRec, NextRec)
  CLOSE CJEditFile
  IF RecNum = 0 THEN
    EdMode = 1
  ELSE
    EdMode = 0
  END IF
  '===
  
  Scr = 1       'Set screen number to one
  LibFile2Scrn "GL", FormName$(Scr), MonoCode, -1, ErrCode      'display screen
  
  IF EdMode = 0 THEN
    Rec2Form Form$(), RecNum, Fld()
  END IF
  
  ShowEdMode EdMode

  Action = 1
  
  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    'This chunk of code deletes a distribution line when Alt+F3 is pressed
    'This needs to occur before prev fld test
    'Uses the data statements in the main mod to maintain list of acct fields.
    SELECT CASE Frm(1).KeyCode
      CASE -106  'Alt+F3
        IF Frm(1).FldNo > Header THEN   'if we're on an account field

          'for each accounting distribution
          FOR Cnt = 1 TO 36
            'if this acct field is less than the current field
            'set the row counter
            IF AcctFlds(Cnt) <= Frm(1).FldNo THEN
              FldRow = FldRow + 1  '+AcctFlds(Cnt)

            ELSE
              FOR Cnt3 = AcctFlds(FldRow) TO AcctFlds(FldRow) + 3
                LSET Form$(Cnt3, 0) = ""
              NEXT
              FOR Cnt3 = FldRow TO 35
                LSET Form$(AcctFlds(Cnt3), 0) = Form$(AcctFlds(Cnt3 + 1), 0)
                LSET Form$(AcctFlds(Cnt3) + 1, 0) = Form$((AcctFlds(Cnt3 + 1) + 1), 0)
                LSET Form$(AcctFlds(Cnt3) + 2, 0) = Form$((AcctFlds(Cnt3 + 1) + 2), 0)
                LSET Form$(AcctFlds(Cnt3) + 3, 0) = Form$((AcctFlds(Cnt3 + 1) + 3), 0)
              NEXT
              'Cnt is at 36 after 35 times thru above loop
              LSET Form$(AcctFlds(Cnt3), 0) = ""
              LSET Form$(AcctFlds(Cnt3) + 1, 0) = ""
              LSET Form$(AcctFlds(Cnt3) + 2, 0) = ""
              LSET Form$(AcctFlds(Cnt3) + 3, 0) = ""
              Frm(1).FldNo = AcctFlds(FldRow)
              Frm(1).PrevFld = Frm(1).FldNo - 1
              FldRow = 0
              Action = 1
              EXIT FOR
            END IF
          NEXT
        END IF
      CASE ELSE

    END SELECT

    IF Frm(1).PrevFld <> Frm(1).FldNo THEN  'User is moving around the form
                                            
      GOSUB ShowDistTot                     'Update Distribution totals
      
      '--Check the field we're in to see if it is an acct field.
      '  if so and if there is data in the field, remove dashes for user
      '  set stripped acct string in form

      IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN   '= AcctField
          LSET Form$(Frm(1).FldNo, 0) = StripAcct$(Form$(Frm(1).FldNo, 0))
          Action = 1
      END IF


      '--the following checks the field we have just left to see
      '  if it is an account field. if it is then
      '  if the field is not blank, the acct number is formatted with dashes
      '  if the field is blank the acct name and distribution amt are blanked
      '  and the amt field is protected to prevent user from entering a
      '  distribution without an account.

      IF QPTrim$(Fld(Frm(1).PrevFld).ScratchS) = "A" THEN  '"A" = AcctField
         IF LEN(QPTrim$(Form$(Frm(1).PrevFld, 0))) THEN
           TempAcct$ = FmtAcct$(RTRIM$(Form$(Frm(1).PrevFld, 0)), FundLen%, AcctLen%, DetLen%)
           LSET Form$(Frm(1).PrevFld, 0) = TempAcct$
           Fld(Frm(1).PrevFld + 3).Protected = False
         ELSE
           FOR Cnt2 = 1 TO 3
             LSET Form$(Frm(1).PrevFld + Cnt2, 0) = ""
             Fld(Frm(1).PrevFld + Cnt2).Protected = True
           NEXT
           IF Frm(1).FldNo > Frm(1).PrevFld THEN
             Frm(1).FldNo = Frm(1).PrevFld
           END IF
         END IF
         Action = 1

      END IF


      '--if we've just left an account field, check to see if
      '  we have a valid acct number

      IF LEN(QPTrim$(Form$(Frm(1).PrevFld, 0))) THEN
        IF QPTrim$(Fld(Frm(1).PrevFld).ScratchS) = "A" THEN '= AcctField
          GOSUB CheckForValidAcctNum
        END IF
      END IF

    END IF   'End of test for PrevFld <> current FldNo
    
    IF SaveFlag THEN GOSUB SaveRec
    
    '--Process key strokes
    SELECT CASE Frm(1).KeyCode

      CASE F2Key                          'Change Editing Mode
        SELECT CASE EdMode
          CASE 0                          'were in Edit Mode
            EdMode = 1                    ' switch to add mode
            ShowEdMode EdMode
            ResetForm Form$(), NumFlds
            Action = 1
            QPrintRC "New  ", 25, 68, -1
            GOSUB ShowDistTot

          CASE 1                          'currently adding records
            'check to see if there are any active transactions
            OpenCJEditFile CJEditFile, NumEdTrans
            CurrRec = 0: NextRec = 1
            RecNum = GetNextRec(CJEditFile, NumEdTrans, CurrRec, NextRec)
            CLOSE CJEditFile

            IF RecNum = 0 THEN            'No active records can't edit
               EdMode = 1
            ELSE
               EdMode = 0                  'switch to edit mode
               ShowEdMode EdMode
               RecNum = ListRecs
               IF RecNum > 0 THEN
                  Rec2Form Form$(), RecNum, Fld()
                  Frm(1).FldNo = 1
                  Action = 1
               ELSE
                  CLOSE
                  EXIT SUB
               END IF
            END IF
            GOSUB ShowDistTot
        END SELECT                        'Edmode
      
      CASE F3Key
        IF EMode = 0 THEN
          Ok = MsgBox%("GL.QSL", "OK2DELTR")
          IF Ok = 2 THEN   'Pressed Button #2
             Form$(1, 0) = "-1"
             SaveField 1, Form$(), Fld(), BadFld
             SaveTrans Form$(), RecNum, EdMode
             ResetForm Form$(), NumFlds
             RecNum = ListRecs
             IF RecNum > 0 THEN
                Rec2Form Form$(), RecNum, Fld()
             ELSE
                EXIT SUB
             END IF
          END IF
          Frm(1).FldNo = 1
       END IF

      CASE F5KEY
        IF EdMode = 0 THEN
          RecNum = ListRecs
          IF RecNum > 0 THEN
            Rec2Form Form$(), RecNum, Fld()
            Frm(1).FldNo = 1
            Action = 1
          END IF
        END IF

      CASE F10Key
        '--Make sure acct fields are formatted correctly if user
        '  Preses F10 with cursor in an account field
        IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN '= AcctField
            IF LEN(QPTrim$(Form$(Frm(1).FldNo, 0))) > 0 THEN
               TempAcct$ = FmtAcct$(RTRIM$(Form$(Frm(1).FldNo, 0)), FundLen%, AcctLen%, DetLen%)
               LSET Form$(Frm(1).FldNo, 0) = TempAcct$
               'Fld(Frm(1).PrevFld + 3).Protected = False
               Frm(1).FldNo = 1
            END IF
        END IF
        'send user back thru loop to verify acct before saving
        Action = 1
        SaveFlag = True
        
      CASE CtrlL  'Ctrl+LeftArrow               'Get the next record
        IF RecNum > 0 THEN
           OpenCJEditFile CJEditFile, NumEdTrans
           CurrRec = RecNum: NextRec = -1
           RecNum = GetNextRec(CJEditFile, NumEdTrans, CurrRec, NextRec)
           CLOSE CJEditFile
           Rec2Form Form$(), RecNum, Fld()
           GOSUB ShowDistTot
           Frm(1).FldNo = 1
           Action = 1
        END IF

      CASE CtrlR  'Ctrl+RightArrow             'Get the previous record
        IF RecNum > 0 THEN
           OpenCJEditFile CJEditFile, NumEdTrans
           CurrRec = RecNum: NextRec = 1
           RecNum = GetNextRec(CJEditFile, NumEdTrans, CurrRec, NextRec)
           CLOSE CJEditFile
           Rec2Form Form$(), RecNum, Fld()
           GOSUB ShowDistTot
           Frm(1).FldNo = 1
           Action = 1
       END IF

    END SELECT
     
    '--Set screen pages
    IF Frm(1).FldNo > PageInfo(Scr).LastFld OR Frm(1).FldNo < PageInfo(Scr).FirstFld AND NOT StayOnPage THEN
      SELECT CASE Frm(1).FldNo
      CASE PageInfo(1).FirstFld TO PageInfo(1).LastFld
        Scr = 1
      CASE PageInfo(2).FirstFld TO PageInfo(2).LastFld
        Scr = 2
      CASE PageInfo(3).FirstFld TO PageInfo(3).LastFld
        Scr = 3
      END SELECT
      Action = 1
      LibFile2Scrn "GL", FormName$(Scr), MonoCode, -2, ErrCode

      GOSUB ShowDistTot

    END IF

    '--Buttons are not attached to the form, the following processes
    '--the rodent
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
        CASE 8
          SELECT CASE Frm(1).MCol
            CASE 71 TO 73
              PressButton -72, 8, 72, 72
          END SELECT

        CASE 19
          SELECT CASE Frm(1).MCol
            CASE 71 TO 73
              PressButton -80, 19, 72, 72
          END SELECT

        CASE 21
          SELECT CASE Frm(1).MCol
            CASE 9 TO 21
              PressButton F10Key, 21, 9, 21
            CASE 22 TO 37
              PressButton F3Key, 21, 22, 37  'Alt+F3
            CASE 38 TO 56
                PressButton -106, 21, 38, 56
            CASE 57 TO 71
                PressButton ESC, 21, 57, 71
          END SELECT

        CASE 25
          SELECT CASE Frm(1).MCol
            CASE 27 TO 39
              PressButton F2Key, 25, 27, 39
            CASE 40 TO 52
              PressButton F5KEY, 25, 40, 52
            CASE 53 TO 59
                PressButton CtrlL, 25, 53, 59
            CASE 73 TO 79
                PressButton CtrlR, 25, 73, 79
          END SELECT
      END SELECT
    END IF

  LOOP UNTIL Frm(1).KeyCode = ESC

  EXIT SUB

  
CheckForValidAcctNum:
  AcctFld = Frm(1).PrevFld
  TempAcct$ = QPTrim$(Form$(AcctFld, 0))
  IF LEN(TempAcct$) THEN
    TempAcct$ = Form$(AcctFld, 0)
    RecordNum = FindAcct(TempAcct$)
    IF RecordNum > 0 THEN       'and RecordNum <> PreviousRecNum THEN         'Eureka!
      Form$(AcctFld + 1, 0) = GetAcctTitle$(RecordNum)
      Action = 1
    ELSEIF RecordNum = 0 THEN
      PLAY "L32O3EC"            'Acct NOT found
      Ok = MsgBox("GL.QSL", "NOACCT")
      Frm(1).FldNo = Frm(1).PrevFld             'Put user back on acct field
      LSET Form$(Frm(1).FldNo, 0) = ""
      LSET Form$(Frm(1).FldNo + 1, 0) = ""
      LSET Form$(Frm(1).FldNo + 3, 0) = ""
      Action = 1
    END IF
  END IF

RETURN


ShowDistTot:
  TotDist# = 0
  FOR Cnt = 1 TO NumFlds                         'look thru acct fld list
    IF QPTrim$(Fld(Cnt).ScratchS) = "S" THEN '= Distribution Field
      TotDist# = Round#(TotDist# + Value(Form$(Cnt, 0), ECode))
    END IF
  NEXT

  AMT# = Round#(Value(Form$(TotalAmtFld, 0), ECode))
  UnDist# = AMT# - TotDist#

  DIST$ = SPACE$(14)
  UnDist$ = SPACE$(14)
  LSET DIST$ = FUsing(STR$(TotDist#), CommaFormat$)
  LSET UnDist$ = FUsing(STR$(UnDist#), CommaFormat$)
  QPrintRC UnDist$, 20, 24, -1
  QPrintRC DIST$, 20, 56, -1

RETURN


SaveRec:
   BadFld = 0
   Check4BlankFlds Form$(), BadFld         'make sure no blank fields
   IF BadFld = 0 THEN                      'if no blanks, make sure
     GOSUB ShowDistTot                     'distribution are in bal

     IF AMT# > 0 AND UnDist# = 0 THEN
       Form$(1, 0) = "0"                   'Set deleted flag to false
       SaveField 1, Form$(), Fld(), BadFld 'Write the flag to the form$
       SaveTrans Form$(), RecNum, EdMode   'Save the transaction
       Ok = MsgBox%("GL.QSL", "SAVED")     'tell user we saved it

       IF EdMode = 1 THEN                  'if in Add Mode, clear
         ResetForm Form$(), NumFlds        ' the form for the next entry
         Frm(1).FldNo = 1
       END IF

     ELSE
       Ok = MsgBox%("GL.QSL", "BADDIST")   'Distributions out of balance
                                           'tell user
     END IF

   ELSE
     Ok = MsgBox%("GL.QSL", "NOBLANK")     'Tell user no blank fields
     Frm(1).FldNo = BadFld

   END IF

   Action = 1
   SaveFlag = False

RETURN



    
END SUB

SUB GetFundList (FundList$(), NumFunds)

  DIM FundIndex AS GLFundIndexType
  OpenFundIdx FundIdxFile, NumFunds

  IF NumFunds = 0 THEN
    Ok = MsgBox("GL", "NOFUNDS")
    CLOSE
    EXIT SUB
  END IF

  REDIM FundList$(1 TO NumFunds)

  FOR Cnt = 1 TO NumFunds
    GET FundIdxFile, Cnt, FundIndex
    FundList$(Cnt) = QPTrim$(FundIndex.FundNum)
  NEXT

  CLOSE FundIdxFileNum

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
    
    Rec = Rec + NextRec         'Set file pointer to next record
    
    IF Rec > NumRecs OR Rec <= 0 THEN           'test for beg or end of file
      Found = 0 'if no more records then get out
      EXIT DO
    END IF
    
    GET FileNum, Rec, CJEdit    'Get the record
    
    IF CJEdit.DELFLAG <> -1 THEN  'Ok if not deleted
      Found = 1
      EXIT DO   'Get out of loop when we find one
    END IF
    
  LOOP
  
  IF Found = 0 THEN
    GetNextRec = CurrRec
  ELSE
    GetNextRec = Rec
  END IF
  
END FUNCTION

FUNCTION ListRecs

   SHARED Cnf AS Config
   REDIM ScrBuf(2000)

   MScrnSave 1, 1, 25, 80, ScrBuf(0)       'Save the existing screen

   QPrintRC "List ", 25, 68, -1

   OpenCJEditFile CJEditFile, NumEdTrans
   
   REDIM TransList$(1 TO NumEdTrans)         'Dim array to hold trans list

   Active = 0                                'Counter for Active Transactions

   FOR Cnt = 1 TO NumEdTrans                 'Read the file
      GET CJEditFile, Cnt, CJEdit
      IF NOT CJEdit.DELFLAG THEN             'if active add to list
         Active = Active + 1

         List$ = SPACE$(75)

         LSET List$ = ConvDateStr1(Num2Date(CJEdit.TrDate))
         MID$(List$, 11) = CJEdit.Desc
         MID$(List$, 36) = STR$(Cnt)'CJEdit.Ref
         MID$(List$, 53) = FUsing$(STR$(CJEdit.AMT), "#########,.##")
         RecNo(Active) = Cnt
         'MID$(List$, 73) = MKI$(Cnt)

         TransList$(Active) = List$

      END IF

   NEXT
   CLOSE

   IF Active = 0 THEN
      'KILL "CJED.DAT"                     '..init file to zero
      EXIT FUNCTION
   END IF

   REDIM PRESERVE TransList$(1 TO Active)    'Redim list to only active trans

   LOCATE 4, 6                               'set up vertmenu
   Action = 0
   MaxWidth = 65
   BoxBot = 22

   ListDesc$ = "   Date     Description               Ref                      Amt   "
   QPrintRC ListDesc$, 3, 6, 112 'Cnf.ActivCh
   'MPaintBox 6, 75, 7, 76, 8              'to add shadow to list
   
   DO                                        'call vertmenu
      VertMenu TransList$(), Choice, MaxWidth, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
         CLOSE
         Abort = True
         EXIT DO
         EXIT FUNCTION
      END IF
      IF Choice > 0 THEN EXIT DO
   LOOP
   
   IF Abort THEN
      Choice = 0
   ELSE
      RecNum = RecNo(Picked(Choice))  'CVI(MID$(TransList$(Choice), 73, 2)) 'Get the record number
   END IF                                           'off the end of the $
                                                 
   CLOSE CJEditFile                       'Clean up after ourselves
   ERASE TransList$
   MScrnRest 1, 1, 25, 80, ScrBuf(0)       'Restore Screen

   ListRecs = RecNum

   
END FUNCTION

SUB OpenCJEditFile (CJEditFile, NumEdTrans)
  
  CJEdLen = LEN(CJEdit)
  CJEditFile = FREEFILE
  OPEN "CJed.dat" FOR RANDOM ACCESS READ WRITE SHARED AS CJEditFile LEN = CJEdLen
  NumEdTrans = LOF(CJEditFile) \ CJEdLen
  
END SUB

SUB PostTrans (CJType)
  
   SHARED CashAcct$

   '--Paint a background
   'QPrintRC SPACE$(80), 1, 1, 112
   FOR Cnt = 2 TO 24
     QPrintRC STRING$(80, CHR$(177)), Cnt, 1, 3
   NEXT
   'QPrintRC SPACE$(80), 1, 1, 112

   '--Verify that there are transactions
   OpenCJEditFile CJEditFile, NumEdTrans
   FOR Cnt = 1 TO NumEdTrans     'summarize the file totals
     GET CJEditFile, Cnt, CJEdit
     IF NOT CJEdit.DELFLAG THEN
       Active = Active + 1
     END IF
   NEXT
   CLOSE

   '--Check for no active transactions
   IF Active = 0 THEN
     '--No active transactions - tell user and get out
     Ok = MsgBox%("GL.QSL", "NOTRANS")
     EXIT SUB
   END IF

   '--make sure we're ready to post
   Ok = MsgBox%("GL.QSL", "POSTGJ")
   IF Ok = 1 THEN EXIT SUB    '1=No

   REDIM FundList$(1)
   GetFundList FundList$(), NumFunds
   REDIM TrFundSum#(1 TO NumFunds)
   
   OpenCJEditFile CJEditFile, NumEdTrans

   DIM Tr2Post AS GLTransRecType
   CJ2Post = FREEFILE
   OPEN "CJ2POST.DAT" FOR RANDOM AS CJ2Post LEN = LEN(Tr2Post)

   RecordNum = 0    'Reset Active counter for posting
  
   FOR Cnt = 1 TO NumEdTrans     'Assign edit file to trans format

      'Reinitialize transaction fund summary elements for next loop
      FOR Fund = 1 TO NumFunds
         TrFundSum#(Fund) = 0
      NEXT
      
      GET CJEditFile, Cnt, CJEdit

      IF NOT CJEdit.DELFLAG THEN
         FOR AcctDist = 1 TO 36
            IF LEN(QPTrim$(CJEdit.DIST(AcctDist).DACN)) = 0 THEN

               '--No more distributions now create other side of transaction
               FOR Fund = 1 TO NumFunds
                  IF TrFundSum#(Fund) <> 0 THEN
                     RecordNum = RecordNum + 1
                     Tr2Post.AcctNum = FundList$(Fund) + CashAcct$
                     Tr2Post.TrDate = CJEdit.TrDate
                     Tr2Post.Desc = CJEdit.Desc
                     'Tr2Post.Ref = CJEdit.Ref

                     SELECT CASE CJType
                        '--Cash Reciepts - cash is debited
                        CASE 1
                          Tr2Post.DrAmt = TrFundSum#(Fund)
                          Tr2Post.CrAmt = 0
                          Tr2Post.Src = "CR" + ConvDateStr$(DATE$)

                        '--Cash Disbursements - cash is credited
                        CASE 2
                          Tr2Post.DrAmt = 0
                          Tr2Post.CrAmt = TrFundSum#(Fund)
                          Tr2Post.Src = "CR" + ConvDateStr$(DATE$)

                      END SELECT

                      PUT CJ2Post, RecordNum, Tr2Post
                  END IF 'Fund summary <> 0
               NEXT 'fund

               EXIT FOR 'NumFunds

            ELSE
               '--Process each accounting distribution
               RecordNum = RecordNum + 1

               Tr2Post.AcctNum = CJEdit.DIST(AcctDist).DACN
               Tr2Post.TrDate = CJEdit.TrDate
               Tr2Post.Desc = CJEdit.Desc
               Tr2Post.Ref = CJEdit.DocRef

               SELECT CASE CJType
                  '--Cash Reciepts - distributions are credits
                  CASE 1
                    Tr2Post.DrAmt = 0
                    Tr2Post.CrAmt = CJEdit.DIST(AcctDist).DAMT
                    Tr2Post.Src = "CR" + ConvDateStr$(DATE$)

                  '--Cash Disbursements - distributions are debits
                  CASE 2
                    Tr2Post.DrAmt = CJEdit.DIST(AcctDist).DAMT
                    Tr2Post.CrAmt = 0
                    Tr2Post.Src = "CR" + ConvDateStr$(DATE$)

               END SELECT

               PUT CJ2Post, RecordNum, Tr2Post

               '--Add this distribution to proper fund
               FOR Fund = 1 TO NumFunds
                  FundNum$ = LEFT$(CJEdit.DIST(AcctDist).DACN, FundLen)
                  IF FundNum$ = FundList$(Fund) THEN
                     TrFundSum#(Fund) = TrFundSum#(Fund) + Round#(CJEdit.DIST(AcctDist).DAMT)
                     'FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(CJEdit.DIST(AcctDist).DAMT)
                     EXIT FOR
                  END IF
               NEXT

            END IF 'test for blank distribution line
         NEXT  'Acct'g Distribution
      END IF  'not deleted
   NEXT  'transaction
   
   CLOSE
    
   '--common post & link sub in GLAUX.BAS
   Post2GL "CJ2POST.DAT", BadTrans%

    IF BadTrans = 0 THEN          'posted ok
      KILL "CJED.DAT"             'kill the temp files
      'KILL "CJ2POST.DAT"
    ELSE
      'Couldn't find an account. Account was possibly deleted after entry made?
      CLS
      PRINT "Error: One or more transactions were not posted."
      PRINT "Press any key to exit"
      k$ = INPUT$(1)
    END IF

   '--Posting Report
   DIM Trans2Post AS GLTransRecType
   RecLen = LEN(Trans2Post)

   CJ2Post = FREEFILE
   OPEN "CJ2POST.DAT" FOR RANDOM AS CJ2Post LEN = RecLen

   NumTrans = LOF(CJ2Post) \ RecLen

   PrnFile = FREEFILE
   OPEN "Trans.prn" FOR OUTPUT AS #PrnFile

    'print field names
   ToPrint$ = SPACE$(80)
   LSET ToPrint$ = "Date"
   MID$(ToPrint$, 11) = "Acct Number"
   MID$(ToPrint$, 29) = "Description"
   MID$(ToPrint$, 46) = "Reference"
   MID$(ToPrint$, 61) = "Debit"
   MID$(ToPrint$, 74) = "Credit"
   PRINT #PrnFile, ToPrint$

   FOR Cnt = 1 TO NumTrans
      GET CJ2Post, Cnt, Trans2Post
         ToPrint$ = SPACE$(80)
         LSET ToPrint$ = ConvDateStr1(Num2Date$(Trans2Post.TrDate))
         MID$(ToPrint$, 11) = Trans2Post.AcctNum
         MID$(ToPrint$, 29) = LEFT$(Trans2Post.Desc, 15)
         MID$(ToPrint$, 46) = Trans2Post.Ref
         MID$(ToPrint$, 56) = FUsing(STR$(Trans2Post.DrAmt), "#######,.##")
         MID$(ToPrint$, 69) = FUsing(STR$(Trans2Post.CrAmt), "#######,.##")
         'MID$(ToPrint$, 80) = STR$(Cnt&)
         'MID$(ToPrint$, 90) = STR$(Trans2Post.NextTran)
         PRINT #PrnFile, ToPrint$
   NEXT

   CLOSE
   SHELL "list Trans.prn"
   KILL "trans.prn"
   KILL "CJ2POST.DAT"

   
END SUB

SUB PrnEditList
  
  SHARED Choice$()              'Multiple choice array
  SHARED Cnf AS Config          'System config & colors
  SHARED CashAcct$

  REDIM Form$(0, 0)             'Holds data from form
  REDIM Fld(0) AS FieldInfo     'Field editing array

  REDIM Choice$(0 TO 5, 1)      'Choices for entry type field
  Choice$(0, 0) = "2"           'Field Number
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn #1"
  Choice$(3, 0) = "Prn #2"
  Choice$(4, 0) = "Prn #3"
  'Choice$(5, 0) = "File"
  
  '--call sub in GLAUX module to load screen & form
  LibLoadDisplayForm "GL.QSL", "PRNGJ", Form$(), Fld()
  PrintHelp "Print Register"

  DIM Frm AS FormInfo
  Action = 1

  '--Get the report options from user
  DO

    EditForm Form$(), Fld(), Frm, Cnf, Action

    '--Check the user's key presses
    SELECT CASE Frm.KeyCode

    CASE F10Key 'SaveButton
      IF LEN(LTRIM$(RTRIM$(Form$(2, 0)))) = 0 THEN   'Fld#2 = OutputDevice
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

  '==Start processing register=>
  LOCATE 1, 1, 0
  
  OpenCJEditFile CJEditFile, NumEdTrans
  PrnFileNum = FREEFILE
  FileName$ = "CJREG.PRN"
  OPEN FileName$ FOR OUTPUT AS #PrnFileNum
              
  CommaFmt$ = "#########,.##"         'ten millions
  DistSumLine$ = "-------------"

  TotCredits# = 0
  TranCnt = 0

  '--Get a list of active funds
  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  REDIM FundSum#(1 TO NumFunds)
  REDIM FundGrdTot#(1 TO NumFunds)
  
  FOR transaction = 1 TO NumEdTrans
    GET CJEditFile, transaction, CJEdit

    IF NOT CJEdit.DELFLAG THEN
      TranCnt = TranCnt + 1
      TotCredits# = TotCredits# + Round#(CJEdit.AMT)

      '--1st Line - Transaction details
      ToPrint$ = SPACE$(80)
      LSET ToPrint$ = ConvDateStr1$(Num2Date(CJEdit.TrDate))
      MID$(ToPrint$, 10) = CJEdit.Desc
      MID$(ToPrint$, 32) = CJEdit.DocRef
      MID$(ToPrint$, 67) = FUsing$(STR$(CJEdit.AMT), CommaFmt$)
      PRINT #PrnFileNum, ToPrint$

      '--Blank line between detail and acct'g distributions
      PRINT #PrnFileNum,

      '--Distribution Label
      ToPrint$ = SPACE$(80)
      MID$(ToPrint$, 2) = "Accounting Distribution:"
      PRINT #PrnFileNum, ToPrint$

      '--Distribution Title Line
      ToPrint$ = SPACE$(80)
      MID$(ToPrint$, 4) = "Account Number   Name                        Debit          Credit"
      PRINT #PrnFileNum, ToPrint$

      '--Distribution Lines
      TotTranDist# = 0

      FOR AcctDist = 1 TO 36
        '--no more distributions when we find a blank Acct Number field
        IF LEN(QPTrim$(CJEdit.DIST(AcctDist).DACN)) = 0 THEN
          
          '--Print and summarize calculated cash entries
          TranCashTot# = 0
          FOR Fund = 1 TO NumFunds
            IF FundSum#(Fund) <> 0 THEN
              TranCashTot# = TranCashTot# + FundSum#(Fund)

              FundCashAcct$ = FundList$(Fund) + CashAcct$

              FundCashAcctRec = FindAcct(FundCashAcct$)
              FundCashAcctName$ = GetAcctTitle(FundCashAcctRec)

              ToPrint$ = SPACE$(80)
              MID$(ToPrint$, 4) = FundCashAcct$
              MID$(ToPrint$, 21) = FundCashAcctName$
              MID$(ToPrint$, 42) = FUsing$(STR$(FundSum#(Fund)), CommaFmt$)
              PRINT #PrnFileNum, ToPrint$
            END IF
          NEXT

          '--Summary line after last distributions
          ToPrint$ = SPACE$(80)
          MID$(ToPrint$, 42) = DistSumLine$
          MID$(ToPrint$, 58) = DistSumLine$
          PRINT #PrnFileNum, ToPrint$

          '--Transaction Distribution Totals
          ToPrint$ = SPACE$(80)
          MID$(ToPrint$, 4) = "Total Distributed"
          MID$(ToPrint$, 42) = FUsing$(STR$(TranCashTot#), CommaFmt$)
          MID$(ToPrint$, 58) = FUsing$(STR$(TotTranDist#), CommaFmt$)
          PRINT #PrnFileNum, ToPrint$

          FOR Cnt = 1 TO NumFunds
            FundSum#(Cnt) = 0
          NEXT

          '--Last distribution for this transaction - exit loop
          EXIT FOR

        ELSE
          TotTranDist# = TotTranDist# + Round#(CJEdit.DIST(AcctDist).DAMT)

          '--Add this distribution to proper fund
          Found = False
          FOR Fund = 1 TO NumFunds
            FundNum$ = LEFT$(CJEdit.DIST(AcctDist).DACN, FundLen)
            IF FundNum$ = FundList$(Fund) THEN
              Found = True
              FundSum#(Fund) = FundSum#(Fund) + Round#(CJEdit.DIST(AcctDist).DAMT)
              FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(CJEdit.DIST(AcctDist).DAMT)
              EXIT FOR
            END IF
          NEXT

          IF NOT Found THEN
            Ok = MsgBox("GL", "NOFUND")
          END IF
          
          '--Print this distribution
          ToPrint$ = SPACE$(80)
          MID$(ToPrint$, 4) = CJEdit.DIST(AcctDist).DACN
          MID$(ToPrint$, 21) = CJEdit.DIST(AcctDist).DACNM
          MID$(ToPrint$, 58) = FUsing$(STR$(CJEdit.DIST(AcctDist).DAMT), CommaFmt$)
          PRINT #PrnFileNum, ToPrint$
        
        END IF 'Active transaction test
        
      NEXT 'Distribution

      '--2 blank lines before next distribution
      ToPrint$ = STRING$(80, "=")
      PRINT #PrnFileNum, ToPrint$

      PRINT #PrnFileNum,

    END IF 'Not deleted test

  NEXT 'Transaction
        
  'PRINT #PrnFileNum, STRING$(80, "-")

  '--Summary
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "File Totals:"
  PRINT #PrnFileNum, ToPrint$

  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "Number of Transactions"
  MID$(ToPrint$, 31) = FUsing$(STR$(TranCnt), "####")
  PRINT #PrnFileNum, ToPrint$
  
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "Grand Totals"
  MID$(ToPrint$, 25) = FUsing$(STR$(TotCredits#), CommaFmt$)
  PRINT #PrnFileNum, ToPrint$

  PRINT #PrnFileNum,

  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "Summary by Fund:"
  PRINT #PrnFileNum, ToPrint$
  
  FOR Cnt = 1 TO NumFunds
    ToPrint$ = SPACE$(80)
    LSET ToPrint$ = "Fund" + " " + FundList$(Cnt)
    MID$(ToPrint$, 25) = FUsing$(STR$(FundGrdTot#(Cnt)), CommaFmt$)
    PRINT #PrnFileNum, ToPrint$
    GrdTot# = GrdTot# + Round#(FundGrdTot#(Cnt))
  NEXT
  
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "Total All Funds"
  MID$(ToPrint$, 25) = FUsing$(STR$(GrdTot#), CommaFmt$)
  PRINT #PrnFileNum, ToPrint$
  
  CLOSE
  
  Title1$ = "Cash Reciepts Journal"
  Title2$ = "Date      Description          Reference                                    Amt"
  Title4$ = STRING$(80, "=")
  CrLf$ = CHR$(13) + CHR$(10)


  SELECT CASE Dev$
  CASE "S"

    HideCursor
    CLS

    TitColor = 15
    RptColor = 15

    QPrintRC Title1$, 1, 1, TitColor
    QPrintRC Title2$, 3, 1, TitColor
'    QPrintRC Title3$, 4, 1, TitColor
'    QPrintRC STRING$(80, CHR$(205)), 4, 1, TitColor
    QPrintRC STRING$(80, "-"), 4, 1, TitColor

'    QPrintRC STRING$(80, CHR$(205)), 23, 1, TitColor

    QPrintRC "Esc=Exit", 25, 1, TitColor

    Prn2Scrn FileName$, 5, 22, RptColor, PrnFlag

    IF PrnFlag = 1 THEN
      LPTNo = 1
      GOSUB PrintCJ
    END IF
    ShowCursor

  CASE "P"
    GOSUB PrintCJ

  END SELECT

  KILL "CJREG.PRN"
  EXIT SUB


PrintCJ:
  FOR Cnt = 1 TO 25             'Paint a background screen
    QPrintRC STRING$(80, CHR$(177)), Cnt, 1, 3
  NEXT

  Title$ = Title1$ + CrLf$ + CrLf$ + Title2$ + CrLf$ + Title3$ + CrLf$ + Title4$ + CrLf$
  TitLines = 5
  MaxLines = 55
  Prn2Prn "CJREG.PRN", Title$, TitLines%, MaxLines%, LPTNo%, CancelFlag%

  RETURN


END SUB

SUB Rec2Form (Form$(), RecordNumber%, Fld() AS FieldInfo)

   OpenCJEditFile CJEditFile, NumEdTrans

   GET CJEditFile, RecordNumber, CJEdit

   Form$(0, 0) = SPACE$(LEN(CJEdit))

   BCopy VARSEG(CJEdit), VARPTR(CJEdit), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), LEN(Form$(0, 0)), 0

   UnPackBuffer 0, 0, Form$(), Fld()

   'Form$(1, 0) = STR$(CJEdit.DelFlag)
   'Form$(2, 0) = Num2Date(CJEdit.TrDate)
   'Form$(3, 0) = CJEdit.Desc
   'Form$(4, 0) = CJEdit.Ref
   'Form$(5, 0) = STR$(CJEdit.Amt)
   'Form$(6, 0) = CJEdit.BatchNum
   'Form$(7, 0) = STR$(CJEdit.TCode)

   CLOSE CJEditFile                       'close the file
   
   Rec$ = SPACE$(5)
   LSET Rec$ = STR$(RecordNumber%)
   QPrintRC Rec$, 25, 67, -1


END SUB

SUB ResetForm (Form$(), NumFlds)

FOR F = 1 TO NumFlds
  Form$(F, 0) = ""
NEXT

END SUB

SUB SaveTrans (Form$(), RecordNumber, EdMode)
    
  OpenCJEditFile CJEditFile, NumEdTrans
  
  IF EdMode = 1 THEN
    Rec = NumEdTrans + 1
  ELSE
    Rec = RecordNumber
  END IF

  'DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
  'CALL bcopy(SEG Form$(0, 0), SEG CJEdit, LEN(Form$(0, 0)), 0)

  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(CJEdit), VARPTR(CJEdit), LEN(Form$(0, 0)), 0
  
  'CJEdit.DelFlag = VAL(Form$(1, 0))
  'CJEdit.TrDate = Date2Num(Form$(2, 0))
  'CJEdit.Desc = Form$(3, 0)
  'CJEdit.Ref = Form$(4, 0)
  'CJEdit.Amt = Round#(Value(Form$(5, 0), ErrCode))
  'CJEdit.BatchNum = Form$(6, 0)
  'CJEdit.TCode = VAL(Form$(7, 0))
  'STOP

  PUT CJEditFile, Rec, CJEdit         'write it
  CLOSE CJEditFile                       'close the file

   
'--stuff--->
'   CJEdit.Src = "CR" + ConvDateStr$(DATE$)
'
'   Amt# = Value(Form$(Amt, 0), ErrCode)      'Convert Amt$ to Double
'   IF ErrCode THEN
'      CLS
'      CLOSE
'      PRINT "An error occured in the value function!"
'      EXIT SUB
'   END IF
'                                             'Only one amount field in form!
'   SELECT CASE CJEdit.EType                  'Is amount debit or credit
'      CASE "D"                               'Debit
'         CJEdit.DrAmt = Amt#                 'assign amt to debit amt
'         CJEdit.CrAmt = 0
'      CASE "C"                               'Credit
'         CJEdit.CrAmt = Amt#                 'assign amt to credit amt
'         CJEdit.DrAmt = 0
'   END SELECT
'
'--< end of stuff.


END SUB

SUB ShowEdMode (EdMode%)

  Status$ = SPACE$(24)
  F2Ed$ = "  F2=Edit  "
  F2New$ = "  F2=New   "
  ActiveColor = 112
  InActiveColor = 115

  SELECT CASE EdMode
    CASE 0
      LSET Status$ = "Edit an Existing Record"
      F2BtnCap$ = F2New$
      F5BtnColor = ActiveColor

    CASE 1
      LSET Status$ = "Add a New Record"
      F2BtnCap$ = F2Ed$
      F5BtnColor = InActiveColor
    
  END SELECT

  QPrintRC Status$, 25, 2, -1
  QPrintRC F2BtnCap$, 25, 28, -1
  MPaintBox 25, 41, 25, 51, F5BtnColor       'F5 Button enable / disable

END SUB

