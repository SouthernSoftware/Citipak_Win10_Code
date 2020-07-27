
    FormName$ = "CANTUNDO"
    NumFlds = LibNumberOfFields(LibName$, FormName$)

    REDIM Frm(1) AS FormInfo
    REDIM Form$(NumFlds, 2)       'DIM the form data array
    REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
    StartEl = 0   'Load first form at array start
    LibGetFldDef LibName$, FormName$, StartEl, Fld(), Form$(), ErrCode

    '----- Set the "Action" flag to force the editor to initialize itself and
    '      display the data on the form.
    Action = 1

    '----- Setup TYPE for setting and reading form editing information.
    Frm(1).FldNo = 1              'Start editing on field #1
    Frm(1).InsStat = False        'Set insert state (True = Insert on)
    Frm(1).StartEl = 0            'Set form starting element to 0 and
    BlockClear
    LibFile2Scrn LibName$, FormName$, MonoCode%, Attribute%, ErrorCode%
    QPrintRC LEFT$(UBCustRec(1).CustName, 34), 10, 34, 78
    QPrintRC Count$, 8, 37, 78

    DO
      EditForm Form$(), Fld(), Frm(1), Cnf, Action
      SELECT CASE Frm(1).KeyCode
      CASE F0Key
        KillFlag = True
        EXIT DO
      CASE EscKey
        KillFlag = False
        EXIT DO
      END SELECT
    LOOP

    IF NOT KillFlag THEN
      CursorOff
      BlockClear
      LibFile2Scrn LibName$, "PRORPT", MonoCode%, Attribute%, ErrorCode%
      QPrintRC "REMOVE TRANSACTIONS ABORTED.", 10, 27, -1
      QPrintRC "Press any key to continue.", 13, 28, -1
      WaitForAction
    ELSE   'remove the transactions
      GOSUB WackEM
    END IF
'**************************

  END IF

