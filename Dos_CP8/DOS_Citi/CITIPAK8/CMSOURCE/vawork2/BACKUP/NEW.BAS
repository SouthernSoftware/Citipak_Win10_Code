DECLARE FUNCTION AskSavePayment% (DefaultFlag$)
DEFINT A-Z

FUNCTION AskSavePayment% (DefaultFlag$)

  REDIM TempScrn(0)
  SaveScrn TempScrn()

  LibName$ = "UB"
  ScrnName$ = "ASKR2PST"

  '--Initialize the form name array
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  DisplayUBScrn ScrnName$

  ShowCursor

  Action = 1
  FirstTime = True

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = NOT FirstTime
      LSET Form$(1, 0) = DefaultFlag$
      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE EscKey
      AskSavePayment% = False
      EXIT DO
    CASE F10Key
      Receipt$ = Form$(1, 0)
      IF Receipt$ = "Y" THEN
        AskSavePayment% = 1
      ELSE
        AskSavePayment% = True
      END IF
      EXIT DO
    END SELECT

    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 14
        SELECT CASE Frm(1).MCol
        CASE 27 TO 40           'ESC Cancel button
          PressButton 27, 14, 27, 40
        CASE 41 TO 56           'F10 Save Button
          PressButton -68, 14, 41, 56
        END SELECT
      END SELECT
    END IF
  LOOP

  RestScrn TempScrn()

END FUNCTION

