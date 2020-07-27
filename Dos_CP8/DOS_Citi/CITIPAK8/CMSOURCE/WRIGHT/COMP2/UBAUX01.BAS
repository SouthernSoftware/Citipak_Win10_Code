DEFINT A-Z
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB SaveScrn (Array())
DECLARE SUB RestScrn (Array())
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)

'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
'$INCLUDE: 'SetCnf.BI'

''$INCLUDE: 'ubfiles.BI'

FUNCTION MsgBox% (LibName$, FormName$)

  REDIM ScrBuff(0)
  '--Initialize the form name array
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, FormName$)

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, FormName$, StartEl, Fld(), Form$(), ErrCode

  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT

  SaveScrn ScrBuff()

   'DisplayUBScrn FormName$
   LibFile2Scrn LibName$, FormName$, MonoCode%, Attribute%, ErrorCode%
   Action = 1

   DO
      EditForm Form$(), Fld(), Frm(1), Cnf, Action

      IF Frm(1).KeyCode THEN                  'If a key was pressed,
                                           'Make it upper case
         IF Frm(1).KeyCode >= 97 AND Frm(1).KeyCode <= 122 THEN
            Frm(1).KeyCode = Frm(1).KeyCode - 32
         END IF

         FOR B = 1 TO Fld(0).Fields        'See if its one of the buttons
            IF Fld(B).ScratchI = Frm(1).KeyCode THEN EXIT DO
         NEXT
      END IF
   LOOP

   RestScrn ScrBuff()

   MsgBox% = B

END FUNCTION

