DEFINT A-Z
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE FUNCTION FldNum% (FldName$, Fld() AS ANY)
DECLARE FUNCTION LibNumberOfFields% (LibName$, ScrName$)
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE FUNCTION Value# (E$, ErrCode%)
DECLARE FUNCTION WaitKey% ()
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
DECLARE SUB CalcFields (Start, LastFldNum, Form$(), Fld() AS ANY)
DECLARE SUB DispLoad ()
DECLARE SUB EditForm (Form$(), Fld() AS ANY, Frm AS ANY, Cnf AS ANY, Action%)
DECLARE SUB FileView (FileName$, Ky, Action, FVI AS ANY, SEG Array)
DECLARE SUB HideCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB LibGetFldDef (LibName$, ScrName$, StartEl, Fld() AS ANY, Form$(), ErrCode)
DECLARE SUB LibLoadDisplayForm (LibName$, FormName$, Form$(), Fld() AS ANY)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB PrintArray (FirstFld%, LastFld%, Form$(), Fld() AS ANY)
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (t$, r%, c%, clr%)
DECLARE SUB SaveField (FldNo%, Form$(), Fld() AS ANY, BadFld%)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB UnPackBuffer (FirstFld%, LastFld%, Form$(), Fld() AS ANY)
DECLARE SUB VertMenu (Item$(), Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)

'$INCLUDE: 'Fileview.bi'
'$INCLUDE: 'DefCnf.bi'        'Defines a TYPE for monitor and color info..
'$INCLUDE: 'FieldInf.bi'      'field information
'$INCLUDE: 'FormEdit.bi'      'form editing information
''$INCLUDE: 'QSCR.BI'          'QuickScreen Declarations

CONST False = 0, True = NOT False

'****************************************************************************
'This procedure simplifies loading and displaying data entry forms.  It will
'size the field information and definition arrays, load them from a form
'library (.QFL) and display the screen from the screen library (.QSL).
'
'Inputs:
'  LibName$ - name of the screen/form library containing the form.
'  FormName$ - name of the form to load and display.
'
'Returns:
'  Form$() - field data array sized to the number of fields in the form and
'          filled with formulas and help messages.
'  Fld() - field information array sized to the number of fields in the form
'          and filled with their information.
'Actions:
'  Displays the specified form.
'****************************************************************************
SUB LibLoadDisplayForm (LibName$, FormName$, Form$(), Fld() AS FieldInfo)

   '----- Flag - Set to 3 if using a monochrome monitor or the program was
   '      started with "/b" in the command line.
   DEF SEG = 0
   MonoCode = ABS(PEEK(&H463) = &HB4 OR INSTR(COMMAND$, "/B")) * 3

   '----- Display the screen
   LibFile2Scrn LibName$, FormName$, MonoCode, 0, ErrorCode

   '----- Load the field definition file
                                        'Get the number of fields in form
   NumFlds = LibNumberOfFields(LibName$, FormName$)

   REDIM Form$(NumFlds, 2)              'DIM the form data array

   REDIM Fld(NumFlds) AS FieldInfo      'DIM the field information array. see
                                        '"FieldInf.BI" for structure
                                        'Load the field information
   LibGetFldDef LibName$, FormName$, 0, Fld(), Form$(), ErrCode


END SUB

'****************************************************************************
'Displays a form containing a message and waits for the user to press a key
'or button to acknowledge the message.
'
'Inputs:
'  LibName$ - Name of the screen/form library file containing the input form.
'  FormName$ - Name of the message form.
'
'Re0turns:
'  The number of the button used to acknowledge the message.
'
'Notes:
'  The underlying screen is automaticaly saved on entry and restored on exit.
'  The Message form should contain button fields only.  If you defined 3
'  buttons, where the first returned 65 ("A"bort), the second 82 ("R"etry)
'  and the third 73 ("I"gnore), and the user pressed "R" or clicked on the
'  second button, this function would return 2.
'****************************************************************************
'
FUNCTION MsgBox% (LibName$, FormName$)

   'SHARED Cnf AS Config
   REDIM Form$(0, 0)
   REDIM Fld(0) AS FieldInfo
   REDIM ScrBuf(2000)
   REDIM Frm(1) AS FormInfo

   '--Save the existing screen
   MScrnSave 1, 1, 25, 80, ScrBuf(0)
   
   '--Display the form
   LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
   
   Action = 1

   DO
      EditForm Form$(), Fld(), Frm(1), Cnf, Action

      IF Frm(1).KeyCode THEN               'If a key was pressed,

         '--Make it upper case
         IF Frm(1).KeyCode >= 97 AND Frm(1).KeyCode <= 122 THEN
            Frm(1).KeyCode = Frm(1).KeyCode - 32
         END IF

         '--See if its one of the buttons
         FOR B = 1 TO Fld(0).Fields
            IF Fld(B).ScratchI = Frm(1).KeyCode THEN EXIT DO
         NEXT
      END IF
   LOOP

   MScrnRest 1, 1, 25, 80, ScrBuf(0)       'Restore the original screen

   MsgBox% = B

   ERASE Form$                             'clean up
   ERASE Fld
   ERASE ScrBuf
   ERASE Frm

END FUNCTION

SUB PrintHelp (Help$) STATIC
   IF NOT BeenHere THEN
     t$ = SPACE$(40)
     BeenHere = True
   END IF
   LSET t$ = Help$
   QPrintRC t$, 25, 2, -1
END SUB

SUB PrintTitle (Title$)
   t$ = SPACE$(40)
   LSET t$ = Title$
   QPrintRC t$, 1, 2, -1
   QPrintRC DATE$, 1, 70, -1
   t$ = ""
END SUB

