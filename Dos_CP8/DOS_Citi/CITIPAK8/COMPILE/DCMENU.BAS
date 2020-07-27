DEFINT A-Z
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB HideCursor ()


'$INCLUDE: 'DefCnf.BI'
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf AS Config)


'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'SetCnf.BI'


 STACK 8000

   '--Dim the choice array to the number of menu items
   REDIM MChoice$(1 TO 5)

   MChoice$(1) = "Vehicle Decal Catagory Maintenance"
   MChoice$(2) = "Customer Maintenance"
   MChoice$(3) = "Customer Reports"
   MChoice$(4) = "Purchase Decal"
   MChoice$(5) = "Exit to OS"

   MaxLen = 0     'Set menu width to zero
   BoxBot = 17    'limit the box length to go no lower than line 20
   Action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

   '--Find max menu width
   FOR Cnt = 1 TO UBOUND(MChoice$)
     TLen = LEN(MChoice$(Cnt))
     IF TLen > MaxLen THEN
       MaxLen = TLen
     END IF
   NEXT

   '--Center Menu within Screen
   Row = ((25 - (UBOUND(MChoice$))) \ 2)
   Col = ((80 - MaxLen) \ 2) - 2
   Help$ = "State of Virginia Vehicle Decals"
   
   DO


      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      LibFile2Scrn "DC.QSL", "MENUBAK", MonoCode, -1, ErrorCode

      TitleBox 3, Col, MaxLen + 3, "State of Virginia Vehicle Decals ", Cnf
      TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

      'PrintTitle user$
      'PrintHelp Help$

      ShowCursor

      VertMenu MChoice$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf


      IF ky$ = CHR$(27) THEN EXIT DO 'choice = 0

      SELECT CASE Choice
          CASE 1
          RUN "dccode"
          CASE 2
          RUN "dccust"
          CASE 3
          RUN "dcreport"
          CASE 4
          RUN "dcpaymen"
          CASE 5
          CLOSE
          CLS
          END
      END SELECT
   LOOP
   RUN "cpmenu"

