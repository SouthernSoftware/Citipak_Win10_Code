'******** PickList.Bas - lets a user pick one or more items from a menu


DEFINT A-Z

'DECLARE SUB VertMenu (Item$(), Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
'DECLARE SUB VertMenu (Item$(), Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)

'---- ASM subs
DECLARE SUB DeleteT (SEG Element, ElSize, NumEls)

TYPE FLen2
   V AS STRING * 64
END TYPE

'$INCLUDE: 'DefCnf.BI'
''$INCLUDE: 'setCnf.BI'

'****************************************************************************
'   File Name: PickList.Bas
'          by: Don Malin,  (C) Crescent Software 1988
'
' Description: Wrapper for Vertical Menu Sub Program.  Allows the User to
'              choose several items from a list.
'
'  Parameters:
'   Items$() -
'     The "Items$" array holds the items to be chosen from.
'
'   Picked() -
'     Returns a table of pointers to the selected items.
'     Picked(1) = Element number of the first selected item in "Items$()"
'     The size of Picked() limits the number of choices that may be selected.
'     (This would usually be the same as the size of the Items$() array.)
'
'   NPicked -
'     Returns the number of items selected
'
'   Cnf -
'     Structure defined in "DefCnf.Bi" as "Config" and initialized in
'     "SetCnf.Bi"
'
' Note: Elements in Items$() should be padded with 3 extra spaces to the
'       right.  If elements are not padded on entry, the routine will add
'       2 spaces to each element.  On exit from the routine all padding to
'       the right will be stripped.
'
' External Routines:  VertMenu
'****************************************************************************
SUB PickList (Items() AS FLen2, Picked(), NPicked, Cnf AS Config) 'STATIC

    Marker = 251                                'Use ASCII code for - "û"
    MaxChoices = UBOUND(Picked)                 'Find Max. number of items
    NPicked = 0                                 'Init. Number of choices
    MaxLen = 0                                  'Init. Longest Item


    '----- Initialize the "Items$" array
    FOR N = 1 TO UBOUND(Items)                 'Examine each item
        Lngth = LEN(Items(N))                  'Get its length
        IF Lngth THEN                           'Is it null?
           IF RIGHT$(Items(N).V, 3) <> "   " THEN     'Is it properly padded?
              Items(N).V = RTRIM$(Items(N).V) + "  "   'Pad it out
              Lngth = LEN(Items(N).V)                 'Get new length
           ELSE                                 'Already padded,
               MID$(Items(N).V, Lngth - 1) = "³" 'Insert a bar character
           END IF
           IF Lngth > MaxLen THEN MaxLen = Lngth 'Update the max. length
           IF NPicked < MaxChoices THEN         'Are we at end of "Picked"?
              IF Picked(NPicked + 1) = N THEN   'Should this Item be marked?
                 MID$(Items(N).V, Lngth) = CHR$(Marker)  'Mark it
                 NPicked = NPicked + 1
              END IF
           END IF
        ELSE                                    'Bail out if null Item
           EXIT FOR
        END IF
    NEXT


    '----- Give menu of items and let User make choices
    Action = 1                                  'Set "Initialize Menu" Flag
                                                '  for multi-tasking mode
    BoxBot = 16                                  'No limit on bottom line

    'BoxBot = 0

    DO
       LOCATE 5, 7, 0
                                                'Poll the menu
        'VertMenu Items$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
        VertMenuT2 Items(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
        IF Ky$ = CHR$(13) OR Ky$ = " " THEN     'Did User press Space or Enter?
                                                'Is Item not marked?
           Char = 32                            'Set Space character code
           IF RIGHT$(Items(Choice).V, 1) = " " THEN
              IF NPicked < MaxChoices THEN      'Are we within the limit?
                 Char = Marker                  'Set marking character code
                 NPicked = NPicked + 1          'Increment # of selections
                 Picked(NPicked) = Choice
              END IF
           ELSE                                 'Un-select an item
              FOR i = 1 TO NPicked
                 IF Picked(i) = Choice THEN     'Find the item in Picked() and
                    DeleteT SEG Picked(i), -1, MaxChoices - i ' delete it
                    EXIT FOR                    'All done,
                 END IF
              NEXT
              NPicked = NPicked - 1             ' decrement # of selections
           END IF
                                                'Put mark character in Item$
           MID$(Items(Choice).V, LEN(Items(Choice).V)) = CHR$(Char)
           Action = 2
           LOCATE , , 0

        ELSEIF Ky$ = CHR$(0) + "2" THEN  'Did User press Alt-M???????????
                                                'Is Item not marked?
           FOR CCnt = 1 TO MaxChoices
             'MID$(Items(CCnt).V, (LEN(Items(CCnt).V)) - 3) = CHR$(Marker)
             MID$(Items(CCnt).V, LEN(Items(CCnt).V)) = CHR$(Marker)
             Picked(CCnt) = CCnt     'Find the item in Picked() and
           NEXT
           NPicked = MaxChoices            ' decrement # of selections
                                                'Put mark character in Item$
           Action = 2
           LOCATE , , 0
       
        ELSEIF Ky$ = CHR$(0) + "." THEN  'Did User press Space or Enter?
                                                'Is Item not marked?
           FOR CCnt = 1 TO MaxChoices
             'MID$(Items(CCnt).V, (LEN(Items(CCnt).V)) - 3) = " "
             MID$(Items(CCnt).V, LEN(Items(CCnt).V)) = " "
             Picked(CCnt) = 0     'Find the item in Picked() and
           NEXT
           NPicked = 0            ' decrement # of selections
                                                'Put mark character in Item$
           Action = 2
           LOCATE , , 0
        END IF
        'IF LEN(Ky$) = 2 THEN PRINT Ky$
    'LOOP UNTIL Ky$ = CHR$(27)                   'Wait for User to press Escape
     LOOP UNTIL Ky$ = CHR$(0) + "D"
    '----- Erase the menu from screen with Action of 5
    VertMenuT2 Items(), Choice, MaxLen, BoxBot, Ky$, 5, Cnf
    'VertMenu Items$(), Choice, MaxLen, BoxBot, Ky$, 5, Cnf
    
    '----- Clean up the Item$() array
    FOR N = 1 TO UBOUND(Items)                 'Examine each item
        Lngth = LEN(Items(N).V)                  'Get it's length
        IF Lngth THEN
           IF MID$(Items(N).V, Lngth, 1) <> " " THEN 'Is it marked?
              MID$(Items(N).V, Lngth - 1) = "  " 'Erase the mark character
           END IF
           Items(N).V = RTRIM$(Items(N).V)        'Trim off the padding
        END IF
    NEXT
   
END SUB

