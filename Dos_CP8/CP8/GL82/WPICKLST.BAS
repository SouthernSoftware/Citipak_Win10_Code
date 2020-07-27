'******** PickList.Bas - lets a user pick one or more items from a menu
DEFINT A-Z
  
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)

'---- ASM subs
DECLARE SUB DeleteT (SEG Element, ElSize, NumEls)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)

  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  CONST False = 0, True = NOT False

  '$INCLUDE: 'DefCnf.BI'

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
' Note: Elements in Items$() should be padded with 3 extra spaces to the
'       right.  If elements are not padded on entry, the routine will add
'       2 spaces to each element.  On exit from the routine all padding to
'       the right will be stripped.
'
' External Routines:  VertMenuT2
'****************************************************************************
SUB PickList (Items() AS FLen2, Picked(), NPicked, DspRow%, UpDateFlag, Cnf AS Config)     'STATIC
  
  BeenEdited = False
  Count$ = SPACE$(5)
  Nul$ = CHR$(0)
  AltM$ = Nul$ + "2"
  AltC$ = Nul$ + "."
  F10$ = Nul$ + "D"
  Esc$ = CHR$(27)

  Marker = 251  'Use ASCII code for - "û"

  MaxChoices = UBOUND(Picked)   'Find Max. number of items
  MaxItem = UBOUND(Items)

  NPicked = 0   'Init. Number of choices
  MaxLen = 0    'Init. Longest Item
  
  
  '----- Initialize the "Items$" array
  FOR N = 1 TO MaxItem  'Examine each item
    Lngth = LEN(Items(N))       'Get its length
    IF Lngth THEN               'Is it null?
      MID$(Items(N).V, Lngth - 1) = "³"         'Insert a bar character
      IF Lngth > MaxLen THEN MaxLen = Lngth     'Update the max. length
      IF NPicked < MaxChoices THEN              'Are we at end of "Picked"?
        IF Picked(NPicked + 1) = N THEN         'Should this Item be marked?
          MID$(Items(N).V, MaxLen) = "û"         'Mark it
          NPicked = NPicked + 1
        END IF
      END IF
    ELSE        'Bail out if null Item
      EXIT FOR
    END IF
  NEXT
  
  
  '----- Give menu of items and let User make choices
  Action = 1    'Set "Initialize Menu" Flag
  '  for multi-tasking mode
  BoxBot = 14   'No limit on bottom line
  
  'BoxBot = 0
  GOSUB DisplayTagCount
  
  DO

    LOCATE DspRow%, 6, 0
    
    VertMenuT2 Items(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(13) OR Ky$ = " " THEN         'Did User press Space or Enter?
      'Is Item not marked?
      Char = 32 'Set Space character code
      IF RIGHT$(Items(Choice).V, 1) = " " THEN
        IF NPicked < MaxChoices THEN            'Are we within the limit?
          Char = Marker         'Set marking character code
          NPicked = NPicked + 1 'Increment # of selections
          Picked(NPicked) = Choice
        END IF
      ELSE      'Un-select an item
        FOR i = 1 TO NPicked
          IF Picked(i) = Choice THEN            'Find the item in Picked() and
            DeleteT SEG Picked(i), -1, MaxChoices - i           ' delete it
            EXIT FOR            'All done,
          END IF
        NEXT
        NPicked = NPicked - 1   ' decrement # of selections
      END IF
      'Put mark character in Item$
      MID$(Items(Choice).V, MaxLen) = CHR$(Char)
      Action = 2
      LOCATE , , 0
      BeenEdited = True
    ELSEIF Ky$ = AltM$ THEN    'Did User press Alt-M???????????
      'Is Item not marked?
      FOR CCnt = 1 TO MaxChoices
        MID$(Items(CCnt).V, MaxLen) = CHR$(Marker)
        Picked(CCnt) = CCnt     'Find the item in Picked() and
      NEXT
      NPicked = MaxChoices      ' decrement # of selections
      'Put mark character in Item$
      Action = 2
      LOCATE , , 0
      BeenEdited = True
    ELSEIF Ky$ = AltC$ THEN     'ALT-C Clear all marked
      FOR CCnt = 1 TO MaxChoices
        MID$(Items(CCnt).V, MaxLen) = " "
        Picked(CCnt) = 0
      NEXT
      NPicked = 0
      Action = 2
      LOCATE , , 0
      BeenEdited = True
    END IF
    
    IF Action = 2 THEN
      GOSUB DisplayTagCount
    END IF

    IF Ky$ = CHR$(27) THEN
      IF BeenEdited THEN
        SaveFlag = PromptSaveData%
        SELECT CASE SaveFlag
        CASE False
          UpDateFlag = False
        CASE 1
          Ky$ = ""
        CASE True
          UpDateFlag = True
        END SELECT
      ELSE
        UpDateFlag = False
      END IF
    END IF

    IF Ky$ = F10$ THEN
      IF BeenEdited THEN
        UpDateFlag = True
      END IF
    END IF
  LOOP UNTIL Ky$ = F10$ OR Ky$ = Esc$
  
  
  '----- Clean up the Item$() array
  FOR N = 1 TO UBOUND(Items)    'Examine each item
    Lngth = LEN(Items(N).V)     'Get it's length
    IF Lngth THEN
      IF MID$(Items(N).V, Lngth, 1) <> " " THEN 'Is it marked?
        MID$(Items(N).V, Lngth - 1) = "  "      'Erase the mark character
      END IF
      Items(N).V = RTRIM$(Items(N).V)           'Trim off the padding
    END IF
  NEXT
  EXIT SUB
  
  
DisplayTagCount:
  RSET Count$ = QPTrim$(STR$(NPicked))
  QPrintRC Count$, 18, 55, 11
  RSET Count$ = QPTrim$(STR$(MaxItem))
  QPrintRC Count$, 18, 30, 11

  RETURN
  
END SUB

