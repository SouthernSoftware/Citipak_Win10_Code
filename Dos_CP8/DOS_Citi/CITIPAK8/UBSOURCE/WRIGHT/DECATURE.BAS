    tA = 2
    tB = 5
    tC = 6
    tD = 9
    tE = 12
    tF = 13
    tG = 21
    tH = 29
    tI = 39
    tJ = 43
    tK = 50
    tL = 54
    tM = 57
    tz = 60


2810 FOR LOP = 1 TO 6: LPRINT "": NEXT LOP



    LPRINT TAB(tA); "XXX";
      LPRINT TAB(tB); "XXXXXXX";
        LPRINT TAB(tF); "XXXXXXX";
          LPRINT TAB(tG); "XXXXXXX";
            LPRINT TAB(tH); "XXXXXXXX";
              LPRINT TAB(tI); STRING$((tz - tI), "X")   'Line7

    LPRINT TAB(tA); "XXX";
      LPRINT TAB(tH); "XXXXXXXX";
        LPRINT TAB(tI); STRING$((tz - tI), "X")         'Line8

    LPRINT TAB(tA); "XXX";
      LPRINT TAB(tH); "XXXXXXXX";
        LPRINT TAB(tI); STRING$((tz - tI), "X")         'Line9

    LPRINT TAB(tA); "XXX";
      LPRINT TAB(tH); "XXXXXXXX";
        LPRINT TAB(tI); STRING$((tz - tI) - 4, "X")     'Line10

    LPRINT TAB(tA); "XXX";
      LPRINT TAB(tH); "XXXXXXXX";
        LPRINT TAB(tz - 10); "XXXXX-XXXX"               'Line11

    LPRINT TAB(tA); "XXX";
      LPRINT TAB(tH); "XXXXXXXX";                       'Line12

    LPRINT TAB(tA); "XXX-XXXXXX XXXXXXXXXXXXXXXXXXXXXX" 'Line13

    LPRINT TAB(tI); "XXX";
      LPRINT TAB(tJ); "XXXXXX";
        LPRINT TAB(tK); "XX";
          LPRINT TAB(tL); "XX";
            LPRINT TAB(tM); "XX"                        'Line14

    LPRINT                                              'Line15

    LPRINT TAB(tA); "XX";
      LPRINT TAB(tC); "XX";
        LPRINT TAB(tD); "XX";
          LPRINT TAB(tE); "XXXXXXXX";
            LPRINT TAB(tG); "XXXXXXXX";
              LPRINT TAB(tH); "XXXXXXXX";
                LPRINT TAB(tI); "XXXXXXXXX";
                  LPRINT TAB(tK); "XXXXXXXXX"           'Line16

    LPRINT                                              'Line17
    LPRINT                                              'Line18
    LPRINT TAB(tI); "***** FINAL BILL *****"            'Line19
    LPRINT                                              'Line20
    LPRINT                                              'Line21

'RETURN    'END OF BILL FORMAT

