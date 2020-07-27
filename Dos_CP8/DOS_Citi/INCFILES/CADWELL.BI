TYPE WBCustType

    lname         AS STRING * 25
    firstname     AS STRING * 25
    add1          AS STRING * 30
    add2          AS STRING * 30
    CITY          AS STRING * 20
    ST            AS STRING * 2
    ZP            AS STRING * 9
    phone         AS STRING * 12
    BOOK          AS STRING * 2
    meter         AS STRING * 12
    rate          AS INTEGER 'STRING * 2
    won           AS STRING * 8
    MINSTALL      AS STRING * 8
    AVGUSE        AS STRING * 8
    AVGCT         AS STRING * 2
    CURBAL        AS STRING * 8
    PASTDUE       AS STRING * 8
    cycle         AS STRING * 1
    NOTUSED       AS STRING * 5
    curread       AS DOUBLE      'STRING * 8
    pastread      AS DOUBLE      'STRING * 8
    firsttr       AS STRING * 8
    lasttr        AS STRING * 8
    waterdep      AS SINGLE
    Filler        AS STRING * 3  'padding to correct fielding error
END TYPE

TYPE WBCust1Type
    RAMT1           AS SINGLE
    RAMT2           AS SINGLE
    RAMT3           AS SINGLE
    RDESC1          AS STRING * 25
    RDESC2          AS STRING * 25
    RDESC3          AS STRING * 25
    RACCT1          AS STRING * 9
    RACCT2          AS STRING * 9
    RACCT3          AS STRING * 9
    OAMT1           AS SINGLE
    OAMT2           AS SINGLE
    ODESC1          AS STRING * 25
    ODESC2          AS STRING * 25
    OACCT1          AS STRING * 9
    OACCT2          AS STRING * 9
    sewercode       AS INTEGER   'STRING * 2
    garbagecode     AS INTEGER   'STRING * 2
    METERNUM        AS STRING * 14
    add3            AS STRING * 30
    sprinklerrate   AS STRING * 2
    minimums        AS STRING * 2
    NOTUSE          AS STRING * 10
    Filler          AS STRING * 4  'padding to correct fielding error
END TYPE

