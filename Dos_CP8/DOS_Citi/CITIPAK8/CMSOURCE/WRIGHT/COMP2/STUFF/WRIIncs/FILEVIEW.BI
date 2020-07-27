TYPE FVInfo
    Colr        AS INTEGER     'Text color (white on black default)
    TRow       AS INTEGER     'these four describe the window's corners
    LCol       AS INTEGER
    BRow       AS INTEGER
    RCol       AS INTEGER
    HorizOffset AS INTEGER      'left margin within the window
    LoPtr       AS INTEGER      'used internally -- do not change!
    FileHandle  AS INTEGER      'internal (the DOS file handle if you care)
    EndOfFile   AS INTEGER      'used internally -- do not change!
    LineNumber  AS LONG         'line number of top line displayed, READ ONLY
    TabStop     AS INTEGER      'Tab stop setting
    FileSeek    AS LONG         'force read anywhere in the file
    FileOffset  AS LONG         'used internally -- do not change!
END TYPE

