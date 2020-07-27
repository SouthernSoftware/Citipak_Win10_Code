DEFINT A-Z
    LPRINT "Test one."
    LPRINT "Type 50: POSTNET"
    LPRINT
    LPRINT CHR$(27); "[;;1}";       'Human readable on
    LPRINT CHR$(27); "[50t"
    LPRINT "123456789"
    LPRINT CHR$(27); "[0t"
    LPRINT
    LPRINT
    LPRINT "Test two."
    LPRINT "Type 04: Code 3 of 9 default"
    LPRINT
    LPRINT CHR$(27); "[;;1}";       'Human readable on
    LPRINT CHR$(27); "[4t"
    LPRINT "987654321"
    LPRINT CHR$(27); "[0t"
    LPRINT
    LPRINT
    LPRINT "End of Test."




