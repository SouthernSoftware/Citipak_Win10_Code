DEFINT A-Z
DECLARE FUNCTION GetZipEDigit$ (Zip$)
DECLARE FUNCTION QPTrim$ (Text$)

GOTO skip1
LPRINT CHR$(27); CHR$(16); "A"; CHR$(8); CHR$(1); CHR$(4); CHR$(1); CHR$(2); CHR$(1); CHR$(1); CHR$(1); CHR$(2);
LPRINT CHR$(27); CHR$(16); "B"; CHR$(6); "123456"
LPRINT
LPRINT
LPRINT
LPRINT
LPRINT
LPRINT
skip1:
Zip$ = "283211791"
EDigit$ = GetZipEDigit$(Zip$)
Zip$ = Zip$ + EDigit$

LPRINT CHR$(27); CHR$(16); "C"; CHR$(10); Zip$

FUNCTION GetZipEDigit$ (Zip$)
ZipLen = LEN(Zip$)
ZipVal = 0

FOR Cnt = 1 TO ZipLen
  ZipVal = ZipVal + VAL(MID$(Zip$, Cnt, 1))
NEXT

Dif = 10 - (ZipVal MOD 10)

GetZipEDigit$ = QPTrim$(STR$(Dif))

END FUNCTION

