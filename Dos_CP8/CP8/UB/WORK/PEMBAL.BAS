DECLARE FUNCTION RemoveDSign$ (This$)
DEFINT A-Z

DIM LineIn$
DIM dpos AS INTEGER
DIM OutLine AS STRING * 93
DIM Bucks AS STRING
DIM Sew AS STRING * 10
DIM Wat AS STRING * 10
DIM Tsh AS STRING * 10
DIM Crd AS STRING * 10

CLS
OPEN "pemubbal.txt" FOR INPUT AS #1
OPEN "pembal.txt" FOR OUTPUT AS #2


  DO
    LINE INPUT #1, LineIn$
    LineIn$ = LTRIM$(RTRIM$(LineIn$))
    IF LEFT$(LineIn$, 1) = "0" THEN
      dpos = INSTR(LineIn$, "$")
      DO WHILE dpos > 0
        MID$(LineIn$, dpos, 1) = " "
        dpos = INSTR(LineIn$, "$")
      LOOP
'GOTO DownHere
      LSET OutLine = LineIn$
      LineIn$ = OutLine$
      Bucks = MID$(LineIn$, 45, 9)
      IF LEFT$(Bucks, 1) = "," THEN
        RSET Bucks = "-0.00"
        MID$(LineIn$, 45, 9) = Bucks
        MID$(LineIn$, 84, 9) = Bucks
      ELSE
        IF INSTR(Bucks, "(") > 0 THEN
          dpos = INSTR(Bucks, "(")
          MID$(Bucks, dpos, 1) = " "
          MID$(Bucks, 9, 1) = " "
          RSET Bucks = "-" + LTRIM$(RTRIM$(Bucks))
          MID$(LineIn$, 45, 9) = Bucks
        END IF
        Bucks = MID$(LineIn$, 84, 9)
        IF INSTR(Bucks, "(") > 0 THEN
          dpos = INSTR(Bucks, "(")
          MID$(Bucks, dpos, 1) = " "
          MID$(Bucks, 9, 1) = " "
          RSET Bucks = "-" + LTRIM$(RTRIM$(Bucks))
          MID$(LineIn$, 84, 9) = Bucks
        END IF
        PRINT VAL(Bucks)
      END IF
      'PRINT VAL(Bucks), Bucks
      LSET OutLine = LineIn$
      'SLEEP
      'PRINT outLine
DownHere:
      PRINT #2, LineIn$
    END IF

  LOOP UNTIL EOF(1)

  CLOSE '#1
  RSET Sew$ = "0.00"
  RSET Wat$ = "0.00"
  RSET Tsh$ = "0.00"
  RSET Crd$ = "0.00"


  OPEN "pemubbal.txt" FOR INPUT AS #1
  OPEN "pemReves.txt" FOR OUTPUT AS #2

  DO
    LINE INPUT #1, LineIn$
    LineIn$ = LTRIM$(RTRIM$(LineIn$))
    SELECT CASE LEFT$(LineIn$, 5)
    CASE "SEWER"
      Sew$ = RIGHT$(LineIn$, 10)
    CASE "WATER"
      Wat$ = RIGHT$(LineIn$, 10)
    CASE "TRASH"
      Tsh$ = RIGHT$(LineIn$, 10)
    CASE "CREDI"
      Crd$ = RIGHT$(LineIn$, 10)
    CASE ELSE 'don't do anything
    END SELECT

    IF LEFT$(LineIn$, 1) = "0" THEN  'after output
      LSET OutLine = LEFT$(LineIn$, 10)

      MID$(OutLine, 20, 10) = RemoveDSign$(Wat$)
      MID$(OutLine, 30, 10) = RemoveDSign$(Sew$)
      MID$(OutLine, 40, 10) = RemoveDSign$(Tsh$)
      MID$(OutLine, 50, 10) = RemoveDSign$(Crd$)
      PRINT #2, OutLine$
      RSET Sew$ = "0.00"
      RSET Wat$ = "0.00"
      RSET Tsh$ = "0.00"
      RSET Crd$ = "0.00"
    END IF

  LOOP UNTIL EOF(1)
  CLOSE


FUNCTION RemoveDSign$ (This$)
DIM Mess$
DIM dpos%
DIM pad AS STRING * 10

Mess$ = This$

dpos = INSTR(Mess$, "$")

DO WHILE dpos > 0
  MID$(Mess$, dpos, 1) = " "
  dpos = INSTR(Mess$, "$")
LOOP

dpos = INSTR(Mess$, "(")

DO WHILE dpos > 0
  MID$(Mess$, dpos, 1) = " "
  dpos = INSTR(Mess$, "(")
LOOP

dpos = INSTR(Mess$, ")")

DO WHILE dpos > 0
  MID$(Mess$, dpos, 1) = " "
  dpos = INSTR(Mess$, ")")
LOOP
RSET pad = LTRIM$(RTRIM$(Mess$))

RemoveDSign = pad

END FUNCTION

