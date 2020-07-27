DEFINT A-Z
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
DIM ZZ AS INTEGER
DIM Dat$
DIM p1$, p2$
DIM t$, Frs$, Lst$, nn$
DIM Cpos%, APos, BPos, RPos
DIM AddSR
DIM VBar$
DIM Cnt%, TLen%

VBar$ = "|"
CLS

OPEN "ubcust.txt" FOR INPUT AS #1 LEN = 512
OPEN "DWubcust.txt" FOR OUTPUT AS #2 LEN = 1024

  DO UNTIL EOF(1)
    LINE INPUT #1, Dat$
    ZZ = ZZ + 1
    IF ZZ > 1 THEN
      GOSUB ParseName
'      IF AddSR > 0 THEN
      PRINT #2, p1$ + nn$ + p2$
'      ELSE
'        PRINT #2, p1$ + nn$ + p2$
'      END IF
      
    ELSE
      PRINT #2, Dat$
    END IF

  '  SLEEP
  LOOP
  CLOSE
  PRINT ZZ

  END

ParseName:
  Cpos = INSTR(Dat$, VBar$)
  Cnt = 1: AddSR = 0
  t$ = "": Lst$ = "": Frs$ = "": nn$ = ""
  
DO
  Cpos = INSTR(Cpos + 1, Dat$, VBar$)
  Cnt = Cnt + 1
  IF Cnt = 7 THEN
    APos = INSTR(Cpos + 1, Dat$, VBar$)
    t$ = MID$(Dat$, Cpos + 1, APos - (Cpos + 1))
    EXIT DO
  END IF
LOOP

TLen = LEN(t$)
IF TLen > 0 THEN
  BPos = INSTR(t$, ",")
  IF BPos > 0 THEN
    Lst$ = UCASE$(QPTrim(LEFT$(t$, BPos - 1)))
    IF INSTR(Lst$, "SR,") > 0 OR INSTR(Lst$, "SR.") > 0 THEN
      STOP
    END IF
    Frs$ = UCASE$(QPTrim(MID$(t$, BPos + 1)))

    RPos = INSTR(Frs$, "SR,")

    IF RPos > 0 THEN
      PRINT Frs$
    END IF
    RPos = INSTR(Frs$, "SR.")
    IF RPos > 0 THEN
      AddSR = 1
      PRINT Frs$
      Frs$ = QPTrim(LEFT$(Frs$, RPos - 1))
      IF RIGHT$(Frs$, 1) = "," THEN
        Frs$ = LEFT$(Frs$, LEN(Frs$) - 1)
      END IF

      PRINT Frs$
    END IF
    IF AddSR = 1 THEN
      nn$ = Frs$ + " " + Lst$ + " SR."
    ELSE
      nn$ = Frs$ + " " + Lst$
    END IF
  ELSE
    nn$ = t$
  END IF
ELSE
  nn$ = t$
END IF

p1$ = LEFT$(Dat$, Cpos)
p2$ = MID$(Dat$, APos)


RETURN


