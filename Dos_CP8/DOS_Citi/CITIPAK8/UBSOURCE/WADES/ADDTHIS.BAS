DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB PrintHelp (Help$)
  t$ = SPACE$(40)
  LSET t$ = Help$
  QPrintRC t$, 25, 2, -1
  t$ = ""
END SUB

