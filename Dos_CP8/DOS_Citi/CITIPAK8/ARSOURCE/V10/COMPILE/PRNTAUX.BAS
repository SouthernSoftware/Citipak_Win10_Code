DEFINT A-Z
DECLARE SUB QPrintRC (t$, r%, c%, clr%)


SUB PrintHelp (Help$)

   t$ = SPACE$(40)
   LSET t$ = Help$
   QPrintRC t$, 25, 2, -1
   t$ = ""
   
END SUB

SUB PrintTitle (Title$)
   t$ = SPACE$(40)
   LSET t$ = Title$
   QPrintRC t$, 1, 2, -1
   QPrintRC DATE$, 1, 70, -1
   t$ = ""
END SUB

